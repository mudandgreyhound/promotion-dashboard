"""
page_frequency.py — หน้า "🔁 ความถี่ลูกค้า" ของ app.py (วิเคราะห์ความถี่การมาซื้อของ Member)

นิยาม
- ดูได้เฉพาะ Member (มีรหัสใน BillHdrTB.Remark) — Guest ไม่มีรหัสให้ตามว่ากลับมาไหม
- 1 ครั้ง (visit) = member × สาขา × วัน (บิลแยกในวันเดียวกันนับครั้งเดียว) · บิลที่นับ = BillStatus ไม่ใช่ 4/5
- ข้อมูลใน gh_promotion เริ่ม 2026-01-01 → "member ใหม่" วัดเฉพาะคนที่เห็นครั้งแรกหลัง ม.ค.
- ขาประจำที่หายไป = มา ≥ เกณฑ์ขาประจำ และไม่มามาแล้วเกิน max(2 × รอบปกติของตัวเอง, 30 วัน)
- รายชื่อแสดงแค่รหัสสมาชิก (ไม่มีชื่อ/เบอร์)
"""
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import pymysql

BANDS = [0, 1, 2, 5, 11, 23, 10 ** 6]
BAND_LABELS = ["1 ครั้ง", "2 ครั้ง", "3–5", "6–11", "12–23", "24+"]
WEEKDAYS = ["จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส.", "อา."]


def _load(mysql_cfg):
    conn = pymysql.connect(**mysql_cfg)
    try:
        bills = pd.read_sql("""
            SELECT h.ShopCode shop, h.BillDate d, TRIM(h.Remark) mem, h.NoPerson pax,
                   h.NetBAmt - CASE WHEN s.VatType='0' THEN COALESCE(h.VatAmt1,0)+COALESCE(h.VatAmt2,0) ELSE 0 END base
            FROM BillHdrTB h JOIN PosShopTB s ON s.ShopCode=h.ShopCode
            WHERE (h.BillStatus IS NULL OR h.BillStatus NOT IN ('4','5'))
              AND TRIM(COALESCE(h.Remark,'')) NOT IN ('','0')""", conn)
        all_bills = pd.read_sql("""
            SELECT ShopCode shop, BillDate d, COUNT(*) n FROM BillHdrTB
            WHERE BillStatus IS NULL OR BillStatus NOT IN ('4','5') GROUP BY ShopCode, BillDate""", conn)
        tiers = pd.read_sql("""SELECT TRIM(m.MemCode) mem, COALESCE(t.MemTNam1, m.MemType) tier
                               FROM MemberTB m LEFT JOIN MemTypetb t ON t.MemType=m.MemType""", conn)
        shops = pd.read_sql("SELECT ShopCode shop, ShopDes1 shop_name, ShTypCode brand FROM PosShopTB", conn)
    finally:
        conn.close()
    bills["d"] = pd.to_datetime(bills["d"])
    bills["base"] = pd.to_numeric(bills["base"], errors="coerce").fillna(0)
    bills["pax"] = pd.to_numeric(bills["pax"], errors="coerce").fillna(0)
    v = bills.groupby(["mem", "shop", "d"], as_index=False).agg(base=("base", "sum"), pax=("pax", "sum"),
                                                                 bills=("base", "size"))
    all_bills["d"] = pd.to_datetime(all_bills["d"])
    return v, bills.groupby(["shop", "d"]).size().rename("mem_bills").reset_index(), all_bills, \
        tiers.drop_duplicates("mem").set_index("mem")["tier"], shops.drop_duplicates("shop").set_index("shop")


def render(st, *, mysql_cfg, shops_sel, brand_label, layout_base, gold, gold_txt, title_fn):
    load = st.cache_data(ttl=3600, show_spinner=False)(_load)
    with st.spinner("กำลังโหลดข้อมูล member ทั้งปี..."):
        v_all, mem_bills, all_bills, tier_map, shop_info = load(mysql_cfg)
    if v_all.empty:
        st.warning("ไม่มีข้อมูล member")
        return
    END = v_all["d"].max()
    names = shop_info["shop_name"].to_dict()

    # ── ตัวเลือก ──
    c1, c2, c3 = st.columns([2, 1, 3])
    period = c1.selectbox("ช่วงวิเคราะห์", ["ตั้งแต่ 1 ม.ค.", "180 วันล่าสุด", "90 วันล่าสุด"], key="fq_period")
    default_reg = {"ตั้งแต่ 1 ม.ค.": 6, "180 วันล่าสุด": 4, "90 วันล่าสุด": 3}[period]
    reg_min = c2.number_input("ขาประจำ = มา ≥ (ครั้ง)", 2, 50, default_reg, key=f"fq_reg_{period}")
    start = (pd.Timestamp("2026-01-01") if period.startswith("ตั้งแต่")
             else END - pd.Timedelta(days=(180 if period.startswith("180") else 90) - 1))
    c3.markdown(f"<div style='padding-top:30px;color:#6b7280;font-size:14px'>📅 {start:%d/%m/%Y} – {END:%d/%m/%Y} "
                f"· 🏷️ {brand_label} · ข้อมูลถึงเมื่อวาน (อัปเดตทุกคืน)</div>", unsafe_allow_html=True)

    v = v_all[(v_all["d"] >= start) & (v_all["d"] <= END)]
    if shops_sel:
        v = v[v["shop"].isin(shops_sel)]
    if v.empty:
        st.info("ไม่มี member ในช่วง/สาขาที่เลือก")
        return

    # ── ต่อ member ──
    pm = v.groupby("mem").agg(visits=("d", "size"), last=("d", "max"), first=("d", "min"),
                              shops=("shop", "nunique"), spend=("base", "sum"))
    vv = v.sort_values(["mem", "d"]).drop_duplicates(["mem", "d"]).copy()
    vv["gap"] = vv.groupby("mem")["d"].diff().dt.days
    pm["med_gap"] = vv.groupby("mem")["gap"].median()
    regular = pm["visits"] >= reg_min
    pm["since"] = (END - pm["last"]).dt.days
    pm["overdue"] = regular & (pm["since"] > np.maximum(2 * pm["med_gap"].fillna(0), 30))

    # member ใหม่ = เห็นครั้งแรก (ในข้อมูลทั้งหมด) ตั้งแต่ ก.พ. และอยู่ในช่วงนี้ — วัดการกลับมาใน 30 วัน
    vscope = v_all[v_all["shop"].isin(shops_sel)] if shops_sel else v_all
    first_all = vscope.groupby("mem")["d"].min()
    _u = vscope.drop_duplicates(["mem", "d"]).sort_values(["mem", "d"])
    _u = _u.assign(k=_u.groupby("mem").cumcount())
    second_all = _u[_u["k"] == 1].set_index("mem")["d"]          # วันที่มาครั้งที่ 2 (ไม่มี = ไม่เคยกลับมา)
    newm = first_all[(first_all >= max(start, pd.Timestamp("2026-02-01"))) & (first_all <= END - pd.Timedelta(days=30))]
    ret30 = ((second_all.reindex(newm.index) - newm).dt.days <= 30).mean() * 100 if len(newm) else None

    # ── KPI ──
    n_mem, n_vis = len(pm), int(pm["visits"].sum())
    reg_share = pm.loc[regular, "spend"].sum() / pm["spend"].sum() * 100 if pm["spend"].sum() else 0
    CARD = ("background:#fff;border:1px solid #ebe5d9;border-radius:14px;padding:14px 18px;flex:1;min-width:160px;"
            "box-shadow:0 1px 3px rgba(0,0,0,.05)")
    LB, BIG, SUB = ("color:#6b5e48;font-size:15px;font-weight:700",
                    "font-size:28px;font-weight:800;color:#1f2328;line-height:1.3", "color:#6b7280;font-size:13.5px")
    cards = [
        ("Member ที่มา", f"{n_mem:,}", f"{n_vis:,} ครั้ง · {n_vis / n_mem:.1f} ครั้ง/คน"),
        ("มาซ้ำ (≥2 ครั้ง)", f"{(pm['visits'] >= 2).mean() * 100:.0f}%", f"มาครั้งเดียว {(pm['visits'] == 1).sum():,} คน"),
        (f"ขาประจำ (≥{reg_min} ครั้ง)", f"{regular.sum():,}", f"{regular.mean() * 100:.0f}% ของคน · {reg_share:.0f}% ของยอด"),
        ("รอบการมา (ค่ากลาง)", f"{vv['gap'].median():.0f} วัน", "ระหว่างครั้งของคนที่มาซ้ำ"),
        ("ใหม่ → กลับมาใน 30 วัน", "—" if ret30 is None else f"{ret30:.0f}%",
         f"จาก member ใหม่ {len(newm):,} คน (เห็นครั้งแรกหลัง ม.ค.)"),
        ("ขาประจำที่หายไป", f"{int(pm['overdue'].sum()):,}",
         f"ยอดปีนี้ ฿{pm.loc[pm['overdue'], 'spend'].sum():,.0f}"),
    ]
    st.markdown("<div style='display:flex;gap:12px;flex-wrap:wrap;margin:6px 0 10px'>" + "".join(
        f"<div style='{CARD}'><div style='{LB}'>{a}</div><div style='{BIG}'>{b}</div><div style='{SUB}'>{c}</div></div>"
        for a, b, c in cards) + "</div>", unsafe_allow_html=True)

    def lay(fig, h, **kw):
        fig.update_layout(**layout_base)
        fig.update_layout(height=h, legend=dict(orientation="h", y=1.12, x=0, bgcolor="rgba(0,0,0,0)"), **kw)
        return fig

    # ── การกระจาย + แนวโน้มรายเดือน ──
    c1, c2 = st.columns(2)
    with c1:
        title_fn("📊 มากี่ครั้ง", "สัดส่วนคน vs สัดส่วนยอดขาย member")
        band = pd.cut(pm["visits"], BANDS, labels=BAND_LABELS)
        g = pm.groupby(band, observed=False).agg(n=("visits", "size"), spend=("spend", "sum"),
                                                 per=("spend", "sum"))
        g["per"] = (pm.assign(b=band).groupby("b", observed=False).apply(
            lambda d: d["spend"].sum() / d["visits"].sum() if d["visits"].sum() else 0))
        fig = go.Figure()
        fig.add_trace(go.Bar(x=g.index.astype(str), y=g["n"] / g["n"].sum() * 100, name="% ของ member",
                             marker_color="#cfd4dc", text=(g["n"] / g["n"].sum() * 100).map("{:.0f}%".format),
                             textposition="outside", hovertemplate="%{x}: %{y:.1f}% ของคน<extra></extra>"))
        fig.add_trace(go.Bar(x=g.index.astype(str), y=g["spend"] / g["spend"].sum() * 100, name="% ของยอดขาย",
                             marker_color=gold, text=(g["spend"] / g["spend"].sum() * 100).map("{:.0f}%".format),
                             textposition="outside", hovertemplate="%{x}: %{y:.1f}% ของยอด<extra></extra>"))
        st.plotly_chart(lay(fig, 330, barmode="group"), use_container_width=True, key="fq_bands")
        st.caption("ยอด/ครั้ง: " + " · ".join(f"{k} ฿{val:,.0f}" for k, val in g["per"].items()))
    with c2:
        title_fn("📈 รายเดือน", "member ที่มา · ใหม่ · กลับมาใน 30 วัน")
        vm = v.assign(m=v["d"].dt.to_period("M"))
        act = vm.groupby("m")["mem"].nunique()
        fm = first_all.reindex(pm.index).dt.to_period("M")
        new_n = fm.value_counts().sort_index()
        new_n = new_n[new_n.index >= pd.Period("2026-02")]
        rr = {}
        for m in new_n.index:
            ids = first_all[(first_all.dt.to_period("M") == m)].index.intersection(pm.index)
            if m.end_time + pd.Timedelta(days=30) <= END and len(ids):
                rr[m] = ((second_all.reindex(ids) - first_all.reindex(ids)).dt.days <= 30).mean() * 100
        fig = go.Figure()
        fig.add_trace(go.Bar(x=act.index.astype(str), y=act.values, name="member ที่มา", marker_color="#e3d7b8"))
        fig.add_trace(go.Bar(x=new_n.index.astype(str), y=new_n.values, name="member ใหม่", marker_color=gold))
        fig.add_trace(go.Scatter(x=[str(k) for k in rr], y=list(rr.values()), name="ใหม่ที่กลับมาใน 30 วัน (%)",
                                 yaxis="y2", mode="lines+markers", line=dict(color="#2563eb", width=2)))
        lay(fig, 330, barmode="overlay",
            yaxis2=dict(overlaying="y", side="right", ticksuffix="%", range=[0, 50], showgrid=False,
                        tickfont=dict(color="#2563eb")))
        st.plotly_chart(fig, use_container_width=True, key="fq_monthly")

    # ── ต่อสาขา ──
    title_fn("🏪 เทียบรายสาขา", "เรียงตามจำนวน member · กดหัวคอลัมน์เพื่อเรียงใหม่")
    msv = v.groupby(["shop", "mem"]).agg(n=("d", "size"), last=("d", "max")).reset_index()
    tot_mem = v.groupby("mem").size()
    msv["share"] = msv["n"] / msv["mem"].map(tot_mem)
    vs = v.sort_values(["shop", "mem", "d"]).copy()
    vs["gap"] = vs.groupby(["shop", "mem"])["d"].diff().dt.days
    ab = all_bills[(all_bills["d"] >= start) & (all_bills["d"] <= END)].groupby("shop")["n"].sum()
    mb = mem_bills[(mem_bills["d"] >= start) & (mem_bills["d"] <= END)].groupby("shop")["mem_bills"].sum()
    home = msv.sort_values("n", ascending=False).drop_duplicates("mem").set_index("mem")["shop"]
    od_home = home.reindex(pm.index[pm["overdue"]]).value_counts()
    ps = msv.groupby("shop").agg(members=("mem", "size"), visits=("n", "sum"),
                                 repeat=("n", lambda s: (s >= 2).mean() * 100),
                                 regulars=("n", lambda s: (s >= reg_min).sum()),
                                 loyalty=("share", "mean"))
    ps["gap"] = vs.groupby("shop")["gap"].median()
    ps["per_visit"] = v.groupby("shop")["base"].sum() / ps["visits"]
    ps["mem_pct"] = mb.reindex(ps.index) / ab.reindex(ps.index) * 100
    ps["overdue"] = od_home.reindex(ps.index).fillna(0)
    ps = ps[ps["visits"] >= 30].sort_values("members", ascending=False)
    tbl = pd.DataFrame({
        "สาขา": [("🔒 " if str(names.get(s, "")).lower().startswith("closed") else "") + str(names.get(s, s)) for s in ps.index],
        "แบรนด์": [shop_info["brand"].get(s, "") for s in ps.index],
        "Member": ps["members"].astype(int).values,
        "ครั้ง/คน": (ps["visits"] / ps["members"]).round(1).values,
        "มาซ้ำ %": ps["repeat"].round(1).values,
        f"ขาประจำ (≥{reg_min})": ps["regulars"].astype(int).values,
        "รอบ (วัน)": ps["gap"].round(0).values,
        "ยอด/ครั้ง ฿": ps["per_visit"].round(0).values,
        "% บิลเป็น member": ps["mem_pct"].round(1).values,
        "มาที่สาขานี้ %": (ps["loyalty"] * 100).round(1).values,
        "ขาประจำหายไป": ps["overdue"].astype(int).values,
    })
    st.dataframe(tbl, use_container_width=True, hide_index=True, height=38 + 35 * min(len(tbl), 20),
                 column_config={
                     "สาขา": st.column_config.TextColumn(width="medium"),
                     "แบรนด์": st.column_config.TextColumn(width=60),
                     "Member": st.column_config.NumberColumn(width=70),
                     "ครั้ง/คน": st.column_config.NumberColumn(format="%.1f", width=70),
                     f"ขาประจำ (≥{reg_min})": st.column_config.NumberColumn(width=90),
                     "รอบ (วัน)": st.column_config.NumberColumn(format="%.0f", width=70),
                     "มาซ้ำ %": st.column_config.ProgressColumn(format="%.0f%%", min_value=0, max_value=100, width=110),
                     "% บิลเป็น member": st.column_config.NumberColumn(format="%.0f%%"),
                     "มาที่สาขานี้ %": st.column_config.NumberColumn(
                         format="%.0f%%", help="เฉลี่ยแล้ว member ของสาขานี้มาที่สาขานี้กี่ % ของครั้งทั้งหมด (ที่เหลือไปสาขาอื่น)"),
                     "ยอด/ครั้ง ฿": st.column_config.NumberColumn(format="localized"),
                     "ขาประจำหายไป": st.column_config.NumberColumn(help="ขาประจำที่สาขานี้เป็นสาขาหลัก และไม่มาเกิน 2 เท่าของรอบปกติ"),
                 })
    st.caption("🔒 = สาขาที่ปิดแล้ว · มาซ้ำ % = member ของสาขานี้ที่มาสาขานี้ ≥2 ครั้ง · "
               "% บิลเป็น member = สัดส่วนบิลที่มีรหัสสมาชิก (ต่ำ = ร้านยังเก็บ member ได้น้อย)")

    # ── tier | วันในสัปดาห์ ──
    c1, c2 = st.columns(2)
    with c1:
        title_fn("🏅 ระดับสมาชิก vs ความถี่")
        pt = pm.assign(tier=tier_map.reindex(pm.index).fillna("ไม่พบในทะเบียน").values)
        tg = pt.groupby("tier").agg(members=("visits", "size"), med=("visits", "median"),
                                    mean=("visits", "mean")).sort_values("members", ascending=False)
        tg = tg[tg["members"] >= 10]
        fig = go.Figure(go.Bar(x=tg.index, y=tg["mean"], marker_color=gold,
                               text=[f"{m:.1f} ครั้ง<br>{n:,} คน" for m, n in zip(tg["mean"], tg["members"])],
                               textposition="outside", hovertemplate="%{x}: เฉลี่ย %{y:.1f} ครั้ง<extra></extra>"))
        st.plotly_chart(lay(fig, 320, yaxis_title="ครั้งต่อคน (เฉลี่ย)"), use_container_width=True, key="fq_tier")
    with c2:
        title_fn("📅 วันที่ member มา", "ขาประจำ vs ทั่วไป")
        wd = v.assign(w=v["d"].dt.dayofweek, reg=v["mem"].map(regular))
        fig = go.Figure()
        for flag, nm, col in [(True, "ขาประจำ", gold), (False, "member ทั่วไป", "#cfd4dc")]:
            s = wd[wd["reg"] == flag]["w"].value_counts(normalize=True).reindex(range(7)).fillna(0) * 100
            fig.add_trace(go.Bar(x=WEEKDAYS, y=s.values, name=nm, marker_color=col,
                                 hovertemplate="%{x}: %{y:.1f}%<extra>" + nm + "</extra>"))
        st.plotly_chart(lay(fig, 320, barmode="group", yaxis_ticksuffix="%"), use_container_width=True, key="fq_wd")

    # ── สาขาที่ปิด: member ย้ายไปไหน ──
    closed = [s for s in ps.index if str(names.get(s, "")).lower().startswith("closed")]
    if closed:
        title_fn("🔒 สาขาที่ปิด — member ไปต่อที่สาขาอื่นไหม")
        home_all = (v_all.groupby(["mem", "shop"]).size().reset_index(name="n").sort_values("n", ascending=False)
                    .drop_duplicates("mem").set_index("mem")["shop"])
        last_day = v_all.groupby("shop")["d"].max()
        rows = []
        for s in closed:
            ms = home_all[home_all == s].index
            after = v_all[v_all["mem"].isin(ms) & (v_all["d"] > last_day[s]) & (v_all["shop"] != s)]
            to = after.groupby("shop")["mem"].nunique().sort_values(ascending=False).head(3)
            rows.append({"สาขาที่ปิด": names.get(s, s), "บิลสุดท้าย": f"{last_day[s]:%d/%m/%Y}",
                         "member ที่สาขานี้เป็นหลัก": len(ms), "ไปสาขาอื่นแล้ว": after["mem"].nunique(),
                         "ไปต่อ %": round(after["mem"].nunique() / len(ms) * 100, 1) if len(ms) else 0,
                         "ไปที่ไหน (คน)": ", ".join(f"{names.get(k, k)} {n}" for k, n in to.items())})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                     column_config={"ไปที่ไหน (คน)": st.column_config.TextColumn(width="large"),
                                    "บิลสุดท้าย": st.column_config.TextColumn(width=90),
                                    "ไปต่อ %": st.column_config.NumberColumn(format="%.0f%%", width=70)})
        st.caption("member ที่ยังไม่ไปสาขาอื่น = กลุ่มที่ควรชวนไปสาขาใกล้ ๆ ก่อนหายไป")

    # ── ขาประจำที่หายไป ──
    title_fn("📞 ขาประจำที่หายไป", "มา ≥ เกณฑ์ขาประจำ แต่ไม่มามาแล้วเกิน 2 เท่าของรอบปกติตัวเอง (อย่างน้อย 30 วัน)")
    od = pm[pm["overdue"]].copy()
    if od.empty:
        st.success("ไม่มีขาประจำที่หายไปในช่วง/สาขานี้")
        return
    od["home"] = home.reindex(od.index).map(lambda s: names.get(s, s))
    od["tier"] = tier_map.reindex(od.index).fillna("—").values
    homes = ["ทุกสาขา"] + od["home"].value_counts().index.tolist()
    pick = st.selectbox("สาขาหลัก", homes, key="fq_od_home",
                        format_func=lambda h: h if h == "ทุกสาขา" else f"{h} ({(od['home'] == h).sum()} คน)")
    lst = od if pick == "ทุกสาขา" else od[od["home"] == pick]
    lst = lst.sort_values("spend", ascending=False)
    out = pd.DataFrame({
        "รหัสสมาชิก": lst.index, "ระดับ": lst["tier"].values, "สาขาหลัก": lst["home"].values,
        "ครั้งที่มา": lst["visits"].astype(int).values, "รอบปกติ (วัน)": lst["med_gap"].round(0).values,
        "ไม่มาแล้ว (วัน)": lst["since"].astype(int).values, "มาล่าสุด": lst["last"].dt.strftime("%d/%m/%Y").values,
        "ยอดรวม ฿": lst["spend"].round(0).values, "ยอด/ครั้ง ฿": (lst["spend"] / lst["visits"]).round(0).values,
    })
    st.markdown(f"**{len(out):,} คน** · ยอดรวมในช่วงนี้ ฿{lst['spend'].sum():,.0f} · เรียงจากยอดมากไปน้อย")
    st.dataframe(out, use_container_width=True, hide_index=True, height=38 + 35 * min(len(out), 15),
                 column_config={"ยอดรวม ฿": st.column_config.NumberColumn(format="localized", width=90),
                                "ยอด/ครั้ง ฿": st.column_config.NumberColumn(format="localized", width=80),
                                "ครั้งที่มา": st.column_config.NumberColumn(width=70),
                                "รอบปกติ (วัน)": st.column_config.NumberColumn(format="%.0f", width=90),
                                "ไม่มาแล้ว (วัน)": st.column_config.NumberColumn(width=90),
                                "มาล่าสุด": st.column_config.TextColumn(width=90)})
    st.download_button("📥 ดาวน์โหลดรายชื่อ (CSV)", out.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"winback_{'all' if pick == 'ทุกสาขา' else pick}_{END:%Y%m%d}.csv",
                       mime="text/csv", key="fq_dl")
    st.caption("รายชื่อเป็นรหัสสมาชิกเท่านั้น (ไม่มีชื่อ/เบอร์) — ค้นรายละเอียดต่อใน CM POS")
