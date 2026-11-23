# -*- coding: utf-8 -*-
"""
Stock Movement Viewer — CM POS
อ่านไฟล์ Excel ที่ export จากหน้า WEB Report (stockmove) แล้วแสดงแบบ interactive
Deploy: /opt/stockmove_viewer/  |  streamlit run app.py --server.port 87
"""
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Stock Movement Viewer", page_icon="📦", layout="wide")

# ---------------- คอลัมน์มาตรฐานของไฟล์ export ----------------
COLS = [
    "WHCode", "WHName", "RMCode", "RMDesc", "UnitCode", "UnitName",
    "BF", "GR", "TIn", "WipIn", "AdjIn", "RTV", "TOut", "WipOut",
    "Damage", "AdjOut", "SalesUsage", "WasteUsage",
    "TheoryUsage", "TheoryStock", "ActualStock", "Diff", "DiffPct",
    "ActualUsage", "Turnover", "StdCost", "InvCost", "COGS",
]
NUM_COLS = COLS[6:]
MOVE_COLS = ["GR", "TIn", "WipIn", "AdjIn", "RTV", "TOut", "WipOut",
             "Damage", "AdjOut", "SalesUsage", "WasteUsage", "ActualUsage", "Diff"]

TH_LABEL = {
    "BF": "ยอดยกมา", "GR": "รับสินค้า", "TIn": "โอนเข้า", "WipIn": "ผลิตได้",
    "AdjIn": "ปรับเพิ่ม", "RTV": "คืนสินค้า", "TOut": "โอนออก", "WipOut": "ผลิตจาก",
    "Damage": "สูญเสีย", "AdjOut": "ปรับลด", "SalesUsage": "ตัดสต๊อก Sales",
    "WasteUsage": "Waste", "TheoryUsage": "ใช้ตามทฤษฎี", "TheoryStock": "คงเหลือทฤษฎี",
    "ActualStock": "นับจริง", "Diff": "ผลต่างนับ", "DiffPct": "ผลต่าง %",
    "ActualUsage": "ใช้จริง", "Turnover": "Turnover", "StdCost": "ต้นทุน/หน่วย",
    "InvCost": "มูลค่าคงเหลือ", "COGS": "COGS",
}


# ---------------- โหลดไฟล์ ----------------
@st.cache_data(show_spinner="กำลังอ่านไฟล์ Excel...")
def load_export(file_bytes: bytes) -> pd.DataFrame:
    """หา header row อัตโนมัติ (แถวที่ขึ้นต้นด้วย 'Warehouse Code') แล้วอ่านข้อมูล"""
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=25)
    header_row = None
    for i, v in raw[0].items():
        if str(v).strip().lower() == "warehouse code":
            header_row = i
            break
    if header_row is None:
        header_row = 10  # โครงสร้างมาตรฐานของ export
    df = pd.read_excel(io.BytesIO(file_bytes), header=None,
                       skiprows=header_row + 2, names=COLS)  # +2 ข้าม header EN/TH
    df = df.dropna(subset=["WHCode", "RMCode"])
    df["WHCode"] = df["WHCode"].astype(str).str.strip()
    df["RMCode"] = df["RMCode"].astype(str).str.strip()
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["Shop"] = df["WHCode"] + " | " + df["WHName"].astype(str)
    df["Item"] = df["RMCode"] + " | " + df["RMDesc"].astype(str)
    df["HasMove"] = df[MOVE_COLS].abs().sum(axis=1) > 0
    # หมวดจาก prefix รหัสวัตถุดิบ เช่น NCGD, NCRF, CPF
    df["Prefix"] = df["RMCode"].str.extract(r"^([A-Za-z]+)")[0].fillna("อื่นๆ")
    return df


def fmt(df: pd.DataFrame, cols=None) -> "st.delta_generator.DataFrame":
    cols = cols or [c for c in df.columns if c in NUM_COLS]
    return df.style.format({c: "{:,.2f}" for c in cols})


def dl_button(df: pd.DataFrame, name: str):
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("⬇️ ดาวน์โหลดตารางนี้ (Excel)", buf.getvalue(),
                       file_name=name, key=f"dl_{name}")


# ---------------- UI ----------------
st.title("📦 Stock Movement Viewer")
st.caption("อัปโหลดไฟล์ Excel ที่ export จากหน้า Stock Movement Report (CM POS Web Report)")

up = st.file_uploader("เลือกไฟล์ StockMovement_*.xlsx", type=["xlsx"])
if not up:
    st.info("⬆️ อัปโหลดไฟล์เพื่อเริ่มใช้งาน")
    st.stop()

df = load_export(up.getvalue())

# ---------------- Sidebar filters ----------------
with st.sidebar:
    st.header("🔎 ตัวกรอง")
    shops = st.multiselect("สาขา / คลัง", sorted(df["Shop"].unique()))
    prefixes = st.multiselect("กลุ่มรหัสวัตถุดิบ", sorted(df["Prefix"].unique()))
    q = st.text_input("ค้นหา รหัส/ชื่อวัตถุดิบ")
    only_move = st.toggle("แสดงเฉพาะแถวที่มี movement", value=True,
                          help="ตัดแถว 0 ทั้งบรรทัดออก (ปกติ ~88% ของไฟล์)")

f = df.copy()
if shops:
    f = f[f["Shop"].isin(shops)]
if prefixes:
    f = f[f["Prefix"].isin(prefixes)]
if q:
    f = f[f["Item"].str.contains(q, case=False, na=False)]
if only_move:
    f = f[f["HasMove"]]

# ---------------- KPI ----------------
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("แถวข้อมูล", f"{len(f):,}", f"จากทั้งหมด {len(df):,}")
k2.metric("สาขา", f["WHCode"].nunique())
k3.metric("วัตถุดิบ", f["RMCode"].nunique())
k4.metric("มูลค่าคงเหลือ", f"{f['InvCost'].sum():,.0f}")
k5.metric("COGS", f"{f['COGS'].sum():,.0f}")
waste_cost = (f["WasteUsage"] * f["StdCost"]).sum() + (f["Damage"] * f["StdCost"]).sum()
k6.metric("มูลค่า Waste+Damage", f"{waste_cost:,.0f}")

tab_ov, tab_ex, tab_item, tab_shop, tab_raw = st.tabs(
    ["📊 ภาพรวมรายสาขา", "⚠️ Exception", "🔍 เจาะรายวัตถุดิบ", "🏪 เจาะรายสาขา", "📄 ข้อมูลดิบ"])

# ---------------- ภาพรวมรายสาขา ----------------
with tab_ov:
    g = (f.groupby(["WHCode", "WHName"], as_index=False)
           .agg(Items=("RMCode", "nunique"),
                InvCost=("InvCost", "sum"), COGS=("COGS", "sum"),
                SalesUsage=("SalesUsage", "sum"),
                WasteQty=("WasteUsage", "sum"), DamageQty=("Damage", "sum"),
                DiffRows=("Diff", lambda s: (s != 0).sum())))
    g["Waste%ofCOGS"] = 0.0
    wcost = (f.assign(WC=(f["WasteUsage"] + f["Damage"]) * f["StdCost"])
               .groupby("WHCode")["WC"].sum())
    g["WasteCost"] = g["WHCode"].map(wcost).fillna(0)
    g.loc[g["COGS"] != 0, "Waste%ofCOGS"] = g["WasteCost"] / g["COGS"] * 100
    g = g.sort_values("COGS", ascending=False)
    st.dataframe(fmt(g, ["InvCost", "COGS", "SalesUsage", "WasteQty",
                         "DamageQty", "WasteCost", "Waste%ofCOGS"]),
                 use_container_width=True, height=560)
    dl_button(g, "summary_by_shop.xlsx")

# ---------------- Exception ----------------
with tab_ex:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("❗ นับสต๊อกมีผลต่าง (Diff ≠ 0)")
        d = f[f["Diff"] != 0][["Shop", "Item", "UnitCode", "TheoryStock",
                               "ActualStock", "Diff", "DiffPct", "StdCost"]]
        d = d.assign(DiffCost=d["Diff"] * d["StdCost"]).sort_values("DiffCost")
        st.dataframe(fmt(d), use_container_width=True)

        st.subheader("🔻 สต๊อกติดลบ (BF / Theory / Actual < 0)")
        neg = f[(f["BF"] < 0) | (f["TheoryStock"] < 0) | (f["ActualStock"] < 0)]
        neg = neg[["Shop", "Item", "UnitCode", "BF", "TheoryStock",
                   "ActualStock", "InvCost"]].sort_values("TheoryStock")
        st.dataframe(fmt(neg), use_container_width=True)
        st.caption(f"พบ {len(neg):,} แถว — สต๊อกติดลบมักเกิดจากตัดสต๊อกก่อนรับของ / สูตร WIP ผิด / ยังไม่ทำ GR")
    with c2:
        st.subheader("🗑️ Waste + Damage สูงสุด (ตามมูลค่า)")
        w = f.assign(WasteCost=(f["WasteUsage"] + f["Damage"]) * f["StdCost"])
        w = w[w["WasteCost"] > 0][["Shop", "Item", "WasteUsage", "Damage",
                                   "StdCost", "WasteCost"]]
        st.dataframe(fmt(w.sort_values("WasteCost", ascending=False).head(50)),
                     use_container_width=True)

        st.subheader("⚖️ ใช้จริง ≠ ทฤษฎี มากสุด (Usage Variance)")
        v = f.assign(VarQty=f["ActualUsage"] - f["TheoryUsage"])
        v = v.assign(VarCost=v["VarQty"] * v["StdCost"])
        v = v[v["VarQty"].abs() > 0][["Shop", "Item", "TheoryUsage",
                                      "ActualUsage", "VarQty", "VarCost"]]
        st.dataframe(fmt(v.reindex(v["VarCost"].abs()
                          .sort_values(ascending=False).index).head(50)),
                     use_container_width=True)

# ---------------- เจาะรายวัตถุดิบ ----------------
with tab_item:
    item = st.selectbox("เลือกวัตถุดิบ", sorted(f["Item"].unique()))
    it = f[f["Item"] == item]
    st.markdown(f"**{item}** — หน่วย: {it['UnitName'].iloc[0] if len(it) else '-'} "
                f"| ต้นทุน/หน่วย: {it['StdCost'].max():,.2f}")
    show = it[["Shop", "BF", "GR", "TIn", "TOut", "SalesUsage", "WasteUsage",
               "Damage", "TheoryUsage", "TheoryStock", "ActualStock", "Diff",
               "ActualUsage", "InvCost", "COGS"]].sort_values("COGS", ascending=False)
    st.dataframe(fmt(show), use_container_width=True, height=500)
    st.bar_chart(it.set_index("WHCode")["ActualUsage"])
    dl_button(show, "item_across_shops.xlsx")

# ---------------- เจาะรายสาขา ----------------
with tab_shop:
    shop = st.selectbox("เลือกสาขา", sorted(f["Shop"].unique()))
    sh = f[f["Shop"] == shop]
    sort_by = st.selectbox("เรียงตาม", ["COGS", "InvCost", "ActualUsage",
                                        "WasteUsage", "Diff"], index=0)
    show = sh[["Item", "UnitCode", "BF", "GR", "TIn", "TOut", "WipIn", "WipOut",
               "SalesUsage", "WasteUsage", "Damage", "TheoryStock", "ActualStock",
               "Diff", "ActualUsage", "StdCost", "InvCost", "COGS"]]
    show = show.reindex(show[sort_by].abs().sort_values(ascending=False).index)
    st.dataframe(fmt(show), use_container_width=True, height=600)
    dl_button(show, "shop_detail.xlsx")

# ---------------- ข้อมูลดิบ ----------------
with tab_raw:
    st.caption("คอลัมน์ครบตามไฟล์ export — กรองด้วย sidebar ได้")
    st.dataframe(fmt(f[COLS]), use_container_width=True, height=650)
    dl_button(f[COLS], "filtered_raw.xlsx")