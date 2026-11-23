"""
🐾 GREYHOUND CAFÉ – Promotion Intelligence Dashboard
ดึงข้อมูลจาก MySQL | Built with Streamlit + Plotly
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pymysql
from datetime import datetime, timedelta

import warnings
warnings.filterwarnings("ignore")
import io as _io
import openpyxl as _opx
from openpyxl.styles import Font as _Font, PatternFill as _Fill, Alignment as _Align, Border as _Border, Side as _Side
from openpyxl.utils import get_column_letter as _gcl
from datetime import date as _date

def build_excel(df, start_date: _date, end_date: _date) -> bytes:
    period_days  = (end_date - start_date).days + 1
    period_label = f"{period_days}-Day Period"
    date_range   = f"{start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}"

    # ── Warm Cream Light Palette ──────────────────────────────
    BG_PAGE   = "FAF8F5"   # warm off-white
    BG_HDR    = "F2EDE3"   # warm cream header
    BG_BAND   = "EDE5D5"   # gold-tinted section bg
    BG_WHITE  = "FFFFFF"
    GOLD      = "7A5A0A"   # deep gold (readable on light)
    GOLD_ACC  = "B48C50"   # tab accent gold
    TEXT_DARK = "1C1812"   # near-black warm
    TEXT_MID  = "6B5E4E"   # muted brown
    TEXT_LITE = "A89880"   # light muted
    BORDER    = "DDD5C8"   # soft warm border
    RULE      = "C49A2A"   # gold rule line

    CAT_COLORS = {
        "Stars":       {"bg":"FFFBEE","fg":"6B4A08","accent":"C49A2A","border":"E8D48A"},
        "Puzzles":     {"bg":"EEF6FA","fg":"1A5070","accent":"2E86AB","border":"A8D4E6"},
        "Plow Horses": {"bg":"FFF2F2","fg":"6B1A1A","accent":"C0392B","border":"F0AAAA"},
        "Dogs":        {"bg":"F5F4F2","fg":"4A4540","accent":"888078","border":"C8C0B8"},
    }
    CAT_ORDER = ["Stars","Plow Horses","Puzzles","Dogs"]
    CAT_EMOJI = {"Stars":"⭐","Plow Horses":"🐴","Puzzles":"🧩","Dogs":"🐶"}

    def _fill(c):     return _Fill("solid", fgColor=c)
    def _align(h="left", v="center"):
        return _Align(horizontal=h, vertical=v)
    def _bb(color, style="thin"):
        return _Border(bottom=_Side(style=style, color=color))
    def _bt(color, style="thin"):
        return _Border(top=_Side(style=style, color=color))
    def _hair():
        return _Border(bottom=_Side(style="hair", color=BORDER))

    wb = _opx.Workbook()

    # ── SHEET 1: Overview ─────────────────────────────────────
    ws = wb.active
    ws.title = "Overview"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = GOLD_ACC

    total_items  = len(df)
    total_qty    = df["จำนวน"].sum()
    total_sales  = df["ยอดขาย"].sum()
    total_profit = df["กำไรรวม"].sum()
    gp_pct       = total_profit / total_sales * 100 if total_sales else 0
    avg_cm       = df["CM/unit"].mean()

    # Title block
    for row, (txt, sz, color, ht) in enumerate([
        ("GREYHOUND COFFEE",               8,  GOLD,      16),
        ("Menu Engineering Report",        22, TEXT_DARK, 38),
        (f"{period_label}   ·   {date_range}", 10, TEXT_MID, 20),
        ("",                               9,  BG_HDR,    6),
    ], start=1):
        ws.merge_cells(f"A{row}:G{row}")
        c = ws[f"A{row}"]
        c.value = txt
        c.font  = _Font(color=color, size=sz, bold=(sz<=9), name="Arial")
        c.fill  = _fill(BG_HDR)
        c.alignment = _align("left")
        ws.row_dimensions[row].height = ht

    # Gold rule under header
    ws["A4"].border = _Border(bottom=_Side(style="medium", color=RULE))

    # Portfolio Summary label
    ws.merge_cells("A5:G5")
    ws["A5"].value = "PORTFOLIO SUMMARY"
    ws["A5"].font  = _Font(color=GOLD, size=8, bold=True, name="Arial")
    ws["A5"].fill  = _fill(BG_BAND)
    ws["A5"].alignment = _align("left")
    ws["A5"].border = _bb(RULE)
    ws.row_dimensions[5].height = 22

    # KPI block
    kpis = [
        ("TOTAL ITEMS",  total_items,  "{:,}"),
        ("QTY SOLD",     total_qty,    "{:,}"),
        ("REVENUE",      total_sales,  "฿{:,.0f}"),
        ("PROFIT",       total_profit, "฿{:,.0f}"),
        ("GP %",         gp_pct,       "{:.1f}%"),
        ("AVG CM/UNIT",  avg_cm,       "฿{:,.2f}"),
    ]
    for col, (lbl, val, fmt) in enumerate(kpis, 1):
        cl = ws.cell(row=6, column=col, value=lbl)
        cl.font = _Font(color=TEXT_LITE, size=7, bold=True, name="Arial")
        cl.fill = _fill("F8F4EC"); cl.alignment = _align("center")

        cv = ws.cell(row=7, column=col, value=fmt.format(val))
        cv.font = _Font(color=TEXT_DARK, size=16, bold=True, name="Arial")
        cv.fill = _fill("F8F4EC"); cv.alignment = _align("center")
        cv.border = _bb(BORDER)
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 30

    # Spacer
    ws.merge_cells("A8:G8")
    ws["A8"].fill = _fill(BG_PAGE)
    ws.row_dimensions[8].height = 10

    # Category Breakdown label
    ws.merge_cells("A9:G9")
    ws["A9"].value = "CATEGORY BREAKDOWN"
    ws["A9"].font  = _Font(color=GOLD, size=8, bold=True, name="Arial")
    ws["A9"].fill  = _fill(BG_BAND)
    ws["A9"].alignment = _align("left")
    ws["A9"].border = _bb(RULE)
    ws.row_dimensions[9].height = 22

    # Table header
    for c, h in enumerate(["Category","Items","Items %","Qty Sold","Revenue (฿)","Profit (฿)","GP %"], 1):
        cell = ws.cell(row=10, column=c, value=h)
        cell.font = _Font(color=TEXT_MID, size=8, bold=True, name="Arial")
        cell.fill = _fill(BG_BAND)
        cell.alignment = _align("center" if c > 1 else "left")
        cell.border = _Border(bottom=_Side(style="medium", color=RULE))
    ws.row_dimensions[10].height = 20

    # Category rows
    for off, cat in enumerate(CAT_ORDER):
        sub = df[df["Category"]==cat]; row = 11+off; cc = CAT_COLORS[cat]
        s = sub["ยอดขาย"].sum(); p = sub["กำไรรวม"].sum()
        for c, val in enumerate([
            f"{CAT_EMOJI[cat]}  {cat}", len(sub), f"{len(sub)/total_items*100:.1f}%",
            f"{sub['จำนวน'].sum():,}", f"฿{s:,.0f}", f"฿{p:,.0f}",
            f"{p/s*100:.1f}%" if s else "0%",
        ], 1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.fill = _fill(cc["bg"])
            cell.font = _Font(color=cc["fg"] if c==1 else TEXT_DARK, size=10, bold=(c==1), name="Arial")
            cell.alignment = _align("left" if c==1 else "center")
            cell.border = _Border(bottom=_Side(style="hair", color=cc["border"]))
        ws.row_dimensions[row].height = 22

    # Total row
    total_row = 11 + len(CAT_ORDER)
    for c, val in enumerate(["TOTAL",total_items,"100%",f"{total_qty:,}",
                              f"฿{total_sales:,.0f}",f"฿{total_profit:,.0f}",f"{gp_pct:.1f}%"], 1):
        cell = ws.cell(row=total_row, column=c, value=val)
        cell.fill = _fill(BG_BAND)
        cell.font = _Font(color=GOLD, size=10, bold=True, name="Arial")
        cell.alignment = _align("left" if c==1 else "center")
        cell.border = _Border(top=_Side(style="medium", color=RULE))
    ws.row_dimensions[total_row].height = 24

    # Column widths
    for i, w in enumerate([22,10,10,14,18,18,10], 1):
        ws.column_dimensions[_gcl(i)].width = w

    # Fill remaining BG
    for r in range(1, total_row+4):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            if cell.fill.fgColor.rgb in ("00000000","FFFFFFFF"):
                cell.fill = _fill(BG_PAGE)

    # ── SHEETS 2–5: Per Category ──────────────────────────────
    COL_HDR2 = ["#","เมนู","จำนวนขาย","CM/Unit (฿)","กำไรรวม (฿)","ยอดขาย (฿)","GP %"]
    COL_W    = [5, 38, 14, 14, 16, 16, 10]

    for cat in CAT_ORDER:
        cc  = CAT_COLORS[cat]
        sub = df[df["Category"]==cat].sort_values("กำไรรวม", ascending=False).reset_index(drop=True)
        ws2 = wb.create_sheet(title=f"{CAT_EMOJI[cat]} {cat}")
        ws2.sheet_view.showGridLines = False
        ws2.sheet_properties.tabColor = cc["accent"]

        # Title block
        for row, (txt, sz, color, ht) in enumerate([
            ("GREYHOUND COFFEE  ·  MENU ENGINEERING",             8,  GOLD,     16),
            (f"{CAT_EMOJI[cat]}  {cat}",                          20, cc["fg"], 34),
            (f"{period_label}   ·   {date_range}   ·   {len(sub)} items", 10, TEXT_MID, 18),
            ("",                                                   9,  BG_HDR,  6),
        ], start=1):
            ws2.merge_cells(f"A{row}:G{row}")
            c = ws2[f"A{row}"]
            c.value = txt
            c.font  = _Font(color=color, size=sz, bold=(sz<=9), name="Arial")
            c.fill  = _fill(BG_HDR)
            c.alignment = _align("left")
            ws2.row_dimensions[row].height = ht

        ws2["A4"].border = _Border(bottom=_Side(style="medium", color=cc["accent"]))

        # Mini KPIs
        s = sub["ยอดขาย"].sum(); p = sub["กำไรรวม"].sum()
        mini = [
            ("ITEMS",       f"{len(sub):,}"),
            ("QTY SOLD",    f"{sub['จำนวน'].sum():,}"),
            ("REVENUE",     f"฿{s:,.0f}"),
            ("PROFIT",      f"฿{p:,.0f}"),
            ("GP %",        f"{p/s*100:.1f}%" if s else "0%"),
            ("AVG CM/UNIT", f"฿{sub['CM/unit'].mean():,.2f}"),
        ]
        for col, (lbl, val) in enumerate(mini, 1):
            cl = ws2.cell(row=5, column=col, value=lbl)
            cl.font = _Font(color=TEXT_LITE, size=7, bold=True, name="Arial")
            cl.fill = _fill(cc["bg"]); cl.alignment = _align("center")

            cv = ws2.cell(row=6, column=col, value=val)
            cv.font = _Font(color=cc["fg"], size=13, bold=True, name="Arial")
            cv.fill = _fill(cc["bg"]); cv.alignment = _align("center")
            cv.border = _Border(bottom=_Side(style="thin", color=cc["border"]))
        ws2.row_dimensions[5].height = 16
        ws2.row_dimensions[6].height = 26

        # Spacer
        ws2.merge_cells("A7:G7")
        ws2["A7"].fill = _fill(BG_PAGE)
        ws2.row_dimensions[7].height = 8

        # Table header
        for c, (h, w) in enumerate(zip(COL_HDR2, COL_W), 1):
            cell = ws2.cell(row=8, column=c, value=h)
            cell.font = _Font(color=cc["fg"], size=9, bold=True, name="Arial")
            cell.fill = _fill(BG_BAND)
            cell.alignment = _align("left" if c==2 else "center")
            cell.border = _Border(bottom=_Side(style="medium", color=cc["accent"]))
            ws2.column_dimensions[_gcl(c)].width = w
        ws2.row_dimensions[8].height = 20

        # Data rows
        for idx, row_data in sub.iterrows():
            r   = idx + 9
            gp  = row_data["กำไรรวม"]/row_data["ยอดขาย"] if row_data["ยอดขาย"] else 0
            row_bg = BG_WHITE if idx%2==0 else BG_PAGE
            vals = [idx+1, row_data["เมนู"], row_data["จำนวน"],
                    row_data["CM/unit"], row_data["กำไรรวม"], row_data["ยอดขาย"], gp]
            fmts = [None, None, '#,##0', '฿#,##0.00', '฿#,##0.00', '฿#,##0.00', '0.0%']
            for c, (val, fmt) in enumerate(zip(vals, fmts), 1):
                cell = ws2.cell(row=r, column=c, value=val)
                cell.fill = _fill(row_bg)
                cell.font = _Font(color=TEXT_LITE if c==1 else TEXT_DARK, size=9, name="Arial")
                cell.alignment = _align("left" if c==2 else "center")
                cell.border = _hair()
                if fmt: cell.number_format = fmt
            ws2.row_dimensions[r].height = 18

        ws2.freeze_panes = "A9"

        # Fill remaining BG
        last = len(sub)+9
        for r in range(1, last+3):
            for c in range(1, 8):
                cell = ws2.cell(row=r, column=c)
                if cell.fill.fgColor.rgb in ("00000000","FFFFFFFF"):
                    cell.fill = _fill(BG_PAGE)

    buf = _io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()
# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))   # หา gh_env.py ข้างไฟล์นี้ (ไม่พึ่ง cwd/launcher)
import gh_env   # รหัส DB อยู่ใน .env (ไม่เข้า git)
MYSQL_CFG = gh_env.mysql_cfg("root")
ETL_ADMIN_URL = "http://10.8.1.88:89"

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="GH Promotion Dashboard",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #0d1117; color: #e6edf3; }

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #161b22 0%, #0d1117 100%);
    border-right: 1px solid #30363d;
}
section[data-testid="stSidebar"] * { color: #c9d1d9 !important; }

[data-testid="stMetric"] {
    background: linear-gradient(135deg, #161b22, #1c2230);
    border: 1px solid #30363d; border-radius: 16px; padding: 20px !important;
    transition: all .3s ease;
}
[data-testid="stMetric"]:hover { border-color: #d4af37; transform: translateY(-2px); }
[data-testid="stMetricLabel"] { color: #8b949e !important; font-size: 13px !important; font-weight: 500 !important; }
[data-testid="stMetricValue"] { color: #d4af37 !important; font-size: 28px !important; font-weight: 700 !important; }
[data-testid="stMetricDelta"]  { font-size: 13px !important; }

.card { background: linear-gradient(135deg,#161b22,#1c2230); border:1px solid #30363d; border-radius:16px; padding:20px; margin:8px 0; }
.card-gold     { border-left: 4px solid #d4af37; }
.card-silver   { border-left: 4px solid #9e9e9e; }
.card-platinum { border-left: 4px solid #e5e4e2; }

.dash-header {
    background: linear-gradient(135deg, #1a0a00, #3d2200, #1a0a00);
    border: 1px solid #d4af37; border-radius: 20px;
    padding: 28px 36px; margin-bottom: 24px; text-align: center;
}
.dash-header h1 { color:#d4af37; font-size:32px; font-weight:800; margin:0; letter-spacing:2px; }
.dash-header p  { color:#a0824a; font-size:14px; margin:4px 0 0; }

.section-title {
    font-size: 18px; font-weight: 700; color: #d4af37;
    border-bottom: 2px solid #30363d; padding-bottom: 8px; margin: 24px 0 16px;
}

.badge { display:inline-block; padding:4px 12px; border-radius:20px; font-size:11px; font-weight:600; letter-spacing:.5px; }
.badge-gold     { background:#d4af3730; color:#d4af37; border:1px solid #d4af37; }
.badge-silver   { background:#9e9e9e30; color:#c9d1d9; border:1px solid #9e9e9e; }
.badge-black    { background:#21262d;   color:#c9d1d9; border:1px solid #30363d; }
.badge-platinum { background:#e5e4e230; color:#e5e4e2; border:1px solid #e5e4e2; }
.badge-senior   { background:#58a6ff30; color:#58a6ff; border:1px solid #58a6ff; }

.rank-row { display:flex; align-items:center; justify-content:space-between; padding:12px 16px; margin:6px 0; background:#161b22; border:1px solid #30363d; border-radius:10px; transition:all .2s; }
.rank-row:hover { border-color: #d4af37; }
.rank-num { font-size:20px; font-weight:800; color:#8b949e; width:36px; }
.rank-num.top1 { color:#d4af37; } .rank-num.top2 { color:#9e9e9e; } .rank-num.top3 { color:#cd7f32; }
.rank-name { flex:1; font-weight:600; color:#e6edf3; font-size:14px; }
.rank-sub  { font-size:11px; color:#8b949e; }
.rank-amt  { font-weight:700; color:#d4af37; font-size:15px; }

.promo-card { background:linear-gradient(135deg,#1a2332,#1c2230); border:1px solid #30363d; border-radius:14px; padding:18px; margin:8px 0; display:flex; align-items:center; gap:16px; }
.promo-icon { font-size:28px; }
.promo-info h4 { margin:0 0 4px; color:#e6edf3; font-size:15px; font-weight:600; }
.promo-info p  { margin:0; color:#8b949e; font-size:12px; }
.promo-stat { margin-left:auto; text-align:right; }
.promo-stat .amt { font-size:20px; font-weight:800; color:#d4af37; }
.promo-stat .lbl { font-size:11px; color:#8b949e; }

.stTabs [data-baseweb="tab-list"] { background:#161b22 !important; border-radius:12px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:8px !important; color:#8b949e !important; font-weight:600 !important; font-size:13px !important; }
.stTabs [aria-selected="true"] { background:#d4af37 !important; color:#0d1117 !important; }

.stDateInput input, .stSelectbox div { background:#161b22 !important; border-color:#30363d !important; }
::-webkit-scrollbar { width:6px; } ::-webkit-scrollbar-track { background:#0d1117; }
::-webkit-scrollbar-thumb { background:#30363d; border-radius:3px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────
GOLD      = "#d4af37"
DARK_BG   = "#0d1117"
BLUE      = "#58a6ff"
GREEN     = "#3fb950"
RED       = "#f85149"
ORANGE    = "#d29922"
SILVER    = "#9e9e9e"
COLOR_SEQ = [GOLD, BLUE, GREEN, ORANGE, RED, SILVER, "#bc8cff", "#ff7b72", "#ffa657"]

LAYOUT_BASE = dict(
    paper_bgcolor = "rgba(0,0,0,0)",
    plot_bgcolor  = "rgba(0,0,0,0)",
    font  = dict(family="Inter", color="#c9d1d9", size=12),
    margin= dict(l=0, r=0, t=36, b=0),
    xaxis = dict(gridcolor="#21262d", zerolinecolor="#30363d", tickfont=dict(color="#8b949e")),
    yaxis = dict(gridcolor="#21262d", zerolinecolor="#30363d", tickfont=dict(color="#8b949e")),
    legend= dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#c9d1d9")),
)

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
def load_data(date_from: str, date_to: str, shop_codes: tuple):
    # ไม่ cache ตรงนี้แล้ว — cache ที่ load_clean() (ผลหลัง clean) แทน จะได้ไม่เก็บ 2 ชุดในหน่วยความจำ
    conn = pymysql.connect(**MYSQL_CFG)
    try:
        shop_filter, params = "", [date_from, date_to]
        if shop_codes:
            shop_filter = f" AND ShopCode IN ({','.join(['%s'] * len(shop_codes))})"
            params += list(shop_codes)

        def q_trans(tbl):
            return pd.read_sql(f"SELECT * FROM {tbl} WHERE BillDate BETWEEN %s AND %s" + shop_filter,
                               conn, params=params)

        hdr  = q_trans("BillHdrTB")
        dtl  = q_trans("BillDtlTB")
        disc = q_trans("ONLBillMultiDiscTB")
        pay  = q_trans("BILLPAYTB")

        memtype = pd.read_sql("SELECT * FROM MemTypetb",  conn)
        shop    = pd.read_sql("SELECT ShopCode, ShopDes1, VatType FROM PosShopTB",  conn)
        itm     = pd.read_sql("SELECT * FROM ItmMasTB",   conn)
        dep     = pd.read_sql("SELECT * FROM ItmDepTB",   conn)
        cat     = pd.read_sql("SELECT * FROM ItmCatTB",   conn)
        membertb = pd.read_sql("SELECT MemCode, MemType, FirstName1, LastName1 FROM MemberTB", conn)
    finally:
        conn.close()

    return hdr, dtl, disc, pay, memtype, shop, itm, dep, cat, membertb


@st.cache_data(ttl=600, show_spinner=False)
def _load_shops():
    conn = pymysql.connect(**MYSQL_CFG)
    try:
        df = pd.read_sql(
            "SELECT ShopCode, ShopDes1, ShTypCode FROM PosShopTB ORDER BY ShopCode", conn)
    finally:
        conn.close()
    return df.drop_duplicates("ShopCode")


def get_shops():
    # cache เฉพาะตอนสำเร็จ (ถ้า DB ล่ม จะไม่ค้างรายการว่างไว้ 10 นาที)
    try:
        return _load_shops()
    except Exception:
        return pd.DataFrame(columns=["ShopCode","ShopDes1","ShTypCode"])


def clean_data(hdr, dtl, disc, pay, memtype, shop, itm, dep, cat, membertb):
    hdr  = hdr.copy();  dtl = dtl.copy()
    disc = disc.copy(); pay = pay.copy()

    # numeric
    for c in ["SubTotal","BillDisc","MemDisc","NetBAmt","OthChrg","NoPerson","VatAmt1","VatAmt2"]:
        if c in hdr.columns: hdr[c] = pd.to_numeric(hdr[c], errors="coerce").fillna(0)
    for c in ["SellPrice","ActAmt","NetAmt","ActQTY","totalDis"]:
        if c in dtl.columns: dtl[c] = pd.to_numeric(dtl[c], errors="coerce").fillna(0)
    for c in ["DiscAmt","TotBillDisc"]:
        if c in disc.columns: disc[c] = pd.to_numeric(disc[c], errors="coerce").fillna(0)

    # date
    for df in [hdr, dtl, disc, pay]:
        if "BillDate" in df.columns:
            df["BillDate"] = pd.to_datetime(df["BillDate"], errors="coerce")

    # hour
    if "BillTime" in hdr.columns:
        hdr["BillHour"] = pd.to_numeric(hdr["BillTime"].astype(str).str[:2], errors="coerce")

    # ไม่กรอง void — รวมทุก BillType (INV + VOID)

    # ✅ snapshot ค่าจาก BillHdrTB ก่อน merge ใดๆ — ใช้เป็น ground truth
    # VatType = 1 → ขายไม่รวม VAT → ใช้ NetBAmt ตรงๆ
    # VatType = 0 → ขายรวม VAT   → หัก VatAmt1+VatAmt2 ออก
    if "VatType" in shop.columns:
        vat_map = shop.drop_duplicates("ShopCode").set_index("ShopCode")["VatType"].to_dict()
        hdr["_ShopVatType"] = hdr["ShopCode"].astype(str).map(vat_map).fillna(1)
    else:
        hdr["_ShopVatType"] = 1  # default: ไม่รวม VAT

    vat1 = pd.to_numeric(hdr.get("VatAmt1", 0), errors="coerce").fillna(0)
    vat2 = pd.to_numeric(hdr.get("VatAmt2", 0), errors="coerce").fillna(0)
    net  = pd.to_numeric(hdr["NetBAmt"], errors="coerce").fillna(0)
    # VatType=0 → หัก VAT, VatType=1 → ใช้เต็ม
    hdr["_NetBAmt"]  = net - ((hdr["_ShopVatType"] == 0).astype(int) * (vat1 + vat2))
    hdr["_NoPerson"] = hdr["NoPerson"] if "NoPerson" in hdr.columns else 0

    # ส่วนลดรวม: BillDisc + MemDisc + CoupDisc + ItemDisc จาก BillHdrTB
    def _disc_col(col):
        if col in hdr.columns:
            return pd.to_numeric(hdr[col], errors="coerce").fillna(0)
        return pd.Series(0, index=hdr.index)

    hdr["_TotalDisc"] = (_disc_col("BillDisc") + _disc_col("MemDisc") +
                         _disc_col("CoupDisc") + _disc_col("ItemDisc"))
    hdr["_BillDisc"]  = _disc_col("BillDisc")
    hdr["_MemDisc"]   = _disc_col("MemDisc")

    # ✅ flag บิลที่นับได้ — BillStatus NOT IN ('4','5')
    if "BillStatus" in hdr.columns:
        hdr["_CountBill"] = ~hdr["BillStatus"].astype(str).isin(["4","5"])
    else:
        hdr["_CountBill"] = True

    # ── Member — นิยามเดียวทั้งแอป (ตรงกับ BI API และ v_bill_summary) ──────────
    # member = บิลที่มีการจ่ายแบบ MEM ใน BILLPAYTB หรือมีรหัสสมาชิกใน BillHdrTB.Remark
    #   (เช็ค 2026-11-23: สองแหล่งได้ชุดบิลเดียวกันเป๊ะ 6,762/6,762 บิลใน 30 วัน)
    # รหัสสมาชิก = Remark (ตรง MemberTB 98%) — BILLPAYTB.CardNo มีแถวว่างปน (เช่น "Black Hound Discount") ตรงแค่ ~30%
    # ประเภท = MemberTB.MemType → MemTypetb.MemTNam1 (White/Silver/Black/Platinum Hound ...)
    #   BILLPAYTB.Prefix เป็น '0.0000' ทั้งหมด ใช้หา tier ไม่ได้ (เดิมทุกบิลได้แค่คำว่า "Member")
    hdr["ShopCode"] = hdr["ShopCode"].astype(str)
    hkey = hdr["ShopCode"] + "|" + hdr["BillNo"].astype(str)
    rem = (hdr["Remark"].fillna("").astype(str).str.strip() if "Remark" in hdr.columns
           else pd.Series("", index=hdr.index))
    rem = rem.where(~rem.isin(["0", "None", "nan"]), "")
    has_mem_pay = pd.Series(False, index=hdr.index)
    card = pd.Series("", index=hdr.index)
    if "tendtype" in pay.columns:
        pm = pay[pay["tendtype"] == "MEM"]
        pm_key = pm["ShopCode"].astype(str) + "|" + pm["BillNo"].astype(str)
        has_mem_pay = hkey.isin(set(pm_key))
        if "CardNo" in pm.columns:
            cards = pd.DataFrame({"k": pm_key, "c": pm["CardNo"].fillna("").astype(str).str.strip()})
            cards = cards[cards["c"] != ""].drop_duplicates("k").set_index("k")["c"]
            card = hkey.map(cards).fillna("")
    hdr["IsMember"]   = has_mem_pay | (rem != "")
    hdr["MemberCode"] = rem.where(rem != "", card).where(hdr["IsMember"], "")
    tier = (membertb[["MemCode", "MemType"]]
            .assign(MemCode=lambda d: d["MemCode"].astype(str).str.strip())
            .drop_duplicates("MemCode")
            .merge(memtype[["MemType", "MemTNam1"]].drop_duplicates("MemType"), on="MemType", how="left")
            .set_index("MemCode")["MemTNam1"])
    hdr["MemberTypeName"] = (hdr["MemberCode"].map(tier)
                             .where(hdr["IsMember"], "Guest")
                             .fillna("Member (ไม่พบในทะเบียน)"))

    # item joins
    itm_cols = [c for c in ["ItemCode","ItemDes1","ItemDes2","DepCode0","CatCode0"] if c in itm.columns]
    dtl = dtl.merge(itm[itm_cols], on="ItemCode", how="left")

    dep2 = dep[["DepCode","DepDes1"]].copy()
    dep2.columns = ["DepCode0","DepName"]
    dep2["DepCode0"] = dep2["DepCode0"].astype(str)
    dtl["DepCode0"]  = dtl["DepCode0"].astype(str)
    dtl = dtl.merge(dep2, on="DepCode0", how="left")
    cat2 = cat[["CatCode","CatDes1"]].copy()
    cat2.columns = ["CatCode0","CatName"]
    cat2["CatCode0"] = cat2["CatCode0"].astype(str)
    dtl["CatCode0"]  = dtl["CatCode0"].astype(str)
    dtl = dtl.merge(cat2, on="CatCode0", how="left")

    # shop name — dedup ก่อน merge ป้องกัน row เพิ่ม
    shop2 = (shop[["ShopCode","ShopDes1"]]
             .drop_duplicates("ShopCode")          # ← key fix
             .rename(columns={"ShopDes1":"ShopName"}))
    hdr["ShopCode"] = hdr["ShopCode"].astype(str)
    shop2["ShopCode"] = shop2["ShopCode"].astype(str)
    n_before = len(hdr)
    hdr = hdr.merge(shop2, on="ShopCode", how="left")
    # safety: ถ้า row เพิ่มขึ้นให้ตัดทิ้ง
    if len(hdr) > n_before:
        pk = [c for c in ["ShopCode","TerminalNo","BillDate","BillType","BillNo"] if c in hdr.columns]
        hdr = hdr.drop_duplicates(subset=pk, keep="first")

    # คีย์บิล = ShopCode|BillNo — เลขบิล (BillNo) ซ้ำข้ามสาขาได้ (~25% ของเลขบิลใน 30 วัน)
    # ห้ามนับ/merge ด้วย BillNo อย่างเดียว ใช้ _BillKey แทน (ShopCode+BillNo ไม่ซ้ำทั้งตาราง)
    for _df in (hdr, dtl, disc, pay):
        if "ShopCode" in _df.columns and "BillNo" in _df.columns:
            _df["_BillKey"] = _df["ShopCode"].astype(str) + "|" + _df["BillNo"].astype(str)

    return hdr, dtl, disc, pay, memtype, shop, itm, dep, cat, membertb


@st.cache_data(ttl=300, show_spinner=False, max_entries=8)
def load_clean(date_from: str, date_to: str, shop_codes: tuple):
    """โหลด + clean แล้ว cache ผลลัพธ์ — เดิม clean_data (merge หลายรอบ) รันใหม่ทุกครั้งที่กด widget"""
    return clean_data(*load_data(date_from, date_to, shop_codes))


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:16px 0 8px;">
        <div style="font-size:40px;">🐾</div>
        <div style="font-size:16px;font-weight:800;color:#d4af37;letter-spacing:2px;">GREYHOUND CAFÉ</div>
        <div style="font-size:11px;color:#8b949e;margin-top:4px;">Promotion Intelligence</div>
    </div>
    <hr style="border-color:#30363d;margin:12px 0;">
    """, unsafe_allow_html=True)

    # ── Date Filter ──
    st.markdown("<div style='font-size:11px;color:#8b949e;font-weight:600;letter-spacing:1px;margin-bottom:8px;'>📅 ช่วงเวลา</div>", unsafe_allow_html=True)

    preset = st.selectbox("Quick Select", [
        "วันนี้", "เมื่อวาน", "7 วันล่าสุด", "30 วันล่าสุด",
        "เดือนนี้", "เดือนที่แล้ว", "กำหนดเอง"
    ], label_visibility="collapsed")

    today = datetime.today().date()
    if preset == "วันนี้":
        d_from, d_to = today, today
    elif preset == "เมื่อวาน":
        d_from = d_to = today - timedelta(days=1)
    elif preset == "7 วันล่าสุด":
        d_from, d_to = today - timedelta(days=6), today
    elif preset == "30 วันล่าสุด":
        d_from, d_to = today - timedelta(days=29), today
    elif preset == "เดือนนี้":
        d_from, d_to = today.replace(day=1), today
    elif preset == "เดือนที่แล้ว":
        first_this = today.replace(day=1)
        d_to   = first_this - timedelta(days=1)
        d_from = d_to.replace(day=1)
    else:
        d_from = st.date_input("ตั้งแต่", today - timedelta(days=6))
        d_to   = st.date_input("ถึง",     today)

    st.markdown(f"<div style='font-size:11px;color:#d4af37;text-align:center;margin:4px 0 8px;'>{d_from.strftime('%d/%m/%Y')} – {d_to.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)

    # ── Brand Filter ──
    st.markdown("<hr style='border-color:#30363d;margin:8px 0;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;color:#8b949e;font-weight:600;letter-spacing:1px;margin-bottom:8px;'>🏷️ แบรนด์</div>", unsafe_allow_html=True)

    BRAND_MAP = {"ทั้งหมด": None, "🐾 Greyhound (GC)": "GC", "🌿 Anotherhound (AH)": "AH", "☕ Beanhound (BE)": "BE"}
    brand_sel = st.selectbox("Brand", list(BRAND_MAP.keys()), label_visibility="collapsed", key="brand_sel")
    selected_brand = BRAND_MAP[brand_sel]

    # ── Shop Filter ──
    st.markdown("<hr style='border-color:#30363d;margin:8px 0;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;color:#8b949e;font-weight:600;letter-spacing:1px;margin-bottom:8px;'>🏪 สาขา</div>", unsafe_allow_html=True)

    shops_df = get_shops()
    # กรองตาม brand ที่เลือก
    if selected_brand:
        shops_filtered = shops_df[shops_df["ShTypCode"] == selected_brand]
    else:
        shops_filtered = shops_df

    shop_opts = ["ทั้งหมด"] + [f"{r.ShopCode} – {r.ShopDes1}" for _, r in shops_filtered.iterrows()]
    shop_sel  = st.multiselect("สาขา", shop_opts, default=["ทั้งหมด"], label_visibility="collapsed")

    if "ทั้งหมด" in shop_sel or not shop_sel:
        # ถ้าเลือก brand → ใช้ทุกสาขาของ brand นั้น
        if selected_brand:
            selected_shops = tuple(shops_filtered["ShopCode"].astype(str).tolist())
        else:
            selected_shops = tuple()
    else:
        selected_shops = tuple(s.split(" – ")[0] for s in shop_sel)

    # ── Navigation ──
    st.markdown("<hr style='border-color:#30363d;margin:8px 0;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;color:#8b949e;font-weight:600;letter-spacing:1px;margin-bottom:8px;'>NAVIGATION</div>", unsafe_allow_html=True)
    page = st.radio("", [
        "🏠 Overview",
        "🍽️ รายการขายดี",
        "⏱️ วิเคราะห์ช่วงเวลา",
        "👥 Member / Guest",
        "🧾 Per Bill Analysis",
        "🎁 วิเคราะห์โปรโมชั่น",
        "🧑‍🏫 Menu Engineering",
        "🥩 Ingredient Forecast",
        "⚙️ Admin",
    ], label_visibility="collapsed")

    st.markdown("<hr style='border-color:#30363d;margin:8px 0;'>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:11px;color:#8b949e;'>อัพเดท: {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# LOAD & CLEAN DATA
# ─────────────────────────────────────────────
if page != "⚙️ Admin":
    with st.spinner("🔄 กำลังโหลดข้อมูล..."):
        try:
            hdr, dtl, disc, pay, memtype, shop, itm, dep, cat, membertb = \
                load_clean(str(d_from), str(d_to), selected_shops)
            data_ok = len(hdr) > 0

            if not data_ok:
                # หาวันล่าสุดที่มีข้อมูล
                conn = pymysql.connect(**MYSQL_CFG)
                try:
                    df_latest = pd.read_sql(
                        "SELECT MAX(BillDate) as d FROM BillHdrTB", conn
                    )
                finally:
                    conn.close()

                latest_date = df_latest["d"].iloc[0]
                if latest_date is not None:
                    latest_date = pd.to_datetime(latest_date).date()
                    st.info(f"📅 ไม่มีข้อมูลในช่วง {d_from} – {d_to} → แสดงข้อมูลล่าสุด: {latest_date.strftime('%d/%m/%Y')}")
                    d_from = latest_date
                    d_to   = latest_date
                    hdr, dtl, disc, pay, memtype, shop, itm, dep, cat, membertb = \
                        load_clean(str(d_from), str(d_to), selected_shops)
                    data_ok = len(hdr) > 0
                else:
                    st.warning("⚠️ ไม่พบข้อมูลในฐานข้อมูลเลย")

        except Exception as e:
            data_ok = False
            st.error(f"❌ เชื่อมต่อ MySQL ไม่ได้: {e}")

# ═══════════════════════════════════════════════════════════════
# PAGE 1: OVERVIEW
# ═══════════════════════════════════════════════════════════════
if page == "🏠 Overview" and data_ok:
    st.markdown(f"""
    <div class="dash-header">
        <h1>🐾 GREYHOUND CAFÉ</h1>
        <p>Promotion Intelligence · {d_from.strftime('%d/%m/%Y')} – {d_to.strftime('%d/%m/%Y')}</p>
    </div>
    """, unsafe_allow_html=True)

    # ใช้ _NetBAmt (snapshot ก่อน merge) เพื่อให้ตรงกับ MySQL sum เสมอ
    total_rev   = hdr["_NetBAmt"].sum()
    total_bills = int(hdr["_CountBill"].sum())
    total_disc  = hdr["_TotalDisc"].sum()
    avg_bill    = total_rev / total_bills if total_bills else 0
    total_pax   = hdr.loc[hdr["_CountBill"], "_NoPerson"].sum()
    mem_bills   = int(hdr.loc[hdr["_CountBill"], "IsMember"].sum())

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("💰 รายรับรวม(Base Sale)",    f"฿{total_rev:,.0f}")
    c2.metric("🧾 จำนวนบิล",     f"{total_bills:,}")
    c3.metric("🎁 ส่วนลดรวม",   f"฿{total_disc:,.0f}")
    c4.metric("📊 Avg/Bill",     f"฿{avg_bill:,.0f}")
    c5.metric("👥 Pax",          f"{int(total_pax):,}")
    c6.metric("🐾 Member Bills", f"{int(mem_bills):,}")

    st.markdown("<br>", unsafe_allow_html=True)
    col_l, col_r = st.columns([3, 2])

    with col_l:
        st.markdown('<div class="section-title">📈 รายรับตามช่วงเวลา</div>', unsafe_allow_html=True)
        if "BillHour" in hdr.columns:
            hourly = (hdr[hdr["_CountBill"]]
                      .groupby("BillHour")["_NetBAmt"].sum().reset_index())
            fig = go.Figure(go.Bar(
                x=hourly["BillHour"], y=hourly["_NetBAmt"],
                marker=dict(color=hourly["_NetBAmt"], colorscale=[[0,"#1c2230"],[1,GOLD]]),
                text=hourly["_NetBAmt"].apply(lambda x: f"฿{x:,.0f}"),
                textposition="outside", textfont=dict(color="#c9d1d9", size=10)
            ))
            fig.update_layout(**LAYOUT_BASE, height=280, title="ยอดขายแต่ละชั่วโมง")
            st.plotly_chart(fig, use_container_width=True)

    with col_r:
        st.markdown('<div class="section-title">🐾 Member vs Guest</div>', unsafe_allow_html=True)
        mv  = hdr.groupby("MemberTypeName")["_NetBAmt"].sum().reset_index()
        fig = go.Figure(go.Pie(
            labels=mv["MemberTypeName"], values=mv["_NetBAmt"], hole=0.55,
            marker=dict(colors=COLOR_SEQ, line=dict(color=DARK_BG, width=2)),
            hovertemplate="%{label}<br>฿%{value:,.0f}<extra></extra>"
        ))
        fig.add_annotation(text=f"<b>฿{total_rev:,.0f}</b>", x=0.5, y=0.5,
                           font=dict(size=13, color=GOLD), showarrow=False)
        fig.update_layout(**LAYOUT_BASE, height=280,
                          legend_orientation="v", legend_x=1, legend_y=0.5)
        st.plotly_chart(fig, use_container_width=True)

    # promotions — กรองเฉพาะบิลที่ _CountBill = True
    if len(disc) > 0:
        st.markdown('<div class="section-title">🎁 โปรโมชั่นที่ใช้งาน</div>', unsafe_allow_html=True)
        valid_bills = hdr.loc[hdr["_CountBill"], ["BillNo","ShopCode"]].drop_duplicates()
        disc_valid  = disc.merge(valid_bills, on=["BillNo","ShopCode"], how="inner") if "ShopCode" in disc.columns \
                      else disc[disc["BillNo"].isin(valid_bills["BillNo"])]
        promo_sum = (disc_valid.groupby("DiscName")
                    .agg(DiscAmt=("DiscAmt","sum"), Bills=("_BillKey","nunique"))
                    .reset_index().sort_values("DiscAmt", ascending=False))
        promo_icons = {"Black Hound Discount":"🖤","White Hound Discount 5%":"🤍",
                       "Silver Hound Discount":"🥈","Senior Discount":"🌟"}
        cols = st.columns(min(max(len(promo_sum), 1), 4))
        for i, (_, row) in enumerate(promo_sum.head(4).iterrows()):
            icon = promo_icons.get(row["DiscName"], "🎫")
            cols[i].markdown(f"""
            <div class="promo-card">
                <div class="promo-icon">{icon}</div>
                <div class="promo-info">
                    <h4>{row['DiscName']}</h4>
                    <p>{int(row['Bills'])} บิล</p>
                </div>
                <div class="promo-stat">
                    <div class="amt">฿{row['DiscAmt']:,.0f}</div>
                    <div class="lbl">ส่วนลดรวม</div>
                </div>
            </div>""", unsafe_allow_html=True)

    # revenue by shop (if multi-shop)
    if "ShopName" in hdr.columns and hdr["ShopCode"].nunique() > 1:
        st.markdown('<div class="section-title">🏪 รายรับตามสาขา</div>', unsafe_allow_html=True)
        shop_rev = (hdr.groupby("ShopName")["_NetBAmt"].sum()
                    .reset_index().sort_values("_NetBAmt", ascending=True).tail(15))
        fig = px.bar(shop_rev, x="_NetBAmt", y="ShopName", orientation="h",
                     color_discrete_sequence=[GOLD])
        fig.update_layout(**LAYOUT_BASE, height=max(300, len(shop_rev)*28))
        st.plotly_chart(fig, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# PAGE 2: รายการขายดี
# ═══════════════════════════════════════════════════════════════
elif page == "🍽️ รายการขายดี" and data_ok:
    st.markdown('<div class="section-title" style="font-size:24px;">🍽️ รายการขายดี</div>', unsafe_allow_html=True)
    sold = dtl[dtl["NetAmt"] > 0].copy()

    tab1, tab2, tab3 = st.tabs(["🏆 Top รายการ", "📊 ตามหมวดหมู่", "💰 Revenue Map"])

    with tab1:
        if "p2_dept" not in st.session_state: st.session_state["p2_dept"] = "ทั้งหมด"
        if "p2_cat"  not in st.session_state: st.session_state["p2_cat"]  = "ทั้งหมด"

        dept_list = ["ทั้งหมด"] + sorted(sold["DepName"].dropna().astype(str).unique().tolist()) \
                    if "DepName" in sold.columns else ["ทั้งหมด"]

        # ── DEPT: chip row ───────────────────────────────────────
        dept_cols = st.columns(len(dept_list))
        for i, d in enumerate(dept_list):
            active = st.session_state["p2_dept"] == d
            if dept_cols[i].button(d, key=f"dept_{i}", use_container_width=True):
                st.session_state["p2_dept"] = d
                st.session_state["p2_cat"]  = "ทั้งหมด"

        # ── CAT: horizontal scrollable pills via HTML/JS ──────────
        sold_by_dept = sold.copy()
        if st.session_state["p2_dept"] != "ทั้งหมด" and "DepName" in sold.columns:
            sold_by_dept = sold[sold["DepName"].astype(str) == st.session_state["p2_dept"]]

        cat_list = ["ทั้งหมด"] + sorted(sold_by_dept["CatName"].dropna().astype(str).unique().tolist()) \
                   if "CatName" in sold_by_dept.columns else ["ทั้งหมด"]

        if st.session_state["p2_cat"] not in cat_list:
            st.session_state["p2_cat"] = "ทั้งหมด"

        # render cat เป็น scrollable row ผ่าน st.selectbox แบบ horizontal
        # ใช้ columns แบบ wrap ไม่ได้ → ใช้ selectbox compact แทน
        # แต่ให้ดูดีขึ้น: แสดง cat เป็น horizontal scroll HTML + selectbox hidden
        selected_cat_idx = cat_list.index(st.session_state["p2_cat"]) \
                           if st.session_state["p2_cat"] in cat_list else 0

        # pill HTML (display only, interactive via selectbox below)
        pills_html = "".join([
            f'<span style="display:inline-block;margin:3px 4px;padding:5px 14px;'
            f'border-radius:20px;font-size:12px;font-weight:{"700" if c == st.session_state["p2_cat"] else "400"};'
            f'background:{"#d4af37" if c == st.session_state["p2_cat"] else "#21262d"};'
            f'color:{"#0d1117" if c == st.session_state["p2_cat"] else "#c9d1d9"};'
            f'border:1px solid {"#d4af37" if c == st.session_state["p2_cat"] else "#30363d"};'
            f'white-space:nowrap;">{c}</span>'
            for c in cat_list
        ])
        st.markdown(
            f'<div style="overflow-x:auto;white-space:nowrap;padding:6px 0 4px;">{pills_html}</div>',
            unsafe_allow_html=True
        )

        # selectbox จริงสำหรับ interaction (compact, ซ่อน label)
        chosen_cat = st.selectbox(
            "📂 เลือก Category",
            cat_list,
            index=selected_cat_idx,
            key="p2_cat_sel",
            label_visibility="collapsed",
        )
        if chosen_cat != st.session_state["p2_cat"]:
            st.session_state["p2_cat"] = chosen_cat
            st.rerun()

        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

        # ── Apply filters ────────────────────────────────────────
        sold_f = sold.copy()
        if st.session_state["p2_dept"] != "ทั้งหมด" and "DepName" in sold_f.columns:
            sold_f = sold_f[sold_f["DepName"].astype(str) == st.session_state["p2_dept"]]
        if st.session_state["p2_cat"] != "ทั้งหมด" and "CatName" in sold_f.columns:
            sold_f = sold_f[sold_f["CatName"].astype(str) == st.session_state["p2_cat"]]

        # ── Top N + filter summary ────────────────────────────────
        sl_col, lbl_col = st.columns([3, 2])
        top_n = sl_col.slider("แสดง Top N", 5, 500, 10, key="p2_topn")
        filter_label = " › ".join(x for x in [st.session_state["p2_dept"],
                                               st.session_state["p2_cat"]] if x != "ทั้งหมด") or "ทั้งหมด"
        lbl_col.markdown(
            f"<div style='padding-top:28px;font-size:12px;color:#d4af37;'>"
            f"🔍 {filter_label} · {len(sold_f):,} rows</div>",
            unsafe_allow_html=True
        )

        # ── Aggregate ────────────────────────────────────────────
        group_cols = ["ItemCode","ItemDes1"]
        if "DepName" in sold_f.columns: group_cols.append("DepName")
        if "CatName" in sold_f.columns: group_cols.append("CatName")

        top_items = (sold_f.groupby(group_cols)
                    .agg(Qty=("ActQTY","sum"), Revenue=("NetAmt","sum"), Txn=("_BillKey","nunique"))
                    .reset_index().sort_values("Revenue", ascending=False).head(top_n))

        if top_items.empty:
            st.info("ไม่มีข้อมูล")
        else:
            for i, (_, row) in enumerate(top_items.iterrows()):
                rank_cls  = ["top1","top2","top3"][i] if i < 3 else ""
                dep_name  = row.get("DepName",""); cat_name = row.get("CatName","")
                dep_badge = f'<span style="background:#21262d;color:#8b949e;border-radius:8px;padding:1px 8px;font-size:10px;margin-right:4px;">{dep_name}</span>' if dep_name else ""
                cat_badge = f'<span style="background:#1c2230;color:#58a6ff;border-radius:8px;padding:1px 8px;font-size:10px;">{cat_name}</span>' if cat_name else ""
                st.markdown(f"""
                <div class="rank-row">
                    <div class="rank-num {rank_cls}">#{i+1}</div>
                    <div class="rank-name">
                        {row['ItemCode']} :{row['ItemDes1']}
                        <br><span class="rank-sub">{int(row['Qty'])} qty · {int(row['Txn'])} บิล</span>
                        {'<br>' + dep_badge + cat_badge if (dep_name or cat_name) else ''}
                    </div>
                    <div class="rank-amt">฿{row['Revenue']:,.0f}</div>
                </div>""", unsafe_allow_html=True)

            # ── Export Excel ──────────────────────────────────────
            def _build_topitems_excel(df_exp, d_from, d_to) -> bytes:
                import io as _io2, openpyxl as _ox2
                from openpyxl.styles import Font as F2, PatternFill as P2, Alignment as A2, Border as B2, Side as S2
                from openpyxl.utils import get_column_letter as gcl2
                wb2 = _ox2.Workbook(); ws = wb2.active; ws.title = "Top Items"
                ws.sheet_view.showGridLines = False
                ws.merge_cells("A1:H1")
                ws["A1"].value = f"รายการขายดี  |  {d_from.strftime('%d/%m/%Y')} – {d_to.strftime('%d/%m/%Y')}  |  {filter_label}"
                ws["A1"].font = F2(color="D4AF37", size=13, bold=True, name="Arial")
                ws["A1"].fill = P2("solid", fgColor="161B22")
                ws["A1"].alignment = A2(horizontal="left", vertical="center")
                ws.row_dimensions[1].height = 30
                for c, (h, w) in enumerate(zip(
                    ["#","ItemCode","ชื่อเมนู","Department","Category","Qty","บิล","Revenue (฿)"],
                    [5,14,42,22,22,10,10,16]), 1):
                    cell = ws.cell(row=2, column=c, value=h)
                    cell.font = F2(color="D4AF37", size=9, bold=True, name="Arial")
                    cell.fill = P2("solid", fgColor="1C2230")
                    cell.alignment = A2(horizontal="left" if c==3 else "center", vertical="center")
                    cell.border = B2(bottom=S2(style="medium", color="D4AF37"))
                    ws.column_dimensions[gcl2(c)].width = w
                ws.row_dimensions[2].height = 20
                for idx, r in df_exp.iterrows():
                    rn = idx+3; bg = "0D1117" if idx%2==0 else "161B22"
                    for c, (val, fmt) in enumerate(zip(
                        [idx+1, r.get("ItemCode",""), r.get("ItemDes1",""),
                         r.get("DepName",""), r.get("CatName",""),
                         int(r.get("Qty",0)), int(r.get("Txn",0)), r.get("Revenue",0)],
                        [None,None,None,None,None,"#,##0","#,##0","#,##0.00"]), 1):
                        cell = ws.cell(row=rn, column=c, value=val)
                        cell.fill = P2("solid", fgColor=bg)
                        cell.font = F2(color="D4AF37" if c==8 else "58A6FF" if c in(4,5) else "E6EDF3", size=9, name="Arial")
                        cell.alignment = A2(horizontal="left" if c==3 else "center", vertical="center")
                        cell.border = B2(bottom=S2(style="hair", color="30363D"))
                        if fmt: cell.number_format = fmt
                    ws.row_dimensions[rn].height = 18
                ws.freeze_panes = "A3"
                buf2 = _io2.BytesIO(); wb2.save(buf2); buf2.seek(0)
                return buf2.read()

            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Export Excel",
                data=_build_topitems_excel(top_items.reset_index(drop=True), d_from, d_to),
                file_name=f"top_items_{filter_label.replace(' › ','_')}_{d_from.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with tab2:
        if "DepName" in sold.columns:
            st.markdown('<div class="section-title">🏷️ Department</div>', unsafe_allow_html=True)
            dep_rev = sold.groupby("DepName")["NetAmt"].sum().reset_index().sort_values("NetAmt", ascending=False)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(dep_rev, names="DepName", values="NetAmt",
                             color_discrete_sequence=COLOR_SEQ, hole=0.4, title="สัดส่วนตาม Department")
                fig.update_layout(**LAYOUT_BASE, height=350)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(dep_rev, x="NetAmt", y="DepName", orientation="h",
                             color_discrete_sequence=[GOLD], title="รายรับตาม Department")
                fig.update_layout(**LAYOUT_BASE, height=350)
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(fig, use_container_width=True)

        if "CatName" in sold.columns:
            st.markdown('<div class="section-title">📂 Category</div>', unsafe_allow_html=True)
            cat_rev = sold.groupby("CatName")["NetAmt"].sum().reset_index().sort_values("NetAmt", ascending=False)
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(cat_rev, names="CatName", values="NetAmt",
                             color_discrete_sequence=COLOR_SEQ, hole=0.4, title="สัดส่วนตาม Category")
                fig.update_layout(**LAYOUT_BASE, height=350)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(cat_rev, x="NetAmt", y="CatName", orientation="h",
                             color_discrete_sequence=[BLUE], title="รายรับตาม Category")
                fig.update_layout(**LAYOUT_BASE, height=350)
                fig.update_yaxes(autorange="reversed")
                st.plotly_chart(fig, use_container_width=True)

        if "DepName" in sold.columns and "CatName" in sold.columns:
            st.markdown('<div class="section-title">📊 Dept × Category Heatmap</div>', unsafe_allow_html=True)
            cross = (sold.groupby(["DepName","CatName"])["NetAmt"].sum()
                     .reset_index().pivot(index="DepName", columns="CatName", values="NetAmt").fillna(0))
            fig = px.imshow(cross, color_continuous_scale=["#1c2230", GOLD],
                            aspect="auto", text_auto=",.0f", title="Revenue Heatmap")
            fig.update_layout(**LAYOUT_BASE, height=max(300, len(cross)*40))
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if "DepName" in sold.columns and "ItemDes1" in sold.columns:
            path = ["DepName","CatName","ItemDes1"] if "CatName" in sold.columns else ["DepName","ItemDes1"]
            tree = sold.groupby(path)["NetAmt"].sum().reset_index()
            fig  = px.treemap(tree, path=path, values="NetAmt",
                              color="NetAmt", color_continuous_scale=["#1c2230","#d4af37"],
                              title="Revenue Treemap")
            fig.update_layout(**LAYOUT_BASE, height=500)
            st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PAGE 3: วิเคราะห์ช่วงเวลา (+ Top Items by Time Slot)
# ═══════════════════════════════════════════════════════════════
elif page == "⏱️ วิเคราะห์ช่วงเวลา" and data_ok:
    st.markdown('<div class="section-title" style="font-size:24px;">⏱️ วิเคราะห์ช่วงเวลา</div>', unsafe_allow_html=True)

    def slot_label(h):
        h = int(h)
        if h < 11:   return "🌅 เช้า"
        elif h < 14: return "☀️ กลางวัน"
        elif h < 17: return "🌤️ บ่าย"
        else:        return "🌆 เย็น/ค่ำ"

    SLOT_RANGES = {
        "🌅 เช้า":       (7,  10),
        "☀️ กลางวัน":   (11, 13),
        "🌤️ บ่าย":      (14, 16),
        "🌆 เย็น/ค่ำ":  (17, 23),
    }

    # ใช้ _NetBAmt (หัก VAT ร้าน VatType=0) + เฉพาะบิลที่นับได้ ให้ตรงกับหน้า Overview
    # (เดิมใช้ NetBAmt ดิบ → ยอดหน้านี้สูงกว่า Overview สำหรับร้านที่ขายรวม VAT)
    hdr_cnt = hdr[hdr["_CountBill"]]

    if "BillHour" in hdr.columns:
        hourly = (hdr_cnt
                  .groupby("BillHour")
                  .agg(Revenue=("_NetBAmt","sum"), Bills=("_BillKey","count"), Pax=("_NoPerson","sum"))
                  .reset_index())

        peak_hour = hourly.loc[hourly["Revenue"].idxmax(), "BillHour"] if len(hourly) else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("⚡ Peak Hour",   f"{int(peak_hour):02d}:00")
        c2.metric("📊 Max Revenue", f"฿{hourly['Revenue'].max():,.0f}" if len(hourly) else "–")
        c3.metric("🧾 Max Bills",   f"{hourly['Bills'].max():,}"        if len(hourly) else "–")

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=hourly["BillHour"], y=hourly["Revenue"],
                             name="Revenue", marker_color=GOLD, opacity=0.85), secondary_y=False)
        fig.add_trace(go.Scatter(x=hourly["BillHour"], y=hourly["Bills"],
                                 name="Bills", line=dict(color=BLUE, width=2),
                                 mode="lines+markers"), secondary_y=True)
        fig.update_layout(**LAYOUT_BASE, height=350, title="Revenue & Bills per Hour")
        st.plotly_chart(fig, use_container_width=True)

        # ── Time Slot Bar ──────────────────────────────────────────
        hdr2 = hdr_cnt.copy()
        hdr2["Slot"] = hdr2["BillHour"].apply(lambda x: slot_label(x) if pd.notna(x) else "–")
        slot_rev = (hdr2.groupby("Slot")["_NetBAmt"].sum().rename("NetBAmt")
                    .reset_index().sort_values("NetBAmt", ascending=False))
        fig2 = px.bar(slot_rev, x="Slot", y="NetBAmt",
                      color_discrete_sequence=[GOLD], title="ยอดขายตาม Time Slot")
        fig2.update_layout(**LAYOUT_BASE, height=280)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Daily Trend ────────────────────────────────────────────────
    if (d_to - d_from).days > 0 and "BillDate" in hdr.columns:
        daily = (hdr_cnt.groupby(hdr_cnt["BillDate"].dt.date)
                 .agg(Revenue=("_NetBAmt","sum"), Bills=("_BillKey","count"))
                 .reset_index())
        daily.columns = ["Date","Revenue","Bills"]
        fig3 = px.line(daily, x="Date", y="Revenue",
                       markers=True, color_discrete_sequence=[GOLD],
                       title="Daily Revenue Trend")
        fig3.update_layout(**LAYOUT_BASE, height=280)
        st.plotly_chart(fig3, use_container_width=True)

    # ══════════════════════════════════════════════════════════════
    # 🆕 TOP ITEMS ตามช่วงเวลา
    # ══════════════════════════════════════════════════════════════
    st.divider()
    st.markdown('<div class="section-title">🍽️ สินค้าขายดีตามช่วงเวลา</div>', unsafe_allow_html=True)

    # ตรวจว่า dtl มีคอลัมน์ที่ต้องการ
    has_item_data = (
        "BillHour" in hdr.columns
        and "ItemDes1" in dtl.columns
        and "ActQTY"  in dtl.columns
    )

    if not has_item_data:
        st.info("ไม่พบข้อมูล item-level (ต้องการ dtl.ItemDes1, dtl.ActQTY, hdr.BillHour)")
    else:
        # join dtl ← BillHour จาก hdr
        valid_bills = hdr.loc[hdr["_CountBill"], ["BillNo","ShopCode","BillHour"]].drop_duplicates()
        dtl_h = dtl.merge(valid_bills, on=["BillNo","ShopCode"], how="inner") \
                if "ShopCode" in dtl.columns \
                else dtl.merge(valid_bills[["BillNo","BillHour"]], on="BillNo", how="inner")

        # ── FILTER: ตัด parent set rows / free items ที่ไม่มีราคา ────
        dtl_h = dtl_h[
    ~       (
            (pd.to_numeric(dtl_h["SellPrice"], errors="coerce").fillna(0) == 0)       
            )
        ].copy()

        # ── คำนวณ Revenue ต่อ row = OrgPrice × ActQTY ────────────────
        dtl_h["_RowRevenue"] = (
            pd.to_numeric(dtl_h["SellPrice"], errors="coerce").fillna(0) *
            pd.to_numeric(dtl_h["ActQTY"],   errors="coerce").fillna(0)
            )
        rev_col = "_RowRevenue"

        # ── SELECTOR ──────────────────────────────────────────────
        sel_col1, sel_col2 = st.columns([2, 3])

        with sel_col1:
            mode = st.radio("เลือกช่วงเวลาแบบ", ["📌 Preset Slot", "🕐 กำหนดเอง (Hour Range)"],
                            horizontal=True)

        with sel_col2:
            if mode == "📌 Preset Slot":
                chosen_slot = st.select_slider(
                    "Slot",
                    options=list(SLOT_RANGES.keys()),
                    value="☀️ กลางวัน",
                )
                h_start, h_end = SLOT_RANGES[chosen_slot]
                slot_label_display = chosen_slot
            else:
                h_start, h_end = st.slider(
                    "ช่วงชั่วโมง",
                    min_value=0, max_value=23,
                    value=(11, 14),
                    format="%02d:00"
                )
                slot_label_display = f"{h_start:02d}:00 – {h_end:02d}:00"

        top_n = st.slider("แสดง Top N items", min_value=5, max_value=30, value=15, step=5)

        # ── FILTER BY HOUR ─────────────────────────────────────────
        mask = dtl_h["BillHour"].between(h_start, h_end)
        dtl_slot = dtl_h[mask]

        if len(dtl_slot) == 0:
            st.warning(f"ไม่มีข้อมูลในช่วง {slot_label_display}")
        else:
            agg_dict = {"Qty": ("ActQTY", "sum")}
            if rev_col:
                agg_dict["Revenue"] = (rev_col, "sum")

            item_slot = (dtl_slot.groupby("ItemDes1")
                         .agg(**agg_dict)
                         .reset_index())

            # ── TABS: Qty / Revenue / Table / Heatmap ─────────────
            tab_labels = ["📦 Top by Qty", "📋 Table"]
            if rev_col:
                tab_labels.insert(1, "💰 Top by Revenue")
            tab_labels.append("🔥 Heatmap item × hour")

            tabs = st.tabs(tab_labels)
            tab_idx = 0

            # ── Tab: Qty ──────────────────────────────────────────
            with tabs[tab_idx]:
                tab_idx += 1
                top_qty = item_slot.sort_values("Qty", ascending=False).head(top_n)
                fig_qty = px.bar(
                    top_qty,
                    x="Qty", y="ItemDes1", orientation="h",
                    text=top_qty["Qty"].map("{:,.0f}".format),
                    color_discrete_sequence=[GOLD],
                    title=f"Top {top_n} Items by Qty — {slot_label_display}",
                )
                fig_qty.update_layout(**LAYOUT_BASE, height=max(350, top_n * 28))
                fig_qty.update_yaxes(autorange="reversed")
                st.plotly_chart(fig_qty, use_container_width=True)

            # ── Tab: Revenue ──────────────────────────────────────
            if rev_col:
                with tabs[tab_idx]:
                    tab_idx += 1
                    top_rev = item_slot.sort_values("Revenue", ascending=False).head(top_n)
                    fig_rev = px.bar(
                        top_rev,
                        x="Revenue", y="ItemDes1", orientation="h",
                        text=top_rev["Revenue"].map("฿{:,.0f}".format),
                        color_discrete_sequence=[BLUE],
                        title=f"Top {top_n} Items by Revenue — {slot_label_display}",
                    )
                    fig_rev.update_layout(**LAYOUT_BASE, height=max(350, top_n * 28))
                    fig_rev.update_yaxes(autorange="reversed")
                    st.plotly_chart(fig_rev, use_container_width=True)

                # ── Tab: Table ────────────────────────────────────────
                with tabs[tab_idx]:
                    tab_idx += 1
                    tbl = item_slot.sort_values("Qty", ascending=False).copy()
                    if rev_col:
                        tbl["Revenue"] = tbl["Revenue"].map("฿{:,.0f}".format)
                    tbl["Qty"] = tbl["Qty"].map("{:,.0f}".format)
                    tbl.columns = (["เมนู", "Qty (ชิ้น)", "Revenue"]
                                    if rev_col else ["เมนู", "Qty (ชิ้น)"])
                    st.dataframe(tbl, use_container_width=True, hide_index=True,
                                 height=min(600, len(tbl) * 36 + 40))

                # ── Tab: Heatmap item × hour ───────────────────────────
                with tabs[tab_idx]:
                    # ใช้ทุก hour ในวัน (ไม่กรอง slot) เพื่อให้เห็น pattern ตลอดวัน
                    # แต่ highlight เฉพาะ item ที่อยู่ใน top_n ของ slot ที่เลือก
                    top_items = item_slot.sort_values("Qty", ascending=False).head(top_n)["ItemDes1"].tolist()
                    dtl_top   = dtl_h[dtl_h["ItemDes1"].isin(top_items)]

                    pivot = (dtl_top.groupby(["ItemDes1","BillHour"])["ActQTY"]
                             .sum()
                             .reset_index()
                             .pivot(index="ItemDes1", columns="BillHour", values="ActQTY")
                             .fillna(0))

                    # เรียง item ตาม Qty ใน slot ที่เลือก
                    pivot = pivot.reindex(top_items)

                    fig_hm = px.imshow(
                        pivot,
                        color_continuous_scale=[[0,"#1a1a2e"], [0.3, GOLD], [1,"#c0392b"]],
                        aspect="auto",
                        labels=dict(x="Hour", y="Item", color="Qty"),
                        title=f"Heatmap: Qty by Item × Hour (top {top_n} items ใน {slot_label_display})",
                    )
                    # เพิ่ม highlight zone ของ slot ที่เลือก
                    fig_hm.add_vrect(
                        x0=h_start - 0.5, x1=h_end + 0.5,
                        fillcolor=GOLD, opacity=0.08,
                        line_width=1.5, line_color=GOLD,
                        annotation_text=slot_label_display,
                        annotation_position="top left",
                        annotation_font_color=GOLD,
                    )
                    fig_hm.update_layout(**LAYOUT_BASE, height=max(400, top_n * 26),
                                         coloraxis_showscale=True)
                    st.plotly_chart(fig_hm, use_container_width=True)

                    st.caption("🟨 กรอบสีทอง = ช่วงเวลาที่เลือก | สีแดงเข้ม = ขายสูงสุด")

# ═══════════════════════════════════════════════════════════════
# PAGE 4: MEMBER / GUEST (FINAL - REMARK BASED)
# ═══════════════════════════════════════════════════════════════
elif page == "👥 Member / Guest" and data_ok:
    # ─────────────────────────────────────────────
    # PREP DATA
    # ─────────────────────────────────────────────
    valid_bills = hdr.loc[hdr["_CountBill"], ["BillNo","ShopCode"]].drop_duplicates()

    hdr_f = hdr[hdr["_CountBill"]].copy()

    dtl_f = dtl.merge(valid_bills, on=["BillNo","ShopCode"], how="inner") \
        if "ShopCode" in dtl.columns else dtl[dtl["BillNo"].isin(valid_bills["BillNo"])]

    # IsMember / MemberCode / MemberTypeName มาจาก clean_data() — นิยามเดียวกับหน้า Overview
    # (เดิมหน้านี้คิดเองจาก Remark ส่วน Overview ใช้ BILLPAYTB — ได้ชุดบิลเดียวกันแต่ tier ไม่ออก)
    member_col = "MemberCode"

    st.markdown('<div class="section-title" style="font-size:24px;">👥 Member / Guest Analysis</div>', unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # KPI
    # ─────────────────────────────────────────────
    mem_hdr = hdr_f[hdr_f["IsMember"]]
    gst_hdr = hdr_f[~hdr_f["IsMember"]]

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("🐾 Member Bills",   f"{mem_hdr['_BillKey'].nunique():,}")
    c2.metric("👤 Guest Bills",    f"{gst_hdr['_BillKey'].nunique():,}")
    c3.metric("💰 Member Revenue", f"฿{mem_hdr['_NetBAmt'].sum():,.0f}")
    c4.metric("💰 Guest Revenue",  f"฿{gst_hdr['_NetBAmt'].sum():,.0f}")

    # ─────────────────────────────────────────────
    # MEMBER TYPE
    # ─────────────────────────────────────────────
    if "MemberTypeName" in hdr_f.columns:

        st.markdown('<div class="section-title">📊 ตามประเภท Member</div>', unsafe_allow_html=True)

        type_stats = (hdr_f.groupby("MemberTypeName")
            .agg(Bills=("_BillKey","nunique"),
                 Revenue=("_NetBAmt","sum"),
                 Disc=("_TotalDisc","sum"))
            .reset_index())

        type_stats["AvgBill"] = type_stats["Revenue"] / type_stats["Bills"].replace(0,1)

        col1, col2 = st.columns(2)

        with col1:
            fig = px.bar(type_stats, x="MemberTypeName", y="Revenue",
                         color_discrete_sequence=[GOLD])
            fig.update_layout(**LAYOUT_BASE, height=320)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig = px.bar(type_stats, x="MemberTypeName", y="AvgBill",
                         color_discrete_sequence=[BLUE])
            fig.update_layout(**LAYOUT_BASE, height=320)
            st.plotly_chart(fig, use_container_width=True)

    # ─────────────────────────────────────────────
    # 🧠 CUSTOMER BEHAVIOR (เฉพาะ MEMBER)
    # ─────────────────────────────────────────────
    st.markdown('<div class="section-title">🧠 Customer Behavior (Member)</div>', unsafe_allow_html=True)

    mem_only = hdr_f[hdr_f["IsMember"]].copy()

    if len(mem_only) == 0:
        st.warning("ไม่มีข้อมูล Member")
    else:

        st.markdown("### 🔁 Member Visit Frequency")
        visits = mem_only.groupby(member_col)["_BillKey"].nunique().reset_index(name="Visits")
        visits["Segment"] = pd.cut(
            visits["Visits"],
            bins=[0, 1, 3, float("inf")],   # เดิม 100 → member ที่มาเกิน 100 บิลหลุดกราฟ
            labels=["visit 1 ครั้ง","visit 2–3 ครั้ง","visit 4+ ครั้ง"]
        )

        seg = visits["Segment"].value_counts().reset_index()
        seg.columns = ["Segment","Count"]    
        fig = px.bar(
            seg.sort_values("Segment"),
            x="Segment", y="Count",
            text="Count",
            color="Segment",
            color_discrete_sequence=["#6ea8fe","#fbbc04","#fb9804"]
        )

        fig.update_layout(
            height=400,
            xaxis_title="กลุ่มการมาใช้บริการ",
            yaxis_title="Num Member",
            showlegend=False
        )

        st.plotly_chart(fig, use_container_width=True)

        # 2️⃣ Member Visit Time
        if "BillTime" in mem_only.columns:
            st.markdown("### ⏰ Member Visit Time")
            mem_only["Hour"] = pd.to_datetime(mem_only["BillTime"], errors="coerce").dt.hour

            visit_time = mem_only.groupby("Hour")["BillNo"].count().reset_index()

            fig = px.bar(visit_time, x="Hour", y="BillNo")
            st.plotly_chart(fig, use_container_width=True)

        # 3️⃣ Member Visit Day
        if "BillDate" in mem_only.columns:
            st.markdown("### 📅 Member Visit Day")
            mem_only["Day"] = pd.to_datetime(mem_only["BillDate"], errors="coerce").dt.day_name()

            visit_day = mem_only.groupby("Day")["BillNo"].count().reset_index()

            fig = px.bar(visit_day, x="Day", y="BillNo")
            fig.update_xaxes(categoryorder="array", categoryarray=[   # เดิมเรียงตามตัวอักษร (Friday ก่อน Monday)
                "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"])
            st.plotly_chart(fig, use_container_width=True)

        # 4️⃣ Member Spending Behavior
        st.markdown("### 💰 Member Spending Behavior")
        spend = mem_only.groupby(member_col)["_NetBAmt"].sum().reset_index()

        fig = px.histogram(spend, x="_NetBAmt", nbins=30)
        st.plotly_chart(fig, use_container_width=True)

        # 5️⃣ Member Top / Low Spenders
        st.markdown("### 🏆 Member Top / Low Spenders")

        spend = mem_only.groupby(member_col)["_NetBAmt"].sum().reset_index()

# ── Join ชื่อ + MemType ──────────────────────────────────────
        membertb_clean = membertb[["MemCode","MemType","FirstName1","LastName1"]].copy()
        membertb_clean["MemCode"] = membertb_clean["MemCode"].astype(str).str.strip()
        membertb_clean["FullName"] = (
        membertb_clean["FirstName1"].fillna("") + " " +
        membertb_clean["LastName1"].fillna("")
        ).str.strip()

        memtype_clean = memtype[["MemType","MemTNam1"]].drop_duplicates("MemType").copy()
        membertb_clean = membertb_clean.merge(memtype_clean, on="MemType", how="left")
        membertb_clean["MemTNam1"] = membertb_clean["MemTNam1"].fillna(membertb_clean["MemType"])

        spend["MemberCode"] = spend["MemberCode"].astype(str).str.strip()
        spend = spend.merge(membertb_clean[["MemCode","FullName","MemTNam1"]],
                    left_on="MemberCode", right_on="MemCode", how="left")
        spend["FullName"]  = spend["FullName"].fillna("–")
        spend["MemTNam1"]  = spend["MemTNam1"].fillna("–")

        top = spend.sort_values("_NetBAmt", ascending=False).head(10)
        low = spend.sort_values("_NetBAmt", ascending=True).head(10)

        BADGE_COLOR = {
            "White Hound":    ("#c9d1d9","#9e9e9e"),
            "Silver Hound":   ("#e5e4e2","#9e9e9e"),
            "Black Hound":    ("#e6edf3","#21262d"),
            "Platinum Hound": ("#e5e4e2","#555"),
            "Senior Card":    ("#58a6ff","#1c3a5e"),
            "GH Wallet":      ("#d4af37","#3d2200"),
        }

        def render_spender_row(i, row, is_top=True):
            rank_cls = ["top1","top2","top3"][i] if i < 3 and is_top else ""
            rank_color = {"top1":GOLD,"top2":SILVER,"top3":"#cd7f32"}.get(rank_cls,"#8b949e")
            bc, bg = BADGE_COLOR.get(row["MemTNam1"], ("#d4af37","#3d2200"))
            amt_color = GOLD if is_top else RED
            st.markdown(f"""
        <div class="rank-row">
            <div class="rank-num" style="color:{rank_color};">#{i+1}</div>
            <div style="flex:1;">
                <div style="font-weight:600;color:#e6edf3;font-size:14px;">{row['FullName']}</div>
                <div style="font-size:11px;color:#8b949e;margin-top:2px;">
                    {row['MemberCode']} &nbsp;
                    <span style="background:{bg};color:{bc};border-radius:10px;
                          padding:1px 8px;font-size:10px;font-weight:600;">
                        {row['MemTNam1']}
                    </span>
                </div>
            </div>
            <div style="font-weight:700;color:{amt_color};font-size:15px;">฿{row['_NetBAmt']:,.0f}</div>
        </div>""", unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            st.markdown(f"<div style='font-size:13px;font-weight:700;color:{GOLD};margin-bottom:8px;'>🔥 Top Spenders</div>", unsafe_allow_html=True)
            for i, (_, row) in enumerate(top.iterrows()):
                render_spender_row(i, row, is_top=True)

        with col2:
            st.markdown(f"<div style='font-size:13px;font-weight:700;color:{RED};margin-bottom:8px;'>⚠️ Low Spenders</div>", unsafe_allow_html=True)
            for i, (_, row) in enumerate(low.iterrows()):
                render_spender_row(i, row, is_top=False)

    # 6️⃣ Favorite Item
    st.markdown("### 🍽️ Favorite Item")
    if "ItemDes1" in dtl_f.columns and "ActQTY" in dtl_f.columns:
            # merge ด้วย _BillKey (ShopCode|BillNo) — เดิมใช้ BillNo อย่างเดียว ทำให้ดึงรายการ
            # ของบิลสาขาอื่นที่บังเอิญเลขบิลตรงกันมาปนเป็นของ member
            dtl_mem = dtl_f.merge(
            mem_only[["_BillKey", member_col]], on="_BillKey", how="inner"
            )
            if "SellPrice" in dtl_mem.columns:
                dtl_mem = dtl_mem[dtl_mem["SellPrice"] > 0]

            if dtl_mem.empty:
                st.info("ไม่มีข้อมูล item สำหรับ member")
            else:
                tab_fav1, tab_fav2, tab_fav3 = st.tabs([
                    "👥 ความนิยม (unique members)",
                    "🔁 สั่งซ้ำ (avg per visit)",
                    "🧺 Basket Analysis"
                ])

                with tab_fav1:
                    st.caption("นับว่ามีกี่คนที่เคยสั่ง item นี้อย่างน้อย 1 ครั้ง")
                    fav_members = (
                        dtl_mem.groupby("ItemDes1")[member_col]
                        .nunique()
                        .reset_index(name="UniqueMem")
                        .sort_values("UniqueMem", ascending=False)
                        .head(20)
                    )
                    fig = px.bar(
                        fav_members, x="UniqueMem", y="ItemDes1", orientation="h",
                        color_discrete_sequence=[GOLD],
                        title="Top 20 — จำนวน Member ที่เคยสั่ง"
                    )
                    fig.update_layout(**LAYOUT_BASE, height=520,
                              xaxis_title="Unique Members", yaxis_title="")
                    st.plotly_chart(fig, use_container_width=True)

                with tab_fav2:
                    st.caption("avg จำนวนชิ้นต่อบิลที่มีการสั่ง item นั้น")
                    fav_repeat = (
                        dtl_mem.groupby(["_BillKey","ItemDes1"])["ActQTY"]
                        .sum().reset_index()
                        .groupby("ItemDes1")["ActQTY"]
                        .mean().reset_index(name="AvgQtyPerVisit")
                        .sort_values("AvgQtyPerVisit", ascending=False)
                        .head(20)
                    )
                    fig = px.bar(
                    fav_repeat, x="AvgQtyPerVisit", y="ItemDes1", orientation="h",
                    color_discrete_sequence=[BLUE],
                    title="Top 20 — Avg Qty / Visit"
                    )
                    fig.update_layout(**LAYOUT_BASE, height=520,
                              xaxis_title="Avg Qty per Bill", yaxis_title="")
                    st.plotly_chart(fig, use_container_width=True)

                with tab_fav3:
                    st.caption("item ที่มักถูกสั่งพร้อมกันในบิลเดียว")
                    top15 = fav_members["ItemDes1"].head(15).tolist()
                    sel_item = st.selectbox("เลือก Item หลัก", top15, key="basket_item")

                    bills_with = dtl_mem.loc[
                        dtl_mem["ItemDes1"] == sel_item, "_BillKey"
                    ].unique()

                    paired = (
                        dtl_mem[
                        dtl_mem["_BillKey"].isin(bills_with) &
                        (dtl_mem["ItemDes1"] != sel_item)
                        ]
                        .groupby("ItemDes1")["_BillKey"].nunique()
                        .reset_index(name="CoOccurBills")
                        .sort_values("CoOccurBills", ascending=False)
                        .head(15)
                    )
                    paired["PairRate"] = paired["CoOccurBills"] / len(bills_with) * 100

                    if paired.empty:
                        st.info("ไม่พบ pair")
                    else:
                        fig = px.bar(
                            paired, x="PairRate", y="ItemDes1", orientation="h",
                            color="PairRate",
                            color_continuous_scale=["#1c2230", GOLD],
                            title=f"Items ที่มักสั่งคู่กับ '{sel_item}'"
                        )
                        fig.update_layout(**LAYOUT_BASE, height=420,
                                  xaxis_title="Co-occur Rate (%)", yaxis_title="",
                                  coloraxis_showscale=False)
                        st.plotly_chart(fig, use_container_width=True)

                        st.dataframe(
                            paired.rename(columns={
                                "ItemDes1": "Item",
                                "CoOccurBills": "บิลที่สั่งคู่กัน",
                                "PairRate": "% Co-occur"
                            }).reset_index(drop=True),
                            use_container_width=True, height=240
                        )
    else:
        st.info("ไม่มีข้อมูล ItemDes1 หรือ ActQTY")
        
# ═══════════════════════════════════════════════════════════════
# PAGE 5: PER BILL ANALYSIS
# ═══════════════════════════════════════════════════════════════
elif page == "🧾 Per Bill Analysis" and data_ok:
    st.markdown('<div class="section-title" style="font-size:24px;">🧾 Per Bill Analysis</div>', unsafe_allow_html=True)

    # เฉพาะบิลที่นับได้ (ตัดคู่บิล void status 4/5 ออก) ให้ตรงกับ Avg/Bill ของ Overview
    hb = hdr[hdr["_CountBill"]]
    avg_bill  = hb["_NetBAmt"].mean()
    avg_pax   = hb["_NoPerson"].mean()
    avg_spend = (hb["_NetBAmt"] / hb["_NoPerson"].replace(0,1)).mean()

    # จัดกลุ่มด้วย _BillKey — เดิม groupby BillNo รวมรายการของหลายสาขาที่เลขบิลตรงกันเป็นบิลเดียว
    dtl_cnt = dtl[dtl["_BillKey"].isin(hb["_BillKey"]) & (dtl["NetAmt"] > 0)]
    items_per_bill = dtl_cnt.groupby("_BillKey")["ActQTY"].sum()
    avg_items = items_per_bill.mean()

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📊 Avg Bill",         f"฿{avg_bill:,.0f}")
    c2.metric("🍽️ Avg Items/Bill",   f"{avg_items:.1f}")
    c3.metric("👥 Avg Pax/Bill",     f"{avg_pax:.1f}")
    c4.metric("💰 Avg Spend/Pax",    f"฿{avg_spend:,.0f}")

    st.markdown('<div class="section-title">📊 การกระจาย Bill Value</div>', unsafe_allow_html=True)
    fig = px.histogram(hb[hb["_NetBAmt"]>0], x="_NetBAmt", nbins=40,
                       color_discrete_sequence=[GOLD], title="Bill Value Distribution")
    fig.add_vline(x=avg_bill, line_color=BLUE, line_dash="dash",
                  annotation_text=f"avg ฿{avg_bill:,.0f}", annotation_font_color=BLUE)
    fig.update_layout(**LAYOUT_BASE, height=320)
    st.plotly_chart(fig, use_container_width=True)

    # payment breakdown
    st.markdown('<div class="section-title">💳 วิธีชำระเงิน</div>', unsafe_allow_html=True)
    pay_real = pay[pay["TendAmt"] > 0].copy()
    type_map = {"MEM":"Member","CRD":"Credit Card","CSH":"Cash","CHR":"Service Charge"}
    pay_real["PayType"] = pay_real["tendtype"].map(type_map).fillna(pay_real["tendtype"])
    pay_sum = (pay_real.groupby("PayType")["TendAmt"].sum()
               .reset_index().sort_values("TendAmt", ascending=False))
    c_l, c_r = st.columns(2)
    with c_l:
        fig = px.pie(pay_sum, names="PayType", values="TendAmt",
                     color_discrete_sequence=COLOR_SEQ, hole=0.4, title="Payment Mix")
        fig.update_layout(**LAYOUT_BASE, height=300)
        st.plotly_chart(fig, use_container_width=True)
    with c_r:
        fig = px.bar(pay_sum, x="TendAmt", y="PayType", orientation="h",
                     color_discrete_sequence=[GOLD], title="Payment Amount")
        fig.update_layout(**LAYOUT_BASE, height=300)
        st.plotly_chart(fig, use_container_width=True)
     
# ═══════════════════════════════════════════════════════════════
# PAGE 6: วิเคราะห์โปรโมชั่น (FINAL PRODUCTION VERSION)
# ═══════════════════════════════════════════════════════════════
elif page == "🎁 วิเคราะห์โปรโมชั่น" and data_ok:
    st.markdown('<div class="section-title" style="font-size:24px;">🎁 วิเคราะห์โปรโมชั่น</div>', unsafe_allow_html=True)

    if len(disc) == 0:
        st.info("ไม่มีข้อมูลโปรโมชั่นในช่วงนี้")
    else:
        # ─────────────────────────────────────────────
        # FILTER VALID BILL
        # ─────────────────────────────────────────────
        valid_bills = hdr.loc[hdr["_CountBill"], ["BillNo","ShopCode"]].drop_duplicates()

        hdr_f = hdr[hdr["_CountBill"]].copy()

        disc_f = disc.merge(valid_bills, on=["BillNo","ShopCode"], how="inner") if "ShopCode" in disc.columns \
                 else disc[disc["BillNo"].isin(valid_bills["BillNo"])]

        # ─────────────────────────────────────────────
        # 🎯 TOTAL DISCOUNT (MATCH OVERVIEW)
        # ─────────────────────────────────────────────
        total_disc = hdr_f["_TotalDisc"].sum()

        total_bills = hdr_f["_BillKey"].nunique()
        disc_bills  = disc_f["_BillKey"].nunique()
        penetration = (disc_bills / total_bills * 100) if total_bills else 0
        avg_disc_bill = total_disc / disc_bills if disc_bills else 0

        # ─────────────────────────────────────────────
        # 🧠 DISCOUNT STRUCTURE (VF / VM / VS)
        # ─────────────────────────────────────────────
        vf_disc = pd.to_numeric(hdr_f.get("ItemDisc", 0), errors="coerce").fillna(0).sum()
        vm_disc = pd.to_numeric(hdr_f.get("MemDisc", 0), errors="coerce").fillna(0).sum()
        vs_disc = pd.to_numeric(hdr_f.get("BillDisc", 0), errors="coerce").fillna(0).sum()
        coupon_disc = pd.to_numeric(hdr_f.get("CoupDisc", 0), errors="coerce").fillna(0).sum()

        # ─────────────────────────────────────────────
        # 🎯 KPI
        # ─────────────────────────────────────────────
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("💰 Total Discount", f"฿{total_disc:,.0f}")
        c2.metric("🧾 Bills w/ Promo", f"{disc_bills:,}")
        c3.metric("📊 Penetration", f"{penetration:.1f}%")
        c4.metric("📉 Avg Disc/Bill", f"฿{avg_disc_bill:,.0f}")

        # ─────────────────────────────────────────────
        # 🧠 STRUCTURE
        # ─────────────────────────────────────────────
        st.markdown('<div class="section-title">🧠 Discount Structure</div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            st.metric("🏷️ Item Discount", f"฿{vf_disc:,.0f}")

        with col2:
            st.metric("🐾 Item Discount by Member", f"฿{vm_disc:,.0f}")
            st.metric("🎫 Item Discount By Bill", f"฿{vs_disc:,.0f}")
            st.metric("🎟️ Item Discount by Coupon", f"฿{coupon_disc:,.0f}")

        # ─────────────────────────────────────────────
        # 📊 PROMOTION PERFORMANCE
        # ─────────────────────────────────────────────
        promo_sum = (disc_f.groupby("DiscName")
            .agg(DiscAmt=("DiscAmt","sum"), Bills=("_BillKey","nunique"))
            .reset_index()
            .sort_values("DiscAmt", ascending=False))

        promo_sum["AvgDiscPerBill"] = promo_sum["DiscAmt"] / promo_sum["Bills"]

        st.markdown('<div class="section-title">🔥 Top Promotion</div>', unsafe_allow_html=True)

        fig = px.bar(
            promo_sum.head(15),
            x="DiscAmt",
            y="DiscName",
            orientation="h",
            color_discrete_sequence=[GOLD]
        )
        fig.update_layout(**LAYOUT_BASE, height=400)
        st.plotly_chart(fig, use_container_width=True)

        # ─────────────────────────────────────────────
        # 🔥 PROMO ROI (NO NetAmt, ใช้ HDR)
        # ─────────────────────────────────────────────
        st.markdown('<div class="section-title">🔥 Promotion Return on Investment (ROI)</div>', unsafe_allow_html=True)

        # ใช้ revenue จาก hdr (ชัวร์สุด)
        # 1) ยุบส่วนลดเป็น 1 แถวต่อ (โปร, บิล) ก่อน — บิลเดียวมีหลายรายการใช้โปรเดียวกันได้
        #    เดิม merge ทุกแถวส่วนลด → ยอดบิลถูกบวกซ้ำตามจำนวนรายการ
        # 2) join ยอดบิลด้วย _BillKey — เดิม groupby/merge ด้วย BillNo รวมยอดบิลหลายสาขาที่เลขตรงกัน
        bill_rev = hdr_f[["_BillKey", "_NetBAmt"]]
        promo_rev = (disc_f.groupby(["DiscName", "_BillKey"], as_index=False)["DiscAmt"].sum()
                     .merge(bill_rev, on="_BillKey", how="left"))

        promo_roi = (promo_rev.groupby("DiscName")
            .agg(
                Discount=("DiscAmt","sum"),
                Revenue=("_NetBAmt","sum"),
                Bills=("_BillKey","nunique")
            )
            .reset_index()
        )

        promo_roi["ROI"] = promo_roi["Revenue"] / promo_roi["Discount"].replace(0,1)

        fig_roi = px.bar(
            promo_roi.sort_values("ROI", ascending=False).head(10),
            x="ROI",
            y="DiscName",
            orientation="h",
            color="ROI",
            color_continuous_scale=["#f85149","#d4af37","#3fb950"]
        )
        fig_roi.update_layout(**LAYOUT_BASE, height=400)
        st.plotly_chart(fig_roi, use_container_width=True)

        # ─────────────────────────────────────────────
        # 💰 MARGIN
        # ─────────────────────────────────────────────
        st.markdown('<div class="section-title">💰 Margin After Discount</div>', unsafe_allow_html=True)

        MARGIN_RATE = 0.6
        promo_roi["Profit"] = promo_roi["Revenue"] * MARGIN_RATE - promo_roi["Discount"]

        fig_margin = px.bar(
            promo_roi.sort_values("Profit").head(10),
            x="Profit",
            y="DiscName",
            orientation="h",
            color="Profit",
            color_continuous_scale=["#f85149","#d4af37","#3fb950"]
        )
        fig_margin.update_layout(**LAYOUT_BASE, height=400)
        st.plotly_chart(fig_margin, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# PAGE 7: MENU ENGINEERING
# ═══════════════════════════════════════════════════════════════
elif page == "🧑‍🏫 Menu Engineering":
    import numpy as np

    st.markdown('<div class="section-title" style="font-size:24px;">🧑‍🏫 Menu Engineering</div>', unsafe_allow_html=True)
    st.markdown("""
    <div style='font-size:13px;color:#8b949e;margin-bottom:16px;'>
        <span style='color:#19db6a;font-weight:700;'>⭐ Star</span> ขายดี+กำไรดี &nbsp;|&nbsp;
        <span style='color:#f1c71e;font-weight:700;'>🧩 Puzzle</span> กำไรดี+ขายน้อย &nbsp;|&nbsp;
        <span style='color:#f77606;font-weight:700;'>🐴 Plowhorse</span> ขายดี+กำไรต่ำ &nbsp;|&nbsp;
        <span style='color:#eb3c28;font-weight:700;'>🐶 Dog</span> ขายน้อย+กำไรต่ำ
    </div>
    """, unsafe_allow_html=True)

    @st.cache_data(ttl=300, show_spinner=False)
    def me_get_periods():
        try:
            conn = pymysql.connect(**MYSQL_CFG)
            df = pd.read_sql("SELECT DISTINCT period_days FROM menu_engineering_summary ORDER BY period_days", conn)
            conn.close()
            return df["period_days"].tolist()
        except:
            return []

    @st.cache_data(ttl=300, show_spinner=False)
    def me_load(period_days):
        conn = pymysql.connect(**MYSQL_CFG)
        df = pd.read_sql("""
            SELECT * FROM menu_engineering_summary
            WHERE period_days = %s
              AND date_from = (
                SELECT MAX(date_from) FROM menu_engineering_summary WHERE period_days = %s
              )
        """, conn, params=(period_days, period_days))
        conn.close()
        return df

    period_list = me_get_periods()

    if not period_list:
        st.warning("⚠️ ไม่มีข้อมูล Menu Engineering หรือเชื่อมต่อ DB ไม่ได้")
    else:
        period = st.selectbox("📅 เลือกช่วงเวลา (วัน)", period_list)
        df_me  = me_load(period)

        if df_me.empty:
            st.warning("ไม่มีข้อมูลสำหรับช่วงนี้")
        else:
            color_map = {"Stars":"#19db6a","Puzzles":"#f1c71e","Plow Horses":"#f77606","Dogs":"#eb3c28"}

            c1,c2,c3,c4 = st.columns(4)
            cat_counts = df_me["menu_category"].value_counts()
            c1.metric("⭐ Stars",       int(cat_counts.get("Stars", 0)))
            c2.metric("🧩 Puzzles",     int(cat_counts.get("Puzzles", 0)))
            c3.metric("🐴 Plow Horses", int(cat_counts.get("Plow Horses", 0)))
            c4.metric("🐶 Dogs",        int(cat_counts.get("Dogs", 0)))

            tab_me1, tab_me2, tab_me3, tab_me4 = st.tabs(["📊 Matrix","📈 Forecast","🔗 Combo","🤖 AI Analysis"])

            # ── TAB 1: MATRIX ───────────────────────────────────────
            with tab_me1:
                f1, f2, f3 = st.columns([2, 2, 1])
                with f1:
                    cat_filter = st.multiselect("🔍 กรอง Category",
                        ["Stars","Puzzles","Plow Horses","Dogs"],
                        default=["Stars","Puzzles","Plow Horses","Dogs"], key="me_cat_filter")
                with f2:
                    search_txt = st.text_input("🔍 ค้นหาชื่อเมนู", "", key="me_search", placeholder="พิมพ์ชื่อเมนู...")
                with f3:
                    top_n_filter = st.number_input("Top N ต่อ quadrant", 1, 50, 10, key="me_topn")

                df_f = df_me[df_me["menu_category"].isin(cat_filter)].copy()
                if search_txt:
                    df_f = df_f[df_f["item_name"].str.contains(search_txt, case=False, na=False)]

                avg_cm  = df_me["avg_cm"].iloc[0]
                avg_qty = df_me["avg_qty"].iloc[0]
                x_max   = np.percentile(df_me["cm_per_unit"], 95)
                y_max   = np.percentile(df_me["total_qty"], 95)
                size_col= np.clip(df_f["total_sales"], 0, np.percentile(df_me["total_sales"], 95))

                # เพิ่ม CM% display สำหรับ tooltip
                df_f = df_f.copy()
                df_f["cm_pct_display"] = (df_f["total_profit"] / df_f["total_sales"] * 100).round(1).astype(str) + "%"

                fig = px.scatter(df_f, x="cm_per_unit", y="total_qty",
                    color="menu_category", color_discrete_map=color_map,
                    size=size_col, size_max=40, hover_name="item_name",
                    hover_data={
                        "cm_per_unit":    ":.0f",
                        "total_qty":      ":,.0f",
                        "total_cost":     ":,.0f",
                        "total_profit":   ":,.0f",
                        "cm_pct_display": True,
                        "menu_category":  True,
                        "total_sales":    False,
                    },
                    title=f"Menu Engineering Matrix — {period} วัน ({len(df_f)} รายการ)", opacity=0.82)
                fig.update_xaxes(range=[0, x_max*1.05])
                fig.update_yaxes(range=[0, y_max*1.05])
                fig.add_vline(x=avg_cm, line_dash="dot", line_color="rgba(255,255,255,0.35)", line_width=1.5,
                    annotation_text=f"avg CM ฿{avg_cm:.0f}",
                    annotation_font=dict(color="rgba(255,255,255,0.5)",size=11),
                    annotation_position="top right")
                fig.add_hline(y=avg_qty, line_dash="dot", line_color="rgba(255,255,255,0.35)", line_width=1.5,
                    annotation_text=f"avg qty {avg_qty:.0f}",
                    annotation_font=dict(color="rgba(255,255,255,0.5)",size=11),
                    annotation_position="top right")
                for text, x, y in [
                    ("⭐ Stars",       x_max*0.75, y_max*0.90),
                    ("🐴 Plow Horses", x_max*0.10, y_max*0.90),
                    ("🧩 Puzzles",     x_max*0.75, y_max*0.12),
                    ("🐶 Dogs",        x_max*0.10, y_max*0.12),
                ]:
                    fig.add_annotation(x=x, y=y, text=text, showarrow=False,
                        font=dict(size=13,color="rgba(255,255,255,0.25)"), bgcolor="rgba(0,0,0,0)")
                fig.update_layout(**LAYOUT_BASE, height=560,
                    xaxis_title="Avg CM / Unit (฿)", yaxis_title="Total Qty Sold")
                fig.update_layout(legend=dict(title="Category",orientation="v",x=1.01,y=1,
                    bgcolor="rgba(22,27,34,0.9)",bordercolor="#30363d",borderwidth=1,font=dict(size=12)))
                st.plotly_chart(fig, use_container_width=True)

                # ── Top/Bottom per Quadrant ──
                st.markdown('<div class="section-title">📊 Top/Bottom ต่อ Quadrant</div>', unsafe_allow_html=True)

                # columns ที่จะแสดง — ใช้เฉพาะที่มีใน df จริง
                disp_cols = [c for c in [
                    "item_name","total_qty","cm_per_unit",
                    "total_cost","total_profit","total_sales","cm_ratio"
                ] if c in df_f.columns]

                disp_rename = {
                    "item_name":   "เมนู",
                    "total_qty":   "จำนวน",
                    "cm_per_unit": "CM/unit",
                    "total_cost":  "ต้นทุนรวม",
                    "total_profit":"CM Profit",
                    "total_sales": "ยอดขาย",
                    "cm_ratio":    "CM%",
                }

                quadrants = [
                    ("⭐ Stars",       "Stars",       "#19db6a"),
                    ("🧩 Puzzles",     "Puzzles",     "#f1c71e"),
                    ("🐴 Plow Horses", "Plow Horses", "#f77606"),
                    ("🐶 Dogs",        "Dogs",        "#eb3c28"),
                ]
                for q_label, q_key, q_color in quadrants:
                    q_df = df_f[df_f["menu_category"]==q_key][disp_cols].copy()
                    if q_df.empty:
                        continue
                    # format cm_ratio เป็น %
                    if "cm_ratio" in q_df.columns:
                        q_df["cm_ratio"] = (q_df["cm_ratio"] * 100).round(1).astype(str) + "%"
                    with st.expander(f"{q_label} — {len(q_df)} รายการ", expanded=(q_key=="Stars")):
                        c_top, c_bot = st.columns(2)
                        with c_top:
                            st.markdown(f"<div style='font-size:12px;font-weight:700;color:{q_color};margin-bottom:6px;'>🔝 Top {top_n_filter} by CM Profit</div>", unsafe_allow_html=True)
                            st.dataframe(
                                q_df.sort_values("total_profit", ascending=False)
                                    .head(int(top_n_filter))
                                    .rename(columns=disp_rename)
                                    .reset_index(drop=True),
                                use_container_width=True, height=280
                            )
                        with c_bot:
                            st.markdown(f"<div style='font-size:12px;font-weight:700;color:#8b949e;margin-bottom:6px;'>🔻 Bottom {top_n_filter} by CM Profit</div>", unsafe_allow_html=True)
                            st.dataframe(
                                q_df.sort_values("total_profit", ascending=True)
                                    .head(int(top_n_filter))
                                    .rename(columns=disp_rename)
                                    .reset_index(drop=True),
                                use_container_width=True, height=280
                            )

                # ── Export Excel ──
                st.markdown('<div class="section-title">📥 Export ข้อมูล</div>', unsafe_allow_html=True)

                export_df = df_f[[c for c in [
                    "item_name","total_qty","cm_per_unit",
                    "total_cost","total_profit","total_sales","cm_ratio","menu_category"
                ] if c in df_f.columns]].copy()

                export_df = export_df.rename(columns={
                    "item_name":    "เมนู",
                    "total_qty":    "จำนวน",
                    "cm_per_unit":  "CM/unit",
                    "total_cost":   "ต้นทุนรวม",
                    "total_profit": "กำไรรวม",
                    "total_sales":  "ยอดขาย",
                    "cm_ratio":     "CM%",
                    "menu_category":"Category",
                })

                _end   = pd.to_datetime(df_me["date_to"].max()).date()
                _start = pd.to_datetime(df_me["date_from"].max()).date()

                _excel_bytes = build_excel(export_df, _start, _end)
                _filename    = f"menu_engineering_{period}d_{_end.strftime('%Y%m%d')}.xlsx"

                st.download_button(
                    label="📥 Download Excel",
                    data=_excel_bytes,
                    file_name=_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            # ── TAB 2: PROPHET FORECAST ─────────────────────────────
            with tab_me2:
                st.markdown('<div class="section-title">📈 Sales Forecast (Prophet)</div>', unsafe_allow_html=True)
                try:
                    from prophet import Prophet

                    @st.cache_data(ttl=600, show_spinner=False)
                    def load_daily_sales(item_code):
                        conn = pymysql.connect(**MYSQL_CFG)
                        df = pd.read_sql(f"""
                            SELECT h.BillDate as ds, SUM(d.ActQTY) as y
                            FROM BillHdrTB h JOIN BillDtlTB d ON h.ShopCode=d.ShopCode AND h.BillNo=d.BillNo
                            WHERE d.ItemCode='{item_code}'
                              AND h.BillDate >= DATE_SUB(CURDATE(), INTERVAL 90 DAY)
                            GROUP BY h.BillDate ORDER BY h.BillDate
                        """, conn)
                        conn.close()
                        df["ds"] = pd.to_datetime(df["ds"])
                        return df

                    top_items_me = df_me.sort_values("total_sales", ascending=False).head(20)
                    item_options = [f"{r['item_code']} — {r['item_name']}" for _, r in top_items_me.iterrows()]
                    sel          = st.selectbox("เลือกเมนู", item_options, key="prophet_item")
                    item_code    = sel.split(" — ")[0]
                    horizon      = st.slider("คาดการณ์กี่วัน", 7, 60, 30, key="prophet_horizon")

                    if st.button("🔮 Run Forecast", key="run_prophet"):
                        with st.spinner("กำลังคาดการณ์..."):
                            df_daily = load_daily_sales(item_code)
                            if len(df_daily) < 14:
                                st.warning("ข้อมูลน้อยเกินไป (ต้องการอย่างน้อย 14 วัน)")
                            else:
                                m = Prophet(yearly_seasonality=False, weekly_seasonality=True,
                                            daily_seasonality=False, changepoint_prior_scale=0.1)
                                m.fit(df_daily)
                                future   = m.make_future_dataframe(periods=horizon)
                                forecast = m.predict(future)
                                fig_f = go.Figure()
                                fig_f.add_trace(go.Scatter(x=df_daily["ds"], y=df_daily["y"],
                                    name="Actual", mode="markers", marker=dict(color=GOLD, size=5)))
                                fig_f.add_trace(go.Scatter(x=forecast["ds"], y=forecast["yhat"],
                                    name="Forecast", line=dict(color=BLUE, width=2)))
                                fig_f.add_trace(go.Scatter(
                                    x=pd.concat([forecast["ds"], forecast["ds"][::-1]]),
                                    y=pd.concat([forecast["yhat_upper"], forecast["yhat_lower"][::-1]]),
                                    fill="toself", fillcolor="rgba(88,166,255,0.1)",
                                    line=dict(color="rgba(0,0,0,0)"), name="Confidence"))
                                fig_f.add_vline(x=str(df_daily["ds"].max()),
                                    line_dash="dash", line_color=GOLD, line_width=1)
                                fig_f.update_layout(**LAYOUT_BASE, height=420,
                                    title=f"Forecast: {sel.split(' — ')[1]} (+{horizon} วัน)")
                                st.plotly_chart(fig_f, use_container_width=True)
                                fut_only = forecast[forecast["ds"] > df_daily["ds"].max()]
                                c1,c2,c3 = st.columns(3)
                                c1.metric("📦 Avg/วัน (คาดการณ์)", f"{fut_only['yhat'].mean():.1f}")
                                c2.metric("📈 Max/วัน", f"{fut_only['yhat_upper'].max():.0f}")
                                c3.metric("📉 Min/วัน", f"{fut_only['yhat_lower'].min():.0f}")
                except ImportError:
                    st.error("❌ ต้องติดตั้ง prophet ก่อน:")
                    st.code("/opt/promotion_dashboard/venv/bin/pip install prophet", language="bash")

            # ── TAB 3: ASSOCIATION RULES ────────────────────────────
            with tab_me3:
                st.markdown('<div class="section-title">🔗 Combo Suggestions (Association Rules)</div>', unsafe_allow_html=True)
                try:
                    from mlxtend.frequent_patterns import apriori, association_rules
                    from mlxtend.preprocessing import TransactionEncoder

                    @st.cache_data(ttl=600, show_spinner=False)
                    def load_baskets(date_from_s, date_to_s):
                        conn = pymysql.connect(**MYSQL_CFG)
                        df = pd.read_sql(f"""
                            SELECT d.BillNo, d.ShopCode, i.ItemDes1
                            FROM BillDtlTB d JOIN BillHdrTB h ON h.ShopCode=d.ShopCode AND h.BillNo=d.BillNo
                            LEFT JOIN ItmMasTB i ON d.ItemCode=i.ItemCode
                            WHERE h.BillDate BETWEEN '{date_from_s}' AND '{date_to_s}' AND d.NetAmt > 0
                        """, conn)
                        conn.close()
                        return df

                    min_support    = st.slider("Min Support (%)", 1, 20, 3, key="apr_sup") / 100
                    min_confidence = st.slider("Min Confidence (%)", 10, 80, 30, key="apr_conf") / 100
                    min_lift       = st.slider("Min Lift", 1.0, 5.0, 1.2, step=0.1, key="apr_lift")

                    if st.button("🔍 วิเคราะห์ Combo", key="run_apriori"):
                        with st.spinner("กำลังวิเคราะห์ basket..."):
                            df_bsk  = load_baskets(str(d_from), str(d_to))
                            baskets = df_bsk.groupby(["BillNo","ShopCode"])["ItemDes1"].apply(list).tolist()
                            te      = TransactionEncoder()
                            df_enc  = pd.DataFrame(te.fit(baskets).transform(baskets), columns=te.columns_)
                            freq    = apriori(df_enc, min_support=min_support, use_colnames=True)
                            if freq.empty:
                                st.warning("ไม่พบ pattern — ลด Min Support ลง")
                            else:
                                rules = association_rules(freq, metric="lift", min_threshold=min_lift)
                                rules = rules[rules["confidence"] >= min_confidence].sort_values("lift", ascending=False)
                                st.success(f"พบ {len(rules)} combo rules")
                                for i, (_, r) in enumerate(rules.head(10).iterrows()):
                                    ant = ", ".join(list(r["antecedents"]))
                                    con = ", ".join(list(r["consequents"]))
                                    st.markdown(f"""<div class="card card-gold" style="margin:6px 0;padding:14px 18px;">
                                        <div style="font-size:13px;color:#e6edf3;">
                                            <span style="color:{GOLD};font-weight:700;">#{i+1}</span> &nbsp;
                                            <b>{ant}</b><span style="color:#8b949e;"> → </span><b style="color:{BLUE};">{con}</b>
                                        </div>
                                        <div style="font-size:11px;color:#8b949e;margin-top:4px;">
                                            Support {r['support']:.1%} | Confidence {r['confidence']:.1%} | Lift <span style="color:{GREEN};">{r['lift']:.2f}x</span>
                                        </div></div>""", unsafe_allow_html=True)
                                st.dataframe(rules[["antecedents","consequents","support","confidence","lift"]].head(50).reset_index(drop=True), use_container_width=True)
                except ImportError:
                    st.error("❌ ต้องติดตั้ง mlxtend ก่อน:")
                    st.code("/opt/promotion_dashboard/venv/bin/pip install mlxtend", language="bash")

            # ── TAB 4: AI ANALYSIS ──────────────────────────────────
            with tab_me4:
                st.markdown('<div class="section-title">🤖 AI Strategic Analysis (Groq)</div>', unsafe_allow_html=True)
                st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:12px;">สมัครฟรีที่ <a href="https://console.groq.com" target="_blank" style="color:#58a6ff;">console.groq.com</a> → API Keys → Create API Key</div>', unsafe_allow_html=True)
                api_key = st.text_input("Groq API Key", type="password", placeholder="gsk_...", key="groq_key")
                model   = st.selectbox("Model", [
                    "llama-3.3-70b-versatile",
                    "llama-3.1-8b-instant",
                    "mixtral-8x7b-32768",
                ], key="groq_model")
                focus   = st.multiselect("วิเคราะห์ด้าน", [
                    "📈 แนะนำราคาที่ควรปรับ","🎁 โปรโมชั่นที่ควรทำ",
                    "🗑️ เมนูที่ควรตัดออก","⭐ เมนูที่ควร push","📦 การจัดการ stock/ต้นทุน",
                ], default=["📈 แนะนำราคาที่ควรปรับ","🎁 โปรโมชั่นที่ควรทำ"], key="ai_focus")

                if st.button("🤖 วิเคราะห์ด้วย AI", key="run_groq"):
                    if not api_key:
                        st.warning("กรุณาใส่ Groq API Key ก่อน")
                    else:
                        with st.spinner("AI กำลังวิเคราะห์..."):
                            try:
                                import requests as req

                                def _me_top(cat):
                                    cols = [c for c in ["item_name","total_qty","cm_per_unit","total_cost","total_profit","cm_ratio"] if c in df_me.columns]
                                    return df_me[df_me["menu_category"]==cat][cols].head(10)

                                total_rev     = df_me["total_sales"].sum()
                                total_cm      = df_me["total_profit"].sum()
                                overall_cm_pct= total_cm / total_rev * 100 if total_rev else 0

                                prompt = f"""คุณเป็น Restaurant Business Consultant ผู้เชี่ยวชาญ Menu Engineering

ข้อมูล Menu Engineering ของ Greyhound Café (ช่วง {period} วัน):
- Stars: {len(df_me[df_me["menu_category"]=="Stars"])} รายการ | Plow Horses: {len(df_me[df_me["menu_category"]=="Plow Horses"])} รายการ
- Puzzles: {len(df_me[df_me["menu_category"]=="Puzzles"])} รายการ | Dogs: {len(df_me[df_me["menu_category"]=="Dogs"])} รายการ
- Avg CM/unit: ฿{df_me["avg_cm"].iloc[0]:.0f} | Total Revenue: ฿{total_rev:,.0f}
- Total CM Profit: ฿{total_cm:,.0f} | Overall CM%: {overall_cm_pct:.1f}%

⭐ Top Stars:
{_me_top("Stars").to_string(index=False)}
🐴 Top Plow Horses:
{_me_top("Plow Horses").to_string(index=False)}
🧩 Top Puzzles:
{_me_top("Puzzles").to_string(index=False)}
🐶 Top Dogs:
{_me_top("Dogs").to_string(index=False)}

กรุณาวิเคราะห์ด้าน: {", ".join(focus)}
ตอบเป็นภาษาไทย structured มี header แต่ละด้าน พร้อม action items ที่ชัดเจน"""

                                resp = req.post(
                                    "https://api.groq.com/openai/v1/chat/completions",
                                    headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                                    json={
                                        "model": model,
                                        "max_tokens": 2048,
                                        "messages": [
                                            {"role": "system", "content": "คุณเป็น Restaurant Business Consultant ผู้เชี่ยวชาญด้าน Menu Engineering และกลยุทธ์ร้านอาหาร ตอบเป็นภาษาไทยเสมอ"},
                                            {"role": "user", "content": prompt}
                                        ]
                                    },
                                    timeout=60
                                )
                                result = resp.json()
                                if "choices" in result:
                                    analysis = result["choices"][0]["message"]["content"]
                                    st.markdown(f'<div class="card card-gold" style="padding:24px;line-height:1.9;white-space:pre-wrap;">{analysis}</div>', unsafe_allow_html=True)
                                    if "usage" in result:
                                        u = result["usage"]
                                        st.caption(f"Tokens: {u.get('prompt_tokens',0)} prompt + {u.get('completion_tokens',0)} output = {u.get('total_tokens',0)} total")
                                else:
                                    st.error(f"API Error: {result.get('error', {}).get('message', str(result))}")
                            except Exception as e:
                                st.error(f"❌ {e}")

# ═══════════════════════════════════════════════════════════════
# PAGE 8: INGREDIENT FORECAST
# ═══════════════════════════════════════════════════════════════
elif page == "🥩 Ingredient Forecast" and data_ok:
    st.markdown('<div class="section-title" style="font-size:24px;">🥩 Ingredient Forecast</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:16px;">คาดการณ์วัตถุดิบที่ต้องใช้จากยอดขาย × Recipe</div>', unsafe_allow_html=True)

    # ── Load recipe & ingredient from MySQL ──
    @st.cache_data(ttl=300, show_spinner=False)
    def load_recipe_data():
        conn = pymysql.connect(**MYSQL_CFG)
        try:
            df_rec = pd.read_sql("""
                SELECT r.rec_code, r.item_code AS ing_code,
                       i.item_name AS ing_name, i.use_unit,
                       r.quantity, i.stnd_cost,
                       COALESCE(s.current_stock,0) AS current_stock,
                       COALESCE(s.min_stock,0) AS min_stock
                FROM recipe_tb r
                JOIN ingredient_tb i ON r.item_code = i.item_code
                LEFT JOIN ingredient_stock_tb s ON i.item_code = s.item_code
            """, conn)
            df_menu = pd.read_sql("""
                SELECT ItemCode, ItemDes1 FROM ItmMasTB
            """, conn)
        finally:
            conn.close()
        return df_rec, df_menu

    try:
        df_rec, df_menu = load_recipe_data()
        has_recipe = len(df_rec) > 0
    except Exception as e:
        st.error(f"❌ โหลด recipe ไม่ได้: {e}")
        has_recipe = False

    if not has_recipe:
        st.warning("⚠️ ยังไม่มีข้อมูล Recipe — ไปที่ ETL Admin → Sync Recipe & Ingredient ก่อน")
        st.stop()

    # ── Forecast horizon ──
    c1, c2, c3 = st.columns(3)
    horizon = c1.selectbox("📅 คาดการณ์ล่วงหน้า", [7, 14, 30], index=0, key="ing_horizon")
    ing_search = c2.text_input("🔍 ค้นหาวัตถุดิบ / รหัส", "", key="ing_search",
                               placeholder="พิมพ์เพื่อกรอง...",
                               on_change=None)
    show_alert_only = c3.checkbox("⚠️ แสดงเฉพาะที่ต้องสั่ง", key="ing_alert_only")

    st.markdown("<hr style='border-color:#30363d;margin:8px 0;'>", unsafe_allow_html=True)

    if st.button("🔮 คำนวณ Ingredient Forecast", use_container_width=True, key="run_ing_forecast"):

        with st.spinner("กำลังคำนวณ..."):

            # ── Step 1: คำนวณ avg daily qty ต่อ rec_code จาก BillDtlTB ──
            # ใช้ 30 วันล้าสุด เป็น baseline (ไม่ต้อง Prophet ก็ได้ถ้าข้อมูลน้อย)
            from datetime import date, timedelta
            days_back = 30
            date_start = (date.today() - timedelta(days=days_back)).strftime("%Y-%m-%d")
            date_end   = date.today().strftime("%Y-%m-%d")

            try:
                conn = pymysql.connect(**MYSQL_CFG)
                df_sales = pd.read_sql(f"""
                    SELECT d.ItemCode, SUM(d.ActQTY) as total_qty,
                           COUNT(DISTINCT h.BillDate) as days_sold
                    FROM BillDtlTB d
                    JOIN BillHdrTB h ON d.ShopCode=h.ShopCode AND d.BillNo=h.BillNo
                    WHERE h.BillDate BETWEEN '{date_start}' AND '{date_end}'
                      AND d.NetAmt > 0
                    GROUP BY d.ItemCode
                """, conn)
                conn.close()
            except Exception as e:
                st.error(f"❌ Query ไม่ได้: {e}")
                st.stop()

            df_sales["avg_daily_qty"] = df_sales["total_qty"] / days_back
            df_sales["forecast_qty"]  = df_sales["avg_daily_qty"] * horizon

            # ── Step 2: Join recipe → คำนวณ ingredient demand ──
            # rec_code ใน recipe_tb ตรงกับ ItemCode ใน BillDtlTB
            df_demand = df_rec.merge(
                df_sales[["ItemCode","forecast_qty"]],
                left_on="rec_code", right_on="ItemCode", how="left"
            ).fillna({"forecast_qty": 0})

            df_demand["demand"] = df_demand["forecast_qty"] * df_demand["quantity"]

            # ── Step 3: รวม demand ต่อวัตถุดิบ ──
            df_result = (df_demand.groupby(["ing_code","ing_name","use_unit","current_stock","min_stock","stnd_cost"])
                        .agg(demand=("demand","sum"))
                        .reset_index()
                        .sort_values("demand", ascending=False))

            df_result["need_to_order"] = (df_result["demand"] - df_result["current_stock"]).clip(lower=0)
            df_result["status"] = df_result.apply(
                lambda r: "🔴 สั่งด่วน" if r["need_to_order"] > 0 and r["current_stock"] <= r["min_stock"]
                          else ("🟡 ใกล้หมด" if r["need_to_order"] > 0 else "🟢 พอ"), axis=1)
            df_result["est_cost"] = df_result["need_to_order"] * df_result["stnd_cost"]

            st.session_state["ing_forecast_result"] = df_result
            st.session_state["ing_horizon_used"]    = horizon

    # ── Display results ──
    if "ing_forecast_result" in st.session_state:
        df_result = st.session_state["ing_forecast_result"].copy()
        horizon_used = st.session_state.get("ing_horizon_used", horizon)

        # filter
        if ing_search:
            df_result = df_result[
                df_result["ing_name"].str.contains(ing_search, case=False, na=False) |
                df_result["ing_code"].str.contains(ing_search, case=False, na=False)
            ]
        if show_alert_only:
            df_result = df_result[df_result["need_to_order"] > 0]

        # KPIs
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("🔴 สั่งด่วน",   int((df_result["status"]=="🔴 สั่งด่วน").sum()))
        c2.metric("🟡 ใกล้หมด",   int((df_result["status"]=="🟡 ใกล้หมด").sum()))
        c3.metric("🟢 พอ",         int((df_result["status"]=="🟢 พอ").sum()))
        c4.metric("💰 ต้นทุนสั่งซื้อ", f"฿{df_result['est_cost'].sum():,.0f}")

        st.markdown(f'<div style="font-size:12px;color:#8b949e;margin:8px 0;">คาดการณ์ {horizon_used} วัน จาก avg 30 วันล้าสุด</div>', unsafe_allow_html=True)

        # Color rows
        for _, row in df_result.iterrows():
            status_color = {"🔴 สั่งด่วน":"#f85149","🟡 ใกล้หมด":"#d29922","🟢 พอ":"#3fb950"}.get(row["status"],"#8b949e")
            st.markdown(f"""
            <div style="display:flex;align-items:center;padding:10px 16px;margin:4px 0;
                        background:#161b22;border:1px solid #30363d;border-left:4px solid {status_color};
                        border-radius:8px;gap:12px;">
                <div style="flex:3;color:#e6edf3;font-size:13px;">
                    <div style="font-weight:600;">{row['ing_name']}</div>
                    <div style="font-size:11px;color:#8b949e;margin-top:2px;">{row['ing_code']}</div>
                </div>
                <div style="flex:1;font-size:12px;color:#8b949e;text-align:center;">{row['use_unit']}</div>
                <div style="flex:1;text-align:right;font-size:12px;color:#8b949e;">
                    ต้องใช้<br><b style="color:#e6edf3;">{row['demand']:.1f}</b>
                </div>
                <div style="flex:1;text-align:right;font-size:12px;color:#8b949e;">
                    stock<br><b style="color:#58a6ff;">{row['current_stock']:.1f}</b>
                </div>
                <div style="flex:1;text-align:right;font-size:12px;color:#8b949e;">
                    ต้องสั่ง<br><b style="color:{status_color};">{row['need_to_order']:.1f}</b>
                </div>
                <div style="flex:1;text-align:right;font-size:12px;color:#8b949e;">
                    ต้นทุน<br><b style="color:#d4af37;">฿{row['est_cost']:,.0f}</b>
                </div>
                <div style="flex:1;text-align:center;font-size:12px;">{row['status']}</div>
            </div>""", unsafe_allow_html=True)

        # Export
        st.markdown("<br>", unsafe_allow_html=True)
        try:
            import io
            out = io.BytesIO()
            export = df_result.rename(columns={
                "ing_code":"รหัส","ing_name":"วัตถุดิบ","use_unit":"หน่วย",
                "demand":"ต้องใช้","current_stock":"stock","min_stock":"min",
                "need_to_order":"ต้องสั่ง","est_cost":"ต้นทุน(฿)","status":"สถานะ"
            })
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                export.to_excel(writer, sheet_name=f"Forecast_{horizon_used}d", index=False)
            out.seek(0)
            st.download_button("📥 Export Excel",
                data=out,
                file_name=f"ingredient_forecast_{horizon_used}d_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        except ImportError:
            pass
    else:
        st.info("กด **คำนวณ Ingredient Forecast** เพื่อดูผล")


# PAGE 9: ADMIN  →  ETL Admin App (embedded)
# ═══════════════════════════════════════════════════════════════
elif page == "⚙️ Admin":
    st.markdown('<div class="section-title" style="font-size:20px;">⚙️ ETL Admin Panel</div>', unsafe_allow_html=True)
    st.components.v1.iframe(ETL_ADMIN_URL, height=900, scrolling=True)