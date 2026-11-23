"""
🐾 GREYHOUND CAFÉ – ETL Admin Dashboard
ระบบจัดการ ETL Pipeline | MSSQL → MySQL
รันที่ port 8801
"""

import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))  # หา gh_env.py ที่โฟลเดอร์แม่
import gh_env   # รหัส DB / API key อยู่ใน /opt/promotion_dashboard/.env (ไม่เข้า git)
import streamlit as st
import pandas as pd
import subprocess
import threading
import queue
import os
import sys
import time
import json
from datetime import datetime, timedelta
from pathlib import Path

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="GH ETL Admin",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
ETL_SCRIPT   = "/opt/promotion_dashboard/etl/etl_sync.py"
VENV_PYTHON  = "/opt/promotion_dashboard/venv/bin/python3"
LOG_FILE     = "/var/log/gh_etl.log"
DASH_URL     = "http://10.8.1.88:89"

# user:password มาจาก .env (ETL_ADMIN_USERS="admin:xxx,it:yyy")
ADMIN_USERS  = dict(p.split(":", 1) for p in gh_env.need("ETL_ADMIN_USERS").split(",") if ":" in p)

# strip ANSI color codes
import re as _re
def strip_ansi(text: str) -> str:
    return _re.sub(r'\x1b\[[0-9;]*m', '', text)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: #0d1117; color: #e6edf3; }

/* LOGIN */
.login-wrap {
    max-width: 420px; margin: 80px auto;
    background: linear-gradient(135deg, #161b22, #1c2230);
    border: 1px solid #d4af37;
    border-radius: 24px; padding: 48px 40px;
    box-shadow: 0 24px 64px rgba(0,0,0,.6);
}
.login-logo { text-align:center; margin-bottom:32px; }
.login-logo .paw { font-size:52px; }
.login-logo h2 { color:#d4af37; font-size:22px; font-weight:800; letter-spacing:2px; margin:8px 0 4px; }
.login-logo p  { color:#8b949e; font-size:12px; margin:0; }

/* HEADER STRIP */
.admin-header {
    background: linear-gradient(135deg, #1a0a00, #3d2200);
    border: 1px solid #d4af37; border-radius:16px;
    padding:20px 28px; margin-bottom:24px;
    display:flex; align-items:center; justify-content:space-between;
}
.admin-header h1 { color:#d4af37; font-size:22px; font-weight:800; margin:0; }
.admin-header p  { color:#a0824a; font-size:12px; margin:4px 0 0; }

/* METRIC */
[data-testid="stMetric"] {
    background: linear-gradient(135deg,#161b22,#1c2230);
    border:1px solid #30363d; border-radius:14px; padding:18px !important;
}
[data-testid="stMetricLabel"] { color:#8b949e !important; font-size:12px !important; }
[data-testid="stMetricValue"] { color:#d4af37 !important; font-size:26px !important; font-weight:700 !important; }

/* CARD */
.card { background:#161b22; border:1px solid #30363d; border-radius:14px; padding:20px; margin:8px 0; }
.card-gold { border-left:4px solid #d4af37; }

/* TABLE PILL */
.tbl-row {
    display:flex; align-items:center; padding:10px 16px; margin:4px 0;
    background:#161b22; border:1px solid #30363d; border-radius:8px;
}
.tbl-row:hover { border-color:#d4af37; }

/* LOG BOX */
.log-box {
    background:#010409; border:1px solid #21262d; border-radius:10px;
    padding:16px; font-family:'Courier New',monospace; font-size:12px;
    color:#3fb950; max-height:400px; overflow-y:auto; white-space:pre-wrap;
}

/* BADGE */
.badge { display:inline-block; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-ok  { background:#3fb95030; color:#3fb950; border:1px solid #3fb950; }
.badge-err { background:#f8514930; color:#f85149; border:1px solid #f85149; }
.badge-run { background:#d4af3730; color:#d4af37; border:1px solid #d4af37; }
.badge-skip{ background:#8b949e30; color:#8b949e; border:1px solid #8b949e; }

/* INPUT */
.stTextInput input, .stDateInput input {
    background:#161b22 !important; border:1px solid #30363d !important;
    color:#e6edf3 !important; border-radius:8px !important;
}
.stTextInput input:focus { border-color:#d4af37 !important; }

/* BUTTON */
.stButton button {
    border-radius:10px !important; font-weight:600 !important;
}

/* TABS */
.stTabs [data-baseweb="tab-list"] { background:#161b22 !important; border-radius:10px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:8px !important; color:#8b949e !important; font-weight:600 !important; }
.stTabs [aria-selected="true"] { background:#d4af37 !important; color:#0d1117 !important; }

section[data-testid="stSidebar"] { display:none; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "logged_in"   not in st.session_state: st.session_state.logged_in   = False
if "username"    not in st.session_state: st.session_state.username    = ""
if "login_tries" not in st.session_state: st.session_state.login_tries = 0
if "etl_running" not in st.session_state: st.session_state.etl_running = False
if "etl_log"     not in st.session_state: st.session_state.etl_log     = []


# ═══════════════════════════════════════════════
# LOGIN PAGE
# ═══════════════════════════════════════════════
def show_login():
    # ล็อกหน้าไว้ไม่ให้ admin content โผล่
    st.markdown("""
    <style>
    section.main > div { padding-top: 0 !important; }
    </style>
    <div class="login-wrap">
        <div class="login-logo">
            <div class="paw">🐾</div>
            <h2>GREYHOUND CAFÉ</h2>
            <p>ETL Admin · Authorized Access Only</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col = st.columns([1, 1.4, 1])[1]
    with col:

        if st.session_state.login_tries >= 5:
            st.error("🔒 ล็อกอินผิดเกิน 5 ครั้ง กรุณารีเฟรชหน้า")
            return

        user = st.text_input("👤 Username", placeholder="username", key="li_user")
        pwd  = st.text_input("🔑 Password", type="password", placeholder="••••••••", key="li_pwd")

        c1, c2 = st.columns([2, 1])
        login_btn = c1.button("🔓 เข้าสู่ระบบ", use_container_width=True, type="primary")
        c2.markdown(
            f"<a href='{DASH_URL}' target='_blank'>"
            "<button style='width:100%;padding:8px;background:#161b22;color:#d4af37;"
            "border:1px solid #d4af37;border-radius:8px;cursor:pointer;font-weight:600;"
            "font-size:13px;'>📊 Dashboard</button></a>",
            unsafe_allow_html=True
        )

        if login_btn:
            if user in ADMIN_USERS and ADMIN_USERS[user] == pwd:
                st.session_state.logged_in   = True
                st.session_state.username    = user
                st.session_state.login_tries = 0
                # clear login widget state
                for k in ["li_user","li_pwd"]:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()
            else:
                st.session_state.login_tries += 1
                remaining = 5 - st.session_state.login_tries
                st.error(f"❌ Username หรือ Password ไม่ถูกต้อง (เหลือ {remaining} ครั้ง)")


# ═══════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════
def read_log_tail(n=100) -> str:
    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        return "".join(lines[-n:])
    except FileNotFoundError:
        return "— ยังไม่มี log —"
    except Exception as e:
        return f"Error reading log: {e}"

def run_etl(args_extra: list) -> list:
    """รัน etl_sync.py แล้วคืน output เป็น list of lines (ANSI stripped)"""
    python = VENV_PYTHON if Path(VENV_PYTHON).exists() else sys.executable
    cmd    = [python, ETL_SCRIPT] + args_extra
    lines  = []
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            cwd=str(Path(ETL_SCRIPT).parent)
        )
        for line in proc.stdout:
            lines.append(strip_ansi(line.rstrip()))
        proc.wait()
        lines.append(f"\n── Exit code: {proc.returncode} ──")
    except Exception as e:
        lines.append(f"ERROR: {e}")
    return lines

@st.cache_data(ttl=60, show_spinner=False)
def get_db_stats():
    """ดึงสถิติจาก MySQL"""
    try:
        import pymysql
        conn = pymysql.connect(**gh_env.mysql_cfg("root", connect_timeout=5))
        cur = conn.cursor()
        stats = {}
        for tbl in ["BillHdrTB","BillDtlTB","ONLBillMultiDiscTB","BILLPAYTB",
                    "MemberTB","MemTypetb","PosShopTB","ItmMasTB","ItmDepTB","ItmCatTB"]:
            try:
                cur.execute(f"SELECT COUNT(*) FROM `{tbl}`")
                stats[tbl] = cur.fetchone()[0]
            except:
                stats[tbl] = None
        # latest bill date
        try:
            cur.execute("SELECT MAX(BillDate) FROM BillHdrTB")
            stats["_latest"] = str(cur.fetchone()[0] or "–")
        except:
            stats["_latest"] = "–"
        conn.close()
        return stats
    except Exception as e:
        return {"_error": str(e)}


# ═══════════════════════════════════════════════
# MAIN ADMIN PAGE
# ═══════════════════════════════════════════════
def show_admin():
    # ซ่อน login form ถ้าโผล่มาซ้อน
    st.markdown("<style>[data-testid='stTextInput'] label:has(+ div input[value='it']) { display:none !important; } </style>", unsafe_allow_html=True)
    # ── Top bar ──
    st.markdown(f"""
    <div class="admin-header">
        <div>
            <h1>⚙️ ETL Admin Dashboard</h1>
            <p>MSSQL (192.168.5.3) → MySQL (10.8.1.88) · GH Promotion</p>
        </div>
        <div style="text-align:right;">
            <div style="font-size:12px;color:#8b949e;">ผู้ใช้งาน</div>
            <div style="font-size:16px;font-weight:700;color:#d4af37;">👤 {st.session_state.username}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Top action bar ──
    bc1, bc2, bc3, bc4 = st.columns([2,2,2,1])
    bc1.markdown(
        f"<a href='{DASH_URL}' target='_blank' style='text-decoration:none;'>"
        "<div style='background:#161b22;border:1px solid #d4af37;border-radius:10px;"
        "padding:10px 16px;text-align:center;color:#d4af37;font-weight:700;font-size:13px;"
        "cursor:pointer;'>📊 ไปที่ Dashboard</div></a>",
        unsafe_allow_html=True
    )
    if bc4.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username  = ""
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ── TABS ──
    tab_run, tab_status, tab_log, tab_conn, tab_me, tab_recipe = st.tabs([
        "🚀 รัน ETL", "📊 สถานะ DB", "📋 Log", "🔗 Connection", "🧑‍🏫 Menu Engineering", "🥩 Recipe & Ingredient"
    ])

    # ════════════════════════════
    # TAB 1 – RUN ETL
    # ════════════════════════════
    with tab_run:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:16px;">🚀 รัน ETL Sync</div>', unsafe_allow_html=True)

        col_l, col_r = st.columns([3, 2])
        with col_l:
            r1c1, r1c2 = st.columns(2)
            date_from = r1c1.date_input("📅 ตั้งแต่วันที่", datetime.today() - timedelta(days=1))
            date_to   = r1c2.date_input("📅 ถึงวันที่",     datetime.today())

            shop_code = st.text_input("🏪 ShopCode (เว้นว่าง = ทั้งหมด)", placeholder="เช่น 0001,0004")

            col_opt1, col_opt2 = st.columns(2)
            inc_master = col_opt1.checkbox("📦 รวม Master Tables (--full)", value=False)
            dry_run    = col_opt2.checkbox("🧪 Dry Run (ทดสอบ ไม่ write)", value=False)

            st.markdown("<br>", unsafe_allow_html=True)

            btn_col1, btn_col2 = st.columns(2)
            run_btn  = btn_col1.button("▶️ เริ่ม Sync", type="primary", use_container_width=True, disabled=st.session_state.etl_running)
            test_btn = btn_col2.button("🔌 Test Connection", use_container_width=True)

        with col_r:
            st.markdown("""
            <div class="card card-gold">
                <div style="font-size:13px;font-weight:700;color:#d4af37;margin-bottom:12px;">⚡ Quick Run</div>
            """, unsafe_allow_html=True)

            presets = {
                "วันนี้":       ["--date_from", str(datetime.today().date()), "--date_to", str(datetime.today().date())],
                "เมื่อวาน":     ["--date_from", str((datetime.today()-timedelta(days=1)).date())],
                "7 วันย้อนหลัง": ["--date_from", str((datetime.today()-timedelta(days=7)).date())],
                "Full Sync":    ["--full","--date_from", str((datetime.today()-timedelta(days=1)).date())],
            }
            for label, preset_args in presets.items():
                if st.button(f"⚡ {label}", use_container_width=True, key=f"preset_{label}"):
                    with st.spinner(f"Running {label}..."):
                        output = run_etl(preset_args)
                        st.session_state.etl_log = output
                    st.success(f"✅ {label} เสร็จสิ้น")
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

        # ── Run logic ──
        if run_btn:
            args = [
                "--date_from", str(date_from),
                "--date_to",   str(date_to),
            ]
            if shop_code.strip():
                args += ["--shop", shop_code.strip()]
            if inc_master:
                args.append("--full")
            if dry_run:
                args.append("--dry-run")

            st.session_state.etl_running = True
            log_box = st.empty()
            log_box.markdown('<div class="log-box">🔄 กำลังรัน ETL...</div>', unsafe_allow_html=True)

            with st.spinner("ETL กำลังทำงาน..."):
                output = run_etl(args)
                st.session_state.etl_log     = output
                st.session_state.etl_running = False

            # colorize output
            colored = ""
            for line in output:
                clean = strip_ansi(line)
                if "✔" in clean or "success" in clean.lower() or "DONE" in clean:
                    colored += f'<span style="color:#3fb950">{clean}</span>\n'
                elif "✘" in clean or "error" in clean.lower() or "ERROR" in clean:
                    colored += f'<span style="color:#f85149">{clean}</span>\n'
                elif "⚠" in clean or "warn" in clean.lower():
                    colored += f'<span style="color:#d29922">{clean}</span>\n'
                elif "→" in clean:
                    colored += f'<span style="color:#58a6ff">{clean}</span>\n'
                else:
                    colored += f'{clean}\n'

            log_box.markdown(f'<div class="log-box">{colored}</div>', unsafe_allow_html=True)
            st.success("✅ ETL เสร็จสิ้น!")

        if test_btn:
            with st.spinner("กำลังทดสอบ..."):
                output = run_etl(["--test-conn"])
                st.session_state.etl_log = output
            for line in output:
                clean = strip_ansi(line).strip()
                if not clean:
                    continue
                if "✔" in clean or "Connected" in clean or "connected" in clean.lower():
                    st.success(clean)
                elif "✘" in clean or "failed" in clean.lower() or "ERROR" in clean:
                    st.error(clean)

    # ════════════════════════════
    # TAB 2 – DB STATUS
    # ════════════════════════════
    with tab_status:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:16px;">📊 สถานะฐานข้อมูล MySQL</div>', unsafe_allow_html=True)

        if st.button("🔄 ดึงข้อมูลสถานะ", key="refresh_stats", use_container_width=True):
            get_db_stats.clear()
            with st.spinner("กำลังดึงข้อมูล..."):
                stats = get_db_stats()
            if "_error" in stats:
                st.error(f"❌ เชื่อมต่อ MySQL ไม่ได้: {stats['_error']}")
            else:
                latest = stats.get("_latest","–")
                c1,c2,c3,c4 = st.columns(4)
                c1.metric("🧾 BillHdrTB",  f"{stats.get('BillHdrTB',0):,}")
                c2.metric("🍽️ BillDtlTB",  f"{stats.get('BillDtlTB',0):,}")
                c3.metric("🎁 DiscTB",     f"{stats.get('ONLBillMultiDiscTB',0):,}")
                c4.metric("📅 ข้อมูลล่าสุด", latest[:10] if latest != "–" else "–")
                st.markdown("<br>", unsafe_allow_html=True)
                TRANS  = ["BillHdrTB","BillDtlTB","ONLBillMultiDiscTB","BILLPAYTB"]
                MASTER = ["MemberTB","MemTypetb","PosShopTB","ItmMasTB","ItmDepTB","ItmCatTB"]
                col_t, col_m = st.columns(2)
                with col_t:
                    st.markdown('<div style="font-size:14px;font-weight:700;color:#58a6ff;margin-bottom:8px;">Transaction Tables</div>', unsafe_allow_html=True)
                    for tbl in TRANS:
                        rows = stats.get(tbl)
                        badge = f'<span class="badge badge-ok">{rows:,} rows</span>' if rows is not None else '<span class="badge badge-err">ไม่มีข้อมูล</span>'
                        st.markdown(f'<div class="tbl-row"><span style="flex:1;font-weight:600;">{tbl}</span>{badge}</div>', unsafe_allow_html=True)
                with col_m:
                    st.markdown('<div style="font-size:14px;font-weight:700;color:#d4af37;margin-bottom:8px;">Master Tables</div>', unsafe_allow_html=True)
                    for tbl in MASTER:
                        rows = stats.get(tbl)
                        badge = f'<span class="badge badge-ok">{rows:,} rows</span>' if rows is not None else '<span class="badge badge-err">ไม่มีข้อมูล</span>'
                        st.markdown(f'<div class="tbl-row"><span style="flex:1;font-weight:600;">{tbl}</span>{badge}</div>', unsafe_allow_html=True)


    # ════════════════════════════
    # TAB 3 – LOG
    # ════════════════════════════
    with tab_log:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:16px;">📋 ETL Log</div>', unsafe_allow_html=True)

        lc1, lc2, lc3 = st.columns([2,2,1])
        n_lines = lc1.selectbox("แสดง", [50, 100, 200, 500], index=1, key="log_lines")
        if lc3.button("🔄 Refresh Log"):
            st.rerun()

        # Session log (last run)
        if st.session_state.etl_log:
            st.markdown("**🕐 Log จากการรันล่าสุด:**")
            log_text = "\n".join(st.session_state.etl_log)
            st.markdown(f'<div class="log-box">{log_text}</div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        # File log
        st.markdown(f"**📁 File Log: `{LOG_FILE}`**")
        file_log = read_log_tail(n_lines)
        st.markdown(f'<div class="log-box">{file_log}</div>', unsafe_allow_html=True)

        # Download log
        st.download_button(
            "⬇️ Download Log",
            data=file_log,
            file_name=f"gh_etl_{datetime.today().strftime('%Y%m%d')}.log",
            mime="text/plain"
        )

    # ════════════════════════════
    # TAB 4 – CONNECTION
    # ════════════════════════════
    with tab_conn:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:16px;">🔗 ทดสอบ Connection</div>', unsafe_allow_html=True)

        c_l, c_r = st.columns(2)
        with c_l:
            st.markdown('<div class="card card-gold"><div style="font-weight:700;color:#58a6ff;margin-bottom:12px;">📡 MSSQL – Data Center</div>', unsafe_allow_html=True)
            ms_srv = st.text_input("Server", "192.168.5.3,1433",  key="c_ms_srv")
            ms_db  = st.text_input("Database", "GH",             key="c_ms_db")
            ms_usr = st.text_input("User", "cmpos01",             key="c_ms_usr")
            ms_pwd = st.text_input("Password", type="password",   key="c_ms_pwd")
            if st.button("🔌 Test MSSQL", use_container_width=True):
                with st.spinner("Connecting..."):
                    try:
                        import pyodbc
                        cs = (f"DRIVER={{ODBC Driver 18 for SQL Server}};"
                              f"SERVER={ms_srv};DATABASE={ms_db};"
                              f"UID={ms_usr};PWD={ms_pwd};"
                              f"TrustServerCertificate=yes;Connection Timeout=5;")
                        conn = pyodbc.connect(cs)
                        cur  = conn.cursor()
                        cur.execute("SELECT @@VERSION")
                        ver  = cur.fetchone()[0].split("\n")[0]
                        conn.close()
                        st.success(f"✅ Connected!\n{ver[:60]}...")
                    except Exception as e:
                        st.error(f"❌ {e}")
            st.markdown("</div>", unsafe_allow_html=True)

        with c_r:
            st.markdown('<div class="card card-gold"><div style="font-weight:700;color:#d4af37;margin-bottom:12px;">🗄️ MySQL – Ubuntu Server</div>', unsafe_allow_html=True)
            my_host = st.text_input("Host",     "10.8.1.88", key="c_my_h")
            my_port = st.number_input("Port",   value=3306,         key="c_my_p")
            my_db   = st.text_input("Database", "gh_promotion",     key="c_my_db")
            my_usr  = st.text_input("User",     "root",             key="c_my_u")
            my_pwd  = st.text_input("Password", type="password",    key="c_my_pw")
            if st.button("🔌 Test MySQL", use_container_width=True):
                with st.spinner("Connecting..."):
                    try:
                        import pymysql
                        conn = pymysql.connect(host=my_host, port=int(my_port),
                                               db=my_db, user=my_usr, password=my_pwd,
                                               connect_timeout=5)
                        cur  = conn.cursor()
                        cur.execute("SELECT VERSION(), DATABASE()")
                        ver, db = cur.fetchone()
                        conn.close()
                        st.success(f"✅ Connected!\nMySQL {ver} | DB: {db}")
                    except Exception as e:
                        st.error(f"❌ {e}")
            st.markdown("</div>", unsafe_allow_html=True)

        # Crontab info
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
        <div class="card">
            <div style="font-weight:700;color:#d4af37;margin-bottom:12px;">⏰ Crontab Schedule</div>
            <pre style="color:#3fb950;font-size:12px;margin:0;">
# ทุก 15 นาที (transaction)
*/15 * * * * cd /opt/promotion_dashboard/etl && /opt/promotion_dashboard/venv/bin/python3 etl_sync.py >> /var/log/gh_etl.log 2>&1

# ทุกคืน 2:00 AM (full sync รวม master)
0 2 * * * cd /opt/promotion_dashboard/etl && /opt/promotion_dashboard/venv/bin/python3 etl_sync.py --full >> /var/log/gh_etl.log 2>&1</pre>
        </div>
        """, unsafe_allow_html=True)

    # ════════════════════════════
    # TAB 5 – MENU ENGINEERING
    # ════════════════════════════
    with tab_me:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:4px;">🧑‍🏫 Build Menu Engineering Summary</div>', unsafe_allow_html=True)
        st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:16px;">คำนวณ Star / Puzzle / Plowhorse / Dog แล้วบันทึกลง menu_engineering_summary</div>', unsafe_allow_html=True)

        me_period = st.radio("📅 เลือกช่วงเวลา", ["30 วัน", "60 วัน", "90 วัน", "ทั้งหมด (30+60+90)"],
                             horizontal=True, key="me_period")

        days_map = {
            "30 วัน":            ["30"],
            "60 วัน":            ["60"],
            "90 วัน":            ["90"],
            "ทั้งหมด (30+60+90)": ["30", "60", "90"],
        }
        me_days = days_map[me_period]

        st.markdown(f"<div style='font-size:12px;color:#8b949e;'>จะรัน: `--menu-engineering --me-days {' '.join(me_days)}`</div>", unsafe_allow_html=True)

        if st.button("🚀 รัน Build Menu Engineering", use_container_width=True, key="me_run"):
            args_extra = ["--menu-engineering", "--me-days"] + me_days
            lines = run_etl(args_extra)
            if lines:
                ok_count  = sum(1 for l in lines if "✔" in l or "items" in l)
                err_count = sum(1 for l in lines if "✘" in l or "Error" in l)
                if err_count:
                    st.error(f"❌ มี {err_count} error")
                else:
                    st.success(f"✅ สำเร็จ — {ok_count} periods")
                with st.expander("📋 Log"):
                    st.code("\n".join(lines), language="text")

    # ════════════════════════════
    # TAB 6 – RECIPE & INGREDIENT
    # ════════════════════════════
    with tab_recipe:
        st.markdown('<div style="font-size:18px;font-weight:700;color:#d4af37;margin-bottom:4px;">🥩 Recipe & Ingredient</div>', unsafe_allow_html=True)
        st.markdown('<div style="font-size:12px;color:#8b949e;margin-bottom:16px;">ข้อมูลมาจาก MSSQL: <b>ItemMast</b> (NC*) → วัตถุดิบ &nbsp;|&nbsp; <b>IRecDet</b> → Recipe</div>', unsafe_allow_html=True)

        # ── Sync button ──
        c1, c2 = st.columns([2,3])
        with c1:
            if st.button("🔄 Sync Recipe & Ingredient จาก MSSQL", use_container_width=True, key="sync_recipe_btn"):
                lines = run_etl(["--sync-recipe"])
                if lines:
                    err_lines = [l for l in lines if "✘" in l or "failed" in l.lower()]
                    if err_lines:
                        st.error(f"❌ {err_lines[0]}")
                    else:
                        st.success("✅ Sync สำเร็จ")
                    with st.expander("📋 Log"):
                        st.code("\n".join(lines), language="text")
        with c2:
            st.markdown('<div style="font-size:12px;color:#8b949e;padding-top:10px;">Cron ทุกคืน: <code>python etl_sync.py --sync-recipe</code></div>', unsafe_allow_html=True)

        st.markdown("<hr style='border-color:#30363d;margin:12px 0;'>", unsafe_allow_html=True)

        rtab1, rtab2, rtab3 = st.tabs(["🧂 วัตถุดิบ", "📋 Recipe", "📦 Stock"])

        import pymysql
        MYSQL_CFG_ETL = gh_env.mysql_cfg("root")

        # ── วัตถุดิบ ──
        with rtab1:
            if st.button("📋 โหลดรายการวัตถุดิบ", key="load_ing", use_container_width=True):
                try:
                    conn = pymysql.connect(**MYSQL_CFG_ETL)
                    df_ing = pd.read_sql("""
                        SELECT i.item_code, i.item_name, i.use_unit,
                               i.stnd_cost, i.dep_code,
                               COALESCE(s.current_stock,0) as current_stock,
                               COALESCE(s.min_stock,0) as min_stock
                        FROM ingredient_tb i
                        LEFT JOIN ingredient_stock_tb s ON i.item_code = s.item_code
                        ORDER BY i.item_code
                    """, conn)
                    conn.close()
                    st.markdown(f'<div style="font-size:12px;color:#8b949e;">พบ {len(df_ing)} รายการ</div>', unsafe_allow_html=True)
                    st.dataframe(df_ing, use_container_width=True, height=400)
                except Exception as e:
                    st.warning(f"ยังไม่มีข้อมูล — กด Sync ก่อน ({e})")

        # ── Recipe ──
        with rtab2:
            search = st.text_input("🔍 ค้นหา RecCode", key="rec_search", placeholder="CPB11001...")
            if st.button("📋 โหลด Recipe", key="load_rec", use_container_width=True):
                try:
                    conn = pymysql.connect(**MYSQL_CFG_ETL)
                    df_rec = pd.read_sql("""
                        SELECT r.rec_code, r.item_code as ingredient_code,
                               i.item_name as ingredient_name,
                               r.quantity, r.unit_code
                        FROM recipe_tb r
                        LEFT JOIN ingredient_tb i ON r.item_code = i.item_code
                        ORDER BY r.rec_code
                    """, conn)
                    conn.close()
                    if search:
                        df_rec = df_rec[df_rec["rec_code"].str.contains(search, case=False, na=False)]
                    st.markdown(f'<div style="font-size:12px;color:#8b949e;">พบ {len(df_rec)} rows</div>', unsafe_allow_html=True)
                    st.dataframe(df_rec, use_container_width=True, height=400)
                except Exception as e:
                    st.warning(f"ยังไม่มีข้อมูล — กด Sync ก่อน ({e})")

        # ── Stock ──
        with rtab3:
            st.markdown('<div style="font-size:13px;font-weight:700;color:#d4af37;margin-bottom:8px;">📦 อัพเดท Stock ปัจจุบัน</div>', unsafe_allow_html=True)
            if st.button("📋 โหลดรายการ Stock", key="load_stock", use_container_width=True):
                try:
                    conn = pymysql.connect(**MYSQL_CFG_ETL)
                    df_stock = pd.read_sql("""
                        SELECT i.item_code, i.item_name, i.use_unit,
                               COALESCE(s.current_stock,0) as current_stock,
                               COALESCE(s.min_stock,0) as min_stock
                        FROM ingredient_tb i
                        LEFT JOIN ingredient_stock_tb s ON i.item_code = s.item_code
                        ORDER BY i.item_code
                    """, conn)
                    conn.close()
                    st.session_state["df_stock_data"] = df_stock
                except Exception as e:
                    st.warning(f"ยังไม่มีข้อมูล — กด Sync ก่อน ({e})")

            if "df_stock_data" in st.session_state:
                df_stock = st.session_state["df_stock_data"]
                for _, row in df_stock.iterrows():
                    c1,c2,c3,c4 = st.columns([3,2,2,1])
                    c1.write(f"**{row['item_name']}** ({row['use_unit']})")
                    new_stock = c2.number_input("stock", value=float(row['current_stock']),
                                               min_value=0.0, step=1.0,
                                               key=f"stk_{row['item_code']}",
                                               label_visibility="collapsed")
                    new_min   = c3.number_input("min", value=float(row['min_stock']),
                                               min_value=0.0, step=1.0,
                                               key=f"min_{row['item_code']}",
                                               label_visibility="collapsed")
                    if c4.button("💾", key=f"save_stk_{row['item_code']}"):
                        try:
                            conn = pymysql.connect(**MYSQL_CFG_ETL)
                            cur = conn.cursor()
                            cur.execute("""INSERT INTO ingredient_stock_tb
                                (item_code,current_stock,min_stock)
                                VALUES (%s,%s,%s)
                                ON DUPLICATE KEY UPDATE current_stock=%s, min_stock=%s""",
                                (row['item_code'], new_stock, new_min, new_stock, new_min))
                            conn.commit(); conn.close()
                            del st.session_state["df_stock_data"]
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ {e}")

# ─────────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────────
_page = st.empty()

if not st.session_state.logged_in:
    with _page.container():
        show_login()
    st.stop()

_page.empty()   # ← ล้าง login form ออกก่อนแสดง admin
show_admin()