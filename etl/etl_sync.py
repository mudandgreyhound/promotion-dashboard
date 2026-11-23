#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════╗
║     🐾  GREYHOUND CAFÉ  –  ETL Sync Pipeline  v2.0          ║
║     MSSQL (Data Center)  →→→  MySQL (Ubuntu Server)         ║
║                                                              ║
║     Author  : GH IT Team                                    ║
║     Engine  : ODBC Driver 18 for SQL Server                  ║
╚══════════════════════════════════════════════════════════════╝

Usage:
  python etl_sync.py                            # ดึงข้อมูลเมื่อวานจนถึงวันนี้
  python etl_sync.py --date_from 2026-03-01     # ระบุวันเริ่มต้น
  python etl_sync.py --date_from 2026-03-01 --date_to 2026-03-07
  python etl_sync.py --full                     # รวม Master Tables ด้วย
  python etl_sync.py --shop 0001,0004           # กรอง ShopCode
  python etl_sync.py --full --dry-run           # ทดสอบโดยไม่ write จริง
"""

import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))  # หา gh_env.py ที่โฟลเดอร์แม่
import gh_env   # รหัส DB / API key อยู่ใน /opt/promotion_dashboard/.env (ไม่เข้า git)
import os
import sys
import time
import argparse
import logging
import traceback
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pyodbc
import pymysql
from pymysql.cursors import DictCursor

# ══════════════════════════════════════════════
#  ANSI COLOR CODES
# ══════════════════════════════════════════════
class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    DIM    = "\033[2m"
    RED    = "\033[91m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    BLUE   = "\033[94m"
    CYAN   = "\033[96m"
    WHITE  = "\033[97m"
    GOLD   = "\033[33m"
    BG_RED = "\033[41m"

def ok(msg):   return f"{C.GREEN}✔{C.RESET}  {msg}"
def err(msg):  return f"{C.RED}✘{C.RESET}  {msg}"
def warn(msg): return f"{C.YELLOW}⚠{C.RESET}  {msg}"
def inf(msg):  return f"{C.CYAN}→{C.RESET}  {msg}"
def bold(msg): return f"{C.BOLD}{msg}{C.RESET}"
def dim(msg):  return f"{C.DIM}{msg}{C.RESET}"
def gold(msg): return f"{C.GOLD}{C.BOLD}{msg}{C.RESET}"

# ══════════════════════════════════════════════
#  LOGGING SETUP
# ══════════════════════════════════════════════
LOG_FILE = os.getenv("ETL_LOG_FILE", "/var/log/gh_etl.log")

class ColorFormatter(logging.Formatter):
    FORMATS = {
        logging.DEBUG:    dim("%(asctime)s") + " DEBUG   %(message)s",
        logging.INFO:     dim("%(asctime)s") + " %(message)s",
        logging.WARNING:  dim("%(asctime)s") + f" {C.YELLOW}WARN{C.RESET}    %(message)s",
        logging.ERROR:    dim("%(asctime)s") + f" {C.RED}ERROR{C.RESET}   %(message)s",
        logging.CRITICAL: dim("%(asctime)s") + f" {C.BG_RED}CRITICAL{C.RESET} %(message)s",
    }
    def format(self, record):
        fmt = self.FORMATS.get(record.levelno, self._fmt)
        formatter = logging.Formatter(fmt, datefmt="%H:%M:%S")
        return formatter.format(record)

def setup_logger() -> logging.Logger:
    logger = logging.getLogger("gh_etl")
    logger.setLevel(logging.INFO)

    # Console handler (color)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(ColorFormatter())
    logger.addHandler(ch)

    # File handler (plain)
    try:
        fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        ))
        logger.addHandler(fh)
    except PermissionError:
        logger.warning(f"Cannot write log to {LOG_FILE} – logging to stdout only")

    return logger

log = setup_logger()

# ══════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════
MSSQL_CONFIG = dict(
    server   = os.getenv("MSSQL_SERVER") or f'{gh_env.need("GH_MSSQL_HOST")},{os.getenv("GH_MSSQL_PORT", "1433")}',
    database = os.getenv("MSSQL_DATABASE") or gh_env.need("GH_MSSQL_DB"),
    uid      = os.getenv("MSSQL_USER")     or gh_env.need("GH_MSSQL_USER"),
    pwd      = os.getenv("MSSQL_PASSWORD") or gh_env.need("GH_MSSQL_PASSWORD"),
    trust    = os.getenv("MSSQL_TRUST",    "yes"),
    timeout  = int(os.getenv("MSSQL_TIMEOUT", "30")),
)

MYSQL_CONFIG = dict(
    host     = os.getenv("MYSQL_HOST")     or gh_env.need("GH_MYSQL_HOST"),
    port     = int(os.getenv("MYSQL_PORT") or os.getenv("GH_MYSQL_PORT", "3306")),
    db       = os.getenv("MYSQL_DATABASE") or gh_env.need("GH_MYSQL_DB"),
    user     = os.getenv("MYSQL_USER")     or gh_env.need("GH_MYSQL_ROOT_USER"),
    password = os.getenv("MYSQL_PASSWORD") or gh_env.need("GH_MYSQL_ROOT_PASSWORD"),
    charset  = "utf8mb4",
    autocommit = False,
    connect_timeout = 10,
)

CHUNK_SIZE  = int(os.getenv("ETL_CHUNK_SIZE",  "1000"))
MAX_RETRIES = int(os.getenv("ETL_MAX_RETRIES", "3"))
RETRY_DELAY = int(os.getenv("ETL_RETRY_DELAY", "5"))

# ══════════════════════════════════════════════
#  TABLE REGISTRY
# ══════════════════════════════════════════════
TRANS_TABLES: Dict[str, dict] = {
    "BillHdrTB": {
        "pk":        ["ShopCode","TerminalNo","BillDate","BillType","BillNo"],
        "date_col":  "BillDate",
        "shop_col":  "ShopCode",
        "desc":      "Bill Header",
    },
    "BillDtlTB": {
        "pk":        ["ShopCode","TerminalNo","BillDate","BillType","BillNo","grdseq","sequence"],
        "date_col":  "BillDate",
        "shop_col":  "ShopCode",
        "desc":      "Bill Detail",
    },
    "ONLBillMultiDiscTB": {
        "pk":        ["ShopCode","TerminalNo","BillDate","BillType","BillNo","DiscCode","RoundSeq","DiscSeq"],
        "date_col":  "BillDate",
        "shop_col":  "ShopCode",
        "desc":      "Promotion Discount",
    },
    "BILLPAYTB": {
        "pk":        ["ShopCode","TerminalNo","BillDate","BillType","BillNo","sequence"],
        "date_col":  "BillDate",
        "shop_col":  "ShopCode",
        "desc":      "Payment",
    },
}

MASTER_TABLES: Dict[str, dict] = {
    "ItmDepTB":  {"pk": ["DepCode"],   "desc": "Department (POS)"},
    "ItmCatTB":  {"pk": ["CatCode"],   "desc": "Category (POS)"},
    "ItmMasTB":  {"pk": ["ItemCode"],  "desc": "Item/Menu (POS)"},
    "ItemCat":   {"pk": ["CatCode"],   "desc": "Category (RM/WIP)"},    # ← เพิ่ม
    "ItemMast":  {"pk": ["ItemCode"],  "desc": "RM/WIP Master"},        # ← เพิ่ม
    "MemTypetb": {"pk": ["MemType"],   "desc": "Member Type"},
    "MemberTB":  {"pk": ["MemCode"],   "desc": "Member"},
    "PosShopTB": {"pk": ["ShopCode"],  "desc": "Shop/Branch"},
}

# ══════════════════════════════════════════════
#  CONNECTIONS
# ══════════════════════════════════════════════
def get_mssql() -> pyodbc.Connection:
    c = MSSQL_CONFIG
    conn_str = (
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={c['server']};DATABASE={c['database']};"
        f"UID={c['uid']};PWD={c['pwd']};"
        f"TrustServerCertificate={c['trust']};"
        f"Connection Timeout={c['timeout']};"
        f"ApplicationIntent=ReadOnly;"
    )
    return pyodbc.connect(conn_str)

def get_mysql() -> pymysql.connections.Connection:
    return pymysql.connect(**MYSQL_CONFIG)

def test_connections() -> Tuple[bool, bool]:
    """ทดสอบ connection ทั้งสองฝั่ง  returns (mssql_ok, mysql_ok)"""
    mssql_ok = mysql_ok = False

    log.info(inf("Testing MSSQL connection..."))
    try:
        conn = get_mssql()
        cur  = conn.cursor()
        cur.execute("SELECT @@VERSION")
        ver  = cur.fetchone()[0].split("\n")[0].strip()
        conn.close()
        log.info(ok(f"MSSQL  {dim(ver)}"))
        mssql_ok = True
    except Exception as e:
        log.error(err(f"MSSQL  {e}"))

    log.info(inf("Testing MySQL connection..."))
    try:
        conn = get_mysql()
        cur  = conn.cursor()
        cur.execute("SELECT VERSION()")
        ver  = cur.fetchone()[0]
        conn.close()
        log.info(ok(f"MySQL  {dim(ver)}"))
        mysql_ok = True
    except Exception as e:
        log.error(err(f"MySQL  {e}"))

    return mssql_ok, mysql_ok

# ══════════════════════════════════════════════
#  SCHEMA HELPER
# ══════════════════════════════════════════════
DTYPE_MAP = {
    "object":           "TEXT",
    "int64":            "BIGINT",
    "int32":            "INT",
    "float64":          "DOUBLE",
    "bool":             "TINYINT(1)",
    "datetime64[ns]":   "DATETIME(3)",
}

def ensure_table(conn, table: str, df: pd.DataFrame, pk=None) -> None:
    """สร้าง table ใน MySQL อัตโนมัติถ้ายังไม่มี — พร้อม PRIMARY KEY จาก registry
       (ถ้าตารางมีอยู่แล้วจะข้าม — การล้างซ้ำตารางเก่าทำผ่าน full-refresh ใน sync_table)"""
    cur = conn.cursor()
    cur.execute(
        "SELECT COUNT(*) FROM information_schema.TABLES "
        "WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = %s",
        (table,)
    )
    if cur.fetchone()[0]:
        cur.close()
        return  # table มีอยู่แล้ว

    pk = [c for c in (pk or []) if c in df.columns]   # เฉพาะ pk ที่มีใน df จริง
    pkset = set(pk)

    cols = []
    for col, dtype in df.dtypes.items():
        sql_type = DTYPE_MAP.get(str(dtype), "TEXT")
        # ปรับ VARCHAR ให้เหมาะสมกับ column ที่รู้จัก
        if str(dtype) == "object":
            max_len = df[col].dropna().astype(str).str.len().max() if not df[col].empty else 50
            max_len = min(int(max_len * 1.5) + 10, 1024) if max_len else 255
            sql_type = f"VARCHAR({max_len})" if max_len <= 255 else "TEXT"
        if col in pkset:
            # pk column ต้องเป็น key ได้ (ไม่เป็น TEXT/ยาวเกิน) + NOT NULL
            if sql_type == "TEXT" or (sql_type.startswith("VARCHAR") and int(sql_type[8:-1]) > 191):
                sql_type = "VARCHAR(191)"
            cols.append(f"  `{col}` {sql_type} NOT NULL")
        else:
            cols.append(f"  `{col}` {sql_type}")

    ddl = (
        f"CREATE TABLE IF NOT EXISTS `{table}` (\n"
        + ",\n".join(cols)
        + (",\n  PRIMARY KEY (" + ", ".join(f"`{c}`" for c in pk) + ")" if pk else "")
        + f"\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;"
    )
    cur.execute(ddl)
    conn.commit()
    cur.close()
    log.info(inf(f"Auto-created table {bold(table)}" + (f"  {dim('PK: '+', '.join(pk))}" if pk else "")))

# ══════════════════════════════════════════════
#  DATA CLEANING
# ══════════════════════════════════════════════
# columns ที่ต้องเป็น numeric – empty string → NULL
INT_COLS  = {"VoidBill","Authorize","Gender","Flag2","RecStatus","ItemStat",
             "DepCode1","DepCode2","DepCode3","FamilyMem","MemStatus","Running1",
             "Running2","UseQuota","CodeLength","PriceMode","itemstat","MoveStat",
             "Seatpost","ReprintNo","BillStatus","XRPTrun","RunNo1","RunNo2","ShopFlag",
             "FreeItem","OverPrc","WeightS","ItemDisp","Subset","Color","FNColor","ItemFlag"}

# columns ที่ MySQL define เป็น TINYINT แต่จริงๆ ค่าอาจเกิน 127
# → map ให้เป็น SMALLINT ใน ensure_table และ clamp ค่าใน clean_df
FORCE_SMALLINT = {"VoidBill","Authorize","CodeLength","ReprintNo","Seatpost",
                  "FreeItem","OverPrc","WeightS","ItemDisp","Subset",
                  "Color","FNColor","ItemFlag","ShopFlag","XRPTrun"}

# columns ที่อาจยาวเกิน – truncate หรือใช้ TEXT
LONG_COLS = {"CardNo": 50, "Remark": 255, "Apprval": 50,
             "RsonCode": 50,
             "Address1": 500, "Address2": 500, "Child2": 64}

def clean_df(table: str, df: pd.DataFrame) -> pd.DataFrame:
    """Clean DataFrame ก่อน upsert:
       - empty string / non-numeric → NaN สำหรับ int columns
       - clamp SMALLINT columns ให้ไม่เกิน range
       - truncate string columns ที่ยาวเกิน
    """
    df = df.copy()

    for col in df.columns:
        # 1. empty string / non-numeric → NaN for numeric columns
        if col in INT_COLS:
            if df[col].dtype == object:
                df[col] = df[col].replace("", None)
            df[col] = pd.to_numeric(df[col], errors="coerce")

        # 2. truncate long strings
        if col in LONG_COLS and df[col].dtype == object:
            max_len = LONG_COLS[col]
            df[col] = df[col].where(df[col].isna(),
                                    df[col].astype(str).str[:max_len])

    return df

# ══════════════════════════════════════════════
#  UPSERT
# ══════════════════════════════════════════════
def upsert_df(conn, table: str, df: pd.DataFrame, dry_run: bool = False) -> int:
    if df.empty:
        return 0

    cols = [f"`{c}`" for c in df.columns]
    ph   = ", ".join(["%s"] * len(df.columns))
    upd  = ", ".join([f"{c}=VALUES({c})" for c in cols])
    sql  = (
        f"INSERT INTO `{table}` ({', '.join(cols)}) "
        f"VALUES ({ph}) "
        f"ON DUPLICATE KEY UPDATE {upd}"
    )

    if dry_run:
        return len(df)

    def safe_val(v):
        if pd.isna(v) if not isinstance(v, (list, dict)) else False:
            return None
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        # convert numpy int/float to python native
        if hasattr(v, 'item'):
            return v.item()
        return v

    cur = conn.cursor()
    # ปิด FK check ระหว่าง upsert – ป้องกัน FK constraint fail
    cur.execute("SET foreign_key_checks = 0")
    total = 0
    try:
        for start in range(0, len(df), CHUNK_SIZE):
            chunk = df.iloc[start : start + CHUNK_SIZE]
            rows  = [tuple(safe_val(v) for v in r) for r in chunk.itertuples(index=False)]
            cur.executemany(sql, rows)
            conn.commit()
            total += len(rows)
            pct  = min(100, int((start + CHUNK_SIZE) / len(df) * 100))
            bar  = "█" * (pct // 5) + "░" * (20 - pct // 5)
            print(f"\r    {C.GOLD}{bar}{C.RESET} {pct:3d}%  {total:,}/{len(df):,}", end="", flush=True)
    finally:
        cur.execute("SET foreign_key_checks = 1")
        conn.commit()

    print()
    cur.close()
    return total

# ══════════════════════════════════════════════
#  SYNC ONE TABLE  (with retry)
# ══════════════════════════════════════════════
def build_query(table: str, cfg: dict,
                date_from: Optional[str], date_to: Optional[str],
                shop_codes: Optional[List[str]]) -> str:
    query = f"SELECT * FROM dbo.[{table}] WITH (NOLOCK)"
    conds = []
    if date_from and "date_col" in cfg:
        conds.append(f"{cfg['date_col']} >= '{date_from} 00:00:00'")
    if date_to and "date_col" in cfg:
        conds.append(f"{cfg['date_col']} <= '{date_to} 23:59:59'")
    if shop_codes and "shop_col" in cfg:
        in_list = ", ".join(f"'{s}'" for s in shop_codes)
        conds.append(f"{cfg['shop_col']} IN ({in_list})")
    if conds:
        query += " WHERE " + " AND ".join(conds)
    return query

def sync_table(
    mssql_conn, mysql_conn,
    table: str, cfg: dict,
    date_from=None, date_to=None,
    shop_codes=None, dry_run=False, full_refresh=False
) -> dict:

    desc    = cfg.get("desc", table)
    t_start = time.perf_counter()
    query   = build_query(table, cfg, date_from, date_to, shop_codes)

    log.info(f"  {C.CYAN}┌─{C.RESET} {bold(table)} {dim(f'({desc})')}")
    log.info(f"  {C.CYAN}│{C.RESET}  {dim(query[:120] + ('...' if len(query) > 120 else ''))}")

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            df = pd.read_sql(query, mssql_conn)
            log.info(f"  {C.CYAN}│{C.RESET}  Fetched  {C.BOLD}{len(df):,}{C.RESET} rows from MSSQL")

            if df.empty:
                log.info(f"  {C.CYAN}└─{C.RESET} {C.YELLOW}SKIP{C.RESET} (no data)")
                return _result(table, 0, "skip", time.perf_counter() - t_start)

            # clean column names
            df.columns = [c.strip() for c in df.columns]

            # 🧹 clean data – empty string → NULL, truncate long strings
            df = clean_df(table, df)

            ensure_table(mysql_conn, table, df, pk=cfg.get("pk"))

            # full-refresh: ล้างข้อมูลเก่าทั้งหมดก่อน insert (สำหรับ master tables)
            # → ล้างแถวซ้ำที่มีอยู่ + ราคา/ข้อมูลจาก CM POS เข้าทับตรงทุกครั้ง
            if full_refresh and not dry_run:
                fc = mysql_conn.cursor()
                fc.execute("SET foreign_key_checks = 0")
                fc.execute(f"DELETE FROM `{table}`")
                fc.execute("SET foreign_key_checks = 1")
                mysql_conn.commit()
                fc.close()
                log.info(f"  {C.CYAN}│{C.RESET}  {dim('full-refresh: cleared old rows')}")

            rows = upsert_df(mysql_conn, table, df, dry_run=dry_run)

            elapsed = time.perf_counter() - t_start
            rate    = rows / elapsed if elapsed else 0
            tag     = f"{C.DIM}dry-run{C.RESET}" if dry_run else ok("OK")
            log.info(
                f"  {C.CYAN}└─{C.RESET} {tag}  "
                f"{C.BOLD}{rows:,}{C.RESET} rows  "
                f"{dim(f'{elapsed:.2f}s  ~{rate:,.0f} rows/s')}"
            )
            return _result(table, rows, "dry-run" if dry_run else "success", elapsed)

        except Exception as e:
            log.warning(warn(f"Attempt {attempt}/{MAX_RETRIES} failed: {e}"))
            if attempt < MAX_RETRIES:
                log.info(inf(f"Retrying in {RETRY_DELAY}s..."))
                time.sleep(RETRY_DELAY)
            else:
                log.error(err(f"FAILED after {MAX_RETRIES} attempts"))
                log.debug(traceback.format_exc())
                return _result(table, 0, f"error: {e}", time.perf_counter() - t_start)

def _result(table, rows, status, duration) -> dict:
    return {"table": table, "rows": rows, "status": status, "duration": duration}

# ══════════════════════════════════════════════
#  PRINT HELPERS
# ══════════════════════════════════════════════
def print_banner():
    print(f"""
{C.GOLD}{C.BOLD}
  ╔══════════════════════════════════════════════════════╗
  ║   🐾  GREYHOUND CAFÉ  –  ETL Sync Pipeline  v2.0   ║
  ║       MSSQL  →→→  MySQL                             ║
  ╚══════════════════════════════════════════════════════╝
{C.RESET}""")

def print_config(args):
    tags = []
    if args.full:     tags.append(f"{C.BLUE}FULL{C.RESET}")
    if args.dry_run:  tags.append(f"{C.YELLOW}DRY-RUN{C.RESET}")
    tag_str = "  " + "  ".join(tags) if tags else ""

    print(f"  {C.BOLD}MSSQL{C.RESET}   {MSSQL_CONFIG['server']} / {MSSQL_CONFIG['database']}")
    print(f"  {C.BOLD}MySQL{C.RESET}   {MYSQL_CONFIG['host']}:{MYSQL_CONFIG['port']} / {MYSQL_CONFIG['db']}")
    print(f"  {C.BOLD}Period{C.RESET}  {args.date_from}  →  {args.date_to}")
    print(f"  {C.BOLD}Shops{C.RESET}   {args.shop or 'ALL'}{tag_str}")
    print(f"  {dim('─' * 52)}")

def print_summary(results: list, total_elapsed: float):
    print(f"\n  {dim('═' * 52)}")
    print(f"  {bold('SYNC SUMMARY')}\n")

    total_rows = errors = skips = 0
    for r in results:
        if r["status"] == "success":
            icon  = f"{C.GREEN}✔{C.RESET}"
            color = C.RESET
        elif r["status"] == "skip":
            icon  = f"{C.DIM}–{C.RESET}"
            color = C.DIM
            skips += 1
        elif r["status"] == "dry-run":
            icon  = f"{C.YELLOW}◉{C.RESET}"
            color = C.YELLOW
        else:
            icon  = f"{C.RED}✘{C.RESET}"
            color = C.RED
            errors += 1

        total_rows += r["rows"]
        print(
            f"  {icon}  {color}{r['table']:<28}{C.RESET}"
            f"  {C.BOLD}{r['rows']:>8,}{C.RESET} rows"
            f"  {dim('{:.2f}s'.format(r['duration']))}"
        )

    print(f"\n  {dim('─' * 52)}")
    status_str = (
        f"{C.GREEN}SUCCESS{C.RESET}" if not errors
        else f"{C.RED}FAILED  ({errors} error(s)){C.RESET}"
    )
    print(f"  Status   {status_str}")
    print(f"  Rows     {C.BOLD}{total_rows:,}{C.RESET}")
    print(f"  Skipped  {skips}")
    print(f"  Time     {total_elapsed:.2f}s")
    print(f"  {dim('═' * 52)}\n")
    return errors

# ══════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════
def parse_args():
    p = argparse.ArgumentParser(
        description="GH Promotion ETL – MSSQL → MySQL",
        formatter_class=argparse.RawTextHelpFormatter
    )
    p.add_argument("--date_from", default=(datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d"),
                   help="วันเริ่มต้น  (default: เมื่อวาน)")
    p.add_argument("--date_to",   default=datetime.today().strftime("%Y-%m-%d"),
                   help="วันสิ้นสุด  (default: วันนี้)")
    p.add_argument("--shop",      default="",
                   help="ShopCode คั่นด้วย comma  เช่น 0001,0004  (default: ทั้งหมด)")
    p.add_argument("--full",      action="store_true",
                   help="Sync Master Tables ด้วย")
    p.add_argument("--dry-run",   action="store_true", dest="dry_run",
                   help="ทดสอบโดยไม่ write ข้อมูลจริง")
    p.add_argument("--test-conn", action="store_true", dest="test_conn",
                   help="ทดสอบ connection แล้วออก")
    p.add_argument("--verbose",   action="store_true",
                   help="แสดง debug log")
    p.add_argument("--menu-engineering", action="store_true", dest="menu_engineering",
                   help="Build Menu Engineering summary (30/60/90 วัน)")
    p.add_argument("--me-days",   nargs="+", type=int, default=[30, 60, 90],
                   dest="me_days",
                   help="ช่วงวันสำหรับ Menu Engineering  (default: 30 60 90)")
    p.add_argument("--sync-recipe", action="store_true", dest="sync_recipe",
                   help="Sync IRecDet + ItemMast(NC) จาก MSSQL → MySQL")
    return p.parse_args()


# ══════════════════════════════════════════════
#  SYNC RECIPE & INGREDIENT
# ══════════════════════════════════════════════
def run_sync_recipe(mssql_conn, mysql_conn, dry_run: bool = False):
    """Sync IRecDet และ ItemMast(NC) จาก MSSQL → MySQL
       คืน list ของ _result เพื่อให้ print_summary แสดงรวมในตารางสรุป"""
    _rs = []          # เก็บผลแต่ละตาราง
    _st = time.perf_counter()
    log.info(inf("Syncing Ingredient (ItemMast NC*)..."))
    try:
        cur = mssql_conn.cursor()
        cur.execute("""
            SELECT ItemCode, ItemDes1, DepCode, UseUnit, StndCost
            FROM ItemMast
            WHERE ItemCode LIKE 'NC%'
              AND (ItemFlag = 0 OR ItemFlag IS NULL)
        """)
        ingredients = cur.fetchall()
        log.info(ok(f"  {len(ingredients)} ingredients from MSSQL"))

        if not dry_run:
            mc = mysql_conn.cursor()
            mc.execute("""CREATE TABLE IF NOT EXISTS ingredient_tb (
                item_code  VARCHAR(20) NOT NULL PRIMARY KEY,
                item_name  VARCHAR(200),
                dep_code   VARCHAR(20),
                use_unit   VARCHAR(20),
                stnd_cost  DECIMAL(12,4) DEFAULT 0,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )""")
            mc.execute("DELETE FROM ingredient_tb")
            if ingredients:
                mc.executemany("""INSERT INTO ingredient_tb
                    (item_code,item_name,dep_code,use_unit,stnd_cost)
                    VALUES (%s,%s,%s,%s,%s)""",
                    [(r.ItemCode, r.ItemDes1, r.DepCode,
                      r.UseUnit, float(r.StndCost or 0)) for r in ingredients])
            mysql_conn.commit()
            log.info(ok(f"  ingredient_tb: {len(ingredients)} rows upserted"))
        _rs.append(_result("ingredient_tb", len(ingredients),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  ingredient sync failed: {e}"))
        _rs.append(_result("ingredient_tb", 0, "error", time.perf_counter()-_st))

    _st = time.perf_counter()
    log.info(inf("Syncing Recipe (IRecDet)..."))
    try:
        cur.execute("""
            SELECT LTRIM(RTRIM(RecCode)) AS RecCode,
               LTRIM(RTRIM(ItemCode)) AS ItemCode,
               Quantity, LTRIM(RTRIM(UnitCode)) AS UnitCode
        FROM IRecDet
        WHERE ItemType = 'I' AND PriceMode = 0
    """)
        recipes = cur.fetchall()
        log.info(ok(f"  {len(recipes)} recipe rows from MSSQL"))

        if not dry_run:
            mc.execute("""CREATE TABLE IF NOT EXISTS recipe_tb (
                rec_code  VARCHAR(20) NOT NULL,
                item_code VARCHAR(20) NOT NULL,
                quantity  DECIMAL(12,4) DEFAULT 0,
                unit_code VARCHAR(20),
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                PRIMARY KEY (rec_code, item_code)
            )""")
            mc.execute("""CREATE TABLE IF NOT EXISTS ingredient_stock_tb (
                item_code     VARCHAR(20) NOT NULL PRIMARY KEY,
                current_stock DECIMAL(12,4) DEFAULT 0,
                min_stock     DECIMAL(12,4) DEFAULT 0,
                updated_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )""")
            if recipes:
                # deduplicate ใน Python ก่อน (กัน duplicate key จาก MSSQL)
                seen = set()
                unique_recipes = []
                for r in recipes:
                    key = (r.RecCode, r.ItemCode)
                    if key not in seen:
                        seen.add(key)
                        unique_recipes.append((
                            r.RecCode, r.ItemCode,
                            float(r.Quantity or 0), r.UnitCode
                        ))
                mc.execute("TRUNCATE TABLE recipe_tb")
                mc.executemany("""INSERT INTO recipe_tb
                    (rec_code,item_code,quantity,unit_code)
                    VALUES (%s,%s,%s,%s)""", unique_recipes)
            mysql_conn.commit()
            log.info(ok(f"  recipe_tb: {len(unique_recipes)} rows inserted"))
        _rs.append(_result("recipe_tb", len(unique_recipes),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  recipe sync failed: {e}"))
# ── Sync PLU (ItmMasTB) → plu_tb  สำหรับแยก POS/WIP + ราคาขาย ──
    _st = time.perf_counter()
    log.info(inf("Syncing PLU (ItmMasTB)..."))
    try:
        cur.execute("""
            SELECT LTRIM(RTRIM(ItemCode)) AS ItemCode,
               ItemDes1, OriginPrc, DepCode0
        FROM ItmMasTB
    """)
        plus = cur.fetchall()
        log.info(ok(f"  {len(plus)} PLU from MSSQL"))

        if not dry_run:
            mc.execute("""CREATE TABLE IF NOT EXISTS plu_tb (
                item_code  VARCHAR(20) NOT NULL PRIMARY KEY,
                item_name  VARCHAR(200),
                origin_prc DECIMAL(12,4) DEFAULT 0,
                dep_code   VARCHAR(20),
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )""")
            mc.execute("DELETE FROM plu_tb")
            if plus:
                mc.executemany("""INSERT INTO plu_tb
                    (item_code,item_name,origin_prc,dep_code)
                    VALUES (%s,%s,%s,%s)""",
                    [(p.ItemCode, p.ItemDes1,
                      float(p.OriginPrc or 0), p.DepCode0) for p in plus])
            mysql_conn.commit()
            log.info(ok(f"  plu_tb: {len(plus)} rows"))
        _rs.append(_result("plu_tb", len(plus),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  PLU sync failed: {e}"))
        

    # ── Sync Unit Conversion (IgrpunitTB) → unit_conv_tb ──
    _st = time.perf_counter()
    log.info(inf("Syncing Unit Conversion (IgrpunitTB)..."))
    try:
        cur.execute("""SELECT LTRIM(RTRIM(ItemCode)) AS ItemCode,
                              LTRIM(RTRIM(WithUnit)) AS WithUnit,
                              LTRIM(RTRIM(FrmUnit))  AS FrmUnit,
                              Multiplier, ConvFact
                       FROM IgrpunitTB""")
        units = cur.fetchall()
        log.info(ok(f"  {len(units)} unit-conv rows from MSSQL"))
        rows = []
        if not dry_run:
            # PK เดิม (item,with) ทำให้แถวที่ with ซ้ำแต่ from ต่างหาย → เปลี่ยนเป็น (item,from,with)
            mc.execute("DROP TABLE IF EXISTS unit_conv_tb")
            mc.execute("""CREATE TABLE unit_conv_tb (
                item_code  VARCHAR(20) NOT NULL,
                with_unit  VARCHAR(20) NOT NULL,
                from_unit  VARCHAR(20) NOT NULL DEFAULT '',
                multiplier DECIMAL(18,6)  DEFAULT 1,
                conv_fact  DECIMAL(18,10) DEFAULT 1,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                PRIMARY KEY (item_code, from_unit, with_unit)
            )""")
            if units:
                seen=set()
                for u in units:
                    fu = (u.FrmUnit or "").strip()
                    wu = (u.WithUnit or "").strip()
                    key=(u.ItemCode, fu, wu)          # key ครบ 3 คอลัมน์ (เดิมใช้ 2 → KG หาย)
                    if key in seen: continue
                    seen.add(key)
                    rows.append((u.ItemCode, wu, fu,
                                 float(u.Multiplier or 1), float(u.ConvFact or 1)))
                mc.executemany("""INSERT INTO unit_conv_tb
                    (item_code,with_unit,from_unit,multiplier,conv_fact)
                    VALUES (%s,%s,%s,%s,%s)""", rows)
            mysql_conn.commit()
            log.info(ok(f"  unit_conv_tb: {len(rows)} rows"))
        _rs.append(_result("unit_conv_tb", len(rows) if not dry_run else len(units),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  unit conversion sync failed: {e}"))
    # ── Sync WIP Recipe (IWipTB + IWip1Det) → wip_recipe_tb ──
    _st = time.perf_counter()
    log.info(inf("Syncing WIP Recipe (IWipTB)..."))
    try:
        cur.execute("""
            SELECT LTRIM(RTRIM(w.WIPCode))   AS WIPCode,
                   w.WIPName1,
                   LTRIM(RTRIM(d.ItemCode))  AS ItemCode,
                   d.Quantity,
                   LTRIM(RTRIM(d.UnitCode))  AS UnitCode
            FROM IWipTB w
            INNER JOIN IWip1Det d ON LTRIM(RTRIM(w.WIPCode)) = LTRIM(RTRIM(d.WIPCode))
        """)
        wrows = cur.fetchall()
        log.info(ok(f"  {len(wrows)} WIP recipe rows from MSSQL"))
        if not dry_run:
            mc.execute("""CREATE TABLE IF NOT EXISTS wip_recipe_tb (
                wip_code   VARCHAR(20) NOT NULL,
                wip_name   VARCHAR(200),
                item_code  VARCHAR(20) NOT NULL,
                quantity   DECIMAL(14,4) DEFAULT 0,
                unit_code  VARCHAR(20),
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                PRIMARY KEY (wip_code, item_code)
            )""")
            mc.execute("TRUNCATE TABLE wip_recipe_tb")
            seen=set(); rows=[]
            for r in wrows:
                key=(r.WIPCode, r.ItemCode)
                if key in seen: continue
                seen.add(key)
                rows.append((r.WIPCode, r.WIPName1, r.ItemCode,
                             float(r.Quantity or 0), r.UnitCode))
            if rows:
                mc.executemany("""INSERT INTO wip_recipe_tb
                    (wip_code,wip_name,item_code,quantity,unit_code)
                    VALUES (%s,%s,%s,%s,%s)""", rows)
            mysql_conn.commit()
            log.info(ok(f"  wip_recipe_tb: {len(rows)} rows"))
        _rs.append(_result("wip_recipe_tb", len(rows),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  WIP recipe sync failed: {e}"))
# ── Sync WIP Output/Yield (Iwip2det) → wip_output_tb ──
    _st = time.perf_counter()
    log.info(inf("Syncing WIP Output (Iwip2det)..."))
    try:
        cur.execute("""
            SELECT LTRIM(RTRIM(WIPCode))  AS WIPCode,
                   LTRIM(RTRIM(ItemCode)) AS ItemCode,
                   Quantity,
                   LTRIM(RTRIM(UnitCode)) AS UnitCode,
                   CostRatio
            FROM Iwip2det
        """)
        orows = cur.fetchall()
        log.info(ok(f"  {len(orows)} WIP output rows from MSSQL"))
        if not dry_run:
            mc.execute("""CREATE TABLE IF NOT EXISTS wip_output_tb (
                wip_code   VARCHAR(20) NOT NULL,
                item_code  VARCHAR(20) NOT NULL,
                quantity   DECIMAL(14,4) DEFAULT 0,
                unit_code  VARCHAR(20),
                cost_ratio DECIMAL(9,4)  DEFAULT 100,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                PRIMARY KEY (wip_code, item_code)
            )""")
            mc.execute("TRUNCATE TABLE wip_output_tb")
            seen=set(); rows=[]
            for r in orows:
                key=(r.WIPCode, r.ItemCode)
                if key in seen: continue
                seen.add(key)
                rows.append((r.WIPCode, r.ItemCode,
                             float(r.Quantity or 0), r.UnitCode,
                             float(r.CostRatio or 100)))
            if rows:
                mc.executemany("""INSERT INTO wip_output_tb
                    (wip_code,item_code,quantity,unit_code,cost_ratio)
                    VALUES (%s,%s,%s,%s,%s)""", rows)
            mysql_conn.commit()
            log.info(ok(f"  wip_output_tb: {len(rows)} rows"))
        _rs.append(_result("wip_output_tb", len(orows),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  WIP output sync failed: {e}"))

    # ── Sync Category (ItmCatTB) → category_tb ──
    _st = time.perf_counter()
    log.info(inf("Syncing Category (ItmCatTB)..."))
    try:
        cur.execute("""
            SELECT LTRIM(RTRIM(CatCode)) AS CatCode, CatDes1, CatDes2
            FROM ItmCatTB
        """)
        cats = cur.fetchall()
        log.info(ok(f"  {len(cats)} categories from MSSQL"))
        if not dry_run:
            mc.execute("""CREATE TABLE IF NOT EXISTS category_tb (
                cat_code   VARCHAR(20) NOT NULL PRIMARY KEY,
                cat_name   VARCHAR(200),
                cat_name_th VARCHAR(200),
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
            )""")
            mc.execute("DELETE FROM category_tb")
            if cats:
                mc.executemany("""INSERT INTO category_tb (cat_code,cat_name,cat_name_th)
                    VALUES (%s,%s,%s)""",
                    [(c.CatCode, c.CatDes1, c.CatDes2) for c in cats])
            mysql_conn.commit()
            log.info(ok(f"  category_tb: {len(cats)} rows"))
        _rs.append(_result("category_tb", len(cats),
                           "dry-run" if dry_run else "success", time.perf_counter()-_st))
    except Exception as e:
        log.error(err(f"  category sync failed: {e}"))
        _rs.append(_result("category_tb", 0, "error", time.perf_counter()-_st))

    log.info(ok("Recipe & Ingredient sync complete"))
    return _rs


# ══════════════════════════════════════════════
#  MENU ENGINEERING
# ══════════════════════════════════════════════
def run_menu_engineering(mssql_conn, mysql_conn, days_list: list, dry_run: bool = False):
    """Build menu_engineering_summary สำหรับแต่ละ period"""
    from datetime import date

    SQL_ME = """
    WITH sales_base AS (
        SELECT d.ItemCode,
               SUM(d.ActQty)              AS TotalQty,
               SUM(d.ActQty * d.Sellprice) AS TotalSales
        FROM BillHdrTB h
        INNER JOIN BillDtlTB d ON h.ShopCode=d.ShopCode AND h.BillNo=d.BillNo
        WHERE h.BillDate BETWEEN ? AND ?
          AND h.BillType = 'INV'
          AND h.PriceMode = '0'
        GROUP BY d.ItemCode
    ),
    final_calc AS (
        SELECT s.ItemCode,
               ISNULL(i.ItemDes1, s.ItemCode+' - No Master') AS ItemDes1,
               s.TotalQty, s.TotalSales,
               (s.TotalQty * ISNULL(i.StndCost,0))                              AS TotalCost,
               (s.TotalSales-(s.TotalQty*ISNULL(i.StndCost,0)))                 AS TotalProfit,
               (s.TotalSales/NULLIF(s.TotalQty,0))                              AS AvgPrice,
               ((s.TotalSales-(s.TotalQty*ISNULL(i.StndCost,0)))/NULLIF(s.TotalQty,0))   AS CM_per_Unit,
               ((s.TotalSales-(s.TotalQty*ISNULL(i.StndCost,0)))/NULLIF(s.TotalSales,0)) AS CM_Ratio
        FROM sales_base s
        LEFT JOIN ItmMasTB i ON s.ItemCode=i.ItemCode
    )
    SELECT *,
           TotalQty*1.0/SUM(TotalQty) OVER()  AS Popularity_Ratio,
           AVG(TotalQty)     OVER()            AS AvgQty,
           AVG(CM_per_Unit)  OVER()            AS AvgCM,
           CASE
               WHEN TotalQty >= AVG(TotalQty) OVER() AND CM_per_Unit >= AVG(CM_per_Unit) OVER() THEN 'Stars'
               WHEN TotalQty >= AVG(TotalQty) OVER() AND CM_per_Unit <  AVG(CM_per_Unit) OVER() THEN 'Plow Horses'
               WHEN TotalQty <  AVG(TotalQty) OVER() AND CM_per_Unit >= AVG(CM_per_Unit) OVER() THEN 'Puzzles'
               ELSE 'Dogs'
           END AS MenuCategory
    FROM final_calc
    ORDER BY TotalSales DESC
    """

    results = []
    end_date = date.today()

    for p_days in days_list:
        start_date = end_date - timedelta(days=p_days - 1)
        log.info(inf(f"Menu Engineering — {p_days} วัน  ({start_date} → {end_date})"))
        try:
            cur = mssql_conn.cursor()
            cur.execute(SQL_ME, (start_date, end_date))
            rows = cur.fetchall()

            if not rows:
                log.warning(warn(f"  ไม่มีข้อมูล period {p_days} วัน"))
                results.append((p_days, 0, False))
                continue

            if dry_run:
                log.info(ok(f"  [dry-run] {len(rows)} items — period {p_days} วัน"))
                results.append((p_days, len(rows), True))
                continue

            mc = mysql_conn.cursor()
            mc.execute("SET foreign_key_checks=0")
            mc.execute("DELETE FROM menu_engineering_summary WHERE date_from=%s AND date_to=%s",
                       (start_date, end_date))

            batch = [(p_days, start_date, end_date,
                      r.ItemCode, r.ItemDes1,
                      float(r.TotalQty or 0), float(r.TotalSales or 0),
                      float(r.TotalCost or 0), float(r.TotalProfit or 0),
                      float(r.AvgPrice or 0), float(r.CM_per_Unit or 0),
                      float(r.CM_Ratio or 0), float(r.Popularity_Ratio or 0),
                      float(r.AvgQty or 0), float(r.AvgCM or 0),
                      r.MenuCategory) for r in rows]

            mc.executemany("""
                INSERT INTO menu_engineering_summary
                (period_days,date_from,date_to,item_code,item_name,
                 total_qty,total_sales,total_cost,total_profit,avg_price,
                 cm_per_unit,cm_ratio,popularity_ratio,avg_qty,avg_cm,menu_category)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, batch)
            mc.execute("SET foreign_key_checks=1")
            mysql_conn.commit()
            log.info(ok(f"  {len(rows)} items → menu_engineering_summary  [{p_days}d]"))
            results.append((p_days, len(rows), True))

        except Exception as e:
            log.error(err(f"  Menu Engineering {p_days}d failed: {e}"))
            results.append((p_days, 0, False))

    # summary
    print()
    log.info(bold("── Menu Engineering Summary ─────────────────────"))
    for p_days, n, success in results:
        status = ok(f"{n} items") if success else err("failed")
        log.info(f"  {p_days:3d} วัน  →  {status}")
    return results

def main():
    args = parse_args()
    if args.verbose:
        log.setLevel(logging.DEBUG)

    print_banner()

    # ── test-only mode ──
    if args.test_conn:
        ms_ok, my_ok = test_connections()
        sys.exit(0 if (ms_ok and my_ok) else 1)

    print_config(args)

    shop_codes = [s.strip() for s in args.shop.split(",") if s.strip()] or None
    run_start  = time.perf_counter()

    # ── Connect ──
    log.info(inf("Connecting to MSSQL..."))
    try:
        mssql_conn = get_mssql()
        log.info(ok("MSSQL connected"))
    except Exception as e:
        log.error(err(f"MSSQL connection failed: {e}"))
        sys.exit(1)

    log.info(inf("Connecting to MySQL..."))
    try:
        mysql_conn = get_mysql()
        log.info(ok("MySQL connected"))
    except Exception as e:
        log.error(err(f"MySQL connection failed: {e}"))
        mssql_conn.close()
        sys.exit(1)

    print()
    results = []

    # ── Transactional tables ──
    log.info(bold("── Transaction Tables ──────────────────────────"))
    for tbl, cfg in TRANS_TABLES.items():
        r = sync_table(
            mssql_conn, mysql_conn, tbl, cfg,
            date_from=args.date_from, date_to=args.date_to,
            shop_codes=shop_codes, dry_run=args.dry_run
        )
        results.append(r)

    # ── Master tables ──
    if args.full:
        print()
        log.info(bold("── Master Tables ───────────────────────────────"))
        for tbl, cfg in MASTER_TABLES.items():
            r = sync_table(
                mssql_conn, mysql_conn, tbl, cfg,
                dry_run=args.dry_run, full_refresh=True
            )
            results.append(r)

    # ── Menu Engineering ──
    if args.menu_engineering:
        print()
        log.info(bold("── Menu Engineering ────────────────────────────"))
        me_results = run_menu_engineering(mssql_conn, mysql_conn,
                             days_list=args.me_days,
                             dry_run=args.dry_run)
        for p_days, n, success in (me_results or []):
            results.append(_result(f"menu_eng_{p_days}d", n,
                                   "success" if success else "error", 0.0))

    # ── Sync Recipe & Ingredient ──
    if args.sync_recipe:
        print()
        log.info(bold("── Sync Recipe & Ingredient ────────────────────"))
        rec_results = run_sync_recipe(mssql_conn, mysql_conn, dry_run=args.dry_run)
        results.extend(rec_results or [])

    mssql_conn.close()
    mysql_conn.close()

    errors = print_summary(results, time.perf_counter() - run_start)
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
    sys.exit(1 if errors else 0)