"""
sale_monitor_etl.py — Greyhound BI : real-time sale monitor
  ETL : CM POS (MSSQL ONL*) --> MySQL sale_monitor_rt / _interval   (cron ทุก 2-5 นาที)
  API : FastAPI /api/overview/sale-monitor --> ป้อน widget หน้า Overview

ติดตั้ง:  pip install pymssql pymysql fastapi uvicorn
รัน API:  uvicorn sale_monitor_etl:app --host 0.0.0.0 --port 8010
cron:
  */3 6-23 * * *  /usr/bin/python3 /opt/promotion_dashboard/sale_monitor_etl.py today
  15  02   * * *  /usr/bin/python3 /opt/promotion_dashboard/sale_monitor_etl.py yesterday   # freeze เมื่อวาน
"""
import sys, datetime as dt
import pymssql, pymysql

import os as _os
sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))  # หา gh_env.py ที่โฟลเดอร์แม่
import gh_env   # รหัส DB / API key อยู่ใน /opt/promotion_dashboard/.env (ไม่เข้า git)
MSSQL = dict(server=gh_env.need("GH_MSSQL_HOST"), user=gh_env.need("GH_MSSQL_USER"),
             password=gh_env.need("GH_MSSQL_PASSWORD"), database=gh_env.need("GH_MSSQL_DB"))
MYSQL = gh_env.mysql_cfg("admin", autocommit=False)
API_KEY = gh_env.need("SALE_MONITOR_API_KEY")   # ใช้ค่าเดียวกับระบบเดิม

# ---- โหลด SQL ทั้ง 2 ตัวจากไฟล์ sale_monitor.sql (block 2 และ 3) หรือฝังตรงนี้ก็ได้ ----
SQL_SHOP = """
SELECT CONVERT(varchar(10), bh.EntDate,120) AS biz_date, bh.ShopCode AS shop_code,
  p.ShopDes1 AS shop_name, brn.brandnam1 AS brand,
  COUNT(CASE WHEN bh.BillType='INV' AND ISNULL(bh.BillStatus,'')='' THEN 1 END) AS bill_count,
  SUM(CASE WHEN bh.BillType='INV' AND ISNULL(bh.BillStatus,'')='' THEN bh.NoPerson ELSE 0 END) AS cust_count,
  SUM(bh.NetBAmt+(bh.BillDisc+bh.MemDisc+bh.Voucher+bh.ItemDisc+bh.CoupDisc)) AS gross_sales,
  SUM(bh.BillDisc+bh.MemDisc+bh.Voucher+bh.ItemDisc+bh.CoupDisc) AS discount,
  SUM(CASE p.VatType WHEN '1' THEN bh.NetBAmt ELSE bh.NetBAmt-(bh.VatAmt1-bh.VatAmt2) END) AS base_sales,
  SUM(bh.VatAmt1-bh.VatAmt2) AS vat,
  SUM(CASE p.VatType WHEN '1' THEN bh.NetBAmt+(bh.VatAmt1-bh.VatAmt2) ELSE bh.NetBAmt END) AS net_sales,
  SUM(bh.OthChrg+bh.VatAmt2) AS service_charge,
  SUM(CASE p.VatType WHEN '1' THEN (bh.NetBAmt+bh.OthChrg+bh.VatAmt1) ELSE (bh.NetBAmt+bh.OthChrg+bh.VatAmt2) END) AS net_total,
  CONVERT(varchar(8), MAX(bh.BillTime),108) AS last_bill_time
FROM ONLBillHdrTB bh WITH (NOLOCK)
INNER JOIN PosShopTB p WITH (NOLOCK) ON p.ShopCode=bh.ShopCode
INNER JOIN posbrntb brn WITH (NOLOCK) ON brn.brandcode=p.brandcode
WHERE bh.EntDate=%(bizdate)s
GROUP BY CONVERT(varchar(10),bh.EntDate,120), bh.ShopCode, p.ShopDes1, brn.brandnam1, p.VatType;
"""

SQL_HOUR = """
SELECT CONVERT(varchar(10), bh.EntDate,120) AS biz_date, bh.ShopCode AS shop_code,
  DATEPART(HOUR, bh.BillTime) AS hour_slot,
  COUNT(CASE WHEN bh.BillType='INV' AND ISNULL(bh.BillStatus,'')='' THEN 1 END) AS bill_count,
  SUM(CASE p.VatType WHEN '1' THEN bh.NetBAmt+(bh.VatAmt1-bh.VatAmt2) ELSE bh.NetBAmt END) AS net_sales,
  SUM(bh.BillDisc+bh.MemDisc+bh.Voucher+bh.ItemDisc+bh.CoupDisc) AS discount
FROM ONLBillHdrTB bh WITH (NOLOCK)
INNER JOIN PosShopTB p WITH (NOLOCK) ON p.ShopCode=bh.ShopCode
WHERE bh.EntDate=%(bizdate)s
GROUP BY CONVERT(varchar(10),bh.EntDate,120), bh.ShopCode, DATEPART(HOUR,bh.BillTime), p.VatType;
"""

# นับเฉพาะบิลขายที่ไม่ถูก void (เหมือน bill_count) — เดิมนับบิลที่ถูกยกเลิกด้วย member เลยเกินจริงเล็กน้อย (2026-11-23)
SQL_MEMBER = """
SELECT biz_date, shop_code, COUNT(*) AS member_bills FROM (
  SELECT DISTINCT CONVERT(varchar(10), p.EntDate,120) AS biz_date,
         p.ShopCode AS shop_code, p.TerminalNo, p.BillNo
  FROM ONLBillPayTB p WITH (NOLOCK)
  INNER JOIN ONLBillHdrTB h WITH (NOLOCK)
    ON h.ShopCode=p.ShopCode AND h.TerminalNo=p.TerminalNo AND h.BillNo=p.BillNo AND h.EntDate=p.EntDate
  WHERE p.EntDate=%(bizdate)s AND p.tendtype='MEM' AND p.TendCode='Member'
    AND h.BillType='INV' AND ISNULL(h.BillStatus,'')=''
) x
GROUP BY biz_date, shop_code;
"""

UPSERT_RT = """
INSERT INTO sale_monitor_rt
 (biz_date,shop_code,shop_name,brand,bill_count,cust_count,gross_sales,discount,
  base_sales,vat,net_sales,service_charge,net_total,member_bills,last_bill_time,snapshot_at)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
ON DUPLICATE KEY UPDATE
  shop_name=VALUES(shop_name), brand=VALUES(brand), bill_count=VALUES(bill_count),
  cust_count=VALUES(cust_count), gross_sales=VALUES(gross_sales), discount=VALUES(discount),
  base_sales=VALUES(base_sales), vat=VALUES(vat), net_sales=VALUES(net_sales),
  service_charge=VALUES(service_charge), net_total=VALUES(net_total),
  member_bills=VALUES(member_bills), last_bill_time=VALUES(last_bill_time), snapshot_at=VALUES(snapshot_at);
"""

UPSERT_INT = """
INSERT INTO sale_monitor_interval
 (biz_date,shop_code,hour_slot,bill_count,net_sales,discount,snapshot_at)
VALUES (%s,%s,%s,%s,%s,%s,%s)
ON DUPLICATE KEY UPDATE
  bill_count=VALUES(bill_count), net_sales=VALUES(net_sales),
  discount=VALUES(discount), snapshot_at=VALUES(snapshot_at);
"""

def sync(bizdate: str, force_extras: bool = False):
    """ดึง 1 วันธุรกิจจาก CM POS แล้ว upsert ลง MySQL"""
    now = dt.datetime.now()
    with pymssql.connect(**MSSQL) as src, src.cursor(as_dict=True) as c:
        c.execute(SQL_SHOP, {"bizdate": bizdate}); shops = c.fetchall()
        c.execute(SQL_HOUR, {"bizdate": bizdate}); hours = c.fetchall()
        c.execute(SQL_MEMBER, {"bizdate": bizdate}); members = c.fetchall()

    mem_map = {(r["biz_date"], r["shop_code"]): r["member_bills"] for r in members}
    rt = [(r["biz_date"], r["shop_code"], r["shop_name"], r["brand"], r["bill_count"],
           r["cust_count"], r["gross_sales"], r["discount"], r["base_sales"], r["vat"],
           r["net_sales"], r["service_charge"], r["net_total"],
           mem_map.get((r["biz_date"], r["shop_code"]), 0), r["last_bill_time"], now)
          for r in shops]
    iv = [(r["biz_date"], r["shop_code"], r["hour_slot"], r["bill_count"],
           r["net_sales"], r["discount"], now) for r in hours]

    with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
        if rt: c.executemany(UPSERT_RT, rt)
        if iv: c.executemany(UPSERT_INT, iv)
        dst.commit()

    try:
        n_min = sync_minute(bizdate)
    except Exception as e:                 # ตารางรายนาทีเป็นส่วนเสริม — พังก็ไม่กระทบ rt/interval
        n_min = -1
        print(f"[{now:%H:%M:%S}] {bizdate}: minute sync failed: {e}")
    try:
        ext = sync_extras(bizdate, force=force_extras)
    except Exception as e:                 # ตารางรายละเอียดเป็นส่วนเสริม — พังก็ไม่กระทบ rt/interval
        ext = f"extras failed: {e}"
    print(f"[{now:%H:%M:%S}] {bizdate}: {len(rt)} shops, {len(iv)} hourly, {n_min} minute rows synced; {ext}")


# ===================== รายละเอียดหน้า Sale Monitor: จ่ายเงิน / โปรโมชั่น / Member-Guest / สินค้า / ตะกร้า =====================
# เก็บรายสาขา (หน้าเว็บกรองแบรนด์/สาขาได้) แบบแทนทั้งวัน (DELETE+INSERT ใน transaction เดียว)
# รอบ today เติมทุก ~EXTRA_EVERY_SEC วิ (เช็คจาก snapshot ล่าสุด — ไม่ต้องแก้ crontab, cron ค้างรอบไหนก็ตามทันรอบถัดไป)
# ไม่ทุก 30 วิเหมือน rt เพราะแถวเยอะกว่า (สินค้า ~4–6k แถว/วัน) และกราฟพวกนี้ไม่ต้องสดระดับวินาที
EXTRA_EVERY_SEC   = 240
EXTRA_KEEP_MONTHS = 2      # tender/promo/custtype/item
BILL_ITEM_KEEP_DAYS = 14   # bill_item (ไว้คิดตะกร้า) โตเร็วกว่า เก็บสั้น

DDL_EXTRA = {
"sale_monitor_tender": """CREATE TABLE IF NOT EXISTS sale_monitor_tender (
    biz_date DATE NOT NULL, shop_code VARCHAR(10) NOT NULL,
    tendtype VARCHAR(10) NOT NULL, tend_code VARCHAR(20) NOT NULL, card_name VARCHAR(100) NULL,
    amt DECIMAL(18,2) NOT NULL DEFAULT 0, bills INT NOT NULL DEFAULT 0, snapshot_at DATETIME NULL,
    PRIMARY KEY (biz_date, shop_code, tendtype, tend_code)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
"sale_monitor_promo": """CREATE TABLE IF NOT EXISTS sale_monitor_promo (
    biz_date DATE NOT NULL, shop_code VARCHAR(10) NOT NULL,
    disc_code VARCHAR(20) NOT NULL, disc_name VARCHAR(150) NULL, disc_type VARCHAR(10) NULL,
    disc_amt DECIMAL(18,2) NOT NULL DEFAULT 0, bills INT NOT NULL DEFAULT 0, snapshot_at DATETIME NULL,
    PRIMARY KEY (biz_date, shop_code, disc_code)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
"sale_monitor_custtype": """CREATE TABLE IF NOT EXISTS sale_monitor_custtype (
    biz_date DATE NOT NULL, shop_code VARCHAR(10) NOT NULL, cust_type VARCHAR(10) NOT NULL,
    bills INT NOT NULL DEFAULT 0, net_sales DECIMAL(18,2) NOT NULL DEFAULT 0, pax INT NOT NULL DEFAULT 0,
    snapshot_at DATETIME NULL,
    PRIMARY KEY (biz_date, shop_code, cust_type)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
"sale_monitor_item": """CREATE TABLE IF NOT EXISTS sale_monitor_item (
    biz_date DATE NOT NULL, shop_code VARCHAR(10) NOT NULL, item_code VARCHAR(30) NOT NULL,
    item_name VARCHAR(150) NULL, dept VARCHAR(50) NULL,
    qty DECIMAL(18,2) NOT NULL DEFAULT 0, net_amt DECIMAL(18,2) NOT NULL DEFAULT 0, bills INT NOT NULL DEFAULT 0,
    snapshot_at DATETIME NULL,
    PRIMARY KEY (biz_date, shop_code, item_code)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
"sale_monitor_bill_item": """CREATE TABLE IF NOT EXISTS sale_monitor_bill_item (
    biz_date DATE NOT NULL, shop_code VARCHAR(10) NOT NULL, bill_key VARCHAR(40) NOT NULL,
    item_code VARCHAR(30) NOT NULL,
    PRIMARY KEY (biz_date, shop_code, bill_key, item_code)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
}

# จ่ายเงิน: เก็บดิบทุก tendtype/TendCode (หน้าเว็บจัดกลุ่ม เงินสด/บัตร/QR/delivery/voucher เอง)
# ตัดแถวที่ไม่ใช่เงิน (OCR ปัดเศษ, CUS/REF/REP) และแถว MEM 'Member' (CardName = ชื่อลูกค้า — ไม่เก็บข้อมูลส่วนตัว)
SQL_X_TENDER = """
SELECT CONVERT(varchar(10), EntDate,120) AS biz_date, ShopCode AS shop_code, tendtype, TendCode AS tend_code,
  MAX(CardName) AS card_name, SUM(TendAmt) AS amt, COUNT(DISTINCT TerminalNo+'|'+BillNo) AS bills
FROM ONLBillPayTB WITH (NOLOCK)
WHERE EntDate=%(bizdate)s AND tendtype NOT IN ('OCR','CUS','REF','REP')
  AND NOT (tendtype='MEM' AND TendCode='Member')
GROUP BY CONVERT(varchar(10), EntDate,120), ShopCode, tendtype, TendCode;
"""
SQL_X_PROMO = """
SELECT CONVERT(varchar(10), EntDate,120) AS biz_date, ShopCode AS shop_code, DiscCode AS disc_code,
  MAX(DiscName) AS disc_name, MAX(DiscType) AS disc_type, SUM(DiscAmt) AS disc_amt,
  COUNT(DISTINCT TerminalNo+'|'+BillNo) AS bills
FROM ONLBillMultiDiscTB WITH (NOLOCK)
WHERE EntDate=%(bizdate)s
GROUP BY CONVERT(varchar(10), EntDate,120), ShopCode, DiscCode;
"""
# member = บิลที่มีแถว MEM 'Member' (นิยามเดียวกับ sale_monitor_rt.member_bills) — net ใช้สูตรเดียวกับ SQL_SHOP
SQL_X_CUST = """
SELECT CONVERT(varchar(10), h.EntDate,120) AS biz_date, h.ShopCode AS shop_code,
  CASE WHEN m.BillNo IS NULL THEN 'Guest' ELSE 'Member' END AS cust_type,
  COUNT(CASE WHEN h.BillType='INV' AND ISNULL(h.BillStatus,'')='' THEN 1 END) AS bills,
  SUM(CASE p.VatType WHEN '1' THEN h.NetBAmt+(h.VatAmt1-h.VatAmt2) ELSE h.NetBAmt END) AS net_sales,
  SUM(CASE WHEN h.BillType='INV' AND ISNULL(h.BillStatus,'')='' THEN h.NoPerson ELSE 0 END) AS pax
FROM ONLBillHdrTB h WITH (NOLOCK)
INNER JOIN PosShopTB p WITH (NOLOCK) ON p.ShopCode=h.ShopCode
LEFT JOIN (SELECT DISTINCT ShopCode, TerminalNo, BillNo FROM ONLBillPayTB WITH (NOLOCK)
           WHERE EntDate=%(bizdate)s AND tendtype='MEM' AND TendCode='Member') m
  ON m.ShopCode=h.ShopCode AND m.TerminalNo=h.TerminalNo AND m.BillNo=h.BillNo
WHERE h.EntDate=%(bizdate)s
GROUP BY CONVERT(varchar(10), h.EntDate,120), h.ShopCode, CASE WHEN m.BillNo IS NULL THEN 'Guest' ELSE 'Member' END;
"""
# สินค้า: เฉพาะบรรทัดที่มีราคา (NetAmt>0) ของบิลขายที่ไม่ถูก void — NetAmt รวมทั้งวัน ≈ Base Sale (ก่อน VAT)
_X_DTL_FROM = """
FROM ONLBillDtlTB d WITH (NOLOCK)
INNER JOIN ONLBillHdrTB h WITH (NOLOCK)
  ON h.ShopCode=d.ShopCode AND h.TerminalNo=d.TerminalNo AND h.BillType=d.BillType
 AND h.BillNo=d.BillNo AND h.EntDate=d.EntDate
WHERE d.EntDate=%(bizdate)s AND h.BillType='INV' AND ISNULL(h.BillStatus,'')='' AND d.NetAmt>0
"""
SQL_X_ITEM = """
SELECT CONVERT(varchar(10), d.EntDate,120) AS biz_date, d.ShopCode AS shop_code, d.ItemCode AS item_code,
  MAX(COALESCE(i.ItemDes1, d.ItemCode)) AS item_name, MAX(RTRIM(COALESCE(dep.DepDes1,''))) AS dept,
  SUM(d.ActQTY) AS qty, SUM(d.NetAmt) AS net_amt, COUNT(DISTINCT d.TerminalNo+'|'+d.BillNo) AS bills
""" + _X_DTL_FROM.replace("WHERE", """LEFT JOIN ItmMasTB i WITH (NOLOCK) ON i.ItemCode=d.ItemCode
LEFT JOIN ItmDepTB dep WITH (NOLOCK) ON dep.DepCode=i.DepCode0
WHERE""") + """
GROUP BY CONVERT(varchar(10), d.EntDate,120), d.ShopCode, d.ItemCode;
"""
SQL_X_BILL_ITEM = """
SELECT DISTINCT CONVERT(varchar(10), d.EntDate,120) AS biz_date, d.ShopCode AS shop_code,
  d.TerminalNo+'|'+d.BillNo AS bill_key, d.ItemCode AS item_code
""" + _X_DTL_FROM + ";"

X_COLS = {
    "sale_monitor_tender":    ["biz_date", "shop_code", "tendtype", "tend_code", "card_name", "amt", "bills", "snapshot_at"],
    "sale_monitor_promo":     ["biz_date", "shop_code", "disc_code", "disc_name", "disc_type", "disc_amt", "bills", "snapshot_at"],
    "sale_monitor_custtype":  ["biz_date", "shop_code", "cust_type", "bills", "net_sales", "pax", "snapshot_at"],
    "sale_monitor_item":      ["biz_date", "shop_code", "item_code", "item_name", "dept", "qty", "net_amt", "bills", "snapshot_at"],
    "sale_monitor_bill_item": ["biz_date", "shop_code", "bill_key", "item_code"],
}
X_SQL = {"sale_monitor_tender": SQL_X_TENDER, "sale_monitor_promo": SQL_X_PROMO,
         "sale_monitor_custtype": SQL_X_CUST, "sale_monitor_item": SQL_X_ITEM,
         "sale_monitor_bill_item": SQL_X_BILL_ITEM}


def sync_extras(bizdate: str, force: bool = False) -> str:
    now = dt.datetime.now().replace(microsecond=0)
    with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
        last = None
        try:
            c.execute("SELECT MAX(snapshot_at) FROM sale_monitor_custtype WHERE biz_date=%s", (bizdate,))
            last = c.fetchone()[0]
        except pymysql.err.ProgrammingError:          # ยังไม่มีตาราง → สร้างครั้งแรก
            for ddl in DDL_EXTRA.values():
                c.execute(ddl)
            dst.commit()
        if not force and last and (now - last).total_seconds() < EXTRA_EVERY_SEC:
            return "extras skip"
    data = {}
    with pymssql.connect(**MSSQL, charset="CP874") as src, src.cursor(as_dict=True) as c:   # ชื่อไทย (เมนู/โปร) ใน CM POS เป็น CP874
        for tbl, sql in X_SQL.items():
            c.execute(sql, {"bizdate": bizdate}); data[tbl] = c.fetchall()
    if not data["sale_monitor_custtype"]:
        return "extras: no bills"                     # source ว่างชั่วคราว — ไม่แตะข้อมูลเดิม
    with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
        for tbl, cols in X_COLS.items():
            rows = []
            for r in data[tbl]:
                row = []
                for col in cols:
                    v = now if col == "snapshot_at" else r.get(col)
                    if isinstance(v, str):
                        v = v.strip()[:150]
                    row.append(v)
                rows.append(tuple(row))
            c.execute(f"DELETE FROM {tbl} WHERE biz_date=%s", (bizdate,))
            if rows:
                c.executemany(f"INSERT INTO {tbl} ({','.join(cols)}) VALUES ({','.join(['%s'] * len(cols))})", rows)
        dst.commit()
    return "extras " + ", ".join(f"{t.replace('sale_monitor_', '')}={len(data[t])}" for t in X_COLS)


def purge_extras():
    try:
        with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
            for tbl in ("sale_monitor_tender", "sale_monitor_promo", "sale_monitor_custtype", "sale_monitor_item"):
                c.execute(f"DELETE FROM {tbl} WHERE biz_date < DATE_SUB(CURDATE(), INTERVAL %s MONTH)", (EXTRA_KEEP_MONTHS,))
            c.execute("DELETE FROM sale_monitor_bill_item WHERE biz_date < DATE_SUB(CURDATE(), INTERVAL %s DAY)",
                      (BILL_ITEM_KEEP_DAYS,))
            dst.commit()
    except Exception as e:
        print(f"purge extras failed: {e}")


# ===================== ยอดรายนาที (เทียบ "เมื่อวาน ณ เวลาเดียวกัน" ตรงนาที) =====================
# แบบเดียวกับ Dunkin (/opt/exception_report/etl/etl_sale_monitor.py) — หน้า Sale Monitor ใน app.py ใช้ตัด
# ยอดเมื่อวานที่เวลา snapshot ของวันนี้ (นาทีก่อนหน้าเต็ม ๆ + นาทีปัจจุบันตามสัดส่วนวินาที)
# net_sales / bill_count ใช้สูตรเดียวกับ SQL_SHOP → รวมทั้งวันแล้วเท่ากับ sale_monitor_rt
DDL_MIN = """
CREATE TABLE IF NOT EXISTS sale_monitor_minute (
    biz_date    DATE          NOT NULL,
    shop_code   VARCHAR(10)   NOT NULL,
    minute_slot SMALLINT      NOT NULL,     -- HHMM เช่น 1325 = 13:25
    bill_count  INT           NOT NULL DEFAULT 0,
    net_sales   DECIMAL(18,2) NOT NULL DEFAULT 0,
    snapshot_at DATETIME      NULL,
    PRIMARY KEY (biz_date, shop_code, minute_slot)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""
SQL_MINUTE = """
SELECT CONVERT(varchar(10), bh.EntDate,120) AS biz_date, bh.ShopCode AS shop_code,
  CAST(LEFT(bh.BillTime,2) AS INT)*100 + CAST(SUBSTRING(bh.BillTime,4,2) AS INT) AS minute_slot,
  COUNT(CASE WHEN bh.BillType='INV' AND ISNULL(bh.BillStatus,'')='' THEN 1 END) AS bill_count,
  SUM(CASE p.VatType WHEN '1' THEN bh.NetBAmt+(bh.VatAmt1-bh.VatAmt2) ELSE bh.NetBAmt END) AS net_sales
FROM ONLBillHdrTB bh WITH (NOLOCK)
INNER JOIN PosShopTB p WITH (NOLOCK) ON p.ShopCode=bh.ShopCode
WHERE bh.EntDate=%(bizdate)s AND bh.BillTime LIKE '[0-2][0-9]:[0-5][0-9]%%'
  AND CONVERT(date, bh.BillDate) <= CONVERT(date, bh.EntDate)
GROUP BY CONVERT(varchar(10),bh.EntDate,120), bh.ShopCode,
  CAST(LEFT(bh.BillTime,2) AS INT)*100 + CAST(SUBSTRING(bh.BillTime,4,2) AS INT);
"""
# ↑ ตัดบิลที่ "คีย์ย้อนหลังข้ามวัน" (BillDate > EntDate) — เจอ 2026-11-23: สาขา 0033 MSMC-Mahanakhon CUBE
#   คีย์ยอดของวันธุรกิจ D ตอน ~11:00 (บางวัน ~20:30) ของวัน D+1 ทุกวัน → BillTime ไม่ใช่เวลาขายจริง
#   ถ้านับ: เมื่อวาน ณ เวลาเดียวกันจะได้ยอดก้อนนี้ที่ 11:xx แต่วันนี้ไม่มีวันมี (มาพรุ่งนี้) → pace ติดลบเกินจริง
#   ยอดทั้งวัน (sale_monitor_rt) ยังนับครบ — ตัดเฉพาะการเทียบรายนาที
UPSERT_MIN = """
INSERT INTO sale_monitor_minute (biz_date,shop_code,minute_slot,bill_count,net_sales,snapshot_at)
VALUES (%s,%s,%s,%s,%s,%s)
ON DUPLICATE KEY UPDATE bill_count=VALUES(bill_count), net_sales=VALUES(net_sales),
  snapshot_at=VALUES(snapshot_at);
"""
MINUTE_KEEP_MONTHS = 2   # เก็บย้อนหลังกี่เดือน (Greyhound ~3–4k แถว/วัน)


def sync_minute(bizdate: str) -> int:
    # ตัด microsecond: DATETIME ใน MySQL ปัดวินาที — ไม่ตัดแล้วแถวที่เพิ่ง upsert อาจโดน DELETE ด้านล่าง
    now = dt.datetime.now().replace(microsecond=0)
    with pymssql.connect(**MSSQL) as src, src.cursor(as_dict=True) as c:
        c.execute(SQL_MINUTE, {"bizdate": bizdate}); mins = c.fetchall()
    rows = [(r["biz_date"], str(r["shop_code"] or "")[:10], int(r["minute_slot"]),
             int(r["bill_count"] or 0), float(r["net_sales"] or 0), now)
            for r in mins if r["minute_slot"] is not None]
    if not rows:
        return 0          # ไม่มีบิล (หรือ source ว่างชั่วคราว) — ไม่แตะข้อมูลเดิม
    with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
        c.execute(DDL_MIN)
        c.executemany(UPSERT_MIN, rows)
        # นาทีที่ไม่ถูก upsert รอบนี้ = ไม่มีบิลเหลือแล้ว (เช่น void ทีหลัง) → ลบ
        c.execute("DELETE FROM sale_monitor_minute WHERE biz_date=%s AND (snapshot_at IS NULL OR snapshot_at<%s)",
                  (bizdate, now))
        dst.commit()
    return len(rows)


def purge_minute() -> int:
    """ลบ sale_monitor_minute ที่เก่ากว่า MINUTE_KEEP_MONTHS เดือน (เรียกจากรอบ 02:15 yesterday)"""
    try:
        with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
            n = c.execute("DELETE FROM sale_monitor_minute WHERE biz_date < DATE_SUB(CURDATE(), INTERVAL %s MONTH)",
                          (MINUTE_KEEP_MONTHS,))
            dst.commit()
        return n
    except Exception as e:
        print(f"purge sale_monitor_minute failed: {e}")
        return -1


# ===================== FastAPI =====================
from fastapi import FastAPI, Header, HTTPException, Query
app = FastAPI(title="Greyhound BI")

def mysql():
    return pymysql.connect(**{**MYSQL, "cursorclass": pymysql.cursors.DictCursor})

@app.get("/api/overview/sale-monitor")
def sale_monitor(brand: str | None = None, topn: int = Query(5, ge=1, le=50),
                 x_api_key: str = Header(None)):
    if x_api_key != API_KEY:
        raise HTTPException(401, "invalid api key")
    today = dt.date.today().isoformat()
    yday  = (dt.date.today() - dt.timedelta(days=1)).isoformat()
    cur_h = dt.datetime.now().hour
    b = brand if brand and brand != "ทุกแบรนด์" else None

    with mysql() as conn, conn.cursor() as c:
        c.execute("""SELECT shop_code,shop_name,brand,
              MAX(CASE WHEN biz_date=%(t)s THEN net_sales END) today,
              MAX(CASE WHEN biz_date=%(y)s THEN net_sales END) yday
            FROM sale_monitor_rt
            WHERE biz_date IN (%(t)s,%(y)s) AND (%(b)s IS NULL OR brand=%(b)s)
            GROUP BY shop_code,shop_name,brand ORDER BY today DESC LIMIT %(n)s""",
            {"t": today, "y": yday, "b": b, "n": topn})
        shops = [{"code": r["shop_code"], "name": r["shop_name"],
                  "today": float(r["today"] or 0), "yday": float(r["yday"] or 0)} for r in c.fetchall()]

        c.execute("""SELECT biz_date, SUM(net_sales) total FROM sale_monitor_rt
            WHERE biz_date IN (%(t)s,%(y)s) AND (%(b)s IS NULL OR brand=%(b)s)
            GROUP BY biz_date""", {"t": today, "y": yday, "b": b})
        tot = {str(r["biz_date"]): float(r["total"] or 0) for r in c.fetchall()}

        c.execute("""SELECT SUM(net_sales) s FROM sale_monitor_interval
            WHERE biz_date=%(y)s AND hour_slot<=%(h)s""", {"y": yday, "h": cur_h})
        yst = float((c.fetchone() or {}).get("s") or 0)

        c.execute("SELECT MAX(snapshot_at) s FROM sale_monitor_rt WHERE biz_date=%(t)s", {"t": today})
        snap = (c.fetchone() or {}).get("s")

    return {  # ← ตรงกับ DATA.monitor ในหน้า greyhound-bi.html
        "snapshot": snap.strftime("%H:%M") if snap else "--:--",
        "totalToday": tot.get(today, 0.0),
        "ydayFull":   tot.get(yday, 0.0),
        "ydaySameTime": yst,
        "shops": shops,
    }


def resync_day(bizdate: str) -> str:
    """ดึงวันย้อนหลังใหม่ทั้งวัน — กันพลาด: ถ้ายอดต้นทางตอนนี้ต่ำกว่าที่เก็บไว้ (เช่น ONL ของ CM POS
    เริ่มลบวันเก่าทิ้งไปบางส่วน) จะไม่เขียนทับ"""
    with pymssql.connect(**MSSQL) as src, src.cursor(as_dict=True) as c:
        c.execute(SQL_SHOP, {"bizdate": bizdate}); shops = c.fetchall()
    src_total = sum(float(r["net_sales"] or 0) for r in shops)
    with pymysql.connect(**MYSQL) as dst, dst.cursor() as c:
        c.execute("SELECT COALESCE(SUM(net_sales),0) FROM sale_monitor_rt WHERE biz_date=%s", (bizdate,))
        cur_total = float(c.fetchone()[0] or 0)
    if not shops or src_total < cur_total * 0.995:
        return f"{bizdate}: SKIP ต้นทาง ฿{src_total:,.0f} < ที่เก็บไว้ ฿{cur_total:,.0f}"
    sync(bizdate, force_extras=True)
    return f"{bizdate}: ฿{cur_total:,.0f} → ฿{src_total:,.0f} ({src_total - cur_total:+,.0f})"


if __name__ == "__main__":
    # today | yesterday (cron) | minute YYYY-MM-DD [YYYY-MM-DD ...] = เติมเฉพาะตารางรายนาทีย้อนหลัง
    # (ONLBillHdrTB ของ CM POS เก็บย้อนหลังแค่ ~2 สัปดาห์)
    arg = sys.argv[1] if len(sys.argv) > 1 else "today"
    if arg == "minute":
        for d in sys.argv[2:]:
            print(d, sync_minute(dt.date.fromisoformat(d).isoformat()), "minute rows")
        sys.exit(0)
    if arg == "extras":                 # extras YYYY-MM-DD ... = เติมตารางรายละเอียดย้อนหลัง/บังคับรอบนี้
        for d in sys.argv[2:]:
            print(d, sync_extras(dt.date.fromisoformat(d).isoformat(), force=True))
        sys.exit(0)
    if arg == "resync":                 # resync YYYY-MM-DD ... = ดึงใหม่ทั้งวัน (rt/interval/minute/extras) ย้อนหลัง
        for d in sys.argv[2:]:
            print(resync_day(dt.date.fromisoformat(d).isoformat()))
        sys.exit(0)
    target = dt.date.today() if arg == "today" else dt.date.today() - dt.timedelta(days=1)
    try:
        sync(target.isoformat(), force_extras=(arg == "yesterday"))
    finally:
        if arg == "yesterday":          # รอบ 02:15 / 12:15 / 23:50 (ดึงเมื่อวานซ้ำให้ได้บิลที่คีย์ย้อนหลัง)
            purge_minute()
            purge_extras()
