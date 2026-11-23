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

SQL_MEMBER = """
SELECT biz_date, shop_code, COUNT(*) AS member_bills FROM (
  SELECT DISTINCT CONVERT(varchar(10), EntDate,120) AS biz_date,
         ShopCode AS shop_code, TerminalNo, BillNo
  FROM ONLBillPayTB WITH (NOLOCK)
  WHERE EntDate=%(bizdate)s AND tendtype='MEM' AND TendCode='Member'
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

def sync(bizdate: str):
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
    print(f"[{now:%H:%M:%S}] {bizdate}: {len(rt)} shops, {len(iv)} hourly, {n_min} minute rows synced")


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
GROUP BY CONVERT(varchar(10),bh.EntDate,120), bh.ShopCode,
  CAST(LEFT(bh.BillTime,2) AS INT)*100 + CAST(SUBSTRING(bh.BillTime,4,2) AS INT);
"""
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


if __name__ == "__main__":
    # today | yesterday (cron) | minute YYYY-MM-DD [YYYY-MM-DD ...] = เติมเฉพาะตารางรายนาทีย้อนหลัง
    # (ONLBillHdrTB ของ CM POS เก็บย้อนหลังแค่ ~2 สัปดาห์)
    arg = sys.argv[1] if len(sys.argv) > 1 else "today"
    if arg == "minute":
        for d in sys.argv[2:]:
            print(d, sync_minute(dt.date.fromisoformat(d).isoformat()), "minute rows")
        sys.exit(0)
    target = dt.date.today() if arg == "today" else dt.date.today() - dt.timedelta(days=1)
    try:
        sync(target.isoformat())
    finally:
        if arg == "yesterday":          # รอบ 02:15 วันละครั้ง
            purge_minute()
