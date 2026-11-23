"""
greyhound_bi_api.py — FastAPI backend (port 9007) สำหรับหน้า greyhound-bi.html
-----------------------------------------------------------------------------
จุดเด่น: รันได้ทันทีแม้ยังไม่มีตาราง/ฐานข้อมูล — endpoint ทุกตัว fallback เป็นข้อมูล
ตัวอย่าง (SAMPLE) ถ้าต่อ MySQL ไม่ได้ จึงเอาไว้เทสต์การ connect ของหน้าเว็บก่อน
แล้วค่อยทยอยแทน query จริงทีละส่วน

ติดตั้ง:   pip install fastapi "uvicorn[standard]" pymysql
รัน:       uvicorn greyhound_bi_api:app --host 0.0.0.0 --port 9007
ทดสอบ:     endpoint ข้อมูลทุกตัวต้อง login ก่อน (ไม่มี token = 401) — ได้ token จาก /api/auth/login ด้านล่าง
           curl "http://10.8.1.88:9007/api/bi" -H "Authorization: Bearer <token>"
           curl "http://10.8.1.88:9007/api/bi?date_from=2026-07-01&date_to=2026-07-31&shop=0004,0015" -H "Authorization: Bearer <token>"
           curl "http://10.8.1.88:9007/api/bi?brand=GC" -H "Authorization: Bearer <token>"   # brand = PosShopTB.ShTypCode (GC/AH/BE)
ค่าลับ:    รหัส DB / API key อยู่ใน .env (โหลดผ่าน gh_env.py) — AUTH_SECRET ตั้งใน systemd unit
Login:     curl -X POST :9007/api/auth/login -d '{"username":"viewer","password":"..."}' -H "Content-Type: application/json"
           curl :9007/api/auth/me -H "Authorization: Bearer <token>"

⚠️ MIXED CONTENT: หน้าเว็บอยู่บน https://mudandgreyhound.github.io แต่ API นี้เป็น http://
   เบราว์เซอร์จะบล็อกการ fetch ข้าม http จาก https — เลือกทางใดทางหนึ่ง:
     (ก) ย้าย index.html มาเสิร์ฟบนเซิร์ฟเวอร์เดียวกับ API ผ่าน http  (ง่ายสุด)
     (ข) ทำ https ให้ API: reverse proxy (nginx + cert) หรือ cloudflared tunnel
   ระหว่างนี้หน้า GitHub Pages จะยังโชว์ข้อมูลตัวอย่างได้ (มี fallback)

systemd (/etc/systemd/system/greyhound-bi-api.service):
   [Service]
   WorkingDirectory=/opt/promotion_dashboard
   Environment=AUTH_SECRET=<random secret — สร้างด้วย python3 -c "import secrets;print(secrets.token_hex(32))">
   ExecStart=/usr/bin/uvicorn greyhound_bi_api:app --host 0.0.0.0 --port 9007
   Restart=always
"""
import os, datetime as dt
import hashlib, hmac, base64, json, time, secrets
from fastapi import FastAPI, Header, Query, Body, Depends, HTTPException
from fastapi.middleware.cors import CORSMiddleware

try:
    import pymysql
    from pymysql.cursors import DictCursor
except Exception:
    pymysql = None

# ---------- config ----------
import gh_env   # รหัส DB / API key อยู่ใน .env (ไม่เข้า git) — env ของ process (MYSQL_* / BI_API_KEY) ยัง override ได้
API_KEY = os.getenv("BI_API_KEY") or gh_env.need("BI_API_KEY")
MYSQL = dict(
    host=os.getenv("MYSQL_HOST") or gh_env.need("GH_MYSQL_HOST"),
    user=os.getenv("MYSQL_USER") or gh_env.need("GH_MYSQL_ADMIN_USER"),
    password=os.getenv("MYSQL_PWD") or gh_env.need("GH_MYSQL_ADMIN_PASSWORD"),
    database=os.getenv("MYSQL_DB") or gh_env.need("GH_MYSQL_DB"),
    charset="utf8mb4",
)
ALLOWED_ORIGINS = [
    "https://mudandgreyhound.github.io",
    "http://localhost:5500", "http://127.0.0.1:5500",
    "http://10.8.1.88:9007", "http://10.8.1.88",
]

app = FastAPI(title="Greyhound BI API", version="1.1")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],   # POST เพิ่มมาสำหรับ /api/auth/login
    allow_headers=["*"],
)

# ---------- db helper ----------
def q(sql, params=None):
    """รัน SELECT คืน list[dict]; ถ้า DB ใช้ไม่ได้ → คืน None เพื่อให้ caller fallback"""
    if pymysql is None:
        return None
    try:
        conn = pymysql.connect(**MYSQL, cursorclass=DictCursor, connect_timeout=4)
        with conn, conn.cursor() as c:
            c.execute(sql, params or ())
            return c.fetchall()
    except Exception as e:
        print("[mysql] fallback to sample:", e)
        return None

def check_key(key, x_api_key):
    """อนุญาตเมื่อ key ถูก หรือไม่ส่ง key มาเลย. ส่ง key ผิด = ปฏิเสธ
    หมายเหตุ: key นี้ฝังอยู่ใน index.html ที่ GitHub Pages (สาธารณะ) จึงไม่ใช่ความลับ
    ด่านจริงคือ login — endpoint ข้อมูลทุกตัวบังคับ Bearer token ผ่าน Depends(current_user)"""
    got = key or x_api_key
    return got is None or got == API_KEY

# =====================================================================
# AUTH — login + role (stdlib ล้วน: pbkdf2 hash + HMAC token)
# ใช้ตาราง `users` เดิมที่มีอยู่แล้วใน gh_promotion (ใช้ร่วมกับ recipe_master)
# role: admin (เห็นทุกอย่าง+users) / editor / viewer
# pattern เดียวกับ /opt/recipe_master/api/api.py เพื่อความสม่ำเสมอในเซิร์ฟเวอร์นี้
# =====================================================================
_AUTH_SECRET_ENV = os.getenv("AUTH_SECRET")
if _AUTH_SECRET_ENV:
    AUTH_SECRET = _AUTH_SECRET_ENV
else:
    AUTH_SECRET = secrets.token_hex(32)
    print("[auth] WARNING: ไม่ได้ตั้ง AUTH_SECRET env var — ใช้ secret ชั่วคราว "
          "(token จะใช้ไม่ได้ทุกครั้งที่ restart service) ตั้งค่าใน systemd unit สำหรับ production")

def verify_pw(pw, stored):
    """ตรวจ password กับ hash รูปแบบ pbkdf2$sha256$iters$salt$hex (หรือ pbkdf2:sha256:... ของ Werkzeug)"""
    try:
        parts = stored.split("$") if "$" in stored else stored.split(":")
        _, _, iters, salt, hexdk = parts
        dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), int(iters))
        return hmac.compare_digest(dk.hex(), hexdk)
    except Exception:
        return False

def _b64(b): return base64.urlsafe_b64encode(b).rstrip(b"=").decode()
def _ub64(s): return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))

def make_token(username, role, ttl=8 * 3600):
    payload = {"sub": username, "role": role, "exp": int(time.time()) + ttl}
    p = _b64(json.dumps(payload).encode())
    sig = _b64(hmac.new(AUTH_SECRET.encode(), p.encode(), hashlib.sha256).digest())
    return p + "." + sig

def read_token(tok):
    try:
        p, sig = tok.split(".")
        exp_sig = _b64(hmac.new(AUTH_SECRET.encode(), p.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(sig, exp_sig):
            return None
        payload = json.loads(_ub64(p))
        if payload.get("exp", 0) < time.time():
            return None
        return payload
    except Exception:
        return None

def current_user(authorization: str = Header(None)):
    """FastAPI dependency — บังคับ login (401 ถ้าไม่มี/หมดอายุ) ใช้กับ endpoint ข้อมูลทุกตัว
    หน้าเว็บ (promotion-intelligence) เจอ 401 จะล้าง token แล้วเด้งหน้า login เอง"""
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="ต้อง login ก่อน")
    p = read_token(authorization[7:])
    if not p:
        raise HTTPException(status_code=401, detail="token หมดอายุ/ไม่ถูกต้อง — login ใหม่")
    return p

def current_user_optional(authorization: str = Header(None)):
    """เหมือน current_user แต่คืน None แทนการ raise — ใช้เช็ค role แบบไม่บังคับ login (เช่น gate เฉพาะ admin block)"""
    if not authorization or not authorization.startswith("Bearer "):
        return None
    return read_token(authorization[7:])

def require_role(*roles):
    """เผื่อใช้ตอนบังคับ auth เต็มรูปแบบใน endpoint (รอบ 2 ตามแผน)"""
    def dep(user=Depends(current_user)):
        if user.get("role") not in roles:
            raise HTTPException(status_code=403, detail="ไม่มีสิทธิ์ทำรายการนี้")
        return user
    return dep

# =====================================================================
# SAMPLE — โครงเดียวกับ DATA ในหน้าเว็บ (ใช้เป็น fallback/ค่าตั้งต้น)
# =====================================================================
SAMPLE = {
  "asOf": "18/06/2026", "updated": "19/06/2026 10:26", "sample": True,
  "kpi": {"baseSale":989203,"bills":1344,"discount":107291,"avgBill":736,"pax":2278,"memberBills":229},
  "memberSplit": {"guest":72.5,"member":27.5},
  "monitor": {"snapshot":"14:32","totalToday":43744.57,"ydayFull":241560.0,"ydaySameTime":39210.0,
    "shops":[{"code":"0004","name":"PG-Siam Paragon","today":2289.80,"yday":58986.84},
             {"code":"0034","name":"GHMP-MedPark","today":7708.28,"yday":57008.26},
             {"code":"0015","name":"EB-Central Embassy","today":4183.70,"yday":48076.17},
             {"code":"2015","name":"GHBE-Head Office","today":6284.00,"yday":8373.00},
             {"code":"2008","name":"BEMSN-Bean Hound Makro","today":4263.00,"yday":6098.00}]},
  "promos":[{"name":"Happy Hour Buy 1 Get 1","amount":66240,"bills":84,"related":198400,"icon":"ti-gift"},
            {"name":"Silver Hound Discount","amount":11021,"bills":105,"related":121800,"icon":"ti-medal-2"},
            {"name":"Black Hound Discount","amount":6909,"bills":47,"related":69400,"icon":"ti-discount-2"},
            {"name":"Value Price 49 Baht","amount":5498,"bills":68,"related":34600,"icon":"ti-tag"}],
  "hours":[6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21],
  "sales":[70,891,4854,4397,11128,36378,96672,147573,85405,51446,41928,63033,79394,117086,137099,114137],
  "discount":[4,60,350,320,1100,3600,9200,13800,9800,9200,8600,9600,9400,12200,12800,7300],
  "branches":[{"name":"PG-Siam Paragon","v":88200},{"name":"EQ-Emquartier","v":86400},
    {"name":"LP-Lad Praw","v":74100},{"name":"PH-Phaholyothin","v":65300},{"name":"CL-Chidlom","v":60200},
    {"name":"EB-Central Embassy","v":55400},{"name":"NL-Nang Linchee","v":50100},{"name":"SC-Siam Center","v":49300},
    {"name":"EM-Emporium","v":47200},{"name":"GHHO-Head Office-001 (POP Mart-IS)","v":44100},
    {"name":"GHRT-Ratchayothin","v":42300},{"name":"GHMP-MedPark","v":41000},
    {"name":"GHEV-Central Eastville","v":40200},{"name":"GHKT-Krungthep Kreetha","v":34100},{"name":"MB-Mega Bangna","v":32000}],
  "topItems":[{"name":"Spaghetti Cake","qty":412,"sales":148320},{"name":"Greyhound Fried Rice","qty":386,"sales":131240},
    {"name":"Tom Yum Spaghetti","qty":341,"sales":122760},{"name":"Crispy Pork Salad","qty":298,"sales":101320},
    {"name":"Iced Latte","qty":540,"sales":97200},{"name":"Complicated Sandwich","qty":265,"sales":92750},
    {"name":"Chocolate Lava","qty":243,"sales":72900},{"name":"Pad Thai Hound","qty":221,"sales":70720},
    {"name":"Green Curry Spaghetti","qty":198,"sales":65340},{"name":"Hound Burger","qty":176,"sales":61600}],
  "itemCats":[{"name":"จานหลัก","pct":46},{"name":"เครื่องดื่ม","pct":22},{"name":"ของหวาน","pct":18},{"name":"ทานเล่น/สลัด","pct":14}],
  "menuEng":[{"name":"Spaghetti Cake","pop":412,"margin":68},{"name":"Tom Yum Spaghetti","pop":341,"margin":64},
    {"name":"Iced Latte","pop":540,"margin":72},{"name":"Chocolate Lava","pop":243,"margin":66},
    {"name":"Greyhound Fried Rice","pop":386,"margin":41},{"name":"Pad Thai Hound","pop":221,"margin":38},
    {"name":"Hound Burger","pop":176,"margin":35},{"name":"Truffle Pasta","pop":128,"margin":71},
    {"name":"Lobster Roll","pop":96,"margin":69},{"name":"Matcha Latte","pop":142,"margin":63},
    {"name":"Fish & Chips","pop":118,"margin":39},{"name":"Caesar Salad","pop":104,"margin":42},
    {"name":"Soda Mocktail","pop":88,"margin":37},{"name":"Veggie Wrap","pop":72,"margin":33}],
  "member":{"guest":{"sales":717172,"bills":1115,"avg":643},"member":{"sales":272031,"bills":229,"avg":1188},
    "tiers":[{"name":"Silver Hound","pct":62},{"name":"Black Hound","pct":30},{"name":"อื่นๆ","pct":8}],"repeat":3.4},
  "billHist":[{"b":"<300","n":210},{"b":"300–500","n":340},{"b":"500–700","n":298},
    {"b":"700–1000","n":256},{"b":"1000–1500","n":148},{"b":">1500","n":92}],
  "billScatter":[[1,260],[1,180],[2,420],[2,520],[2,610],[3,690],[3,840],[4,920],[2,350],[3,760],
    [4,1180],[5,1420],[2,480],[3,580],[1,310],[4,1050],[6,1680],[2,540],[3,720],[5,1260]],
  "dailyTrend":[142000,138000,151000,160000,168000,212000,198000],
  "forecast":{"ingredient":"เนื้อไก่ (kg)","actual":[42,38,45,51,47,63,58],"predict":[44,40,47,52,49,65,60],
    "reorder":[{"name":"เนื้อไก่","unit":"kg","onhand":120,"perday":48,"days":2.5,"reorder":200,"risk":False},
      {"name":"กุ้ง","unit":"kg","onhand":18,"perday":15,"days":1.2,"reorder":80,"risk":True},
      {"name":"เส้นสปาเกตตี","unit":"kg","onhand":85,"perday":22,"days":3.9,"reorder":100,"risk":False},
      {"name":"นมสด","unit":"L","onhand":60,"perday":25,"days":2.4,"reorder":120,"risk":False},
      {"name":"โหระพา","unit":"kg","onhand":4,"perday":5,"days":0.8,"reorder":15,"risk":True}]},
  "admin":{"users":[{"name":"สมชาย ผ.จัดการ","role":"Admin","branch":"HQ","status":"active"},
      {"name":"นภา ก.การเงิน","role":"Finance","branch":"HQ","status":"active"},
      {"name":"PG Manager","role":"Branch","branch":"PG-Siam Paragon","status":"active"},
      {"name":"EQ Manager","role":"Branch","branch":"EQ-Emquartier","status":"inactive"}],
    "promoMap":[{"code":"HH-B1G1","promo":"Happy Hour Buy 1 Get 1","status":"on"},
      {"code":"SLV-10","promo":"Silver Hound Discount","status":"on"},
      {"code":"BLK-15","promo":"Black Hound Discount","status":"on"},
      {"code":"VP-49","promo":"Value Price 49 Baht","status":"off"}]},
}

# =====================================================================
# FILTER HELPERS — ช่วงวันที่ / สาขา / แบรนด์
# =====================================================================
def _range(date_from, date_to):
    """ไม่ส่งมา = default วันนี้ (frontend เก่าที่ยังไม่ส่ง param จะเห็นข้อมูลวันนี้เหมือนเดิม ไม่พัง)"""
    today = dt.date.today().isoformat()
    return date_from or today, date_to or today

def _codes_clause(column, codes):
    """คืน (' AND col IN (%s,...)', params) หรือ ('', ()) ถ้าไม่กรอง"""
    if not codes:
        return "", ()
    ph = ",".join(["%s"] * len(codes))
    return f" AND {column} IN ({ph})", tuple(codes)

# ---------------------------------------------------------------------
# ระดับบิล — อ่านจากตารางจริง ไม่ผ่าน v_bill_summary / v_promotion_summary
# view ทั้งสอง join BILLPAYTB ด้วย BillNo อย่างเดียว (ไม่มี ShopCode/BillDate) แต่เลขบิลซ้ำข้ามสาขา/วันได้
# → แถวงอก (30 วัน: บิลจริง 34,725 แต่ view ได้ 76,166 แถว, ยอด 23.8M กลายเป็น 56.3M, ส่วนลด 2.6M → 7.5M)
# คีย์บิลที่ถูกคือ (ShopCode, BillNo) — ไม่ซ้ำทั้งตาราง (เช็ค 2026-11-23)
# ---------------------------------------------------------------------
# บิลที่นับได้ = ตัดคู่บิล void (BillStatus 4 = บิลเดิมที่ถูก void, 5 = บิล VOID) — ตรงกับ app.py (_CountBill)
BILL_OK = "(h.BillStatus IS NULL OR h.BillStatus NOT IN ('4','5'))"
# member ต่อบิล: BILLPAYTB tendtype=MEM มีได้หลายแถวต่อบิล → ยุบเหลือ 1 แถว/บิล ก่อน join (ใช้ 2 params: date_from, date_to)
# รหัสสมาชิก = BillHdrTB.Remark (ตรง MemberTB 98%) สำรองด้วย CardNo ที่ไม่ว่าง — CardNo มีแถวว่างปน
# (เช่นแถว "Black Hound Discount") เดิม MIN(CardNo) ได้ '' → tier หาไม่เจอ ขึ้น "อื่นๆ" 70%
# นิยาม member เดียวกับ app.py และ v_bill_summary: มีแถว MEM ใน BILLPAYTB (= มี Remark — ชุดบิลเดียวกัน)
MEM_JOIN = """LEFT JOIN (SELECT ShopCode, BillNo, MAX(NULLIF(CardNo,'')) CardNo FROM BILLPAYTB
                     WHERE tendtype='MEM' AND BillDate BETWEEN %s AND %s
                     GROUP BY ShopCode, BillNo) p
              ON p.ShopCode = h.ShopCode AND p.BillNo = h.BillNo"""
IS_MEM = "p.ShopCode IS NOT NULL"
MEM_CODE = "COALESCE(NULLIF(TRIM(h.Remark),''), p.CardNo)"

def resolve_shops(shop, brand):
    """shop: CSV รหัสสาขา, brand: GC/AH/BE (PosShopTB.ShTypCode) → รวมเป็น tuple รหัสสาขาเดียว หรือ None ถ้าไม่กรองเลย"""
    codes = {c.strip() for c in (shop or "").split(",") if c.strip()}
    if brand:
        rows = q("SELECT ShopCode FROM PosShopTB WHERE ShTypCode=%s", (brand,)) or []
        brand_codes = {r["ShopCode"] for r in rows}
        codes = (codes & brand_codes) if codes else brand_codes
    return tuple(codes) if codes else None

# =====================================================================
# REAL builders
# =====================================================================
def member_range(date_from, date_to, shop_codes=None):
    clause, params = _codes_clause("shop_code", shop_codes)
    r = q(f"""SELECT SUM(member_bills) m FROM sale_monitor_rt
              WHERE biz_date BETWEEN %s AND %s{clause}""", (date_from, date_to, *params))
    if r and r[0].get("m") is not None:
        return int(r[0]["m"])
    return SAMPLE["kpi"]["memberBills"]

def kpi_range(date_from, date_to, shop_codes=None):
    """KPI รวมของช่วงวันที่ที่เลือก จาก sale_monitor_rt — ใช้ทั้ง /api/bi และ /api/overview/sale-monitor"""
    clause, params = _codes_clause("shop_code", shop_codes)
    rt = q(f"""SELECT SUM(net_sales) base, SUM(discount) disc, SUM(bill_count) bills,
              SUM(cust_count) pax FROM sale_monitor_rt
              WHERE biz_date BETWEEN %s AND %s{clause}""", (date_from, date_to, *params))
    if not rt or not rt[0].get("bills"):
        return None
    r = rt[0]; base = float(r["base"] or 0); bills = int(r["bills"] or 0)
    return {"baseSale": round(base), "bills": bills,
            "discount": round(float(r["disc"] or 0)),
            "avgBill": round(base / bills) if bills else 0,
            "pax": int(r["pax"] or 0), "memberBills": member_range(date_from, date_to, shop_codes)}

def build_monitor(shop_codes=None):
    """Sale Monitor เรียลไทม์ (วันนี้ vs เมื่อวาน) — ไม่ผูกกับช่วงวันที่ที่เลือกในหน้าเว็บ กรองได้แค่สาขา/แบรนด์"""
    today = dt.date.today().isoformat()
    yday  = (dt.date.today() - dt.timedelta(days=1)).isoformat()
    clause, params = _codes_clause("shop_code", shop_codes)
    rows = q(f"""SELECT shop_code,shop_name,
              MAX(CASE WHEN biz_date=%s THEN net_sales END) today,
              MAX(CASE WHEN biz_date=%s THEN net_sales END) yday
            FROM sale_monitor_rt WHERE biz_date IN (%s,%s){clause}
            GROUP BY shop_code,shop_name ORDER BY today DESC LIMIT 5""",
            (today, yday, today, yday, *params))
    if rows is None:
        return None
    shops = [{"code":r["shop_code"],"name":r["shop_name"],
              "today":float(r["today"] or 0),"yday":float(r["yday"] or 0)} for r in rows]
    tot = q(f"""SELECT biz_date, SUM(net_sales) total FROM sale_monitor_rt
              WHERE biz_date IN (%s,%s){clause} GROUP BY biz_date""", (today, yday, *params)) or []
    tmap = {str(r["biz_date"]): float(r["total"] or 0) for r in tot}
    snap = q(f"SELECT MAX(snapshot_at) s FROM sale_monitor_rt WHERE biz_date=%s{clause}", (today, *params)) or [{}]
    s = snap[0].get("s")
    # "เมื่อวาน ณ เวลาเดียวกัน" ตัดที่เวลา snapshot ของข้อมูลวันนี้ (ไม่ใช่นาฬิกา server → ETL ช้าก็ยังเทียบช่วงเท่ากัน)
    # ชั่วโมงปัจจุบันคิดตามสัดส่วนนาที — เดิม hour_slot<=ชม.ปัจจุบัน นับเมื่อวานเต็มชั่วโมง → pace ติดลบเกินจริง
    cut = s or dt.datetime.now()
    # 1) ตรงนาทีจาก sale_monitor_minute (เหมือนหน้า Sale Monitor ใน app.py) 2) ไม่มี → ประมาณจากรายชั่วโมง
    cut_hhmm = cut.hour * 100 + cut.minute
    yst = q(f"""SELECT COUNT(*) n, SUM(CASE WHEN minute_slot < %s THEN net_sales
                               WHEN minute_slot = %s THEN net_sales * %s / 60 ELSE 0 END) s
              FROM sale_monitor_minute WHERE biz_date=%s{clause}""",
            (cut_hhmm, cut_hhmm, cut.second, yday, *params))
    if not yst or not int(yst[0].get("n") or 0):
        cut_h, frac = cut.hour, (cut.minute * 60 + cut.second) / 3600
        yst = q(f"""SELECT SUM(CASE WHEN hour_slot < %s THEN net_sales
                                   WHEN hour_slot = %s THEN net_sales * %s ELSE 0 END) s
                  FROM sale_monitor_interval WHERE biz_date=%s{clause}""",
                (cut_h, cut_h, frac, yday, *params)) or [{}]
    return {"snapshot": s.strftime("%H:%M") if s else dt.datetime.now().strftime("%H:%M"),
            "totalToday": tmap.get(today,0.0), "ydayFull": tmap.get(yday,0.0),
            "ydaySameTime": float((yst[0] or {}).get("s") or 0),
            "kpi": kpi_range(today, today, shop_codes) or SAMPLE["kpi"],
            "shops": shops}

def build_overview(date_from, date_to, shop_codes=None):
    """KPI + hourly + branches ของช่วงวันที่ที่เลือก จาก sale_monitor_*"""
    clause, params = _codes_clause("shop_code", shop_codes)
    rt = q(f"""SELECT shop_code, shop_name, SUM(net_sales) net_sales
              FROM sale_monitor_rt WHERE biz_date BETWEEN %s AND %s{clause}
              GROUP BY shop_code, shop_name""", (date_from, date_to, *params))
    if not rt:
        return None
    out = {}
    out["kpi"] = kpi_range(date_from, date_to, shop_codes) or SAMPLE["kpi"]
    out["branches"] = sorted(
        [{"name": r["shop_name"], "code": r["shop_code"], "v": round(float(r["net_sales"] or 0))} for r in rt],
        key=lambda x: -x["v"])
    hr = q(f"""SELECT hour_slot, SUM(net_sales) s, SUM(discount) d
              FROM sale_monitor_interval WHERE biz_date BETWEEN %s AND %s{clause}
              GROUP BY hour_slot ORDER BY hour_slot""", (date_from, date_to, *params))
    if hr:
        out["hours"]    = [r["hour_slot"] for r in hr]
        out["sales"]    = [round(float(r["s"] or 0)) for r in hr]
        out["discount"] = [round(float(r["d"] or 0)) for r in hr]
    d1, d2 = dt.date.fromisoformat(date_from), dt.date.fromisoformat(date_to)
    out["asOf"] = d1.strftime("%d/%m/%Y") if d1 == d2 else f'{d1.strftime("%d/%m/%Y")} – {d2.strftime("%d/%m/%Y")}'
    out["updated"] = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
    out["sample"] = False
    return out

def build_promos(date_from, date_to, shop_codes=None):
    """รวมเป็น 1 แถวต่อ (โปรโมชั่น, บิล) ก่อน แล้วค่อย join ยอดขายบิล — กัน NetBAmt ถูกนับซ้ำ
    ตอนบิลเดียวมีหลายรายการที่ใช้โปรโมชั่นเดียวกัน (ONLBillMultiDiscTB คือ 1 แถวต่อ 1 รายการสินค้า ไม่ใช่ต่อบิล)
    คีย์บิล = (ShopCode, BillNo) — ไม่ใช้ v_promotion_summary/v_bill_summary เพราะ join ด้วย BillNo อย่างเดียว แถวงอก"""
    clause, params = _codes_clause("ShopCode", shop_codes)
    rows = q(f"""SELECT bp.name, SUM(bp.disc_amt) amount, COUNT(*) bills, SUM(h.NetBAmt) related
            FROM (
                SELECT DiscName name, ShopCode, BillNo, SUM(DiscAmt) disc_amt
                FROM ONLBillMultiDiscTB
                WHERE BillDate BETWEEN %s AND %s{clause}
                GROUP BY DiscName, ShopCode, BillNo
            ) bp
            JOIN BillHdrTB h ON h.ShopCode=bp.ShopCode AND h.BillNo=bp.BillNo
            WHERE {BILL_OK}
            GROUP BY bp.name ORDER BY amount DESC LIMIT 12""", (date_from, date_to, *params))
    if rows is None:
        return None
    icons = ["ti-gift","ti-medal-2","ti-discount-2","ti-tag","ti-ticket","ti-percentage"]
    return [{"name": r["name"] or "(ไม่ระบุ)", "amount": round(float(r["amount"] or 0)),
             "bills": int(r["bills"] or 0), "related": round(float(r["related"] or 0)),
             "icon": icons[i % len(icons)]} for i, r in enumerate(rows)]

def build_member(date_from, date_to, shop_codes=None):
    """คืน dict {memberSplit, member} รวมกัน — เผื่อ /api/bi ไปแยก key เอง"""
    clause, params = _codes_clause("h.ShopCode", shop_codes)
    where = f"WHERE h.BillDate BETWEEN %s AND %s AND {BILL_OK}{clause}"
    bp = (date_from, date_to, date_from, date_to, *params)   # MEM_JOIN 2 ตัว + WHERE
    rows = q(f"""SELECT CASE WHEN {IS_MEM} THEN 'Member' ELSE 'Guest' END CustomerType,
                     COUNT(*) bills, SUM(h.NetBAmt) sales
              FROM BillHdrTB h {MEM_JOIN} {where}
              GROUP BY CustomerType""", bp)
    if not rows:
        return None
    g = next((r for r in rows if r["CustomerType"] == "Guest"), {"bills": 0, "sales": 0})
    m = next((r for r in rows if r["CustomerType"] == "Member"), {"bills": 0, "sales": 0})
    g_bills, g_sales = int(g["bills"] or 0), float(g["sales"] or 0)
    m_bills, m_sales = int(m["bills"] or 0), float(m["sales"] or 0)
    tot_bills = g_bills + m_bills
    # BILLPAYTB.Prefix ไม่ตรง MemTypetb ในข้อมูลจริง → ใช้ MemberTB.MemCode → MemType แทน ได้ tier จริง
    # (Silver/Black/White/Platinum Hound ฯลฯ) — ระบุ h.ShopCode ชัดเจน เพราะ MemTypetb ก็มีคอลัมน์ Shopcode
    tiers = q(f"""SELECT COALESCE(mt.MemTNam1,'อื่นๆ') name, COUNT(*) n
              FROM BillHdrTB h {MEM_JOIN}
              LEFT JOIN MemberTB m ON m.MemCode = {MEM_CODE}
              LEFT JOIN MemTypetb mt ON mt.MemType = m.MemType
              {where} AND {IS_MEM}
              GROUP BY name ORDER BY n DESC LIMIT 8""", bp) or []
    tier_total = sum(int(t["n"]) for t in tiers) or 1
    rep = q(f"""SELECT COUNT(*)/NULLIF(COUNT(DISTINCT {MEM_CODE}),0) r
              FROM BillHdrTB h {MEM_JOIN}
              {where} AND {IS_MEM}""", bp)
    repeat = float((rep[0] or {}).get("r") or 0) if rep else 0.0
    return {
        "memberSplit": {"guest": round(g_bills / tot_bills * 100, 1) if tot_bills else 0,
                         "member": round(m_bills / tot_bills * 100, 1) if tot_bills else 0},
        "member": {
            "guest": {"sales": round(g_sales), "bills": g_bills, "avg": round(g_sales / g_bills) if g_bills else 0},
            "member": {"sales": round(m_sales), "bills": m_bills, "avg": round(m_sales / m_bills) if m_bills else 0},
            "tiers": [{"name": t["name"], "pct": round(int(t["n"]) / tier_total * 100)} for t in tiers],
            "repeat": round(repeat, 1),
        },
    }

def build_top_items(date_from, date_to, shop_codes=None):
    clause, params = _codes_clause("ShopCode", shop_codes)
    rows = q(f"""SELECT ItemDes1 name, SUM(Qty) qty, SUM(NetAmt) sales
              FROM v_bill_items WHERE BillDate BETWEEN %s AND %s{clause}
              GROUP BY ItemCode, ItemDes1 ORDER BY sales DESC LIMIT 10""", (date_from, date_to, *params))
    if rows is None:
        return None
    return [{"name": r["name"], "qty": int(r["qty"] or 0), "sales": round(float(r["sales"] or 0))} for r in rows]

def build_top_items_full(date_from, date_to, shop_codes=None):
    """รายการขายทั้งหมด (ไม่ limit) พร้อม Department/Category/จำนวนบิล — ใช้กับหน้า 'รายการขายดี' ที่กรอง/เรียงฝั่ง client"""
    clause, params = _codes_clause("ShopCode", shop_codes)
    rows = q(f"""SELECT ItemDes1 name, COALESCE(NULLIF(TRIM(Department),''),'อื่นๆ') dept,
              COALESCE(NULLIF(TRIM(Category),''),'อื่นๆ') cat,
              SUM(Qty) qty, SUM(NetAmt) sales, COUNT(DISTINCT ShopCode, BillNo) txn
              FROM v_bill_items WHERE BillDate BETWEEN %s AND %s{clause}
              GROUP BY ItemCode, ItemDes1, dept, cat ORDER BY sales DESC""", (date_from, date_to, *params))
    if rows is None:
        return None
    return [{"name": r["name"], "dept": r["dept"], "cat": r["cat"],
              "qty": int(r["qty"] or 0), "sales": round(float(r["sales"] or 0)),
              "txn": int(r["txn"] or 0)} for r in rows]

def build_item_cats(date_from, date_to, shop_codes=None):
    clause, params = _codes_clause("ShopCode", shop_codes)
    rows = q(f"""SELECT COALESCE(Category,'อื่นๆ') name, SUM(NetAmt) sales
              FROM v_bill_items WHERE BillDate BETWEEN %s AND %s{clause}
              GROUP BY Category ORDER BY sales DESC LIMIT 8""", (date_from, date_to, *params))
    if not rows:
        return None
    total = sum(float(r["sales"] or 0) for r in rows) or 1
    return [{"name": r["name"], "pct": round(float(r["sales"] or 0) / total * 100)} for r in rows]

def build_menu_eng():
    """menu_engineering_summary คำนวณรวมทั้งเชน ไม่มี ShopCode ให้กรอง — อัปเดตจาก cron รายวัน (ดู etl_sync.py --menu-engineering)"""
    latest = q("SELECT MAX(date_to) d FROM menu_engineering_summary")
    if not latest or not latest[0].get("d"):
        return None, None
    d_to = latest[0]["d"]
    rows = q("""SELECT item_name name, total_qty qty, cm_ratio margin
              FROM menu_engineering_summary WHERE date_to=%s
              ORDER BY total_sales DESC LIMIT 20""", (d_to,))
    if not rows:
        return None, None
    items = [{"name": r["name"], "pop": int(r["qty"] or 0),
              "margin": round(float(r["margin"] or 0) * 100)} for r in rows]
    return items, str(d_to)

def build_bill_hist(date_from, date_to, shop_codes=None):
    clause, params = _codes_clause("h.ShopCode", shop_codes)
    rows = q(f"""SELECT CASE
                WHEN h.NetBAmt < 300 THEN '<300'
                WHEN h.NetBAmt < 500 THEN '300–500'
                WHEN h.NetBAmt < 700 THEN '500–700'
                WHEN h.NetBAmt < 1000 THEN '700–1000'
                WHEN h.NetBAmt < 1500 THEN '1000–1500'
                ELSE '>1500' END AS b, COUNT(*) n
              FROM BillHdrTB h WHERE h.BillDate BETWEEN %s AND %s AND {BILL_OK}{clause}
              GROUP BY b""", (date_from, date_to, *params))
    if rows is None:
        return None
    order = ["<300", "300–500", "500–700", "700–1000", "1000–1500", ">1500"]
    m = {r["b"]: int(r["n"]) for r in rows}
    return [{"b": b, "n": m.get(b, 0)} for b in order]

def build_bill_scatter(date_from, date_to, shop_codes=None):
    clause, params = _codes_clause("h.ShopCode", shop_codes)
    rows = q(f"""SELECT h.NoPerson, h.NetBAmt FROM BillHdrTB h
              WHERE h.BillDate BETWEEN %s AND %s AND h.NoPerson>0 AND {BILL_OK}{clause}
              ORDER BY h.BillDate DESC, h.BillTime DESC LIMIT 300""", (date_from, date_to, *params))
    if rows is None:
        return None
    return [[int(r["NoPerson"]), round(float(r["NetBAmt"] or 0))] for r in rows]

def build_daily_trend(date_to, shop_codes=None):
    """เทรนด์ย้อนหลัง 7 วัน นับถอยจาก date_to (ไม่ใช่ default วันนี้เหมือนส่วนอื่น)"""
    end = dt.date.fromisoformat(date_to) if date_to else dt.date.today()
    start = end - dt.timedelta(days=6)
    clause, params = _codes_clause("shop_code", shop_codes)
    rows = q(f"""SELECT biz_date, SUM(net_sales) s FROM sale_monitor_rt
              WHERE biz_date BETWEEN %s AND %s{clause}
              GROUP BY biz_date""", (start.isoformat(), end.isoformat(), *params))
    if rows is None:
        return None
    m = {str(r["biz_date"]): float(r["s"] or 0) for r in rows}
    days = [(start + dt.timedelta(days=i)).isoformat() for i in range(7)]
    return [round(m.get(d, 0)) for d in days]

def build_forecast_reorder(shop_codes=None):
    """เฉพาะตาราง reorder จริงจาก v_ingredient_forecast — กราฟจริง-vs-พยากรณ์ 7 วันตัดออก
    (v_ingredient_forecast เป็น snapshot ปัจจุบัน ไม่ใช่ time-series ไม่มีข้อมูลจริงมาทำกราฟ ไม่ปั้นตัวเลขเทียม)"""
    clause, params = _codes_clause("shop_code", shop_codes)
    rows = q(f"""SELECT item_name name, unit, current_stock onhand, suggest_qty reorder, status
              FROM v_ingredient_forecast WHERE 1=1{clause}
              ORDER BY FIELD(status,'REORDER','WATCH','AUDIT','OK','MANUAL','NODATA'), item_name
              LIMIT 30""", tuple(params))
    if rows is None:
        return None
    out = []
    for r in rows:
        status = (r.get("status") or "").upper()
        out.append({"name": r["name"], "unit": r["unit"],
                     "onhand": round(float(r["onhand"] or 0), 1),
                     "reorder": round(float(r["reorder"] or 0), 1),
                     "risk": status in ("REORDER", "WATCH")})
    return out

def build_admin_users():
    rows = q("SELECT username, full_name, role, active FROM users ORDER BY username") or []
    return [{"username": r["username"], "full_name": r["full_name"] or r["username"],
              "role": r["role"], "active": bool(r["active"])} for r in rows]

# =====================================================================
# ENDPOINTS
# =====================================================================
@app.get("/api/health")
def health():
    return {"ok": True, "db": "up" if q("SELECT 1") is not None else "down (sample mode)",
            "time": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

@app.get("/healthz")
def healthz():
    """alias สั้น สำหรับ health-check ของ tunnel/uptime monitor"""
    return health()

@app.post("/api/auth/login")
def auth_login(payload: dict = Body(...)):
    u = (payload.get("username") or "").strip()
    pw = payload.get("password") or ""
    rows = q("SELECT username,password,full_name,role,active FROM users WHERE username=%s", (u,))
    row = rows[0] if rows else None
    if not row or not row["active"] or not verify_pw(pw, row["password"]):
        return {"ok": False, "error": "username หรือ password ไม่ถูกต้อง"}
    token = make_token(row["username"], row["role"])
    return {"ok": True, "token": token,
            "user": {"username": row["username"], "full_name": row["full_name"], "role": row["role"]}}

@app.get("/api/auth/me")
def auth_me(user=Depends(current_user)):
    return {"username": user["sub"], "role": user["role"]}

@app.get("/api/overview/sale-monitor")
def sale_monitor(key: str = Query(None), x_api_key: str = Header(None),
                  shop: str = Query(None), brand: str = Query(None),
                  user=Depends(current_user)):
    if not check_key(key, x_api_key):
        return {"error": "invalid api key"}
    shop_codes = resolve_shops(shop, brand)
    mon = build_monitor(shop_codes) or dict(SAMPLE["monitor"])
    mon.setdefault("kpi", SAMPLE["kpi"])   # ให้หน้าเว็บอัปเดต Hero/KPI สดได้เสมอ
    return mon

@app.get("/api/bi")
def bi(key: str = Query(None), x_api_key: str = Header(None),
       date_from: str = Query(None), date_to: str = Query(None),
       shop: str = Query(None), brand: str = Query(None),
       user=Depends(current_user)):          # บังคับ login (2026-11-23) — เดิมไม่ส่งอะไรมาก็ได้ข้อมูล
    if not check_key(key, x_api_key):
        return {"error": "invalid api key"}

    d_from, d_to = _range(date_from, date_to)
    shop_codes = resolve_shops(shop, brand)

    data = dict(SAMPLE)                            # เริ่มจากตัวอย่าง (กันพังถ้า query ไหนล้ม)
    data["dateFrom"], data["dateTo"] = d_from, d_to

    ov = build_overview(d_from, d_to, shop_codes)
    if ov:
        data.update(ov)
    data["monitor"] = build_monitor(shop_codes) or SAMPLE["monitor"]

    promos = build_promos(d_from, d_to, shop_codes)
    if promos is not None:
        data["promos"] = promos

    mem = build_member(d_from, d_to, shop_codes)
    if mem is not None:
        data["memberSplit"] = mem["memberSplit"]
        data["member"] = mem["member"]

    top = build_top_items(d_from, d_to, shop_codes)
    if top is not None:
        data["topItems"] = top

    cats = build_item_cats(d_from, d_to, shop_codes)
    if cats is not None:
        data["itemCats"] = cats

    me_items, me_as_of = build_menu_eng()
    if me_items is not None:
        data["menuEng"] = me_items
        data["menuEngAsOf"] = me_as_of

    hist = build_bill_hist(d_from, d_to, shop_codes)
    if hist is not None:
        data["billHist"] = hist

    scatter = build_bill_scatter(d_from, d_to, shop_codes)
    if scatter is not None:
        data["billScatter"] = scatter

    trend = build_daily_trend(d_to, shop_codes)
    if trend is not None:
        data["dailyTrend"] = trend

    reorder = build_forecast_reorder(shop_codes)
    if reorder is not None:
        # เก็บ ingredient/actual/predict จาก SAMPLE ไว้ก่อน (ยังไม่มีต้นทางจริง) ทับแค่ reorder ที่เป็นของจริง
        fc = dict(SAMPLE["forecast"])
        fc["reorder"] = reorder
        data["forecast"] = fc

    # admin.users ส่งเฉพาะ token ที่ role=admin เท่านั้น — กันข้อมูล user จริงหลุดไปกับ JSON สาธารณะ
    admin_users = build_admin_users() if (user and user.get("role") == "admin") else None
    data["admin"] = {"users": admin_users if admin_users is not None else [],
                      "promoMap": SAMPLE["admin"]["promoMap"]}

    return data

@app.get("/api/bi/items")
def bi_items(key: str = Query(None), x_api_key: str = Header(None),
             date_from: str = Query(None), date_to: str = Query(None),
             shop: str = Query(None), brand: str = Query(None),
             user=Depends(current_user)):
    if not check_key(key, x_api_key):
        return {"error": "invalid api key"}
    d_from, d_to = _range(date_from, date_to)
    shop_codes = resolve_shops(shop, brand)
    items = build_top_items_full(d_from, d_to, shop_codes)
    if items is None:
        return {"dateFrom": d_from, "dateTo": d_to, "items": [], "error": "db unavailable"}
    return {"dateFrom": d_from, "dateTo": d_to, "items": items}

@app.get("/")
def root():
    return {"service": "Greyhound BI API", "endpoints": [
        "/api/bi", "/api/bi/items", "/api/overview/sale-monitor", "/api/health",
        "/api/auth/login", "/api/auth/me",
    ]}
