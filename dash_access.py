"""
dash_access.py — สิทธิ์ดูเมนูราย user ของหน้า Streamlit (app_main.py พอร์ต 83)

- เก็บในตาราง promo_dash_users (ตารางเดียวกับ login ใน dash_auth.py) 2 คอลัมน์:
    role          'admin' = เห็นทุกหน้า + หน้า Admin (จัดสิทธิ์ได้) · 'user' = เห็นเฉพาะ allowed_pages
    allowed_pages รหัสหน้าคั่นด้วย comma เช่น "sale_monitor,overview"
- user ใหม่ = role 'user' + ไม่มีหน้า → ไม่เห็นเมนูใดจนกว่า admin จะติ๊กให้ (หน้า Admin › สิทธิ์ผู้ใช้)
- ตอนเพิ่มคอลัมน์ครั้งแรก user เดิมทุกคนเป็น admin (คงสิทธิ์เดิมที่เห็นทุกหน้า)
- app_main.py อ่านสิทธิ์ใหม่ทุก rerun → เปลี่ยนสิทธิ์แล้วมีผลทันที ไม่ต้อง login ใหม่
- app_main.py reload ไฟล์นี้เองเมื่อแก้ (service ไม่มี file watcher — ดูเหตุผลใน app.py)
- จัดการบัญชี (เพิ่ม / ตั้งรหัสใหม่ / เปิด-ปิด): หน้า Admin › 👤 เพิ่ม/จัดการผู้ใช้ (create_user, reset_password, set_active)
- CLI: python manage_users.py role|pages ...
"""
import re
import pymysql
import dash_auth

TABLE = dash_auth.TABLE
ADMIN_PAGE = "⚙️ Admin"     # เห็นเฉพาะ role admin — ไม่อยู่ในรายการให้ติ๊ก

# (รหัส, label ในเมนู) ตามลำดับเมนู — label ต้องตรงกับ `page == "..."` ใน app_main.py ทุกตัวอักษร (รวม emoji)
PAGES = [
    ("sale_monitor",  "📡 Sale Monitor"),
    ("overview",      "🏠 Overview"),
    ("top_items",     "🍽️ รายการขายดี"),
    ("time_analysis", "⏱️ วิเคราะห์ช่วงเวลา"),
    ("member",        "👥 Member / Guest"),
    ("frequency",     "🔁 ความถี่ลูกค้า"),
    ("per_bill",      "🧾 Per Bill Analysis"),
    ("promotion",     "🎁 วิเคราะห์โปรโมชั่น"),
    ("menu_eng",      "🧑‍🏫 Menu Engineering"),
    ("ing_forecast",  "🥩 Ingredient Forecast"),
]
PAGE_KEYS = [k for k, _ in PAGES]
LABEL = dict(PAGES)
ROLES = ("admin", "user")


def ensure_columns(cur):
    """เพิ่ม role / allowed_pages ถ้ายังไม่มี (MySQL 8 ไม่มี ADD COLUMN IF NOT EXISTS)"""
    cur.execute("SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=DATABASE() "
                "AND TABLE_NAME=%s AND COLUMN_NAME IN ('role','allowed_pages')", (TABLE,))
    have = {r[0] for r in cur.fetchall()}
    if "role" not in have:
        cur.execute(f"ALTER TABLE {TABLE} ADD COLUMN role VARCHAR(10) NOT NULL DEFAULT 'user' AFTER full_name")
        cur.execute(f"UPDATE {TABLE} SET role='admin'")      # user เดิมเห็นทุกหน้าอยู่แล้ว → คงไว้
    if "allowed_pages" not in have:
        cur.execute(f"ALTER TABLE {TABLE} ADD COLUMN allowed_pages VARCHAR(500) NULL AFTER role")


def parse_pages(s):
    """"a,b,x" → ['a','b'] (ตัดรหัสที่ไม่รู้จัก, เรียงตามลำดับเมนู)"""
    got = {k.strip() for k in (s or "").split(",")}
    return [k for k in PAGE_KEYS if k in got]


def join_pages(keys):
    return ",".join(parse_pages(",".join(keys))) or None


def get_access(username):
    """คืน (role, set ของรหัสหน้า) — บัญชีไม่พบ/อ่านไม่ได้ = ('user', set()) คือไม่เห็นอะไร (fail closed)"""
    sql = f"SELECT role, allowed_pages FROM {TABLE} WHERE username=%s AND active=1"
    try:
        with dash_auth._conn() as c, c.cursor() as cur:
            try:
                cur.execute(sql, (username,))
            except pymysql.err.OperationalError as e:
                if e.args[0] != 1054:                       # 1054 = ยังไม่มีคอลัมน์ → เพิ่มแล้วลองใหม่
                    raise
                ensure_columns(cur)
                cur.execute(sql, (username,))
            row = cur.fetchone()
    except Exception:
        return "user", set()
    if not row:
        return "user", set()
    role = row[0] if row[0] in ROLES else "user"
    return role, set(parse_pages(row[1]))


def nav_labels(role, pages):
    """label เมนูที่ user นี้เห็น (admin = ทุกหน้า + Admin)"""
    if role == "admin":
        return [lbl for _, lbl in PAGES] + [ADMIN_PAGE]
    return [lbl for k, lbl in PAGES if k in pages]


def list_access():
    with dash_auth._conn() as c, c.cursor() as cur:
        ensure_columns(cur)
        cur.execute(f"SELECT username, full_name, active, role, allowed_pages, last_login, locked_until "
                    f"FROM {TABLE} ORDER BY username")
        return [dict(username=r[0], full_name=r[1], active=bool(r[2]), role=r[3] if r[3] in ROLES else "user",
                     pages=set(parse_pages(r[4])), last_login=r[5], locked_until=r[6]) for r in cur.fetchall()]


# ── จัดการบัญชี (หน้า Admin › ผู้ใช้) — กติกาเดียวกับ manage_users.py ──
MIN_PW = 4                                   # เท่ากับ manage_users.py add
USERNAME_RE = re.compile(r"^[A-Za-z0-9._-]{3,50}$")


def _check_pw(pw, pw2):
    if pw != pw2:
        return "รหัสผ่านทั้ง 2 ช่องไม่ตรงกัน"
    if len(pw or "") < MIN_PW:
        return f"รหัสผ่านสั้นเกินไป (อย่างน้อย {MIN_PW} ตัว)"
    return None


def create_user(username, full_name, pw, pw2, role="user", pages=()):
    """เพิ่มบัญชีใหม่ — username ซ้ำ = ไม่ทับของเดิม (ใช้ reset_password แทน) · คืน (ok, ข้อความ)"""
    username = (username or "").strip()
    full_name = (full_name or "").strip() or None
    if not USERNAME_RE.match(username):
        return False, "Username ต้องเป็น a-z A-Z 0-9 . _ - ยาว 3–50 ตัว (ไม่มีเว้นวรรค/ภาษาไทย)"
    err = _check_pw(pw, pw2)
    if err:
        return False, err
    role = role if role in ROLES else "user"
    with dash_auth._conn() as c, c.cursor() as cur:
        cur.execute(dash_auth.DDL)
        ensure_columns(cur)
        try:
            cur.execute(f"INSERT INTO {TABLE} (username, password_hash, full_name, role, allowed_pages) "
                        f"VALUES (%s,%s,%s,%s,%s)",
                        (username, dash_auth.hash_pw(pw), full_name, role, join_pages(pages)))
        except pymysql.err.IntegrityError:
            return False, f"มี Username '{username}' อยู่แล้ว — ถ้าจะเปลี่ยนรหัสให้ใช้ส่วนตั้งรหัสผ่านใหม่ด้านล่าง"
    what = "admin (เห็นทุกหน้า)" if role == "admin" else (f"{len(parse_pages(','.join(pages)))} หน้า" if pages else "ยังไม่มีสิทธิ์ดูเมนู")
    return True, f"เพิ่มผู้ใช้ {username} แล้ว — {what}"


def reset_password(username, pw, pw2):
    """ตั้งรหัสใหม่ + ปลดล็อกที่ใส่รหัสผิด"""
    err = _check_pw(pw, pw2)
    if err:
        return False, err
    with dash_auth._conn() as c, c.cursor() as cur:
        n = cur.execute(f"UPDATE {TABLE} SET password_hash=%s, failed_count=0, locked_until=NULL, lock_count=0 "
                        f"WHERE username=%s", (dash_auth.hash_pw(pw), username))
    return (True, f"ตั้งรหัสผ่านใหม่ให้ {username} แล้ว (ปลดล็อกด้วย)") if n else (False, f"ไม่พบ {username}")


def set_active(username, active, by):
    """เปิด/ปิดบัญชี — ปิดแล้ว session ที่เปิดค้างจะไม่เห็นเมนูตั้งแต่ rerun ถัดไป (get_access กรอง active=1)"""
    if not active and username == by:
        return False, "ปิดบัญชีของตัวเองไม่ได้"
    with dash_auth._conn() as c, c.cursor() as cur:
        if not active:
            cur.execute(f"SELECT COUNT(*) FROM {TABLE} WHERE role='admin' AND active=1 AND username<>%s", (username,))
            if not cur.fetchone()[0]:
                return False, "ต้องเหลือ admin ที่เปิดใช้งานอย่างน้อย 1 คน"
        n = cur.execute(f"UPDATE {TABLE} SET active=%s WHERE username=%s", (1 if active else 0, username))
    return (True, f"{'เปิด' if active else 'ปิด'}บัญชี {username} แล้ว") if n else (False, f"ไม่พบ {username} (หรือเป็นค่านี้อยู่แล้ว)")


def save_access(rows, by):
    """rows = [{username, role, pages(iterable)}] — คืน (ok, ข้อความ)
    กันล็อกตัวเองออก: ห้ามถอด admin ของคนที่กำลังบันทึก และต้องเหลือ admin ที่ active อย่างน้อย 1 คน"""
    cur_state = {r["username"]: r for r in list_access()}
    new = {}
    for r in rows:
        u = r["username"]
        if u not in cur_state:
            continue
        role = r["role"] if r["role"] in ROLES else "user"
        new[u] = (role, join_pages(r.get("pages") or []))
    if by in new and new[by][0] != "admin":
        return False, "ถอดสิทธิ์ admin ของตัวเองไม่ได้ (กันล็อกตัวเองออกจากหน้า Admin)"
    admins = [u for u, s in cur_state.items() if s["active"] and new.get(u, (s["role"],))[0] == "admin"]
    if not admins:
        return False, "ต้องมี admin ที่เปิดใช้งานอย่างน้อย 1 คน"
    changed = [u for u, (role, pages) in new.items()
               if role != cur_state[u]["role"] or (pages or None) != join_pages(cur_state[u]["pages"])]
    if not changed:
        return True, "ไม่มีอะไรเปลี่ยน"
    with dash_auth._conn() as c, c.cursor() as cur:
        cur.executemany(f"UPDATE {TABLE} SET role=%s, allowed_pages=%s WHERE username=%s",
                        [(new[u][0], new[u][1], u) for u in changed])
    return True, f"บันทึกแล้ว {len(changed)} บัญชี: {', '.join(changed)}"
