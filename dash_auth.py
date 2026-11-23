"""
dash_auth.py — login หน้า Streamlit (app.py พอร์ต 83)

- บัญชีอยู่ในตาราง gh_promotion.promo_dash_users (แยกจาก `users` ของเว็บ BI ที่ใช้ร่วมกับ recipe_master)
- รหัสผ่านเก็บเป็น pbkdf2$sha256$iters$salt$hex (รูปแบบเดียวกับ greyhound_bi_api.verify_pw)
- ผิดเกิน MAX_FAILED ครั้ง → ล็อกบัญชี LOCK_MINUTES นาที
- จำการ login ด้วย cookie ที่เซ็น HMAC (DASH_SESSION_SECRET ใน .env) อายุ SESSION_HOURS ชม.
  → กด refresh / เปิดแท็บใหม่ไม่ต้อง login ซ้ำ; ปิดบัญชีใน DB แล้ว cookie ใช้ไม่ได้ทันที
- จัดการบัญชี: python manage_users.py add|passwd|disable|enable|list
"""
import base64, hashlib, hmac, os, secrets, time
import datetime as dt

import pymysql
import gh_env

TABLE = "promo_dash_users"
MAX_FAILED = 5
LOCK_MINUTES = 5
SESSION_HOURS = 12
COOKIE = "gh_dash_session"

DDL = f"""CREATE TABLE IF NOT EXISTS {TABLE} (
    username      VARCHAR(50)  NOT NULL PRIMARY KEY,
    password_hash VARCHAR(200) NOT NULL,
    full_name     VARCHAR(100) NULL,
    active        TINYINT(1)   NOT NULL DEFAULT 1,
    failed_count  INT          NOT NULL DEFAULT 0,
    locked_until  DATETIME     NULL,
    last_login    DATETIME     NULL,
    created_at    DATETIME     NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"""


def _conn():
    return pymysql.connect(**gh_env.mysql_cfg("root", autocommit=True))


# ── password hash ──
def hash_pw(pw, iters=200_000):
    salt = secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), iters)
    return f"pbkdf2$sha256${iters}${salt}${dk.hex()}"


def verify_pw(pw, stored):
    try:
        _, _, iters, salt, hexdk = stored.split("$")
        dk = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), int(iters))
        return hmac.compare_digest(dk.hex(), hexdk)
    except Exception:
        return False


# ── login ──
def check_login(username, pw):
    """คืน (ok, ข้อความ, full_name)"""
    username = (username or "").strip()
    if not username or not pw:
        return False, "กรอก Username และ Password", None
    with _conn() as c, c.cursor() as cur:
        cur.execute(DDL)
        cur.execute(f"SELECT password_hash, full_name, active, failed_count, locked_until FROM {TABLE} "
                    f"WHERE username=%s", (username,))
        row = cur.fetchone()
        now = dt.datetime.now()
        if not row:
            return False, "Username หรือ Password ไม่ถูกต้อง", None
        pw_hash, full_name, active, failed, locked_until = row
        if not active:
            return False, "บัญชีนี้ถูกปิดใช้งาน", None
        if locked_until and locked_until > now:
            mins = max(1, round((locked_until - now).total_seconds() / 60))
            return False, f"ใส่รหัสผิดหลายครั้ง — ลองใหม่ในอีก {mins} นาที", None
        if not verify_pw(pw, pw_hash):
            failed = (failed or 0) + 1
            lock = now + dt.timedelta(minutes=LOCK_MINUTES) if failed >= MAX_FAILED else None
            cur.execute(f"UPDATE {TABLE} SET failed_count=%s, locked_until=%s WHERE username=%s",
                        (0 if lock else failed, lock, username))
            if lock:
                return False, f"ใส่รหัสผิด {MAX_FAILED} ครั้ง — ล็อก {LOCK_MINUTES} นาที", None
            return False, "Username หรือ Password ไม่ถูกต้อง", None
        cur.execute(f"UPDATE {TABLE} SET failed_count=0, locked_until=NULL, last_login=%s WHERE username=%s",
                    (now, username))
        return True, "", full_name or username


def user_active(username):
    try:
        with _conn() as c, c.cursor() as cur:
            cur.execute(f"SELECT full_name, active FROM {TABLE} WHERE username=%s", (username,))
            row = cur.fetchone()
        return (row[0] or username) if row and row[1] else None
    except Exception:
        return None


# ── session cookie (เซ็น HMAC) ──
def _secret():
    return gh_env.need("DASH_SESSION_SECRET").encode()


def _b64(b):
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def make_token(username, hours=SESSION_HOURS):
    exp = int(time.time()) + hours * 3600
    body = f"{_b64(username.encode())}.{exp}"
    sig = _b64(hmac.new(_secret(), body.encode(), hashlib.sha256).digest())
    return f"{body}.{sig}"


def read_token(token):
    """คืน username ถ้า token ถูกต้องและยังไม่หมดอายุ"""
    try:
        u64, exp, sig = token.split(".")
        body = f"{u64}.{exp}"
        good = _b64(hmac.new(_secret(), body.encode(), hashlib.sha256).digest())
        if not hmac.compare_digest(sig, good) or int(exp) < time.time():
            return None
        return base64.urlsafe_b64decode(u64 + "=" * (-len(u64) % 4)).decode()
    except Exception:
        return None


# ── UI (Streamlit) ──
_COOKIE_JS = r"""
export default function({ data }) {
  if (!data || !data.cmd) return;
  if (data.cmd === "set") {
    document.cookie = data.name + "=" + data.value + "; max-age=" + data.maxAge + "; path=/; SameSite=Lax";
  } else if (data.cmd === "clear") {
    document.cookie = data.name + "=; max-age=0; path=/; SameSite=Lax";
  }
}
"""


def _cookie_cmd(st, cmd, value=""):
    """ตั้ง/ลบ cookie ฝั่งเบราว์เซอร์ (component v2 รันใน DOM หน้าหลัก) — พังก็แค่ไม่จำ login"""
    try:
        comp = st.components.v2.component("gh_dash_cookie", html="<span></span>\n<i></i>",
                                          js=_COOKIE_JS, isolate_styles=True)
        comp(data={"cmd": cmd, "name": COOKIE, "value": value, "maxAge": SESSION_HOURS * 3600},
             key=f"_gh_cookie_{cmd}_{int(time.time())}", height=0)
    except Exception:
        pass


def require_login(st):
    """เรียกต้น app.py — ยังไม่ login = แสดงฟอร์มแล้ว st.stop(); login แล้วคืน (username, full_name)"""
    ss = st.session_state
    if ss.get("auth_user"):
        return ss["auth_user"], ss.get("auth_name", ss["auth_user"])

    # มี cookie ที่ยังไม่หมดอายุ → เข้าเลย (ตรวจว่าบัญชียังเปิดอยู่)
    # หลังกดออกจากระบบ: ไม่ใช้ cookie เดิมใน session นี้ (st.context.cookies ยังเป็นค่าตอนเปิดหน้า) + สั่งลบ cookie
    if ss.get("auth_logout"):
        _cookie_cmd(st, "clear")
    else:
        try:
            tok = st.context.cookies.get(COOKIE)
        except Exception:
            tok = None
        u = read_token(tok) if tok else None
        name = user_active(u) if u else None
        if u and name:
            ss["auth_user"], ss["auth_name"] = u, name
            return u, name

    st.markdown("<style>section[data-testid='stSidebar']{display:none}</style>", unsafe_allow_html=True)
    box = st.empty()
    with box.container():
        _, mid, _ = st.columns([1, 1.2, 1])
        with mid:
            st.markdown(
                "<div style='text-align:center;margin:60px 0 18px'>"
                "<div style='font-size:42px'>🐾</div>"
                "<div style='font-size:22px;font-weight:800;color:#1f2328;letter-spacing:2px'>GREYHOUND CAFÉ</div>"
                "<div style='color:#6b7280;font-size:15px'>Promotion Intelligence · เข้าสู่ระบบ</div></div>",
                unsafe_allow_html=True)
            with st.form("gh_login", clear_on_submit=False):
                user = st.text_input("Username", key="login_user")
                pw = st.text_input("Password", type="password", key="login_pw")
                ok = st.form_submit_button("เข้าสู่ระบบ", use_container_width=True, type="primary")
            msg = st.empty()
    if ok:
        good, err, full_name = check_login(user, pw)
        if good:
            ss["auth_user"], ss["auth_name"] = user.strip(), full_name
            ss.pop("auth_logout", None)
            box.empty()
            st.markdown("<style>section[data-testid='stSidebar']{display:flex}</style>", unsafe_allow_html=True)
            _cookie_cmd(st, "set", make_token(user.strip()))
            return ss["auth_user"], full_name
        msg.error(err)
    st.stop()


def logout(st):
    """เรียกแล้ว st.rerun() — หน้า login รอบถัดไปจะลบ cookie และไม่ auto-login ด้วย cookie เดิม"""
    for k in ("auth_user", "auth_name"):
        st.session_state.pop(k, None)
    st.session_state["auth_logout"] = True
