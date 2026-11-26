"""
manage_users.py — จัดการบัญชี login ของหน้า Streamlit (app.py พอร์ต 83) ในตาราง promo_dash_users

  python manage_users.py list
  python manage_users.py add <username> "ชื่อ-นามสกุล"     # ถามรหัสผ่าน (ไม่ขึ้นจอ) หรือส่งทาง stdin
  python manage_users.py passwd <username>
  python manage_users.py disable <username>
  python manage_users.py enable <username>
  python manage_users.py unlock <username>               # ปลดล็อกที่ใส่รหัสผิดเกิน 5 ครั้ง

สิทธิ์ดูเมนู (dash_access.py — ทำในหน้า Admin › สิทธิ์ผู้ใช้ ก็ได้):
  python manage_users.py role <username> admin|user     # admin = เห็นทุกหน้า + หน้า Admin
  python manage_users.py pages                          # รายการรหัสหน้า
  python manage_users.py pages <username> overview,sale_monitor | all | none

user ใหม่ = role user + ไม่มีหน้า → ไม่เห็นเมนูใดจนกว่าจะให้สิทธิ์
รหัสผ่านไม่ถูกเก็บเป็นข้อความ (pbkdf2) และไม่ควรใส่เป็น argument (จะค้างใน shell history)
"""
import sys, getpass
import dash_auth
import dash_access


def _read_pw():
    if sys.stdin.isatty():
        a = getpass.getpass("Password: ")
        b = getpass.getpass("Password (อีกครั้ง): ")
        if a != b:
            sys.exit("รหัสผ่านไม่ตรงกัน")
        return a
    return sys.stdin.readline().rstrip("\n")


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    cmd, args = sys.argv[1], sys.argv[2:]
    with dash_auth._conn() as c, c.cursor() as cur:
        cur.execute(dash_auth.DDL)
        dash_access.ensure_columns(cur)
        T = dash_auth.TABLE
        if cmd == "list":
            cur.execute(f"SELECT username, full_name, active, role, allowed_pages, failed_count, locked_until, last_login "
                        f"FROM {T} ORDER BY username")
            for r in cur.fetchall():
                print(r)
        elif cmd == "role" and len(args) == 2 and args[1] in dash_access.ROLES:
            if args[1] == "user":
                cur.execute(f"SELECT COUNT(*) FROM {T} WHERE role='admin' AND active=1 AND username<>%s", (args[0],))
                if not cur.fetchone()[0]:
                    sys.exit("ต้องเหลือ admin ที่เปิดใช้งานอย่างน้อย 1 คน")
            n = cur.execute(f"UPDATE {T} SET role=%s WHERE username=%s", (args[1], args[0]))
            print(f"role={args[1]}" if n else "ไม่พบ username (หรือค่าเดิมอยู่แล้ว)", args[0])
        elif cmd == "pages" and not args:
            for k, lbl in dash_access.PAGES:
                print(f"  {k:15} {lbl}")
        elif cmd == "pages" and len(args) == 2:
            spec = args[1].strip().lower()
            keys = dash_access.PAGE_KEYS if spec == "all" else [] if spec == "none" else spec.split(",")
            bad = [k for k in keys if k and k not in dash_access.LABEL]
            if bad:
                sys.exit(f"ไม่รู้จักรหัสหน้า: {', '.join(bad)} — ดูรายการด้วย: python manage_users.py pages")
            val = dash_access.join_pages(keys)
            n = cur.execute(f"UPDATE {T} SET allowed_pages=%s WHERE username=%s", (val, args[0]))
            print(f"pages={val or '(ไม่มี)'}" if n else "ไม่พบ username (หรือค่าเดิมอยู่แล้ว)", args[0])
        elif cmd == "add" and args:
            pw = _read_pw()
            if len(pw) < 4:
                sys.exit("รหัสผ่านสั้นเกินไป (อย่างน้อย 4 ตัว)")
            cur.execute(f"INSERT INTO {T} (username, password_hash, full_name) VALUES (%s,%s,%s) "
                        f"ON DUPLICATE KEY UPDATE password_hash=VALUES(password_hash), full_name=VALUES(full_name), "
                        f"active=1, failed_count=0, locked_until=NULL",
                        (args[0], dash_auth.hash_pw(pw), args[1] if len(args) > 1 else None))
            print("saved", args[0])
        elif cmd == "passwd" and args:
            pw = _read_pw()
            n = cur.execute(f"UPDATE {T} SET password_hash=%s, failed_count=0, locked_until=NULL WHERE username=%s",
                            (dash_auth.hash_pw(pw), args[0]))
            print("updated" if n else "ไม่พบ username", args[0])
        elif cmd in ("disable", "enable") and args:
            n = cur.execute(f"UPDATE {T} SET active=%s WHERE username=%s", (1 if cmd == "enable" else 0, args[0]))
            print(cmd, "ok" if n else "ไม่พบ username", args[0])
        elif cmd == "unlock" and args:
            dash_auth._ensure_lock_count(cur)
            n = cur.execute(f"UPDATE {T} SET failed_count=0, locked_until=NULL, lock_count=0 WHERE username=%s", (args[0],))
            print("unlocked" if n else "ไม่พบ username", args[0])
        else:
            sys.exit(__doc__)


if __name__ == "__main__":
    main()
