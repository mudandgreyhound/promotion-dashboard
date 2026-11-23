"""
manage_users.py — จัดการบัญชี login ของหน้า Streamlit (app.py พอร์ต 83) ในตาราง promo_dash_users

  python manage_users.py list
  python manage_users.py add <username> "ชื่อ-นามสกุล"     # ถามรหัสผ่าน (ไม่ขึ้นจอ) หรือส่งทาง stdin
  python manage_users.py passwd <username>
  python manage_users.py disable <username>
  python manage_users.py enable <username>
  python manage_users.py unlock <username>               # ปลดล็อกที่ใส่รหัสผิดเกิน 5 ครั้ง

รหัสผ่านไม่ถูกเก็บเป็นข้อความ (pbkdf2) และไม่ควรใส่เป็น argument (จะค้างใน shell history)
"""
import sys, getpass
import dash_auth


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
        T = dash_auth.TABLE
        if cmd == "list":
            cur.execute(f"SELECT username, full_name, active, failed_count, locked_until, last_login FROM {T} ORDER BY username")
            for r in cur.fetchall():
                print(r)
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
            n = cur.execute(f"UPDATE {T} SET failed_count=0, locked_until=NULL WHERE username=%s", (args[0],))
            print("unlocked" if n else "ไม่พบ username", args[0])
        else:
            sys.exit(__doc__)


if __name__ == "__main__":
    main()
