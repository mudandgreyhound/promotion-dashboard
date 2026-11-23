"""
gh_env.py — โหลดค่าลับ (รหัส DB / API key) จากไฟล์ .env ข้างไฟล์นี้เข้า os.environ

- รูปแบบไฟล์: KEY=VALUE ทีละบรรทัด, บรรทัดว่าง/ขึ้นต้น # ข้าม, ครอบค่าด้วย "..." หรือ '...' ได้
- ค่าที่ตั้งใน environment อยู่แล้ว (เช่น systemd Environment=) ชนะค่าในไฟล์
- .env ไม่เข้า git (ดู .gitignore) — ตัวอย่างคีย์ทั้งหมดอยู่ใน .env.example

ใช้:  import gh_env;  pw = gh_env.need("GH_MYSQL_ROOT_PASSWORD")
สคริปต์ใน etl/ ต้องเพิ่มโฟลเดอร์แม่เข้า sys.path ก่อน import (ดูหัวไฟล์ etl/*.py)
"""
import os

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")


def load(path=ENV_PATH):
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
                    v = v[1:-1]
                os.environ.setdefault(k, v)
    except FileNotFoundError:
        pass


def need(key):
    """คืนค่า env ที่ต้องมี — ไม่มี = error ชัด ๆ (แทนการ fallback ไปรหัสที่ฝังในโค้ด)"""
    v = os.environ.get(key)
    if not v:
        raise RuntimeError(f"ไม่พบค่า {key} — ตั้งใน {ENV_PATH} (ดูตัวอย่างใน .env.example)")
    return v


def mysql_cfg(role="root", **extra):
    """dict สำหรับ pymysql.connect — role: 'root' (app/ETL) หรือ 'admin' (BI API / sale monitor ETL)"""
    r = role.upper()
    cfg = dict(host=need("GH_MYSQL_HOST"), port=int(os.environ.get("GH_MYSQL_PORT", "3306")),
               user=need(f"GH_MYSQL_{r}_USER"), password=need(f"GH_MYSQL_{r}_PASSWORD"),
               database=need("GH_MYSQL_DB"), charset="utf8mb4")
    cfg.update(extra)
    return cfg


load()
