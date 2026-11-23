"""
app.py — ตัวเรียก (loader) ของ Streamlit พอร์ต 83 (promotion_dashboard.service รันไฟล์นี้)
รันโค้ดจริงใน app_main.py ใหม่ทุกครั้งที่หน้า rerun

เหตุผล (2026-11-23): Streamlit cache bytecode ของ main script ไว้ จนกว่า file watcher จะเห็นว่าไฟล์เปลี่ยน
หรือทุก session ปิดหมด — server นี้ไม่มี watchdog และ polling watcher ไม่ทำงาน + มีจอ Sale Monitor เปิดค้าง
→ แก้ app.py แล้วหน้าเว็บยังเป็นโค้ดเก่าจนกว่าจะ restart service (root)
ไฟล์นี้แทบไม่ต้องแก้อีก — แก้ app_main.py แล้วมีผลรอบ rerun ถัดไปทันที ไม่ต้อง restart
"""
import os
import runpy

runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_main.py"), run_name="__main__")
