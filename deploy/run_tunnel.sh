#!/bin/bash
# เปิด Cloudflare tunnel ของ Greyhound Dashboard (greyhound-dashboard.mobile1234.site → Streamlit พอร์ต 83)
# ถ้ายังไม่รัน — เรียกจาก cron ของ ubuntu1: @reboot และทุก 5 นาที (ล้มเมื่อไรก็ขึ้นใหม่เอง ไม่ต้องใช้ sudo)
# config: /home/ubuntu1/.cloudflared/greyhound-dashboard.yml (ไม่เก็บใน git — มี path ไฟล์ credentials)
CFG=/home/ubuntu1/.cloudflared/greyhound-dashboard.yml
LOG=/home/ubuntu1/.cloudflared/greyhound-dashboard.log
pgrep -u "$(id -u)" -f "^/usr/bin/cloudflared .*--config $CFG tunnel run" >/dev/null && exit 0   # ^ = ต้องเป็นตัว cloudflared เอง (กัน match shell อื่น)
nohup /usr/bin/cloudflared --no-autoupdate --config "$CFG" tunnel run >>"$LOG" 2>&1 &
