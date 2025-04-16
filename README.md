# HDMI Monitor

> โปรแกรมตรวจสอบสถานะการเชื่อมต่อ HDMI และ Driver  
> HDMI connection & driver status monitoring tool

---

## 🧰 วิธีการติดตั้ง | Installation

**ภาษาไทย:**

1. ติดตั้ง Python 3.12 หรือสูงกว่า
2. ติดตั้ง package ที่จำเป็นโดยใช้คำสั่ง:
   ```
   pip install -r requirements.txt
   ```

**ENGLISH:**

1. Install Python 3.12 or higher  
2. Install required packages using:
   ```
   pip install -r requirements.txt
   ```

---

## ▶️ วิธีการใช้งาน | How to Use

**ภาษาไทย:**

1. เปิดโฟลเดอร์ `dist/HDMI Monitor`
2. รันไฟล์ `HDMI Monitor.exe`
3. โปรแกรมจะแสดงไอคอนในถรายะบะบบ (System Tray)
4. คลิกขวาขวาไอคอนเพื่อเปิดเมนู:
   - แสดงหน้าต่างสถานะ
   - ออกจากโปรแกรม

**ENGLISH:**

1. Open the folder `dist/HDMI Monitor`
2. Run `HDMI Monitor.exe`
3. The app will appear as an icon in the system tray
4. Right-click the icon for menu options:
   - Show status window
   - Exit program

---

## 🧯 การแก้ไขปัญหา | Troubleshooting

**ภาษาไทย:**

1. หากโปรแกรมไม่ทำงาน ให้ลอง "Run as Administrator"
2. ตรวจการติดตั้ง package จาก `requirements.txt`
3. หากเจอปัญหาเกี่ยวกับ Driver ให้อัปเดต Driver การ์ดจอให้เป็นเวอร์ชันล่าสุด

**ENGLISH:**

1. If the app doesn’t work, try running it as **Administrator**
2. Check if all packages in `requirements.txt` are properly installed
3. If there are driver issues, update your **graphics card driver** to the latest version

---

## 🔄 หมายเหตุ | Notes

**ภาษาไทย:**

- โปรแกรมจะตรวจสอบสถานะ HDMI และ Driver อัตโนมัติทุก 5 วินาที
- สามารถตั้งค่าให้เปิดอัตโนมัติเมื่อเริ่ม Windows ได้ โดยการสร้าง Shortcut ไปที่โฟลเดอร์ `Startup`

**ENGLISH:**

- The app checks HDMI and driver status every 5 seconds automatically
- To auto-start on Windows boot, create a shortcut in the `Startup` folder

