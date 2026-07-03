# 🖥️ IT ServiceDesk Enterprise

ระบบ IT Helpdesk ครบวงจร — Single HTML File + Google Sheets Database

![Version](https://img.shields.io/badge/version-5.0-blue)
![Database](https://img.shields.io/badge/database-Google%20Sheets-green)
![Language](https://img.shields.io/badge/language-Thai-orange)

---

## 📦 ไฟล์ในโปรเจกต์

| ไฟล์ | คำอธิบาย |
|------|-----------|
| `IT-ServiceDesk-v5.html` | แอปหลัก — เปิดในเบราว์เซอร์ได้เลย |
| `installer.html` | ตัวช่วยติดตั้ง Google Sheets อัตโนมัติ |
| `Code.gs` | Google Apps Script Backend |

---

## 🚀 วิธีติดตั้ง

### วิธีง่ายที่สุด (แนะนำ)
1. ดาวน์โหลดทั้ง 3 ไฟล์ไว้ในโฟลเดอร์เดียวกัน
2. เปิด **`installer.html`** ในเบราว์เซอร์
3. ทำตามขั้นตอน 4 ขั้นใน Installer (~2 นาที)

### วิธี Manual
1. สร้าง Google Sheet ใหม่ → Extensions → Apps Script
2. วาง `Code.gs` → Run `initSheets()` → อนุมัติ Permission
3. Deploy → Web App (Execute as: Me, Access: Anyone) → Copy URL
4. เปิด `IT-ServiceDesk-v5.html` → คลิก badge "Local · คลิกเชื่อม" → วาง URL

---

## ✨ ฟีเจอร์หลัก

### 👤 User Portal (พนักงาน)
- แจ้งปัญหา IT พร้อม category icon grid
- ติดตามสถานะ Ticket ของตัวเอง
- คลังความรู้แก้ปัญหาเบื้องต้น
- ETA ตามระดับความเร่งด่วน

### 🛡️ Admin Portal (ฝ่าย IT)
- Dashboard + Charts แบบ Real-time
- จัดการ Ticket ครบ lifecycle (Draft→New→Open→Assigned→In Progress→Testing→Resolved→Closed)
- **Asset Management (CMDB)** — ครบทุกฟิลด์ + Warranty tracking
- **Document Repository** — อัปโหลด/จัดการเอกสาร (ใบเสร็จ/ใบรับประกัน/คู่มือ)
- SLA Management + Escalation Rules
- Approval Workflow
- Executive Dashboard (Radar Chart, CSAT, Priority Distribution)
- Audit Logs

### 🗄️ Database (Google Sheets)
- ซิงค์ Real-time ผ่าน Google Apps Script Web App
- 6 Sheet tabs: Tickets, Assets, Documents, Users, AuditLog, Settings
- Auto SLA calculation (Trigger ทุก 15 นาที)
- Export CSV ทุก table
- Offline fallback ด้วย localStorage
- Optimistic UI updates

---

## 🏗️ Architecture

```
IT-ServiceDesk-v5.html
        ↕ fetch() REST API
Google Apps Script Web App
        ↕ SpreadsheetApp
Google Sheets Database
  ├── Tickets    (28 columns)
  ├── Assets     (26 columns)
  ├── Documents  (15 columns)
  ├── Users      (13 columns)
  ├── AuditLog   (10 columns)
  └── Settings   (3 columns)
```

---

## 🔐 ความปลอดภัย

- Google OAuth 2.0 (ผ่าน Apps Script)
- SLA auto-update ทุก 15 นาที (Time-based Trigger)
- Audit Log ทุก action
- Session จัดการโดย Google

---

## 📋 Requirements

- Google Account
- Google Chrome / Edge
- ไม่ต้อง install อะไรเพิ่ม

---

## 📸 Screenshots

| Login Screen | Admin Dashboard | Asset Management |
|---|---|---|
| Role picker (User/Admin) | KPI + Charts | CMDB + Document Upload |

---

## 🔧 Tech Stack

- **Frontend**: Vanilla HTML/CSS/JS (Single File, ~230KB)
- **Icons**: Tabler Icons v3.24
- **Charts**: Chart.js 4.4.1
- **Database**: Google Sheets via Apps Script REST API
- **Fonts**: Inter + Sarabun (Google Fonts)

---

## 📄 License

MIT License — ใช้งานได้ฟรีสำหรับองค์กร

