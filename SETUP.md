# Merch Order Form — Setup Guide
OPH 2026 × MWIT 36 ปี

## ขั้นตอนการ deploy (ทำครั้งเดียว ~15 นาที)

### 1. สร้าง Google Sheet

1. ไปที่ [sheets.new](https://sheets.new) → สร้าง Sheet ใหม่
2. ตั้งชื่อ: `Merch Orders OPH 2026`
3. Copy **Sheet ID** จาก URL:
   `https://docs.google.com/spreadsheets/d/**[COPY_THIS]**/edit`

### 2. สร้าง Google Drive Folder สำหรับสลิป

1. เปิด Google Drive → New Folder → ตั้งชื่อ `Merch Slips`
2. เปิด folder → Copy **Folder ID** จาก URL:
   `https://drive.google.com/drive/folders/**[COPY_THIS]**`

### 3. Deploy Google Apps Script

1. เปิด Google Sheet ที่สร้างไว้
2. Extensions → Apps Script
3. ลบ code เดิม → วาง code จากไฟล์ `gas/Code.gs`
4. แก้ค่าใน CONFIG:
   ```js
   const CONFIG = {
     SHEET_ID:        'วาง Sheet ID ที่ copy ไว้',
     DRIVE_FOLDER_ID: 'วาง Folder ID ที่ copy ไว้'
   };
   ```
5. กด **Deploy → New deployment**
   - Type: Web app
   - Execute as: **Me**
   - Who has access: **Anyone**
6. กด Deploy → Copy **Deployment URL**

> ⚠ ถ้า Google ขอ Permission ให้ Allow ทั้งหมด

### 4. ตั้งค่า Summary Sheet (optional แต่แนะนำ)

1. ใน GAS editor → เลือก function `setupSummarySheet` → กด Run
2. จะมี Summary tab ปรากฏในชีต — เปิดดูภาพรวมได้เลย

### 5. อัปเดต index.html และ admin.html

แก้ค่าเหล่านี้ใน **index.html** (บรรทัดบนสุดของ `<script>`):
```js
const GAS_URL        = 'วาง Deployment URL';
const PROMPTPAY_NUM  = 'เบอร์โทร หรือ เลขบัตรประชาชน';
const PROMPTPAY_NAME = 'ชื่อ-นามสกุล ผู้รับเงิน';
```

แก้ค่าเดียวกันใน **admin.html**:
```js
const GAS_URL = 'วาง Deployment URL';
```

### 6. Host บน GitHub Pages

```bash
cd D:\mik\merch-order-form
git init
git add .
git commit -m "feat: merch order form OPH 2026"
git remote add origin https://github.com/YOUR_USERNAME/merch-oph26.git
git push -u origin main
```

จากนั้น:
- GitHub repo → Settings → Pages → Source: `main` branch
- URL จะได้: `https://YOUR_USERNAME.github.io/merch-oph26/`

**ส่ง link นี้ใน Google Classroom ✓**

---

## โครงสร้างไฟล์

```
merch-order-form/
├── index.html      ← ฟอร์ม pre-order (ส่ง link นี้)
├── admin.html      ← Dashboard สำหรับกรรมการ
└── gas/
    └── Code.gs     ← วางใน Google Apps Script
```

## การใช้งาน Admin Dashboard

เปิด `admin.html` → กด 🔄 รีเฟรช → เห็นข้อมูลรวมทั้งหมด

- **สรุปสินค้า**: จำนวนแต่ละ variant แยกรับเอง / ส่งไปรษณีย์
- **Export CSV**: กด ⬇ Export CSV → เปิดใน Excel ได้เลย
- **Filter**: ค้นหาชื่อ / กรองตามวิธีรับ / สถานะ

> แชร์ `admin.html` link ให้เฉพาะกรรมการที่รับผิดชอบจัดส่ง
