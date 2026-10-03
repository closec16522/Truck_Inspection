# ระบบตรวจรถบรรทุก หจก. ทั่วไทยขนส่งมงคล (TTMK Truck Inspection)

ระบบ Web Application สำหรับตรวจสอบสภาพรถบรรทุก 6 ล้อ ประจำวัน (รองรับทั้ง **รถน้ำมัน / ดีเซล (ICE)** และ **รถบรรทุกไฟฟ้า (EV)**)
พัฒนาสำหรับ: **หจก. ทั่วไทยขนส่งมงคล**
ผู้ดูแล / เจ้าของลิขสิทธิ์: **Mr.Taweesak.kom (062-3285963)**

---

## 🏗️ ภาพรวมโครงสร้างระบบ (System Architecture)

```
[ Frontend: GitHub Pages (Single Page Application) ]
  ├── 1. Modern UI: TailwindCSS (Mobile-First ออกแบบเน้นหน้างานจริง ใช้งานง่าย)
  ├── 2. Live Camera Engine: HTML5 MediaStream API + Canvas
  │      - บังคับเปิดกล้องหลังแบบสด (Live Stream เท่านั้น)
  │      - ห้ามเลือกรูปจากแกลเลอรี ป้องกันการทุจริตรูปเก่าย้อนหลัง 100%
  ├── 3. Dynamic Vehicle Mode:
  │      - รถน้ำมัน (ICE): เช็กระดับน้ำมันเครื่อง, น้ำมันเกียร์, หม้อน้ำ/สายพาน, รอยรั่วใต้ท้อง
  │      - รถไฟฟ้า (EV): เช็กระดับ SoC (%), พอร์ตชาร์จ CCS2/GBT, สายไฟแรงสูง (HV), น้ำหล่อเย็นแบตเตอรี่
  └── 4. SweetAlert2: แจ้งเตือน Loading, Success, Error สวยงามระดับมืออาชีพ
       │
       │ HTTPS POST (JSON Payload พร้อมภาพ Base64)
       ▼
[ Backend: Google Apps Script (Web App) ]
  ├── 1. Authentication & Security:
  │      - ตรวจสอบ Email กับ Sheet `Registered_Drivers`
  │      - เช็กสถานะ Status ต้องเป็น `ACTIVE`
  │      - ดึง Driver_ID, Driver_Name, Phone, Vehicle_Plate อัตโนมัติ
  ├── 2. Image Processing & Cloud Storage:
  │      - แปลง Base64 เป็นไฟล์ภาพ JPG พร้อมตั้งชื่อ Timestamp + DriverID
  │      - อัปโหลดเข้า Google Drive โฟลเดอร์ ID: `1ryLhwkO1lnv-2qgLbDDl0XyVaq8S2Szs`
  │      - สร้าง Public Link สำหรับเรียกดูภาพ
  └── 3. Database Logging:
         - บันทึกข้อมูลและ Link ภาพลง Sheet `Truck_Inspection_Logs`
         - Google Spreadsheet ID: `1JM-i8_nrGR7-VDEY82QZ5l5JMJTIBOIsuqOSQSrcD3Y`
```

---

## 📊 โครงสร้างตาราง Google Sheet (`Truck_Inspection_Logs`)

| คอลัมน์ | ชื่อฟิลด์ | รายละเอียด |
|:---:|:---|:---|
| **A** | `Timestamp` | วัน-เวลาที่ตรวจ (DD/MM/YYYY HH:mm:ss) |
| **B** | `Driver_ID` | รหัสพนักงาน (ตรวจสอบจาก Registered_Drivers) |
| **C** | `Driver_Name` | ชื่อพนักงานขับรถ |
| **D** | `Email` | อีเมลที่ใช้ส่งรายการ |
| **E** | `Vehicle_Plate` | ทะเบียนรถ |
| **F** | `Vehicle_Type` | ประเภทรถ (ICE (ดีเซล/น้ำมัน) หรือ EV (ไฟฟ้า)) |
| **G** | `Odometer` | เลขไมล์ปัจจุบัน |
| **H** | `Overall_Status` | ผลประเมินรวม (พร้อมใช้งาน / แจ้งซ่อม/ไม่พร้อม) |
| **I** | `Checklist_Summary` | สรุปผลการตรวจจุดต่างๆ (JSON Text) |
| **J** | `Defects_Note` | จุดชำรุด/หมายเหตุเพิ่มเติม |
| **K** | `Img_Front_Tire` | ลิงก์รูปถ่าย: ยางหน้า-ดอกยาง |
| **L** | `Img_Rear_Tire` | ลิงก์รูปถ่าย: ยางหลัง-กระทะล้อ |
| **M** | `Img_Powertrain` | ลิงก์รูปถ่าย: ห้องเครื่อง (ICE) หรือ แบตเตอรี่/สายส้ม HV (EV) |
| **N** | `Img_Brake_Fluid` | ลิงก์รูปถ่าย: ระดับน้ำมันเบรก/ลมเบรก หรือ ปั๊มลม EV |
| **O** | `Img_Lights_Body` | ลิงก์รูปถ่าย: ไฟสัญญาณรอบคันและสภาพรอบตัวถัง |
| **P** | `Img_Dashboard` | ลิงก์รูปถ่าย: หน้าปัดไมล์/ไฟเตือน (และระดับแบตเตอรี่ SoC) |

---

## 🚀 ขั้นตอนการติดตั้งและ Deploy ใช้งาน

### ขั้นตอนที่ 1: ติดตั้ง Google Apps Script (Backend)
1. เปิด Google Sheet ID: `1JM-i8_nrGR7-VDEY82QZ5l5JMJTIBOIsuqOSQSrcD3Y`
2. ไปที่เมนู **ส่วนขยาย (Extensions)** > **Apps Script**
3. คัดลอกโค้ดจากไฟล์ `Code.gs` ไปวางแทนที่โค้ดเดิมทั้งหมด
4. คลิกไอคอน **บันทึก (Save)**
5. คลิกปุ่ม **Deploy (การทำให้ใช้งานได้)** > **New deployment (การทำให้ใช้งานได้ใหม่)**
   - เลือกประเภท: **Web App**
   - Execute as (ดำเนินการในฐานะ): **Me (ฉัน - เจ้าของบัญชี)**
   - Who has access (ผู้มีสิทธิ์เข้าถึง): **Anyone (ทุกคน)**
6. คลิก **Deploy** แล้วอนุญาตสิทธิ์ (Authorize access) ให้เรียบร้อย
7. คัดลอก **Web App URL** ที่ได้ (เช่น `https://script.google.com/macros/s/AKfycb.../exec`)

### ขั้นตอนที่ 2: ตั้งค่า Frontend บน GitHub Pages
1. เปิดไฟล์ `index.html`
2. นำ Web App URL ที่ได้จากขั้นตอนที่ 1 ไปวางแทนที่ในตัวแปร:
   ```javascript
   const SCRIPT_URL = "https://script.google.com/macros/s/AKfycb.../exec";
   ```
3. บันทึกและ Push โค้ดขึ้น GitHub Repository: `https://github.com/closec16522/Truck_Inspection`
4. ไปที่ **Settings** ของ Repository บน GitHub > เมนู **Pages**
5. เลือก Source เป็น **Deploy from a branch** (Branch: `main` / Folder: `/root`) แล้วกด **Save**
6. เข้าใช้งานผ่าน URL ของ GitHub Pages ได้ทันทีจากสมาร์ตโฟนของพนักงานขับรถ

---

## 🔒 ลิขสิทธิ์และการติดต่อ
- **หจก. ทั่วไทยขนส่งมงคล**
- ลิขสิทธิ์โดย: **Mr.Taweesak.kom (062-3285963)**
