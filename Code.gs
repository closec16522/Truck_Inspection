/**
 * ระบบตรวจเช็กรถบรรทุก 6 ล้อ - หจก. ทั่วไทยขนส่งมงคล
 * Google Apps Script Backend (Code.gs)
 * ผู้พัฒนา/ลิขสิทธิ์: Mr.Taweesak.kom (062-3285963)
 */

const SPREADSHEET_ID = "1JM-i8_nrGR7-VDEY82QZ5l5JMJTIBOIsuqOSQSrcD3Y";
const DRIVE_FOLDER_ID = "1ryLhwkO1lnv-2qgLbDDl0XyVaq8S2Szs";

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    message: "TTMK Truck Inspection API is active"
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // รอล็อคสูงสุด 25 วินาที เพื่อป้องกัน Race Condition ในการบันทึก Row ในสเปรดชีต
  lock.tryLock(25000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "error", message: "ไม่พบข้อมูลที่ส่งเข้ามา (Empty payload)" });
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. ตรวจสอบ Email กับ Sheet 'Registered_Drivers'
    const regSheet = ss.getSheetByName("Registered_Drivers");
    if (!regSheet) {
      return responseJSON({ status: "error", message: "ไม่พบ Sheet 'Registered_Drivers' ในระบบฐานข้อมูล" });
    }

    const regData = regSheet.getDataRange().getValues();
    let driverInfo = null;

    // โครงสร้างชีต Registered_Drivers:
    // Col A: Driver_ID (0)
    // Col B: Driver_Name (1)
    // Col C: Email (2)
    // Col D: Phone (3)
    // Col E: Vehicle_Plate (4)
    // Col G: Status (6)
    const targetEmail = (data.email || "").toString().trim().toLowerCase();

    for (let i = 1; i < regData.length; i++) {
      const emailInSheet = (regData[i][2] || "").toString().trim().toLowerCase();
      if (emailInSheet === targetEmail) {
        const status = (regData[i][6] || "ACTIVE").toString().trim().toUpperCase();
        if (status !== "ACTIVE") {
          return responseJSON({ status: "error", message: "บัญชีของท่านถูกระงับ (Status ไม่ใช่ ACTIVE)" });
        }
        driverInfo = {
          driverId: regData[i][0] || "EMP-N/A",
          name: regData[i][1] || "ไม่ระบุชื่อ",
          phone: regData[i][3] || "-",
          plate: regData[i][4] || data.vehiclePlate
        };
        break;
      }
    }

    if (!driverInfo) {
      return responseJSON({ 
        status: "error", 
        message: `อีเมล [${data.email}] ยังไม่ได้ลงทะเบียนในตาราง Registered_Drivers กรุณาติดต่อผู้ดูแลระบบ` 
      });
    }

    // 2. ตรวจสอบหรือสร้าง Sheet ใหม่ 'Truck_Inspection_Logs'
    let logSheet = ss.getSheetByName("Truck_Inspection_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Truck_Inspection_Logs");
      const headers = [
        "Timestamp", "Driver_ID", "Driver_Name", "Email", "Vehicle_Plate", 
        "Vehicle_Type", "Odometer", "Overall_Status", "Checklist_Summary", "Defects_Note",
        "Img_Front_Tire", "Img_Rear_Tire", "Img_Powertrain", "Img_Brake_Fluid", "Img_Lights_Body", "Img_Dashboard"
      ];
      logSheet.appendRow(headers);
      logSheet.setFrozenRows(1);
      
      // จัดรูปแบบหัวตาราง
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#1e293b");
      headerRange.setFontColor("#f8fafc");
      headerRange.setFontWeight("bold");
    }

    // 3. จัดการอัปโหลดไฟล์ภาพเข้า Google Drive
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const timeCode = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss");
    const photoUrls = {};
    const photoKeys = ['frontTire', 'rearTire', 'powertrain', 'brakeFluid', 'lightsBody', 'dashboard'];

    for (let key of photoKeys) {
      if (data.images && data.images[key] && data.images[key].includes(",")) {
        try {
          const base64Data = data.images[key].split(",")[1];
          const decodedBytes = Utilities.base64Decode(base64Data);
          const fileName = `${timeCode}_${driverInfo.driverId}_${key}.jpg`;
          const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoUrls[key] = file.getUrl();
        } catch (imgErr) {
          photoUrls[key] = "Upload Error: " + imgErr.toString();
        }
      } else {
        photoUrls[key] = "-";
      }
    }

    // 4. บันทึกผลลง Sheet 'Truck_Inspection_Logs'
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
    logSheet.appendRow([
      timestamp,
      driverInfo.driverId,
      driverInfo.name,
      data.email,
      data.vehiclePlate || driverInfo.plate,
      data.vehicleType,
      data.odometer,
      data.overallStatus,
      JSON.stringify(data.checklist),
      data.defectsNote || "ไม่มี",
      photoUrls['frontTire'],
      photoUrls['rearTire'],
      photoUrls['powertrain'],
      photoUrls['brakeFluid'],
      photoUrls['lightsBody'],
      photoUrls['dashboard']
    ]);

    return responseJSON({
      status: "success",
      message: `บันทึกข้อมูลการตรวจรถสำเร็จเรียบร้อย (${driverInfo.name})`,
      driverName: driverInfo.name,
      timestamp: timestamp
    });

  } catch (err) {
    return responseJSON({ status: "error", message: "Server Error: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
