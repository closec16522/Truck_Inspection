/**
 * ระบบตรวจเช็กรถบรรทุก 6 ล้อ - หจก. ทั่วไทยขนส่งมงคล
 * Google Apps Script Backend (Code.gs)
 * รองรับทั้ง Standard WebApp และ 3D Digital Twin Simulation
 * ผู้พัฒนา/ลิขสิทธิ์: Mr.Taweesak.kom (062-3285963)
 */

const SPREADSHEET_ID = "1JM-i8_nrGR7-VDEY82QZ5l5JMJTIBOIsuqOSQSrcD3Y";
const DRIVE_FOLDER_ID = "1ryLhwkO1lnv-2qgLbDDl0XyVaq8S2Szs";

function doGet(e) {
  // รองรับการตรวจสอบ Email พนักงานแบบ Real-time API
  if (e && e.parameter && e.parameter.action === "verify_driver") {
    const email = (e.parameter.email || "").trim().toLowerCase();
    return responseJSON(checkDriverRegistration(email));
  }

  return responseJSON({
    status: "online",
    message: "TTMK Truck Inspection API & 3D Twin Backend is active",
    timestamp: new Date().toISOString()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(25000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseJSON({ status: "error", message: "ไม่พบข้อมูลที่ส่งเข้ามา (Empty payload)" });
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. ตรวจสอบ Driver Registration
    const driverCheck = checkDriverRegistration(data.email, ss);
    if (driverCheck.status !== "success") {
      return responseJSON(driverCheck);
    }
    const driverInfo = driverCheck.data;

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
      
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0f172a");
      headerRange.setFontColor("#38bdf8");
      headerRange.setFontWeight("bold");
    }

    // 3. จัดการอัปโหลดไฟล์ภาพเข้า Google Drive
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const timeCode = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss");
    const photoUrls = {};
    
    // Mapping key ทั้งแบบ Standard และ 3D Simulation
    // 1. Front Tire
    const frontKey = data.images.frontTire ? 'frontTire' : (data.images.front_tires ? 'front_tires' : null);
    // 2. Rear Tire
    const rearKey = data.images.rearTire ? 'rearTire' : (data.images.rear_dually_tires ? 'rear_dually_tires' : null);
    // 3. Powertrain (ICE or EV)
    const powerKey = data.images.powertrain ? 'powertrain' : (data.images.powertrain_ice ? 'powertrain_ice' : (data.images.hv_battery_pack ? 'hv_battery_pack' : null));
    // 4. Brake / Fuel / Charging
    const brakeKey = data.images.brakeFluid ? 'brakeFluid' : (data.images.air_brake_tanks ? 'air_brake_tanks' : (data.images.charging_port ? 'charging_port' : null));
    // 5. Lights / Body / Cargo
    const lightsKey = data.images.lightsBody ? 'lightsBody' : (data.images.cargo_rear_lights ? 'cargo_rear_lights' : null);
    // 6. Dashboard / Cockpit
    const dashKey = data.images.dashboard ? 'dashboard' : (data.images.cockpit_dash ? 'cockpit_dash' : (data.images.cockpit_ev_dash ? 'cockpit_ev_dash' : null));

    const mapping = {
      Img_Front_Tire: frontKey,
      Img_Rear_Tire: rearKey,
      Img_Powertrain: powerKey,
      Img_Brake_Fluid: brakeKey,
      Img_Lights_Body: lightsKey,
      Img_Dashboard: dashKey
    };

    for (let col in mapping) {
      const srcKey = mapping[col];
      if (srcKey && data.images[srcKey] && data.images[srcKey].includes(",")) {
        try {
          const base64Data = data.images[srcKey].split(",")[1];
          const decodedBytes = Utilities.base64Decode(base64Data);
          const fileName = `${timeCode}_${driverInfo.driverId}_${col}.jpg`;
          const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoUrls[col] = file.getUrl();
        } catch (imgErr) {
          photoUrls[col] = "Upload Error: " + imgErr.toString();
        }
      } else {
        photoUrls[col] = "-";
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
      photoUrls['Img_Front_Tire'],
      photoUrls['Img_Rear_Tire'],
      photoUrls['Img_Powertrain'],
      photoUrls['Img_Brake_Fluid'],
      photoUrls['Img_Lights_Body'],
      photoUrls['Img_Dashboard']
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

function checkDriverRegistration(email, spreadsheetInstance) {
  try {
    const ss = spreadsheetInstance || SpreadsheetApp.openById(SPREADSHEET_ID);
    const regSheet = ss.getSheetByName("Registered_Drivers");
    if (!regSheet) {
      return { status: "error", message: "ไม่พบ Sheet 'Registered_Drivers' ในระบบฐานข้อมูล" };
    }

    const regData = regSheet.getDataRange().getValues();
    const targetEmail = (email || "").toString().trim().toLowerCase();

    for (let i = 1; i < regData.length; i++) {
      const emailInSheet = (regData[i][2] || "").toString().trim().toLowerCase();
      if (emailInSheet === targetEmail) {
        const status = (regData[i][6] || "ACTIVE").toString().trim().toUpperCase();
        if (status !== "ACTIVE") {
          return { status: "error", message: "บัญชีของท่านถูกระงับ (Status ไม่ใช่ ACTIVE)" };
        }
        return {
          status: "success",
          data: {
            driverId: regData[i][0] || "EMP-N/A",
            name: regData[i][1] || "ไม่ระบุชื่อ",
            email: regData[i][2],
            phone: regData[i][3] || "-",
            plate: regData[i][4] || ""
          }
        };
      }
    }

    return { 
      status: "error", 
      message: `อีเมล [${email}] ยังไม่ได้ลงทะเบียนในระบบ Registered_Drivers กรุณาติดต่อแอดมิน` 
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
