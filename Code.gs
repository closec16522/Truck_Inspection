/**
 * ระบบตรวจเช็กรถบรรทุก 6 ล้อ - หจก. ทั่วไทยขนส่งมงคล
 * Google Apps Script Backend (Code.gs)
 * อ้างอิงมาตรฐานแบบฟอร์ม: ใบรายการตรวจสภาพรถบรรทุกประจำวัน (Fleet TCC ชลบุรี)
 * พร้อมระบบ PDF Generator ส่งออกไฟล์ใบตรวจเข้า Google Drive อัตโนมัติ (รายวัน & รายเดือน)
 * ลิขสิทธิ์โดย: Mr.Taweesak.kom (062-3285963)
 */

const SPREADSHEET_ID = "1JM-i8_nrGR7-VDEY82QZ5l5JMJTIBOIsuqOSQSrcD3Y";
const DRIVE_FOLDER_ID = "1ryLhwkO1lnv-2qgLbDDl0XyVaq8S2Szs";

function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.action === "verify_driver") {
      const email = (e.parameter.email || "").trim().toLowerCase();
      return responseJSON(checkDriverRegistration(email));
    }

    // ฟังก์ชันสร้างรายงานสรุปรายเดือนแบบตารางรวม Actual ประจำวัน
    if (e.parameter.action === "generate_monthly_pdf") {
      const plate = (e.parameter.plate || "").trim();
      const monthYear = (e.parameter.month || Utilities.formatDate(new Date(), "Asia/Bangkok", "MM/yyyy")).trim();
      return generateMonthlyTruckInspectionPdf(plate, monthYear);
    }
  }

  return responseJSON({
    status: "online",
    message: "TTMK Truck Inspection API & PDF Generator is active",
    timestamp: new Date().toISOString()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

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

    // หากเป็นการ Verify Driver ผ่าน POST
    if (data.action === "verify_driver") {
      return responseJSON({
        status: "success",
        driverName: driverInfo.name,
        data: driverInfo
      });
    }

    // หากเป็นการสั่งสร้างรายงานรายเดือนผ่าน POST
    if (data.action === "generate_monthly_pdf") {
      const targetPlate = String(data.vehiclePlate || driverInfo.plate || "");
      const targetMonth = String(data.month || Utilities.formatDate(new Date(), "Asia/Bangkok", "MM/yyyy"));
      return generateMonthlyTruckInspectionPdf(targetPlate, targetMonth);
    }

    // 2. ตรวจสอบหรือสร้าง Sheet 'Truck_Inspection_Logs'
    let logSheet = ss.getSheetByName("Truck_Inspection_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Truck_Inspection_Logs");
      const headers = [
        "Timestamp", "Driver_ID", "Driver_Name", "Email", "Vehicle_Plate", 
        "Vehicle_Type", "Odometer", "Overall_Status", "Checklist_Summary", "Defects_Note",
        "Img_Dashboard", "Img_Front_Body", "Img_Defects_1", "Img_Defects_2", "PDF_Report_URL"
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
    const photoBlobs = {};

    const imageKeys = ['dashboard', 'frontVehicle', 'defect_1', 'defect_2', 'defect_3'];
    for (let key of imageKeys) {
      if (data.images && data.images[key] && data.images[key].includes(",")) {
        try {
          const base64Data = data.images[key].split(",")[1];
          const decodedBytes = Utilities.base64Decode(base64Data);
          const fileName = `${timeCode}_${driverInfo.driverId}_${key}.jpg`;
          const blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoUrls[key] = file.getUrl();
          photoBlobs[key] = blob;
        } catch (imgErr) {
          photoUrls[key] = "Upload Error: " + imgErr.toString();
        }
      } else {
        photoUrls[key] = "-";
      }
    }

    // 4. สร้างเอกสาร PDF รายงานแบบฟอร์มประจำวัน อัตโนมัติ
    let pdfUrl = "-";
    try {
      const pdfBlob = generateFCAM032PdfReport(data, driverInfo, photoBlobs);
      const safePlate = String(data.vehiclePlate || driverInfo.plate || "70XXXX").replace(/[^a-zA-Z0-9ก-๙]/g, '');
      const pdfFileName = `TTMK_Daily_${timeCode}_${driverInfo.driverId}_${safePlate}.pdf`;
      const pdfFile = folder.createFile(pdfBlob.setName(pdfFileName));
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = pdfFile.getUrl();
    } catch (pdfErr) {
      pdfUrl = "PDF Gen Error: " + pdfErr.toString();
    }

    // 5. บันทึกผลลง Sheet 'Truck_Inspection_Logs'
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
      JSON.stringify(data.checklist || []),
      data.defectsNote || "ไม่มี",
      photoUrls['dashboard'],
      photoUrls['frontVehicle'],
      photoUrls['defect_1'],
      photoUrls['defect_2'],
      pdfUrl
    ]);

    return responseJSON({
      status: "success",
      message: `บันทึกข้อมูลและสร้างใบตรวจ PDF สำเร็จ (${driverInfo.name})`,
      driverName: driverInfo.name,
      pdfUrl: pdfUrl,
      timestamp: timestamp
    });

  } catch (err) {
    return responseJSON({ status: "error", message: "Server Error: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1. สร้างไฟล์ PDF ใบรายงานตรวจสภาพรถประจำวัน (Daily Inspection Report)
 */
function generateFCAM032PdfReport(data, driverInfo, photoBlobs) {
  const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
  const plate = String(data.vehiclePlate || driverInfo.plate || "70-XXXX");
  const driverName = String(driverInfo.name || "ไม่ระบุ");
  const driverId = String(driverInfo.driverId || "EMP-N/A");
  const odo = data.odometer || "-";
  const vType = data.vehicleType || "ICE (ดีเซล/น้ำมัน)";
  const status = data.overallStatus || "พร้อมใช้งาน";
  const note = data.defectsNote || "ไม่มี";

  // แถวตารางเช็กลิสต์
  let checklistRowsHtml = "";
  if (Array.isArray(data.checklist)) {
    data.checklist.forEach((item, index) => {
      const isDefect = item.status === "DEFECT";
      const icon = isDefect ? "<b style='color:#dc2626;'>[X] ชำรุด/บกพร่อง</b>" : "<span style='color:#16a34a;'>[/] ปกติดี</span>";
      const rowBg = isDefect ? "#fee2e2" : (index % 2 === 0 ? "#f8fafc" : "#ffffff");
      checklistRowsHtml += `
        <tr style="background-color: ${rowBg};">
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-size: 10px;">${item.no || (index + 1)}</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; font-size: 10px;"><b>[${item.zone || 'ทั่วไป'}]</b> ${item.text || item.question || ''}</td>
          <td style="border: 1px solid #cbd5e1; padding: 4px; text-align: center; font-size: 10px;">${icon}</td>
        </tr>
      `;
    });
  }

  // รูปภาพประกอบ (แสดงเฉพาะรูปที่มีจริง)
  let imagesHtml = "";
  if (data.images && (data.images.dashboard || data.images.frontVehicle || data.images.defect_1)) {
    imagesHtml += "<div style='display: flex; gap: 10px; margin-top: 10px;'>";
    if (data.images.dashboard) {
      imagesHtml += `
        <div style="flex: 1; border: 1px solid #94a3b8; padding: 5px; text-align: center;">
          <div style="font-size: 9px; font-weight: bold; margin-bottom: 3px;">1. รูปหน้าปัดไมล์/ไฟเตือน (Dashboard)</div>
          <img src="${data.images.dashboard}" style="max-width: 100%; height: 130px; object-fit: contain;" />
        </div>`;
    }
    if (data.images.frontVehicle) {
      imagesHtml += `
        <div style="flex: 1; border: 1px solid #94a3b8; padding: 5px; text-align: center;">
          <div style="font-size: 9px; font-weight: bold; margin-bottom: 3px;">2. รูปหน้ารถ/รอบคัน (Front Vehicle)</div>
          <img src="${data.images.frontVehicle}" style="max-width: 100%; height: 130px; object-fit: contain;" />
        </div>`;
    }
    if (data.images.defect_1) {
      imagesHtml += `
        <div style="flex: 1; border: 1px solid #dc2626; padding: 5px; text-align: center; background-color: #fef2f2;">
          <div style="font-size: 9px; font-weight: bold; color: #dc2626; margin-bottom: 3px;">3. จุดชำรุดแจ้งซ่อม (Defect Photo)</div>
          <img src="${data.images.defect_1}" style="max-width: 100%; height: 130px; object-fit: contain;" />
        </div>`;
    }
    imagesHtml += "</div>";
  }

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body { font-family: 'Garuda', 'Sarabun', sans-serif; color: #0f172a; padding: 15px; margin: 0; }
        .header-box { border-bottom: 2px solid #0284c7; padding-bottom: 6px; margin-bottom: 8px; }
        .title { font-size: 15px; font-weight: bold; color: #0f172a; margin: 0; }
        .sub-title { font-size: 11px; color: #475569; margin-top: 2px; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 10px; }
        .info-table td { padding: 3px 6px; border: 1px solid #cbd5e1; }
        .info-header { background-color: #f1f5f9; font-weight: bold; color: #334155; }
        .check-table { width: 100%; border-collapse: collapse; margin-top: 6px; }
        .check-table th { background-color: #0f172a; color: #ffffff; padding: 5px; font-size: 10px; border: 1px solid #0f172a; }
        .footer-sign { width: 100%; margin-top: 15px; border-collapse: collapse; font-size: 10px; }
        .footer-sign td { width: 33%; text-align: center; padding: 8px; border: 1px solid #cbd5e1; }
      </style>
    </head>
    <body>
      <div class="header-box">
        <h1 class="title">หจก. ทั่วไทยขนส่งมงคล (TTMK LOGISTICS)</h1>
        <div class="sub-title">ใบรายการตรวจสภาพรถบรรทุกประจำวัน</div>
      </div>

      <table class="info-table">
        <tr>
          <td class="info-header" style="width: 18%;">พนักงานขับรถ:</td>
          <td style="width: 32%;"><b>${driverName}</b> (${driverId})</td>
          <td class="info-header" style="width: 18%;">ทะเบียนรถ 6 ล้อ:</td>
          <td style="width: 32%; color: #b45309; font-weight: bold;">${plate}</td>
        </tr>
        <tr>
          <td class="info-header">วัน-เวลาที่ตรวจ:</td>
          <td>${timestamp}</td>
          <td class="info-header">เลขไมล์ปัจจุบัน:</td>
          <td><b>${odo}</b> กม.</td>
        </tr>
        <tr>
          <td class="info-header">ระบบรถยนต์:</td>
          <td>${vType}</td>
          <td class="info-header">ผลการประเมินรวม:</td>
          <td><b style="color: ${status === 'พร้อมใช้งาน' ? '#16a34a' : '#dc2626'};">${status}</b></td>
        </tr>
      </table>

      <table class="check-table">
        <thead>
          <tr>
            <th style="width: 10%;">ข้อที่</th>
            <th style="width: 65%;">รายการตรวจสภาพรถบรรทุก</th>
            <th style="width: 25%;">ผลการตรวจ</th>
          </tr>
        </thead>
        <tbody>
          ${checklistRowsHtml}
        </tbody>
      </table>

      <div style="margin-top: 8px; font-size: 10px; border: 1px solid #cbd5e1; padding: 6px; background-color: #f8fafc;">
        <b>หมายเหตุ / จุดชำรุดที่พบเพิ่มเติม:</b> ${note}
      </div>

      ${imagesHtml}

      <table class="footer-sign">
        <tr>
          <td>
            ลงชื่อ ...................................................<br>
            ( <b>${driverName}</b> )<br>
            พนักงานขับรถผู้ตรวจ
          </td>
          <td>
            ลงชื่อ ...................................................<br>
            ( ................................................... )<br>
            หัวหน้างาน / ผู้ตรวจสอบ (Spot Check)
          </td>
          <td>
            ลงชื่อ ...................................................<br>
            ( <b>นางสาว.นิชานันท์ เอื้อจิรพรชัย</b> )<br>
            ผู้รับเหมาขนส่ง / TTMK
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;

  return Utilities.newBlob(htmlContent, 'text/html', 'report.html').getAs('application/pdf');
}

/**
 * 2. สร้างไฟล์ PDF รายงานสรุปรายเดือนแบบตารางเมทริกซ์ 31 วัน (Monthly Matrix PDF Report)
 * ดึงข้อมูล Actual ประจำวันของรถคันเดียวกันจาก Sheet 'Truck_Inspection_Logs' มาสรุปลงตารางวันที่ 1 - 31
 */
function generateMonthlyTruckInspectionPdf(targetPlate, targetMonthYear) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName("Truck_Inspection_Logs");
    if (!logSheet) {
      return responseJSON({ status: "error", message: "ไม่พบข้อมูลในระบบ Truck_Inspection_Logs" });
    }

    const logData = logSheet.getDataRange().getValues();
    const plateToSearch = targetPlate.replace(/[^a-zA-Z0-9ก-๙]/g, '').toLowerCase();

    // กรองประวัติการตรวจเฉพาะทะเบียนรถคันนี้ ในเดือน-ปี ที่เลือก
    // Col A: Timestamp (dd/MM/yyyy HH:mm:ss), Col E: Vehicle_Plate, Col I: Checklist_Summary (JSON)
    const dailyInspections = {}; // { '1': { status: 'a', items: {...} }, '2': ... }
    let matchedDriverName = "-";

    for (let i = 1; i < logData.length; i++) {
      const rowTimestamp = String(logData[i][0] || "");
      const rowPlate = String(logData[i][4] || "").replace(/[^a-zA-Z0-9ก-๙]/g, '').toLowerCase();
      
      if (rowPlate.includes(plateToSearch) || plateToSearch.includes(rowPlate) || !targetPlate) {
        // เช็คว่าอยู่ในเดือนที่ระบุหรือไม่ เช่น "10/2026"
        if (rowTimestamp.includes(targetMonthYear)) {
          matchedDriverName = logData[i][2] || matchedDriverName;
          const day = parseInt(rowTimestamp.substring(0, 2), 10);
          
          let parsedChecklist = [];
          try {
            parsedChecklist = JSON.parse(logData[i][8] || "[]");
          } catch(e) {}

          dailyInspections[day] = {
            status: logData[i][7] === "พร้อมใช้งาน" ? "a" : "r",
            checklist: parsedChecklist,
            odometer: logData[i][6]
          };
        }
      }
    }

    // รายการตรวจมาตรฐานตามตาราง (1 ถึง 21 ข้อหลัก)
    const standardItems = [
      "1. ระดับน้ำมันเบรคและคลัตซ์ พร้อมใช้งาน",
      "2. ที่ปัดน้ำฝนพร้อมใช้งาน",
      "3. เบรคมือใช้งานได้ปกติ",
      "4. ระบบเบรคใช้งานได้ปกติ",
      "5. Safety Belt เข็มขัดนิรภัย ใช้งานได้ปกติ",
      "6. กระจกหน้า / มองข้าง / มองหลัง / มองมุม",
      "7. ไฟส่องสว่างหน้ารถ และไฟสัญญาณพร้อมใช้งาน",
      "8. สภาพยางรถ (ดอกยางหน้า 3 มม., หลัง 2 มม.)",
      "9. สัญญาณแตร",
      "10. หมอนรองล้อ 2 อัน & กรวยยาง 2 อัน",
      "11. ถังดับเพลิง พร้อมใช้งาน เกจ์เขียว",
      "12. กล้องหน้ารถ & GPS ทำงานปกติ",
      "13. ตรวจเช็คไฟหน้าปัด ไม่มีไฟโชว์ผิดปกติ",
      "14. ระดับน้ำมันเครื่อง อยู่ในเกณฑ์ปกติ",
      "15. ระดับน้ำในหม้อน้ำระบายความร้อน & ท่อยาง",
      "16. ท่อยางและสายพานพัดลม",
      "17. ถังอัดอากาศพร้อมวาล์วระบายน้ำทิ้ง",
      "18. ผ้าใบ/สภาพตู้ อยู่ในสภาพดี ไม่บุบ ไม่รั่ว",
      "19. บานพับ, กลอน/ตัวล็อค แข็งแรง ไม่ชำรุด",
      "20. ผ้าแดงแขวนท้ายรถ & ป้ายสามเหลี่ยมสะท้อนแสง",
      "21. สภาพแบตเตอรี่ / สายไฟแรงสูง (EV)"
    ];

    // สร้าง Header วันที่ 1 ถึง 31
    let dayHeadersHtml = "";
    for (let d = 1; d <= 31; d++) {
      dayHeadersHtml += `<th style="width: 2.2%; font-size: 8px; padding: 2px 0; border: 1px solid #94a3b8;">${d}</th>`;
    }

    // สร้างเนื้อหาตารางแต่ละข้อ
    let rowsHtml = "";
    standardItems.forEach((title, idx) => {
      let cellsHtml = "";
      for (let d = 1; d <= 31; d++) {
        let cellVal = "";
        if (dailyInspections[d]) {
          // ถ้าวันนั้นตรวจแล้วผ่าน ใส่ 'a' (หรือ '/'), ถ้าชำรุดใส่ 'r' (หรือ 'X')
          cellVal = dailyInspections[d].status === "a" ? "<span style='color:#15803d; font-weight:bold;'>/</span>" : "<span style='color:#b91c1c; font-weight:bold;'>X</span>";
        }
        cellsHtml += `<td style="text-align: center; font-size: 8px; padding: 1px 0; border: 1px solid #cbd5e1;">${cellVal}</td>`;
      }

      rowsHtml += `
        <tr style="background-color: ${idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
          <td style="font-size: 8.5px; padding: 2.5px 4px; border: 1px solid #cbd5e1; white-space: nowrap;">${title}</td>
          ${cellsHtml}
        </tr>
      `;
    });

    const monthlyHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body { font-family: 'Garuda', 'Sarabun', sans-serif; color: #0f172a; padding: 10px; margin: 0; }
          .header-title { text-align: center; font-size: 14px; font-weight: bold; margin-bottom: 2px; }
          .header-meta { font-size: 9.5px; margin-bottom: 6px; display: flex; justify-content: space-between; border-bottom: 1.5px solid #0f172a; padding-bottom: 4px; }
          .matrix-table { width: 100%; border-collapse: collapse; margin-top: 4px; }
          .matrix-table th { background-color: #0f172a; color: #ffffff; text-align: center; }
          .sign-table { width: 100%; margin-top: 10px; border-collapse: collapse; font-size: 8.5px; }
          .sign-table td { width: 33.33%; text-align: center; padding: 6px; border: 1px solid #cbd5e1; }
        </style>
      </head>
      <body>
        <div class="header-title">ใบรายการตรวจสภาพรถบรรทุกประจำวัน (สรุปประจำเดือน ${targetMonthYear})</div>
        <table style="width: 100%; font-size: 9.5px; margin-bottom: 4px;">
          <tr>
            <td><b>หจก. ทั่วไทยขนส่งมงคล</b></td>
            <td style="text-align: right;">
              พขร.: <b>${matchedDriverName}</b> | ทะเบียนรถ: <b style="color: #b45309;">${targetPlate || '-'}</b> | เดือน: <b>${targetMonthYear}</b>
            </td>
          </tr>
          <tr>
            <td colspan="2" style="font-size: 8px; color: #475569;">สัญลักษณ์: / = ดี (ปกติ) , X = บกพร่อง (ต้องแจ้งซ่อม)</td>
          </tr>
        </table>

        <table class="matrix-table">
          <thead>
            <tr>
              <th style="width: 31%; font-size: 9px; padding: 3px; border: 1px solid #0f172a; text-align: left;">รายการตรวจสภาพรถ (Actual ประจำวัน)</th>
              ${dayHeadersHtml}
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>

        <table class="sign-table">
          <tr>
            <td>
              ลงชื่อ ...................................................<br>
              ( <b>${matchedDriverName}</b> )<br>
              พนักงานขับรถประจำคัน
            </td>
            <td>
              ลงชื่อ ...................................................<br>
              ( ................................................... )<br>
              หัวหน้างาน / ผู้ตรวจสอบ (Spot Check)
            </td>
            <td>
              ลงชื่อ ...................................................<br>
              ( <b>นางสาว.นิชานันท์ เอื้อจิรพรชัย</b> )<br>
              ผู้รับเหมาขนส่ง / TTMK
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const pdfBlob = Utilities.newBlob(monthlyHtml, 'text/html', 'monthly.html').getAs('application/pdf');
    const safePlate = targetPlate.replace(/[^a-zA-Z0-9ก-๙]/g, '');
    const cleanMonth = targetMonthYear.replace(/[^0-9]/g, '');
    const fileName = `TTMK_Monthly_${cleanMonth}_${safePlate}.pdf`;
    const pdfFile = folder.createFile(pdfBlob.setName(fileName));
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return responseJSON({
      status: "success",
      message: `สร้างรายงานสรุปรายเดือนสำเร็จ (${targetPlate})`,
      pdfUrl: pdfFile.getUrl()
    });

  } catch(err) {
    return responseJSON({ status: "error", message: "Monthly PDF Error: " + err.toString() });
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
