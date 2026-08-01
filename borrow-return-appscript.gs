/**
 * ===================================================================
 * ระบบยืม-คืน / เบิกใช้ อุปกรณ์แผนก — Apps Script Backend
 * ===================================================================
 * วิธี Deploy:
 * 1. สร้าง Google Sheet ใหม่ (หรือใช้ที่มีอยู่) แล้ว Extensions > Apps Script
 * 2. วางโค้ดนี้ทั้งหมด
 * 3. เปลี่ยนค่า SHEET id ไม่ต้อง — สคริปต์นี้สร้างชีตย่อยที่ต้องใช้ให้เอง
 *    อัตโนมัติตอนรันครั้งแรก (ทั้ง "รายการอุปกรณ์" และ "ประวัติการทำรายการ")
 * 4. รันฟังก์ชัน seedItems() ครั้งเดียวก่อน (เลือกจาก dropdown ด้านบน กด Run ▶)
 *    เพื่อใส่รายการอุปกรณ์เริ่มต้นลงชีต "รายการอุปกรณ์"
 * 5. Deploy > New deployment > Web app > Execute as: Me, Who has access:
 *    Anyone with the link > Deploy > คัดลอก URL ไปใส่ใน CONFIG.GAS_URL ของ
 *    ไฟล์ borrow-return.html
 * ===================================================================
 * v1 — 2026-07: เวอร์ชันแรก — รองรับอุปกรณ์ 2 ประเภท:
 *      "borrow" (ยืม-คืนได้ เช่น ไม้ถูพื้น รองเท้าบูท) ใช้กลไกสแกน QR
 *      สลับสถานะ ว่าง <-> ถูกยืม ตามลำดับการสแกน, และ "issue" (เบิกใช้
 *      ครั้งเดียวจบ เช่น น้ำยาทำความสะอาด ถุงมือ) ที่แค่บันทึกรายการ
 *      เบิกไว้เป็นประวัติ ไม่มีสถานะค้างรอคืน
 * ===================================================================
 */

const SHEET_ITEMS = "รายการอุปกรณ์";
const SHEET_LOG = "ประวัติการทำรายการ";

const ITEM_HEADERS = ["รหัส", "ชื่ออุปกรณ์", "หน่วยนับ", "ประเภท", "จำนวนทั้งหมด", "สถานะ", "ผู้ยืมปัจจุบัน", "วันที่ยืม"];
const LOG_HEADERS = ["เวลา", "รหัส", "ชื่ออุปกรณ์", "การทำรายการ", "ผู้ทำรายการ", "จำนวน", "หมายเหตุ"];

function getItemSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ITEMS);
  if (!sheet){
    sheet = ss.insertSheet(SHEET_ITEMS);
    sheet.appendRow(ITEM_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLogSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet){
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(LOG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * ใส่รายการอุปกรณ์เริ่มต้นลงชีต "รายการอุปกรณ์" — รันด้วยมือครั้งเดียว
 * ตอนตั้งระบบใหม่ (เลือกฟังก์ชันนี้จาก dropdown แล้วกด Run ▶)
 * ถ้ารันซ้ำจะข้ามรหัสที่มีอยู่แล้ว ไม่สร้างซ้ำ
 */
function seedItems(){
  const sheet = getItemSheet();
  const existing = sheet.getDataRange().getValues();
  const existingCodes = existing.slice(1).map(r => r[0]);

  const seed = [
    ["ITM001","กระบอกฉีดพ่นแรงดัน POLO 1.5L","อัน","borrow",1],
    ["ITM002","ไม้กวาดดอกหญ้า","อัน","borrow",5],
    ["ITM003","ไม้กวาดทางมะพร้าว","อัน","borrow",2],
    ["ITM004","ไม้ถูพื้นแบบเส้นด้าย (string mop)","อัน","borrow",2],
    ["ITM005","ไม้ถูพื้นแบบผ้า (cloth mop หัวเหล็ก)","อัน","borrow",1],
    ["ITM006","ไม้ปาดน้ำ/ยางปาดพื้น (squeegee)","อัน","borrow",1],
    ["ITM007","ที่ตักผง","อัน","borrow",3],
    ["ITM008","ถังพลาสติกสีดำ (Hi Price)","ใบ","borrow",1],
    ["ITM009","รองเท้าบูทยางกันน้ำ","คู่","borrow",2],
    ["ITM010","ถุงมือผ้าฝ้าย","คู่","issue",1],
    ["ITM011","น้ำยาฆ่าเชื้อพื้นผิว Bossklein","แกลลอน","issue",1],
    ["ITM012","น้ำยาทำความสะอาด (แกลลอนสีม่วง)","แกลลอน","issue",1],
    ["ITM013","ม้วนพลาสติกใส","ม้วน","issue",1],
    ["ITM014","ผ้าเช็ดทำความสะอาด","ผืน","issue",1]
  ];

  let added = 0;
  seed.forEach(row => {
    if (existingCodes.indexOf(row[0]) === -1){
      // [รหัส, ชื่อ, หน่วยนับ, ประเภท, จำนวนทั้งหมด, สถานะ, ผู้ยืมปัจจุบัน, วันที่ยืม]
      sheet.appendRow([row[0], row[1], row[2], row[3], row[4], row[3]==="borrow" ? "ว่าง" : "", "", ""]);
      added++;
    }
  });
  Logger.log("เพิ่มอุปกรณ์ใหม่ " + added + " รายการ (ข้ามรหัสที่มีอยู่แล้ว)");
}

function findItemRow(sheet, code){
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === String(code).trim()){
      return { rowIndex: i + 1, values: data[i] };
    }
  }
  return null;
}

function doGet(e){
  const action = e.parameter.action;
  if (action === "iteminfo"){
    return jsonOut(getItemInfo(e.parameter.code));
  }
  if (action === "items"){
    return jsonOut(getAllItems());
  }
  if (action === "log"){
    return jsonOut(getLog());
  }
  return jsonOut({status:"ok", message:"ระบบยืม-คืนอุปกรณ์ backend กำลังทำงาน"});
}

function getItemInfo(code){
  const sheet = getItemSheet();
  const found = findItemRow(sheet, code);
  if (!found) return { found: false, code: code };
  const v = found.values;
  return {
    found: true,
    code: v[0], name: v[1], unit: v[2], type: v[3], totalQty: v[4],
    status: v[5], currentBorrower: v[6], borrowedAt: v[7]
  };
}

function getAllItems(){
  const sheet = getItemSheet();
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(v => ({
    code: v[0], name: v[1], unit: v[2], type: v[3], totalQty: v[4],
    status: v[5], currentBorrower: v[6], borrowedAt: v[7]
  }));
}

function getLog(){
  const sheet = getLogSheet();
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  return data.map(r => ({
    time: formatDT(r[0]), code: r[1], name: r[2], action: r[3], user: r[4], qty: r[5], note: r[6] || ""
  })).reverse(); // ล่าสุดขึ้นก่อน
}

function formatDT(v){
  if (v && typeof v.getFullYear === "function"){
    return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
  }
  return String(v || "");
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);
    const action = data.action; // "borrow" | "return" | "issue"
    const code = data.code;
    const user = (data.user || "").trim();
    const qty = data.qty || 1;
    const note = data.note || "";

    if (!code) throw new Error("ไม่พบรหัสอุปกรณ์");
    if (!user) throw new Error("กรุณาระบุชื่อผู้ทำรายการ");

    const sheet = getItemSheet();
    const found = findItemRow(sheet, code);
    if (!found) throw new Error("ไม่พบอุปกรณ์รหัส " + code + " ในระบบ");

    const v = found.values;
    const type = v[3];
    const currentStatus = v[5];

    if (type === "borrow"){
      if (action === "borrow"){
        if (currentStatus === "ถูกยืม"){
          throw new Error("อุปกรณ์นี้ถูกยืมอยู่แล้วโดย " + v[6]);
        }
        sheet.getRange(found.rowIndex, 6, 1, 3).setValues([["ถูกยืม", user, new Date()]]);
        appendLog(code, v[1], "ยืม", user, qty, note);
      } else if (action === "return"){
        if (currentStatus !== "ถูกยืม"){
          throw new Error("อุปกรณ์นี้ไม่ได้อยู่ในสถานะถูกยืม");
        }
        const borrower = v[6];
        sheet.getRange(found.rowIndex, 6, 1, 3).setValues([["ว่าง", "", ""]]);
        appendLog(code, v[1], "คืน", user, qty, note + (borrower ? (" (ผู้ยืมเดิม: " + borrower + ")") : ""));
      } else {
        throw new Error("การกระทำไม่ถูกต้องสำหรับอุปกรณ์ประเภทยืม-คืน");
      }
    } else if (type === "issue"){
      if (action !== "issue"){
        throw new Error("อุปกรณ์นี้เป็นประเภทเบิกใช้ ไม่มีการยืม-คืน");
      }
      appendLog(code, v[1], "เบิกใช้", user, qty, note);
    } else {
      throw new Error("ไม่รู้จักประเภทอุปกรณ์: " + type);
    }

    return jsonOut({ status: "ok", item: getItemInfo(code) });
  }catch(err){
    return jsonOut({ status: "error", message: err.message });
  }
}

function appendLog(code, name, actionText, user, qty, note){
  const sheet = getLogSheet();
  sheet.appendRow([new Date(), code, name, actionText, user, qty, note]);
}
