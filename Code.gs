/**
 * ระบบเบิก-คืนอุปกรณ์ (Equipment Borrow-Return Store System)
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีติดตั้ง:
 * 1. สร้าง Google Sheet ใหม่ ตั้งชื่อว่า "BorrowReturn_DB"
 * 2. สร้าง 2 ชีตในไฟล์นี้ ชื่อ "Items" และ "Log" ตามโครงสร้างด้านล่าง
 * 3. เปิด Extensions > Apps Script วางโค้ดนี้ทับ
 * 4. แก้ SHEET_ID ด้านล่างให้ตรงกับ Google Sheet ของคุณ
 * 5. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 * 6. คัดลอก URL ที่ได้ไปใส่ใน GS_URL ของ index.html
 *
 * โครงสร้างชีต "Items" (แถวที่ 1 เป็นหัวตาราง):
 * รหัสQR | ชื่ออุปกรณ์ | หมวดหมู่ | หน่วยนับ | จำนวนคงเหลือ
 *
 * โครงสร้างชีต "Log" (แถวที่ 1 เป็นหัวตาราง):
 * Timestamp | รหัสQR | ชื่ออุปกรณ์ | Action | จำนวน | ผู้ทำรายการ | แผนก/งาน | คงเหลือหลังทำรายการ
 */

const SHEET_ID = 'PUT_YOUR_GOOGLE_SHEET_ID_HERE'; // เอาจาก URL ของ Google Sheet
const ITEMS_SHEET = 'Items';
const LOG_SHEET = 'Log';

function doGet(e) {
  const action = e.parameter.action || 'getStock';
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (action === 'getStock') {
    return jsonOut(getStockData(ss));
  }
  if (action === 'getLog') {
    return jsonOut(getLogData(ss));
  }
  return jsonOut({ error: 'unknown action' });
}

function doPost(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === 'transaction') {
    return jsonOut(handleTransaction(ss, data));
  }
  return jsonOut({ error: 'unknown action' });
}

function getStockData(ss) {
  const sheet = ss.getSheetByName(ITEMS_SHEET);
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const items = rows
    .filter(r => r[0]) // skip empty rows
    .map(r => ({
      id: String(r[0]),
      name: r[1],
      category: r[2],
      unit: r[3],
      qty: Number(r[4]) || 0
    }));
  return { items: items };
}

function getLogData(ss) {
  const sheet = ss.getSheetByName(LOG_SHEET);
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const logs = rows
    .filter(r => r[0])
    .map(r => ({
      timestamp: r[0],
      id: String(r[1]),
      name: r[2],
      action: r[3],
      amount: Number(r[4]),
      user: r[5],
      note: r[6],
      remaining: Number(r[7])
    }))
    .reverse(); // newest first
  return { logs: logs };
}

function handleTransaction(ss, data) {
  const itemsSheet = ss.getSheetByName(ITEMS_SHEET);
  const logSheet = ss.getSheetByName(LOG_SHEET);
  const itemsRows = itemsSheet.getDataRange().getValues();

  // หาแถวของอุปกรณ์ตาม id (คอลัมน์ A)
  let rowIndex = -1;
  for (let i = 1; i < itemsRows.length; i++) {
    if (String(itemsRows[i][0]) === String(data.id)) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) {
    return { error: 'ไม่พบรหัสอุปกรณ์นี้' };
  }

  const currentQty = Number(itemsRows[rowIndex][4]) || 0;
  const amount = Number(data.amount) || 0;
  let newQty;

  if (data.txnAction === 'เบิก') {
    if (amount > currentQty) {
      return { error: 'จำนวนคงเหลือไม่พอ (เหลือ ' + currentQty + ' ' + itemsRows[rowIndex][3] + ')' };
    }
    newQty = currentQty - amount;
  } else if (data.txnAction === 'คืน') {
    newQty = currentQty + amount;
  } else {
    return { error: 'ไม่ระบุประเภทรายการ' };
  }

  // อัปเดตจำนวนคงเหลือ (คอลัมน์ E = index 5 ใน sheet, rowIndex+1 เพราะ header)
  itemsSheet.getRange(rowIndex + 2, 5).setValue(newQty);

  // บันทึกประวัติ
  logSheet.appendRow([
    new Date(),
    data.id,
    itemsRows[rowIndex][1], // ชื่ออุปกรณ์
    data.txnAction,
    amount,
    data.user || '',
    data.note || '',
    newQty
  ]);

  return { success: true, newQty: newQty };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
