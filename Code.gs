// ===== Google Apps Script (วางใน Extensions > Apps Script ของ Google Sheet) =====
// Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone
// ชีตที่ใช้: Products (ข้อมูลสินค้า) / Orders (ผู้รับ) / OrderItems (รายการต่อ SKU)
//           PickList (ใบหยิบสินค้า) / Labels (ใบจ่าหน้าพัสดุ)  <- 2 ชีตหลังสร้างจากเมนู "ร้านช็อกโกแลต"

const TZ = "Asia/Bangkok";
const SS = () => SpreadsheetApp.getActiveSpreadsheet();

function sheet_(name, headers, seed) {
  let sh = SS().getSheetByName(name);
  if (!sh) {
    sh = SS().insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0a83c");
    sh.setFrozenRows(1);
    (seed || []).forEach(r => sh.appendRow(r));
  }
  return sh;
}
const productsSh_ = () => sheet_("Products",
  ["sku", "name", "desc", "price", "image", "stock"],
  [["CH-01", "ดาร์กช็อกโกแลต 70%", "เข้ม ขมนุ่ม หอมผลไม้ 50 g", 89, "dark70.jpg", 100],
   ["CH-02", "มิลค์ช็อกโกแลต 40%", "หวานมัน ละลายในปาก 50 g", 79, "milk40.jpg", 100]]);

// หน้าเว็บเรียกอ่านสินค้าจากชีต Products
function doGet() {
  const v = productsSh_().getDataRange().getValues();
  const h = v.shift();
  const list = v.filter(r => r[0]).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
  return ContentService.createTextOutput(JSON.stringify(list)).setMimeType(ContentService.MimeType.JSON);
}

// หน้าเว็บส่งออเดอร์มาบันทึก
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const p = e.parameter, items = JSON.parse(p.items || "[]");
    const orders = sheet_("Orders", ["เวลา", "เลขออเดอร์", "ชื่อผู้รับ", "เบอร์โทร", "ที่อยู่จัดส่ง", "ยอดรวม (฿)", "สถานะ"]);
    const lines = sheet_("OrderItems", ["เลขออเดอร์", "sku", "ชื่อสินค้า", "จำนวน", "ราคา/ชิ้น"]);
    const prod = productsSh_(), pv = prod.getDataRange().getValues();
    const id = "ORD-" + Utilities.formatDate(new Date(), TZ, "yyMMdd-HHmmss");
    let total = 0;
    items.forEach(it => {
      const row = pv.findIndex((r, i) => i > 0 && r[0] === it.sku);
      if (row < 0) return;
      const price = Number(pv[row][3]); total += price * it.qty;      // ใช้ราคาจากชีต ไม่เชื่อราคาจากหน้าเว็บ
      lines.appendRow([id, it.sku, pv[row][1], it.qty, price]);
      prod.getRange(row + 1, 6).setValue(Number(pv[row][5]) - it.qty); // ตัดสต็อก
    });
    orders.appendRow([new Date(), id, p.name, "'" + p.phone, p.address, total, "รอแพ็ค"]);
    return ContentService.createTextOutput("ok");
  } finally { lock.release(); }
}

// เมนูในชีต
function onOpen() {
  SpreadsheetApp.getUi().createMenu("ร้านช็อกโกแลต")
    .addItem("สร้างใบหยิบสินค้า + ใบจ่าหน้าพัสดุ", "buildPickAndLabels")
    .addItem("เปลี่ยนสถานะออเดอร์ที่รอแพ็ค → แพ็คแล้ว", "markPacked").addToUi();
}

// สร้างใบหยิบ (รวมต่อ SKU + แยกต่อออเดอร์) และใบจ่าหน้า จากออเดอร์ที่ "รอแพ็ค"
function buildPickAndLabels() {
  const ov = sheet_("Orders", ["x"]).getDataRange().getValues().slice(1).filter(r => r[6] === "รอแพ็ค");
  const ids = ov.map(r => r[1]);
  const lv = sheet_("OrderItems", ["x"]).getDataRange().getValues().slice(1).filter(r => ids.includes(r[0]));

  // PickList
  let pk = SS().getSheetByName("PickList") || SS().insertSheet("PickList"); pk.clear();
  const tot = {}; lv.forEach(r => { const k = r[1] + " " + r[2]; tot[k] = (tot[k] || 0) + Number(r[3]); });
  const pr = [["สินค้า (SKU)", "จำนวนที่ต้องหยิบรวม"], ...Object.entries(tot)];
  pk.getRange(1, 1, pr.length, 2).setValues(pr);
  pk.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#e0a83c");

  // Labels
  let lb = SS().getSheetByName("Labels") || SS().insertSheet("Labels"); lb.clear();
  const rows = [["เลขออเดอร์", "ผู้รับ", "เบอร์โทร", "ที่อยู่จัดส่ง", "รายการในพัสดุ (ตรวจตอนแพ็ค)", "แพ็คแล้ว ✓"]];
  ov.forEach(o => {
    const its = lv.filter(r => r[0] === o[1]).map(r => r[1] + " " + r[2] + " x" + r[3]).join("\n");
    rows.push([o[1], o[2], "'" + o[3], o[4], its, "☐"]);
  });
  lb.getRange(1, 1, rows.length, 6).setValues(rows).setVerticalAlignment("top").setWrap(true);
  lb.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e0a83c");
  lb.setColumnWidths(1, 3, 140); lb.setColumnWidth(4, 300); lb.setColumnWidth(5, 260);
  SpreadsheetApp.getUi().alert("สร้างแล้ว: " + ov.length + " ออเดอร์ (ดูชีต PickList และ Labels)");
}

function markPacked() {
  const sh = SS().getSheetByName("Orders"), v = sh.getDataRange().getValues();
  v.forEach((r, i) => { if (i > 0 && r[6] === "รอแพ็ค") sh.getRange(i + 1, 7).setValue("แพ็คแล้ว"); });
}
