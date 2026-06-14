// ═══════════════════════════════════════════════════════
// DC FACTORY STOCK — Apps Script API  v2.0
// ═══════════════════════════════════════════════════════
const SS_NAME      = "Dc Factory Stock";
const SH_STOCK     = "Stock";
const SH_LOG       = "Log";
const SH_PRODUCTS  = "Products";

// ─── ROUTER ───────────────────────────────────────────
function doGet(e) {
  const a = e.parameter.action;
  let r;
  try {
    if      (a === "getStock")    r = getStock();
    else if (a === "getLog")      r = getLog();
    else if (a === "getProducts") r = getProducts();
    else r = { error: "Unknown action: " + a };
  } catch(err) { r = { error: err.message }; }
  return out(r);
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const a = d.action;
  let r;
  try {
    if      (a === "deductStock") r = deductStock(d);
    else if (a === "addStock")    r = addStock(d);
    else if (a === "cancelLog")   r = cancelLog(d);
    else if (a === "setMinimum")  r = setMinimum(d);
    else r = { error: "Unknown action: " + a };
  } catch(err) { r = { error: err.message }; }
  return out(r);
}

function out(r) {
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET SPREADSHEET ──────────────────────────────────
function getSS() {
  const files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return createSS();
}

// ─── CREATE SPREADSHEET ───────────────────────────────
function createSS() {
  const ss = SpreadsheetApp.create(SS_NAME);

  // Sheet: Stock — sku | qty | minimum | unit | unit_cost | note
  const stk = ss.getActiveSheet().setName(SH_STOCK);
  stk.getRange(1,1,1,6).setValues([["sku","qty","minimum","unit","unit_cost","note"]]);

  // Sheet: Log — log_id | timestamp | date | operator | action | sku | qty_change | ref | reason | status
  const log = ss.insertSheet(SH_LOG);
  log.getRange(1,1,1,10).setValues([["log_id","timestamp","date","operator","action","sku","qty_change","ref","reason","status"]]);

  // Sheet: Products — sku | display_name | convert_from | convert_qty | track_stock
  const prd = ss.insertSheet(SH_PRODUCTS);
  prd.getRange(1,1,1,5).setValues([["sku","display_name","convert_from","convert_qty","track_stock"]]);

  // Default products
  const prods = [
    ["ฝาปิดกลม","ฝาปิดกลม","",1,true],
    ["วงกบประตู 80x200","วงกบประตู 80x200","",1,true],
    ["วงกบประตู 80x180","วงกบประตู 80x180","วงกบประตู 80x200",1,true],
    ["วงกบประตู 100x200","วงกบประตู 100x200","",1,true],
    ["วงกบประตู 160x200","วงกบประตู 160x200","",1,true],
    ["ประตูสายฝน 80x200","ประตูสายฝน 80x200","",1,true],
    ["ประตูโมเดิร์น 80x180","ประตูโมเดิร์น 80x180","",1,true],
    ["ประตูโมเดิร์น 80x180 ขีดตรง","ประตูโมเดิร์น 80x180 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 80x180 ขีดบวก+","ประตูโมเดิร์น 80x180 ขีดบวก+","",1,true],
    ["ประตูโมเดิร์น 80x200 ขีดตรง","ประตูโมเดิร์น 80x200 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 80x200 ขีดบวก+","ประตูโมเดิร์น 80x200 ขีดบวก+","",1,true],
    ["ประตูโมเดิร์น 90x200 ขีดบวก+","ประตูโมเดิร์น 90x200 ขีดบวก+","",1,true],
    ["บานเฟี้ยมไม้สัก","บานเฟี้ยมไม้สัก","",1,true],
    ["กาวร้อน","กาวร้อน","",1,false],
    ["ขี้เลื่อยไม้สัก","ขี้เลื่อยไม้สัก","",1,false],
    ["กระต่ายขูดมะพร้าวไม้สัก รุ่นพิเศษ(เคลือบใส)","กระต่ายขูดมะพร้าวไม้สัก รุ่นพิเศษ(เคลือบใส)","",1,true]
  ];
  prd.getRange(2,1,prods.length,5).setValues(prods);

  // Default stock (unlimited = 999999, unit_cost = 0)
  const stks = prods.filter(p=>p[4]).map(p=>[p[0],999999,10,"ชิ้น",0,""]);
  stk.getRange(2,1,stks.length,6).setValues(stks);

  return ss;
}

// ─── GET STOCK ────────────────────────────────────────
function getStock() {
  const data = getSS().getSheetByName(SH_STOCK).getDataRange().getValues();
  const hdrs = data[0];
  const stock = {};
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const obj = {};
    hdrs.forEach((h,i) => obj[h] = row[i]);
    stock[row[0]] = obj;
  });
  return { success:true, stock };
}

// ─── GET LOG ──────────────────────────────────────────
function getLog() {
  const data = getSS().getSheetByName(SH_LOG).getDataRange().getValues();
  if (data.length <= 1) return { success:true, log:[] };
  const hdrs = data[0];
  const log = data.slice(1).reverse().map(row => {
    const obj = {};
    hdrs.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
  return { success:true, log };
}

// ─── GET PRODUCTS ─────────────────────────────────────
function getProducts() {
  const data = getSS().getSheetByName(SH_PRODUCTS).getDataRange().getValues();
  const hdrs = data[0];
  const products = data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
  return { success:true, products };
}

// ─── DEDUCT STOCK (called from picking list) ──────────
function deductStock(d) {
  // d: { operator, sku, qty, ref, action }
  const ss  = getSS();
  const stk = ss.getSheetByName(SH_STOCK);
  const prd = ss.getSheetByName(SH_PRODUCTS);
  const log = ss.getSheetByName(SH_LOG);

  const sku      = d.sku;
  const qty      = parseInt(d.qty) || 1;
  const operator = d.operator || "ระบบ";
  const ref      = d.ref || "";

  // Check for auto-convert rule
  const pData = prd.getDataRange().getValues();
  const pRow  = pData.slice(1).find(r => r[0] === sku);
  const convertFrom = pRow ? (pRow[2] || "") : "";

  const now = new Date();
  const ts  = fmt(now);
  const ds  = fmtDate(now);
  const lid = "LOG" + now.getTime();

  const entries = convertFrom
    ? [{ sku: convertFrom, qtyChange: -qty }]
    : [{ sku, qtyChange: -qty }];

  const sData = stk.getDataRange().getValues();
  entries.forEach(en => {
    for (let i=1; i<sData.length; i++) {
      if (sData[i][0] === en.sku) {
        const cur = sData[i][1];
        stk.getRange(i+1,2).setValue(cur===999999 ? 999999 : cur + en.qtyChange);
        break;
      }
    }
    log.appendRow([lid, ts, ds, operator,
      convertFrom ? "ตัดสต๊อก (auto-convert จาก "+convertFrom+")" : "ตัดสต๊อก",
      en.sku, en.qtyChange, ref, "", "active"]);
  });

  return { success:true, logId:lid, entries, convertFrom };
}

// ─── ADD STOCK (เติม / คืน) ───────────────────────────
function addStock(d) {
  // d: { operator, sku, qty, ref, type, reason }
  const ss  = getSS();
  const stk = ss.getSheetByName(SH_STOCK);
  const log = ss.getSheetByName(SH_LOG);

  const sku      = d.sku;
  const qty      = parseInt(d.qty) || 1;
  const operator = d.operator || "ไม่ระบุ";
  const ref      = d.ref || "";
  const type     = d.type || "เติมสต๊อก";
  const reason   = d.reason || "";

  const now = new Date();
  const ts  = fmt(now);
  const ds  = fmtDate(now);
  const lid = "LOG" + now.getTime();

  const sData = stk.getDataRange().getValues();
  for (let i=1; i<sData.length; i++) {
    if (sData[i][0] === sku) {
      const cur = sData[i][1];
      stk.getRange(i+1,2).setValue(cur===999999 ? 999999 : cur + qty);
      break;
    }
  }
  log.appendRow([lid, ts, ds, operator, type, sku, +qty, ref, reason, "active"]);
  return { success:true, logId:lid };
}

// ─── CANCEL LOG ───────────────────────────────────────
function cancelLog(d) {
  // d: { logId, operator, reason }
  const ss  = getSS();
  const log = ss.getSheetByName(SH_LOG);
  const stk = ss.getSheetByName(SH_STOCK);

  const lid      = d.logId;
  const operator = d.operator || "ไม่ระบุ";
  const reason   = d.reason || "";

  const lData = log.getDataRange().getValues();
  const toRev = [];

  for (let i=1; i<lData.length; i++) {
    if (lData[i][0] === lid && lData[i][9] === "active") {
      toRev.push({ rowIdx: i+1, sku: lData[i][5], qtyChange: lData[i][6] });
      log.getRange(i+1,10).setValue("cancelled");
    }
  }
  if (!toRev.length) return { error: "ไม่พบ Log หรือถูกยกเลิกแล้ว" };

  const now  = new Date();
  const ts   = fmt(now);
  const ds   = fmtDate(now);
  const nlid = "LOG" + now.getTime();
  const sData = stk.getDataRange().getValues();

  toRev.forEach(en => {
    const rev = -en.qtyChange;
    for (let i=1; i<sData.length; i++) {
      if (sData[i][0] === en.sku) {
        const cur = sData[i][1];
        stk.getRange(i+1,2).setValue(cur===999999 ? 999999 : cur + rev);
        break;
      }
    }
    log.appendRow([nlid, ts, ds, operator,
      "ยกเลิก (อ้างอิง "+lid+")", en.sku, rev, lid, reason, "active"]);
  });

  return { success:true, reversed: toRev.length };
}

// ─── SET MINIMUM + UNIT COST ──────────────────────────
function setMinimum(d) {
  // d: { sku, minimum, unit_cost }
  const stk   = getSS().getSheetByName(SH_STOCK);
  const sData = stk.getDataRange().getValues();
  for (let i=1; i<sData.length; i++) {
    if (sData[i][0] === d.sku) {
      stk.getRange(i+1,3).setValue(parseInt(d.minimum)||0);
      stk.getRange(i+1,5).setValue(parseFloat(d.unit_cost)||0);
      return { success:true };
    }
  }
  return { error: "ไม่พบ SKU: " + d.sku };
}

// ─── DATE HELPERS ─────────────────────────────────────
function fmt(d) {
  return Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
}
function fmtDate(d) {
  return Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy");
}
