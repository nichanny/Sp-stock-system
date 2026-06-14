// ═══════════════════════════════════════════════════════
// DC FACTORY STOCK — Apps Script API  v2.1
// ═══════════════════════════════════════════════════════
const SS_NAME      = "Dc Factory Stock";
const SH_STOCK     = "Stock";
const SH_LOG       = "Log";
const SH_PRODUCTS  = "Products";

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

function getSS() {
  const files = DriveApp.getFilesByName(SS_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return createSS();
}

function createSS() {
  const ss = SpreadsheetApp.create(SS_NAME);

  const stk = ss.getActiveSheet().setName(SH_STOCK);
  stk.getRange(1,1,1,6).setValues([["sku","qty","minimum","unit","unit_cost","note"]]);

  const log = ss.insertSheet(SH_LOG);
  log.getRange(1,1,1,10).setValues([["log_id","timestamp","date","operator","action","sku","qty_change","ref","reason","status"]]);

  const prd = ss.insertSheet(SH_PRODUCTS);
  prd.getRange(1,1,1,5).setValues([["sku","display_name","convert_from","convert_qty","track_stock"]]);

  // ── สินค้าทั้งหมด (track=true นับสต๊อก, false=ไม่นับ) ──
  const prods = [
    // กระต่าย
    ["กระต่ายขูดมะพร้าว รุ่นธรรมดา","กระต่ายขูดมะพร้าว รุ่นธรรมดา","",1,true],
    ["กระต่ายขูดมะพร้าว รุ่นธรรมดา(เคลือบใส)","กระต่ายขูดมะพร้าว รุ่นธรรมดา(เคลือบใส)","",1,true],
    ["กระต่ายขูดมะพร้าว รุ่นพิเศษ","กระต่ายขูดมะพร้าว รุ่นพิเศษ","",1,true],
    ["กระต่ายขูดมะพร้าว รุ่นพิเศษ(เคลือบใส)","กระต่ายขูดมะพร้าว รุ่นพิเศษ(เคลือบใส)","",1,true],
    // กาว/ขี้เลื่อย/หนวดกุ้ง — ไม่นับ
    ["กาวร้อน","กาวร้อน","",1,false],
    ["กาวร้อน (ยกกล่อง 30 ขวด)","กาวร้อน (ยกกล่อง 30 ขวด)","",1,false],
    ["ขี้เลื่อยไม้สัก 1000g","ขี้เลื่อยไม้สัก 1000g","",1,false],
    ["ขี้เลื่อยไม้สัก 250g","ขี้เลื่อยไม้สัก 250g","",1,false],
    ["ขี้เลื่อยไม้สัก 500g","ขี้เลื่อยไม้สัก 500g","",1,false],
    ["หนวดกุ้งกาวร้อน","หนวดกุ้งกาวร้อน","",1,false],
    // ชั้นวาง
    ["ชั้นวาง 25x50x20","ชั้นวาง 25x50x20","",1,true],
    ["ชั้นวาง 25x50x20(เคลือบใส)","ชั้นวาง 25x50x20(เคลือบใส)","",1,true],
    // บานเฟี้ยม
    ["บานเฟี้ยม 40x150","บานเฟี้ยม 40x150","",1,true],
    ["บานเฟี้ยม 40x200","บานเฟี้ยม 40x200","",1,true],
    ["บานเฟี้ยม 50x100","บานเฟี้ยม 50x100","",1,true],
    ["บานเฟี้ยม 50x150","บานเฟี้ยม 50x150","",1,true],
    ["บานเฟี้ยม 50x200","บานเฟี้ยม 50x200","",1,true],
    // ประตู
    ["ประตู 5ฟักปีกนก 80x200","ประตู 5ฟักปีกนก 80x200","",1,true],
    ["ประตู 8ฟัก 80x200","ประตู 8ฟัก 80x200","",1,true],
    ["ประตูสายฝน 80x200","ประตูสายฝน 80x200","",1,true],
    ["ประตูสายฝน 80x180","ประตูสายฝน 80x180","",1,true],
    ["ประตูโมเดิร์น 70x200","ประตูโมเดิร์น 70x200","",1,true],
    ["ประตูโมเดิร์น 70x200 ขีดตรง","ประตูโมเดิร์น 70x200 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 70x200 ขีดบวก+","ประตูโมเดิร์น 70x200 ขีดบวก+","",1,true],
    ["ประตูโมเดิร์น 80x180","ประตูโมเดิร์น 80x180","",1,true],
    ["ประตูโมเดิร์น 80x180 ขีดตรง","ประตูโมเดิร์น 80x180 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 80x180 ขีดบวก+","ประตูโมเดิร์น 80x180 ขีดบวก+","",1,true],
    ["ประตูโมเดิร์น 80x200","ประตูโมเดิร์น 80x200","",1,true],
    ["ประตูโมเดิร์น 80x200 ขีดตรง","ประตูโมเดิร์น 80x200 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 80x200 ขีดบวก+","ประตูโมเดิร์น 80x200 ขีดบวก+","",1,true],
    ["ประตูโมเดิร์น 80x200(สักทอง)","ประตูโมเดิร์น 80x200(สักทอง)","",1,true],
    ["ประตูโมเดิร์น 90x200","ประตูโมเดิร์น 90x200","",1,true],
    ["ประตูโมเดิร์น 90x200 ขีดตรง","ประตูโมเดิร์น 90x200 ขีดตรง","",1,true],
    ["ประตูโมเดิร์น 90x200 ขีดบวก+","ประตูโมเดิร์น 90x200 ขีดบวก+","",1,true],
    // ปุ่มมือจับ
    ["ปุ่มมือจับ (29 มม.)","ปุ่มมือจับ (29 มม.)","",1,true],
    ["ปุ่มมือจับ (29 มม.,เคลือบใส)","ปุ่มมือจับ (29 มม.,เคลือบใส)","",1,true],
    ["ปุ่มมือจับ (33 มม.)","ปุ่มมือจับ (33 มม.)","",1,true],
    ["ปุ่มมือจับ (33 มม.,เคลือบใส)","ปุ่มมือจับ (33 มม.,เคลือบใส)","",1,true],
    // ฝาปิด
    ["ฝาปิดกลม","ฝาปิดกลม","",1,true],
    ["ฝาปิดกลม(สักทอง)","ฝาปิดกลม(สักทอง)","",1,true],
    ["ฝาปิดกลม(เคลือบใส)","ฝาปิดกลม(เคลือบใส)","",1,true],
    ["ฝาปิดเหลี่ยม 30x40","ฝาปิดเหลี่ยม 30x40","",1,true],
    ["ฝาปิดเหลี่ยม 30x40(สักทอง)","ฝาปิดเหลี่ยม 30x40(สักทอง)","",1,true],
    ["ฝาปิดเหลี่ยม 30x40(เคลือบใส)","ฝาปิดเหลี่ยม 30x40(เคลือบใส)","",1,true],
    ["ฝาปิดเหลี่ยม 35x40","ฝาปิดเหลี่ยม 35x40","",1,true],
    ["ฝาปิดเหลี่ยม 35x40(สักทอง)","ฝาปิดเหลี่ยม 35x40(สักทอง)","",1,true],
    ["ฝาปิดเหลี่ยม 35x40(เคลือบใส)","ฝาปิดเหลี่ยม 35x40(เคลือบใส)","",1,true],
    ["ฝาปิดเหลี่ยม 40x50","ฝาปิดเหลี่ยม 40x50","",1,true],
    ["ฝาปิดเหลี่ยม 40x50(สักทอง)","ฝาปิดเหลี่ยม 40x50(สักทอง)","",1,true],
    ["ฝาปิดเหลี่ยม 40x50(เคลือบใส)","ฝาปิดเหลี่ยม 40x50(เคลือบใส)","",1,true],
    ["ฝาปิดเหลี่ยม 50x60","ฝาปิดเหลี่ยม 50x60","",1,true],
    ["ฝาปิดเหลี่ยม 50x60(สักทอง)","ฝาปิดเหลี่ยม 50x60(สักทอง)","",1,true],
    ["ฝาปิดเหลี่ยม 50x60(เคลือบใส)","ฝาปิดเหลี่ยม 50x60(เคลือบใส)","",1,true],
    // ราวบันได (ราวบันได 20 ไม่นับ)
    ["ราวบันได 20","ราวบันได 20","",1,false],
    ["ราวบันได 50","ราวบันได 50","ราวบันได 200",1,true],
    ["ราวบันได 90","ราวบันได 90","ราวบันได 200",1,true],
    ["ราวบันได 180","ราวบันได 180","ราวบันได 200",1,true],
    ["ราวบันได 180(สักทอง)","ราวบันได 180(สักทอง)","",1,true],
    ["ราวบันได 180(เคลือบใส)","ราวบันได 180(เคลือบใส)","",1,true],
    ["ราวบันได 200","ราวบันได 200","",1,true],
    // วงกบกลม
    ["วงกบกลม 40","วงกบกลม 40","",1,true],
    ["วงกบกลม 50","วงกบกลม 50","",1,true],
    ["วงกบกลม 60","วงกบกลม 60","",1,true],
    ["วงกบกลม 80","วงกบกลม 80","",1,true],
    ["วงกบกลม 100","วงกบกลม 100","",1,true],
    ["วงกบกลม 120","วงกบกลม 120","",1,true],
    ["วงกบกลม 150","วงกบกลม 150","",1,true],
    ["วงกบกลม 3 ช่อง 40","วงกบกลม 3 ช่อง 40","",1,true],
    ["วงกบกลม 3 ช่อง 50","วงกบกลม 3 ช่อง 50","",1,true],
    ["วงกบกลม 3 ช่อง 60","วงกบกลม 3 ช่อง 60","",1,true],
    ["วงกบกลม 3 ช่อง 80","วงกบกลม 3 ช่อง 80","",1,true],
    ["วงกบกลม 3 ช่อง 100","วงกบกลม 3 ช่อง 100","",1,true],
    ["วงกบกลม 3 ช่อง 120","วงกบกลม 3 ช่อง 120","",1,true],
    ["วงกบกลม 3 ช่อง 150","วงกบกลม 3 ช่อง 150","",1,true],
    // วงกบประตู
    ["วงกบประตู 40x200","วงกบประตู 40x200","",1,true],
    ["วงกบประตู 50x50","วงกบประตู 50x50","",1,true],
    ["วงกบประตู 70x200","วงกบประตู 70x200","",1,true],
    ["วงกบประตู 80x180","วงกบประตู 80x180","วงกบประตู 80x200",1,true],
    ["วงกบประตู 80x200","วงกบประตู 80x200","",1,true],
    ["วงกบประตู 90x200","วงกบประตู 90x200","",1,true],
    ["วงกบประตู 100x120","วงกบประตู 100x120","",1,true],
    ["วงกบประตู 100x200","วงกบประตู 100x200","",1,true],
    ["วงกบประตู 160x200","วงกบประตู 160x200","",1,true],
    ["วงกบประตู 180x200","วงกบประตู 180x200","",1,true],
    ["วงกบประตู 200x200","วงกบประตู 200x200","",1,true],
    // หน้าต่าง
    ["หน้าต่างสายฝน 50x100","หน้าต่างสายฝน 50x100","",1,true],
    ["หน้าต่างสายฝน 50x110","หน้าต่างสายฝน 50x110","",1,true],
    ["หน้าต่างสายฝน 60x100","หน้าต่างสายฝน 60x100","",1,true],
    ["หน้าต่างสายฝน 60x110","หน้าต่างสายฝน 60x110","",1,true],
    // เก้าอี้
    ["เก้าอี้น้อย 30x17x10","เก้าอี้น้อย 30x17x10","",1,true],
    ["เก้าอี้น้อย 30x17x10(เคลือบใส)","เก้าอี้น้อย 30x17x10(เคลือบใส)","",1,true],
    ["เก้าอี้น้อย 33x22x10","เก้าอี้น้อย 33x22x10","",1,true],
    ["เก้าอี้น้อย 33x22x10(เคลือบใส)","เก้าอี้น้อย 33x22x10(เคลือบใส)","",1,true],
    // ฝาปิดบ่อบำบัด
    ["ฝาปิดบ่อบำบัด ไม้สัก","ฝาปิดบ่อบำบัด ไม้สัก","",1,true]
  ];

  prd.getRange(2,1,prods.length,5).setValues(prods);

  // stock เริ่มต้น unlimited เฉพาะ track=true
  const stks = prods.filter(p=>p[4]).map(p=>[p[0],999999,10,"ชิ้น",0,""]);
  stk.getRange(2,1,stks.length,6).setValues(stks);

  return ss;
}

function getStock() {
  const data = getSS().getSheetByName(SH_STOCK).getDataRange().getValues();
  const hdrs = data[0], stock = {};
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const obj = {};
    hdrs.forEach((h,i) => obj[h] = row[i]);
    stock[row[0]] = obj;
  });
  return { success:true, stock };
}

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

function getProducts() {
  const data = getSS().getSheetByName(SH_PRODUCTS).getDataRange().getValues();
  const hdrs = data[0];
  return { success:true, products: data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h,i) => obj[h] = row[i]);
    return obj;
  })};
}

function deductStock(d) {
  const ss=getSS(), stk=ss.getSheetByName(SH_STOCK);
  const prd=ss.getSheetByName(SH_PRODUCTS), log=ss.getSheetByName(SH_LOG);
  const sku=d.sku, qty=parseInt(d.qty)||1;
  const operator=d.operator||"ระบบ", ref=d.ref||"";
  const pData=prd.getDataRange().getValues();
  const pRow=pData.slice(1).find(r=>r[0]===sku);
  const convertFrom=pRow?(pRow[2]||""):"";
  const now=new Date(), ts=fmt(now), ds=fmtDate(now), lid="LOG"+now.getTime();
  const entries=convertFrom?[{sku:convertFrom,qtyChange:-qty}]:[{sku,qtyChange:-qty}];
  const sData=stk.getDataRange().getValues();
  entries.forEach(en=>{
    for(let i=1;i<sData.length;i++){
      if(sData[i][0]===en.sku){
        const cur=sData[i][1];
        stk.getRange(i+1,2).setValue(cur===999999?999999:cur+en.qtyChange);
        break;
      }
    }
    log.appendRow([lid,ts,ds,operator,
      convertFrom?"ตัดสต๊อก (auto-convert จาก "+convertFrom+")":"ตัดสต๊อก",
      en.sku,en.qtyChange,ref,"","active"]);
  });
  return {success:true,logId:lid,entries,convertFrom};
}

function addStock(d) {
  const ss=getSS(), stk=ss.getSheetByName(SH_STOCK), log=ss.getSheetByName(SH_LOG);
  const sku=d.sku, qty=parseInt(d.qty)||1, operator=d.operator||"ไม่ระบุ";
  const ref=d.ref||"", type=d.type||"เติมสต๊อก", reason=d.reason||"";
  const now=new Date(), ts=fmt(now), ds=fmtDate(now), lid="LOG"+now.getTime();
  const sData=stk.getDataRange().getValues();
  for(let i=1;i<sData.length;i++){
    if(sData[i][0]===sku){
      const cur=sData[i][1];
      stk.getRange(i+1,2).setValue(cur===999999?999999:cur+qty);
      break;
    }
  }
  log.appendRow([lid,ts,ds,operator,type,sku,+qty,ref,reason,"active"]);
  return {success:true,logId:lid};
}

function cancelLog(d) {
  const ss=getSS(), log=ss.getSheetByName(SH_LOG), stk=ss.getSheetByName(SH_STOCK);
  const lid=d.logId, operator=d.operator||"ไม่ระบุ", reason=d.reason||"";
  const lData=log.getDataRange().getValues(), toRev=[];
  for(let i=1;i<lData.length;i++){
    if(lData[i][0]===lid&&lData[i][9]==="active"){
      toRev.push({rowIdx:i+1,sku:lData[i][5],qtyChange:lData[i][6]});
      log.getRange(i+1,10).setValue("cancelled");
    }
  }
  if(!toRev.length) return {error:"ไม่พบ Log หรือถูกยกเลิกแล้ว"};
  const now=new Date(), ts=fmt(now), ds=fmtDate(now), nlid="LOG"+now.getTime();
  const sData=stk.getDataRange().getValues();
  toRev.forEach(en=>{
    const rev=-en.qtyChange;
    for(let i=1;i<sData.length;i++){
      if(sData[i][0]===en.sku){
        const cur=sData[i][1];
        stk.getRange(i+1,2).setValue(cur===999999?999999:cur+rev);
        break;
      }
    }
    log.appendRow([nlid,ts,ds,operator,"ยกเลิก (อ้างอิง "+lid+")",en.sku,rev,lid,reason,"active"]);
  });
  return {success:true,reversed:toRev.length};
}

function setMinimum(d) {
  const stk=getSS().getSheetByName(SH_STOCK), sData=stk.getDataRange().getValues();
  for(let i=1;i<sData.length;i++){
    if(sData[i][0]===d.sku){
      stk.getRange(i+1,3).setValue(parseInt(d.minimum)||0);
      stk.getRange(i+1,5).setValue(parseFloat(d.unit_cost)||0);
      return {success:true};
    }
  }
  return {error:"ไม่พบ SKU: "+d.sku};
}

function fmt(d){return Utilities.formatDate(d,"Asia/Bangkok","dd/MM/yyyy HH:mm:ss");}
function fmtDate(d){return Utilities.formatDate(d,"Asia/Bangkok","dd/MM/yyyy");}
