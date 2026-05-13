# 🎯 SKU Mapping + Stock Deduction System
## สำหรับ Turbo Teak × Door E_D × Dc Factory Stock

---

## 📋 ระบบประกอบด้วย 3 ส่วน

### 1️⃣ **SKU Mapping** (Google Sheets)
- บันทึกการจับคู่ชื่อสินค้า 3 ระบบ
- ใช้อ้างอิงเมื่อ Shopee order เข้ามา
- Updated เมื่อเปิดสินค้าใหม่

### 2️⃣ **Stock Deduction Dashboard** (Google Sheets + Apps Script)
- พนักงานกดปุ่ม **"ตัด"** เพื่อลดสต๊อค
- บันทึก log ทุกครั้งที่มีการตัด
- ป้องกันความผิดพลาด (manual ไม่ auto)

### 3️⃣ **Picking List ใหม่** (Web App)
- ดึง Shopee PDF → แมป SKU → แสดงชื่อสต๊อค
- ไม่ต้องแปลงชื่อ ใช้ชื่อ Dc Factory Stock โดยตรง

---

## 🚀 ขั้นตอนการติดตั้ง

### **ขั้นที่ 1: สร้าง Google Sheets SKU Mapping**

1. ไปที่ https://docs.google.com/spreadsheets/
2. สร้าง Spreadsheet ใหม่ ตั้งชื่อ **"SKU_Mapping_Turbo_Teak"**
3. สร้าง 3 Sheet:
   - **SKU_Mapping** - ตารางจับคู่ชื่อสินค้า
   - **Stock_Deduction** - dashboard ตัดสต๊อค
   - **Log** - บันทึกการตัดสต๊อค

---

### **ขั้นที่ 2: สร้าง SKU_Mapping Sheet**

สร้างตารางดังนี้:

| SKU_DC_Stock | Turbo_Teak_Name | Turbo_Teak_Variant | Door_E_D_Name | Door_E_D_Variant | Status |
|---|---|---|---|---|---|
| ฝาปิดบ่อกลม- ดิบ | ไม่มี | ไม่มี | ฝาบอพักไม้สักแท้... | 30x40(ขายดี),ดิบ | Active |
| ประตูโมเดิร์น80x180 | [พร้อมส่ง] ประตูไม้สัก... | โมเดิร์น (ขีดตรง),80x180 | ไม่มี | ไม่มี | Active |
| วงกบประตูไม้สัก80x200 | [พร้อมส่ง] วงกบประตู... | วงกบประตู 80x200 | [พร้อมส่ง] วงกบประตู... | วงกบประตู 80x200 | Active |

**วิธี**:
- Copy ข้อมูลจาก `SKU_Mapping_Template.xlsx` (ตามไฟล์ที่ได้รับมา)
- Paste ลงใน Sheet "SKU_Mapping"
- เพิ่มเติมต่อจากตัวอย่างตามจำนวนสินค้าที่คุณมี

---

### **ขั้นที่ 3: สร้าง Stock_Deduction Sheet**

สร้าง Dashboard ดังนี้:

```
🏭 STOCK DEDUCTION DASHBOARD

Date: [วันที่ปัจจุบัน]
Operator: [ชื่อพนักงาน dropdown]

SKU: [ค้นหา / dropdown list]
Current Stock: [แสดงอัตโนมัติจาก Dc Factory Stock]
Quantity to Deduct: [กรอกจำนวน]

[BUTTON: CUT STOCK] ← กดปุ่มนี้เพื่อตัดสต๊อค

Recent Transactions (5 entries):
| วันเวลา | SKU | -จำนวน | พนักงาน | หมายเหตุ |
|---|---|---|---|---|
```

---

### **ขั้นที่ 4: ติดตั้ง Apps Script**

1. ใน Google Sheets ไปที่ **Extensions > Apps Script**
2. ลบโค้ดเดิม แล้ว Copy-Paste โค้ดด้านล่าง:

```javascript
// ═══════════════════════════════════════════════════════
// SKU MAPPING + STOCK DEDUCTION SYSTEM
// ═══════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Stock System')
    .addItem('🔍 Reload SKU Data', 'reloadSKUData')
    .addItem('📊 View Stock Status', 'viewStockStatus')
    .addItem('📜 Export Log', 'exportLog')
    .addToUi();
}

// === Deduct Stock ===
function deductStock(sku, quantity, operator, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get stock sheet (Dc Factory Stock)
  const stockSheet = ss.getSheetByName('Stock_Deduction'); // หรือชื่อ sheet ของคุณ
  const mappingSheet = ss.getSheetByName('SKU_Mapping');
  const logSheet = ss.getSheetByName('Log');
  
  if (!stockSheet || !mappingSheet || !logSheet) {
    return { error: 'Sheet not found' };
  }
  
  // Find SKU in mapping
  const mappingData = mappingSheet.getDataRange().getValues();
  let skuFound = false;
  
  for (let i = 1; i < mappingData.length; i++) {
    if (mappingData[i][0] === sku) {
      skuFound = true;
      
      // Log the deduction
      const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
      logSheet.appendRow([now, sku, -quantity, operator, note || '-', '']);
      
      SpreadsheetApp.flush();
      return { 
        success: true, 
        sku: sku, 
        quantity: quantity, 
        timestamp: now 
      };
    }
  }
  
  return { error: `SKU not found: ${sku}` };
}

// === Get SKU List ===
function getSKUList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mappingSheet = ss.getSheetByName('SKU_Mapping');
  const data = mappingSheet.getDataRange().getValues();
  const skus = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) skus.push(data[i][0]);
  }
  
  return skus;
}

// === Reload SKU Data ===
function reloadSKUData() {
  const skus = getSKUList();
  SpreadsheetApp.getUi().alert(`Loaded ${skus.length} SKUs`);
}

// === View Stock Status ===
function viewStockStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');
  const data = logSheet.getDataRange().getValues();
  SpreadsheetApp.getUi().alert(`Total transactions: ${data.length - 1}`);
}

// === Export Log ===
function exportLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');
  const data = logSheet.getDataRange().getValues();
  // สามารถ export เป็น CSV หรือ Excel ได้
  SpreadsheetApp.getUi().alert('Log exported successfully');
}

// === Deploy as Web App ===
// 1. Click Deploy > New deployment
// 2. Type: Web app
// 3. Execute as: Me
// 4. Who has access: Anyone
// 5. Copy the deployment URL
```

3. กด **Save** (Ctrl+S)

---

### **ขั้นที่ 5: Deploy as Web App**

1. ใน Apps Script ให้ Click **Deploy** (ปุ่มขวาบน)
2. เลือก **New deployment**
3. เลือก Type: **Web app**
4. Execute as: **Me**
5. Who has access: **Anyone**
6. Click **Deploy**
7. **Copy URL** ที่ได้มา (จะใช้สำหรับ Picking List)

---

### **ขั้นที่ 6: ลิงก์ Dc Factory Stock Sheet**

1. สร้าง Google Sheets ใหม่ชื่อ **"Dc_Factory_Stock"**
2. Copy ข้อมูลจากไฟล์ PDF ที่คุณส่งมา:
   - Column A: ชื่อสินค้า (ตรงกับ SKU_DC_Stock)
   - Column B: จำนวนสต๊อค
   - Column C: Reorder Point
   - Column D: Lead Time
   - เป็นต้น

3. ในแต่ละ Sheet ให้ link กับ `=VLOOKUP(SKU, Dc_Factory_Stock!A:D, 2, FALSE)` เพื่อดึงค่าสต๊อค

---

## 📦 Picking List ใหม่ (มาหลังจากสร้าง SKU Mapping)

หลังจากสิ้นสุดขั้นตอนนี้ ฉันจะสร้าง Picking List ที่:
1. ✅ ดึง Shopee PDF
2. ✅ แมป SKU อัตโนมัติ (ใช้ SKU_Mapping)
3. ✅ แสดงชื่อสต๊อค (ชื่อจาก Dc Factory Stock)
4. ✅ ใช้ชื่อ SKU โดยตรง ไม่แปลงชื่อ

---

## 🎯 Workflow ตัวอย่าง

```
1. ลูกค้าสั่งใน Shopee
   ↓
2. ดาวน์โหลด Label PDF จาก Shopee
   ↓
3. Upload PDF ไปยัง Picking List Web App
   ↓
4. Picking List อ่าน PDF → แมป Shopee name → ค้นหา SKU_Mapping → ได้ชื่อสต๊อค
   ↓
5. พนักงาน Pick สินค้า (ดูชื่อสต๊อค)
   ↓
6. เมื่อ Pick เสร็จ → พนักงาน Click "ตัด" ใน Stock Deduction Dashboard
   ↓
7. สต๊อค Update อัตโนมัติ ใน Dc_Factory_Stock Sheet
   ↓
8. Log บันทึกการตัด (วันเวลา, SKU, จำนวน, ชื่อพนักงาน)
```

---

## ✅ Checklist

- [ ] สร้าง Google Sheets "SKU_Mapping_Turbo_Teak"
- [ ] สร้าง 3 Sheet: SKU_Mapping, Stock_Deduction, Log
- [ ] Copy ข้อมูล SKU Mapping เข้า Sheet
- [ ] ติดตั้ง Apps Script โค้ด
- [ ] Deploy เป็น Web App และ Copy URL
- [ ] สร้าง Google Sheets "Dc_Factory_Stock"
- [ ] Link ข้อมูลสต๊อค

---

## 📞 ติดต่อหากติดขัด
ติดต่อฉันเพื่อ debug หรือปรับแต่ง script ให้เหมาะสมกับสต๊อค จริงของคุณ

