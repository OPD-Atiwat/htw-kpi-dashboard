// ═══════════════════════════════════════════════════════════════════
//  HTW CONTENT KPI DASHBOARD — Google Apps Script
//  OpenDurian How-To | Digital Marketing Team
//  Version 1.0
// ═══════════════════════════════════════════════════════════════════
//
//  วิธีใช้ (ทำตามลำดับ):
//  1. เปิด Google Sheets ใหม่
//  2. Extensions → Apps Script
//  3. วาง Code ทั้งหมดนี้ แทน code เดิม
//  4. กด Save (Ctrl+S)
//  5. เลือก Function: setupDashboard → กด Run
//  6. อนุญาต Permission ที่ขอ
//  7. กลับไป Google Sheets จะเห็นเมนู "📊 HTW Dashboard" ขึ้นมา
// ═══════════════════════════════════════════════════════════════════


// ── ⚙️ GLOBAL SETTINGS (แก้ไขได้) ──────────────────────────────────
const SETTINGS = {
  TARGET_ROAS    : 2.5,    // ROAS ขั้นต่ำที่ต้องการ
  TARGET_CPA     : 150,    // CPA สูงสุดที่ยอมรับได้ (บาท)
  TARGET_CTR     : 0.015,  // CTR ขั้นต่ำ (1.5%)
  TARGET_CPM     : 80,     // CPM สูงสุด (บาท)
  CREATORS       : ['ริว', 'มิ้น', 'หมิว'],
  PLATFORMS      : ['Meta', 'TikTok'],
  PRODUCTS: [
    'กับดักความรู้สึกผิด',
    "The Witches' Club",
    'Ghostly Brews',
    'Ghostly Remains',
    'All Product',
  ],
};

// ── 🎨 COLOR PALETTE ────────────────────────────────────────────────
const C = {
  DARK_PURPLE : '#2D1B69',
  PURPLE      : '#6B3FA0',
  LIGHT_PURPLE: '#EDE7F6',
  DARK_GREEN  : '#1B5E20',
  LIGHT_GREEN : '#E8F5E9',
  DARK_BLUE   : '#0D47A1',
  LIGHT_BLUE  : '#E3F2FD',
  ORANGE      : '#E65100',
  LIGHT_ORANGE: '#FFF8E1',
  PASS_GREEN  : '#C8F7D8',
  FAIL_RED    : '#FFCDD2',
  WAIT_YELLOW : '#FFF9C4',
  WHITE       : '#FFFFFF',
  LIGHT_GRAY  : '#F5F5F5',
  MID_GRAY    : '#EEEEEE',
  DARK_GRAY   : '#757575',
  // Per-creator colors
  RIW_COLOR   : '#EDE7F6',
  MIN_COLOR   : '#E3F2FD',
  MIUW_COLOR  : '#FFF8E1',
};

// ── 📋 SHEET NAMES ───────────────────────────────────────────────────
const SN = {
  CONFIG  : 'Config',
  CONTENT : 'Content_DB',
  ADS     : 'Ads_Performance',
  KPI     : 'KPI_Calc',
  DASH    : 'Dashboard',
};

// ── 🔍 FIND SHEET (ค้นหา Sheet แม้ชื่อจะมี Emoji หรือ ต่างตัวพิมพ์) ───
function _findSheet(ss, name) {
  // ลองหาตรงๆ ก่อน
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  // ถ้าไม่เจอ ลองหา Sheet ที่ชื่อมีคำนี้อยู่ (เผื่อมี Emoji นำหน้า)
  const all = ss.getSheets();
  const lower = name.toLowerCase().replace(/\s/g,'');
  for (const s of all) {
    const sname = s.getName().toLowerCase().replace(/\s/g,'').replace(/[^\w]/g,'');
    const target = lower.replace(/[^\w]/g,'');
    if (sname.includes(target) || target.includes(sname)) return s;
  }
  return null;
}


// ═══════════════════════════════════════════════════════════════════
// 🚀 ENTRY POINT — Run นี้ก่อนเลย!
// ═══════════════════════════════════════════════════════════════════
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '🚀 สร้าง HTW Content KPI Dashboard',
    'Script จะสร้าง Sheet ทั้งหมดอัตโนมัติ\nใช้เวลาประมาณ 30-60 วินาที\n\nพร้อมแล้วกด OK',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  try {
    _createConfigSheet(ss);
    _createContentDBSheet(ss);
    _createAdsSheet(ss);
    _createKPISheet(ss);
    _createDashboardShell(ss);
    _cleanDefaultSheets(ss);

    ss.setActiveSheet(ss.getSheetByName(SN.DASH));
    ui.alert('✅ สร้าง Dashboard สำเร็จ!\n\nขั้นตอนต่อไป:\n→ ไปที่ Sheet "📥 Ads_Performance"\n→ นำเข้าไฟล์ HTW_AdsPerformance_Feb26_IMPORT.csv\n→ กลับมาที่เมนู HTW Dashboard → "สร้าง KPI & Charts"');
  } catch(e) {
    ui.alert('❌ เกิดข้อผิดพลาด: ' + e.message);
  }
}


// ═══════════════════════════════════════════════════════════════════
// 📊 BUILD KPI & CHARTS — Run หลังจาก Import CSV แล้ว
// ═══════════════════════════════════════════════════════════════════
function buildKPIAndCharts() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const ads = _findSheet(ss, SN.ADS);

  if (!ads || ads.getLastRow() < 2) {
    // แสดงชื่อ Sheet ทั้งหมดเพื่อ Debug
    const sheetList = ss.getSheets().map(s => s.getName()).join(', ');
    ui.alert('⚠️ หา Sheet "Ads_Performance" ไม่เจอ\n\nSheet ที่มีอยู่: ' + sheetList + '\n\nกรุณาตรวจสอบว่า Import CSV ลงใน Sheet ที่ถูกต้องแล้ว');
    return;
  }

  try {
    _buildKPICalc(ss);
    SpreadsheetApp.flush();
    _buildDashboard(ss);
    SpreadsheetApp.flush();
    try { _addConditionalFormatting(ads); } catch(fe) { /* non-critical */ }
    const dashSh = _findSheet(ss, SN.DASH);
    if (dashSh) ss.setActiveSheet(dashSh);
    ui.alert('✅ สร้าง KPI & Dashboard สำเร็จ!\n\nดูผลลัพธ์ที่ Sheet Dashboard\n\n💡 ถ้าต้องการกราฟ → เมนู HTW Dashboard → "📊 เพิ่มกราฟ"');
  } catch(e) {
    ui.alert('❌ Error: ' + e.message + '\n\nตรวจสอบว่า Import CSV เรียบร้อยแล้วหรือยัง');
  }
}


// ═══════════════════════════════════════════════════════════════════
// 🔍 DEBUG — ใช้ตรวจสอบว่า Script อ่านข้อมูลถูกไหม
// ═══════════════════════════════════════════════════════════════════
function debugData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const ads = _findSheet(ss, SN.ADS);
  if (!ads) { ui.alert('❌ หา Ads_Performance ไม่เจอ'); return; }

  const data = ads.getDataRange().getValues();
  let headerRow = -1, colMap = {};
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const row = data[r].map(v => String(v).toLowerCase().trim());
    if (row.includes('creator') || row.some(c => c.includes('creator'))) {
      headerRow = r;
      row.forEach((h,i) => colMap[h] = i);
      break;
    }
  }

  const firstDataRow = data[headerRow + 1] || [];
  const spendIdx  = colMap['spend_thb'];
  const revIdx    = colMap['revenue_thb'];
  const roasIdx   = colMap['roas'];
  const creatorIdx = colMap['creator'];

  const msg = [
    '📋 Sheet: ' + ads.getName(),
    '📊 Total rows: ' + data.length,
    '🔢 Header row index: ' + headerRow,
    '',
    '🗂 Column Mapping:',
    '  creator → col ' + creatorIdx,
    '  spend_thb → col ' + spendIdx,
    '  revenue_thb → col ' + revIdx,
    '  roas → col ' + roasIdx,
    '',
    '📌 Row 1 (header): ' + data[0].slice(0,5).join(' | '),
    '📌 Row 2 (data): ' + (data[1]||[]).slice(0,5).join(' | '),
    '',
    '💰 First data row spend_thb value: ' + firstDataRow[spendIdx],
    '💵 First data row revenue_thb value: ' + firstDataRow[revIdx],
    '📈 First data row roas value: ' + firstDataRow[roasIdx],
  ].join('\n');

  ui.alert('🔍 Debug Info', msg, ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════
// 🔄 REFRESH DASHBOARD — Run เพื่ออัปเดตข้อมูล
// ═══════════════════════════════════════════════════════════════════
function refreshDashboard() {
  _buildKPICalc(SpreadsheetApp.getActiveSpreadsheet());
  _buildDashboard(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ อัปเดต Dashboard สำเร็จ!');
}


// ═══════════════════════════════════════════════════════════════════
// 📋 SHEET 1: CONFIG
// ═══════════════════════════════════════════════════════════════════
function _createConfigSheet(ss) {
  let ws = ss.getSheetByName(SN.CONFIG) || ss.insertSheet(SN.CONFIG);
  ws.clear(); ws.setTabColor(C.DARK_PURPLE);

  // Title
  _setCell(ws, 1, 1, '⚙️ Config — HTW Content KPI Dashboard', {
    bold:true, size:14, color:C.DARK_PURPLE
  });
  ws.getRange(1,1,1,3).merge();
  ws.setRowHeight(1, 36);

  // Sections
  const sections = [
    ['📌 เกณฑ์ KPI', null, null],
    ['Target ROAS ขั้นต่ำ',   SETTINGS.TARGET_ROAS,  'ROAS หลัง 7 วัน ต้องมากกว่านี้ (ตั้งไว้ 2.5)'],
    ['Target CPA สูงสุด (฿)', SETTINGS.TARGET_CPA,   'ต้นทุนต่อ Order ไม่ควรเกิน'],
    ['Target CTR ขั้นต่ำ',    SETTINGS.TARGET_CTR,   'Click-Through Rate (1.5%)'],
    ['Target CPM สูงสุด (฿)', SETTINGS.TARGET_CPM,   'Cost per 1,000 Impressions'],
    [null, null, null],
    ['📌 ทีม Creator', null, null],
    ['Creator 1', 'ริว',  ''],
    ['Creator 2', 'มิ้น', ''],
    ['Creator 3', 'หมิว', ''],
    [null, null, null],
    ['📌 Products / หนังสือ', null, null],
    ['Product 1', 'กับดักความรู้สึกผิด',  'Trap'],
    ['Product 2', "The Witches' Club",     'Witches'],
    ['Product 3', 'Ghostly Brews',         'Ghostly'],
    ['Product 4', 'Ghostly Remains',       'Remains'],
    ['Product 5', 'All Product',           ''],
  ];

  let row = 3;
  sections.forEach(([key, val, note]) => {
    if (!key) { row++; return; }
    if (key.startsWith('📌')) {
      _setCell(ws, row, 1, key, {bold:true, size:10, color:'#FFFFFF', bg:C.DARK_PURPLE});
      ws.getRange(row, 1, 1, 3).merge().setBackground(C.DARK_PURPLE);
      ws.setRowHeight(row, 26);
    } else {
      _setCell(ws, row, 1, key,  {size:9});
      _setCell(ws, row, 2, val,  {size:9, bold:true, color:C.DARK_BLUE, bg:C.WAIT_YELLOW});
      _setCell(ws, row, 3, note, {size:9, color:C.DARK_GRAY, italic:true});
      ws.setRowHeight(row, 22);
    }
    row++;
  });

  ws.setColumnWidth(1, 220); ws.setColumnWidth(2, 260); ws.setColumnWidth(3, 320);
}


// ═══════════════════════════════════════════════════════════════════
// 📋 SHEET 2: CONTENT_DB
// ═══════════════════════════════════════════════════════════════════
function _createContentDBSheet(ss) {
  let ws = ss.getSheetByName(SN.CONTENT) || ss.insertSheet(SN.CONTENT);
  ws.clear(); ws.setTabColor(C.PURPLE);

  const TH = [
    'รหัส Content','วันที่วางแผน','Creator','รูปแบบ','วัตถุประสงค์',
    'หนังสือ / Product','มีปก?','Key Message','ประเภท Message','Brief / Idea',
    'สถานะ','วันที่ Publish จริง','FB Ad Name\n(ต้องตรงกับ Ads Manager)',
    'Link FB Post','Link TikTok','Target ROAS','ROAS (หลัง 7 วัน)',
    'วันเช็ค ROAS','ผ่านเกณฑ์?','New Message?','หมายเหตุ'
  ];
  const EN = [
    'content_id','date_planned','creator','content_format','content_objective',
    'product','is_cover','key_message','message_type','idea_brief',
    'status','published_date','fb_ad_name',
    'fb_post_link','tiktok_link','target_roas','roas_7d',
    'roas_check_date','hit_target','is_new_message','notes'
  ];
  const WIDTHS = [
    130,110,80,80,110,
    180,70,320,110,320,
    90,110,220,
    280,280,100,100,
    110,100,110,200
  ];

  // Header row 1 (Thai)
  TH.forEach((h,i) => {
    const cell = ws.getRange(1, i+1);
    cell.setValue(h)
        .setBackground(C.DARK_PURPLE).setFontColor('#FFFFFF')
        .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrap(true);
    ws.setColumnWidth(i+1, WIDTHS[i]);
  });
  ws.setRowHeight(1, 40);

  // Header row 2 (English field names)
  EN.forEach((h,i) => {
    ws.getRange(2, i+1).setValue(h)
      .setBackground(C.MID_GRAY).setFontColor(C.DARK_GRAY)
      .setFontFamily('Arial').setFontSize(7).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(2, 14);

  // Sample row
  const sample = [
    'HTW-FEB26-001','2026-02-01','ริว','VDO','Sale',
    'กับดักความรู้สึกผิด','No','คนดีมักรู้สึกผิด (Trap)','New','ให้ความรู้เรื่องความรู้สึกผิดเรื้อรัง',
    'Published','2026-02-12','[12.02.26][ริว] คนดีมักรู้สึกผิด (Trap)',
    '','',2.5,'',
    '=IF(L3<>"",L3+7,"")','=IF(Q3="","⏳ รอ",IF(Q3>=P3,"✅ ผ่าน","❌ ไม่ผ่าน"))','TRUE',''
  ];
  sample.forEach((v,i) => {
    ws.getRange(3, i+1).setValue(v)
      .setFontFamily('Arial').setFontSize(9)
      .setBackground(C.LIGHT_PURPLE)
      .setVerticalAlignment('middle');
  });
  ws.setRowHeight(3, 20);

  // Format formula cells
  ws.getRange('R3').setNumberFormat('DD-MMM-YY');

  // Data Validations (Dropdowns)
  const dv = (col, values) => {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true).setAllowInvalid(false).build();
    ws.getRange(3, col, 997).setDataValidation(rule);
  };
  dv(3,  SETTINGS.CREATORS);
  dv(4,  ['VDO','IMG','Carousel']);
  dv(5,  ['Sale','Engage','Live','Promotion','Book Fair']);
  dv(7,  ['Yes','No']);
  dv(9,  ['New','Revised','Evergreen']);
  dv(11, ['Draft','Review','Done','Published']);
  dv(20, ['TRUE','FALSE']);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(1);
  ws.getRange('A1:U2').setBorder(true,true,true,true,false,false,'#BBBBBB',SpreadsheetApp.BorderStyle.SOLID);
}


// ═══════════════════════════════════════════════════════════════════
// 📥 SHEET 3: ADS_PERFORMANCE
// ═══════════════════════════════════════════════════════════════════
function _createAdsSheet(ss) {
  let ws = ss.getSheetByName(SN.ADS) || ss.insertSheet(SN.ADS);
  ws.clear(); ws.setTabColor(C.DARK_GREEN);

  const TH = [
    'Creator','วันที่ Post','Content Name','Product','Platform',
    'Reach','Impressions','Spend (THB)','Purchases','Revenue (THB)',
    'ROAS','CPM (THB)','CPA (THB)','Clicks','CTR (%)',
    '2s View Rate','6s View Rate','Ad Variants','ผ่านเกณฑ์?'
  ];
  const EN = [
    'creator','ad_date','content_name','product','platform',
    'reach','impressions','spend_thb','purchases','revenue_thb',
    'roas','cpm_thb','cpa_thb','clicks','ctr',
    'v2s','v6s','ad_variants','hit_roas'
  ];
  const WIDTHS = [
    80,100,320,180,75,
    90,100,100,85,105,
    75,90,90,75,75,
    90,90,90,100
  ];

  // Instructions row
  const instrCell = ws.getRange(1,1);
  instrCell.setValue('📥 วิธี Import ข้อมูล: File → Import → Upload → HTW_AdsPerformance_Feb26_IMPORT.csv → Replace current sheet → Import data')
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold').setFontColor('#B71C1C')
    .setBackground('#FFEBEE').setVerticalAlignment('middle');
  ws.getRange(1,1,1,19).merge().setBackground('#FFEBEE');
  ws.setRowHeight(1, 28);

  // Header rows
  TH.forEach((h,i) => {
    ws.getRange(2, i+1).setValue(h)
      .setBackground(C.DARK_GREEN).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    ws.setColumnWidth(i+1, WIDTHS[i]);
  });
  ws.setRowHeight(2, 36);

  EN.forEach((h,i) => {
    ws.getRange(3, i+1).setValue(h)
      .setBackground('#DCEDC8').setFontColor(C.DARK_GRAY)
      .setFontFamily('Arial').setFontSize(7).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(3, 14);

  ws.setFrozenRows(3);
  ws.setFrozenColumns(1);

  // Note: actual data starts at row 4 after CSV import
  // CSV import will replace from row 1, so we rebuild headers after import
}


// ═══════════════════════════════════════════════════════════════════
// 🧮 KPI CALC — Built after data import
// ═══════════════════════════════════════════════════════════════════
function _createKPISheet(ss) {
  let ws = ss.getSheetByName(SN.KPI) || ss.insertSheet(SN.KPI);
  ws.clear(); ws.setTabColor(C.ORANGE);
  _setCell(ws, 1, 1, '🧮 KPI_Calc — สร้างอัตโนมัติหลัง Import ข้อมูล', {
    bold:true, size:12, color:C.ORANGE
  });
  ws.getRange(1,1,1,4).merge();
  _setCell(ws, 2, 1, 'กด "สร้าง KPI & Charts" จากเมนู HTW Dashboard หลัง Import CSV แล้ว', {
    size:9, color:C.DARK_GRAY, italic:true
  });
  ws.getRange(2,1,1,4).merge();
  ws.setRowHeight(1,32); ws.setRowHeight(2,20);
}


// ═══════════════════════════════════════════════════════════════════
// 🧮 BUILD KPI CALCULATIONS (runs after import)
// ═══════════════════════════════════════════════════════════════════
function _buildKPICalc(ss) {
  const ads = _findSheet(ss, SN.ADS);
  const ws  = _findSheet(ss, SN.KPI) || ss.insertSheet(SN.KPI);
  ws.clear(); ws.setTabColor(C.ORANGE);

  // Read all ads data
  const data = ads.getDataRange().getValues();
  // Find header row (look for 'creator' column)
  let headerRow = -1, colMap = {};
  // ค้นหา row ที่มี spend_thb ก่อน (English field names row)
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const row = data[r].map(v => String(v).toLowerCase().trim());
    if (row.includes('spend_thb')) {
      headerRow = r;
      row.forEach((h,i) => colMap[h] = i);
      break;
    }
  }
  // Fallback: ถ้าไม่เจอ spend_thb ให้หา creator แทน
  if (headerRow < 0) {
    for (let r = 0; r < Math.min(5, data.length); r++) {
      const row = data[r].map(v => String(v).toLowerCase().trim());
      if (row.some(c => c.includes('creator'))) {
        headerRow = r;
        row.forEach((h,i) => colMap[h] = i);
        break;
      }
    }
  }
  if (headerRow < 0) throw new Error('ไม่พบ Header row ใน Ads_Performance');

  const rows = data.slice(headerRow + 1).filter(r => r[colMap['creator']] !== '');

  // Helper: get numeric value
  const num = (row, col) => {
    const v = row[colMap[col] ?? -1];
    return (v === '' || v === null || v === undefined) ? 0 : Number(v) || 0;
  };
  const str = (row, col) => String(row[colMap[col] ?? -1] || '').trim();

  // Build summary by creator x platform
  const summary = {};
  rows.forEach(row => {
    const cr = str(row,'creator'), pl = str(row,'platform');
    if (!cr) return;
    const key = `${cr}||${pl}`;
    if (!summary[key]) summary[key] = {
      creator:cr, platform:pl, count:0, spend:0, revenue:0,
      reach:0, impressions:0, purchases:0, pass:0, fail:0, wait:0,
      cpm_sum:0, cpa_sum:0, cpm_n:0, cpa_n:0,
      v2s_sum:0, v2s_n:0, v6s_sum:0, v6s_n:0
    };
    const s = summary[key];
    s.count++;
    s.spend    += num(row,'spend_thb');
    s.revenue  += num(row,'revenue_thb');
    s.reach    += num(row,'reach');
    s.impressions += num(row,'impressions');
    s.purchases   += num(row,'purchases');
    const roas = num(row,'roas');
    if (roas >= SETTINGS.TARGET_ROAS) s.pass++;
    else if (roas > 0) s.fail++;
    else s.wait++;
    const cpm = num(row,'cpm_thb');
    if (cpm > 0) { s.cpm_sum += cpm; s.cpm_n++; }
    const cpa = num(row,'cpa_thb');
    if (cpa > 0) { s.cpa_sum += cpa; s.cpa_n++; }
    const v2 = num(row,'v2s');
    if (v2 > 0) { s.v2s_sum += v2; s.v2s_n++; }
    const v6 = num(row,'v6s');
    if (v6 > 0) { s.v6s_sum += v6; s.v6s_n++; }
  });

  // Also by product
  const byProduct = {};
  rows.forEach(row => {
    const prod = str(row,'product') || 'Other';
    if (!byProduct[prod]) byProduct[prod] = {spend:0,revenue:0,purchases:0,count:0};
    byProduct[prod].spend    += num(row,'spend_thb');
    byProduct[prod].revenue  += num(row,'revenue_thb');
    byProduct[prod].purchases+= num(row,'purchases');
    byProduct[prod].count++;
  });

  // Write KPI sheet
  // Title
  _setCell(ws,1,1,'🧮 KPI Summary — Feb 2026',{bold:true,size:14,color:C.DARK_PURPLE});
  ws.getRange(1,1,1,10).merge(); ws.setRowHeight(1,32);
  _setCell(ws,2,1,`อัปเดตล่าสุด: ${new Date().toLocaleString('th-TH')}  |  Total rows: ${rows.length}`,
    {size:9,color:C.DARK_GRAY,italic:true});
  ws.getRange(2,1,1,10).merge(); ws.setRowHeight(2,18);

  // Section 1: By Creator × Platform
  _sectionHeader(ws, 4, '📌 KPI รายบุคคล × Platform');
  const kh1 = ['Creator','Platform','Content\nPieces','Spend (฿)','Revenue (฿)',
                'ROAS','CPM avg (฿)','CPA avg (฿)','Avg 2s%','Avg 6s%','✅ ผ่าน','❌ ไม่ผ่าน','⏳ รอ','% ผ่าน'];
  kh1.forEach((h,i) => {
    ws.getRange(5,i+1).setValue(h)
      .setBackground(C.DARK_PURPLE).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  });
  ws.setRowHeight(5, 36);

  let r = 6;
  const ORDER = ['ริว','มิ้น','หมิว'];
  const PLAT_ORDER = ['Meta','TikTok'];
  const CR_COLORS = {ริว:C.RIW_COLOR, 'มิ้น':C.MIN_COLOR, หมิว:C.MIUW_COLOR};

  ORDER.forEach(cr => {
    PLAT_ORDER.forEach(pl => {
      const s = summary[`${cr}||${pl}`];
      if (!s) return;
      const roas   = s.spend > 0 ? s.revenue/s.spend : 0;
      const cpm    = s.cpm_n > 0 ? s.cpm_sum/s.cpm_n : 0;
      const cpa    = s.cpa_n > 0 ? s.cpa_sum/s.cpa_n : 0;
      const v2avg  = s.v2s_n > 0 ? s.v2s_sum/s.v2s_n : '';
      const v6avg  = s.v6s_n > 0 ? s.v6s_sum/s.v6s_n : '';
      const pct    = s.count > 0 ? (s.pass/s.count*100).toFixed(1)+'%' : '0%';
      const bg     = CR_COLORS[cr] || C.WHITE;
      const vals   = [cr,pl,s.count,s.spend,s.revenue,
                      roas.toFixed(2),cpm.toFixed(0),cpa.toFixed(0),
                      v2avg ? (v2avg*100).toFixed(1)+'%' : '-',
                      v6avg ? (v6avg*100).toFixed(1)+'%' : '-',
                      s.pass,s.fail,s.wait,pct];
      vals.forEach((v,i) => {
        const cell = ws.getRange(r, i+1);
        cell.setValue(v).setFontFamily('Arial').setFontSize(9)
          .setVerticalAlignment('middle').setHorizontalAlignment(i<2?'left':'center')
          .setBackground(i===5 ? (roas>=SETTINGS.TARGET_ROAS ? C.PASS_GREEN : C.FAIL_RED) : bg);
        if (i===3||i===4) cell.setNumberFormat('#,##0.00');
        if (i===5) cell.setFontWeight('bold');
      });
      ws.setRowHeight(r, 22); r++;
    });
  });

  // Total row
  const allSpend   = rows.reduce((a,row)=>a+num(row,'spend_thb'),0);
  const allRevenue = rows.reduce((a,row)=>a+num(row,'revenue_thb'),0);
  const allROAS    = allSpend > 0 ? allRevenue/allSpend : 0;
  const allPass    = Object.values(summary).reduce((a,s)=>a+s.pass,0);
  const allFail    = Object.values(summary).reduce((a,s)=>a+s.fail,0);
  const allWait    = Object.values(summary).reduce((a,s)=>a+s.wait,0);
  const allPct     = rows.length > 0 ? (allPass/rows.length*100).toFixed(1)+'%' : '0%';
  const totalVals  = ['รวมทั้งหมด','Meta+TikTok',rows.length,
                      allSpend,allRevenue,allROAS.toFixed(2),'','','','',allPass,allFail,allWait,allPct];
  totalVals.forEach((v,i) => {
    ws.getRange(r,i+1).setValue(v).setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setBackground(C.LIGHT_PURPLE).setHorizontalAlignment(i<2?'left':'center')
      .setVerticalAlignment('middle');
    if (i===3||i===4) ws.getRange(r,i+1).setNumberFormat('#,##0.00');
  });
  ws.setRowHeight(r, 24); r+=2;

  // Section 2: By Product
  _sectionHeader(ws, r, '📌 KPI รายสินค้า / Product'); r++;
  const ph = ['Product','Spend (฿)','Revenue (฿)','ROAS','Purchases','Content Pieces'];
  ph.forEach((h,i) => {
    ws.getRange(r,i+1).setValue(h)
      .setBackground(C.ORANGE).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(r, 28); r++;

  const prodSorted = Object.entries(byProduct).sort((a,b)=>b[1].revenue-a[1].revenue);
  prodSorted.forEach(([prod, s]) => {
    const roas = s.spend > 0 ? s.revenue/s.spend : 0;
    const vals = [prod,s.spend,s.revenue,roas.toFixed(2),s.purchases,s.count];
    vals.forEach((v,i) => {
      ws.getRange(r,i+1).setValue(v).setFontFamily('Arial').setFontSize(9)
        .setBackground(C.LIGHT_ORANGE).setVerticalAlignment('middle')
        .setHorizontalAlignment(i===0?'left':'center');
      if (i===1||i===2) ws.getRange(r,i+1).setNumberFormat('#,##0.00');
    });
    ws.setRowHeight(r, 20); r++;
  });

  // Column widths
  const kpiWidths = [80,80,80,110,110,80,90,90,80,80,75,80,75,80];
  kpiWidths.forEach((w,i) => ws.setColumnWidth(i+1, w));
  ws.setFrozenRows(5);
}


// ═══════════════════════════════════════════════════════════════════
// 📊 DASHBOARD SHELL (placeholder before data)
// ═══════════════════════════════════════════════════════════════════
function _createDashboardShell(ss) {
  let ws = ss.getSheetByName(SN.DASH) || ss.insertSheet(SN.DASH);
  ws.clear(); ws.setTabColor('#E91E63');

  // Title
  const title = ws.getRange('A1');
  title.setValue('📊 HTW Content KPI Dashboard — Feb 2026')
    .setFontFamily('Arial').setFontSize(18).setFontWeight('bold')
    .setFontColor(C.DARK_PURPLE).setVerticalAlignment('middle');
  ws.getRange('A1:L1').merge().setBackground(C.LIGHT_PURPLE);
  ws.setRowHeight(1, 50);

  // Sub-title
  ws.getRange('A2').setValue('⚡ Import ข้อมูล CSV แล้วกด "สร้าง KPI & Charts" จากเมนู HTW Dashboard')
    .setFontFamily('Arial').setFontSize(10).setFontColor('#B71C1C').setFontStyle('italic');
  ws.getRange('A2:L2').merge().setBackground('#FFF3E0');
  ws.setRowHeight(2, 28);

  // Placeholder KPI cards area
  const cardLabels = ['💰 Total Spend','📈 Total Revenue','🎯 ROAS รวม','✅ ผ่านเกณฑ์','👤 Best Creator','📦 Best Product'];
  cardLabels.forEach((label, i) => {
    const col = (i*2) + 1;
    ws.getRange(4, col).setValue(label)
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
      .setFontColor(C.DARK_PURPLE).setHorizontalAlignment('center')
      .setBackground(C.DARK_PURPLE).setFontColor('#FFFFFF');
    ws.getRange(4, col, 1, 2).merge().setBackground(C.DARK_PURPLE);

    ws.getRange(5, col).setValue('รอข้อมูล...')
      .setFontFamily('Arial').setFontSize(18).setFontWeight('bold')
      .setFontColor(C.DARK_PURPLE).setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    ws.getRange(5, col, 3, 2).merge().setBackground(C.LIGHT_PURPLE);
    ws.setRowHeight(5, 60);
  });

  ws.setColumnWidth(1, 130);
  for (let c = 1; c <= 12; c++) if (c > 1) ws.setColumnWidth(c, 120);
}


// ═══════════════════════════════════════════════════════════════════
// 📊 BUILD FULL DASHBOARD (runs after KPI calc)
// ═══════════════════════════════════════════════════════════════════
function _buildDashboard(ss) {
  const ads = _findSheet(ss, SN.ADS);
  const kpi = _findSheet(ss, SN.KPI) || ss.insertSheet(SN.KPI);
  const ws  = _findSheet(ss, SN.DASH) || ss.insertSheet(SN.DASH);
  ws.clear(); ws.setTabColor('#E91E63');

  // Read data
  const data    = ads.getDataRange().getValues();
  let headerRow = -1, colMap = {};
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const row = data[r].map(v => String(v).toLowerCase().trim());
    if (row.includes('spend_thb')) { headerRow = r; row.forEach((h,i) => colMap[h]=i); break; }
  }
  const rows = data.slice(headerRow+1).filter(r => r[colMap['creator']] !== '');
  const num  = (row,col) => { const v=row[colMap[col]??-1]; return (v===''||v==null)?0:Number(v)||0; };
  const str  = (row,col) => String(row[colMap[col]??-1]||'').trim();

  // Aggregate
  let totalSpend=0, totalRev=0, totalPass=0, totalPieces=rows.length;
  const byCreator={}, byPlatform={}, byProduct={};
  rows.forEach(row => {
    totalSpend += num(row,'spend_thb');
    totalRev   += num(row,'revenue_thb');
    const roas  = num(row,'roas');
    if (roas >= SETTINGS.TARGET_ROAS) totalPass++;
    const cr = str(row,'creator'), pl = str(row,'platform'), prod = str(row,'product')||'Other';
    if (!byCreator[cr])  byCreator[cr]  = {spend:0,revenue:0,count:0,pass:0};
    if (!byPlatform[pl]) byPlatform[pl] = {spend:0,revenue:0,count:0};
    if (!byProduct[prod])byProduct[prod]= {spend:0,revenue:0,count:0};
    byCreator[cr].spend  += num(row,'spend_thb'); byCreator[cr].revenue  += num(row,'revenue_thb');
    byCreator[cr].count++;  if(roas>=SETTINGS.TARGET_ROAS) byCreator[cr].pass++;
    byPlatform[pl].spend += num(row,'spend_thb'); byPlatform[pl].revenue += num(row,'revenue_thb'); byPlatform[pl].count++;
    byProduct[prod].spend+= num(row,'spend_thb'); byProduct[prod].revenue+= num(row,'revenue_thb'); byProduct[prod].count++;
  });

  const totalROAS   = totalSpend>0 ? totalRev/totalSpend : 0;
  const passPct     = totalPieces>0 ? (totalPass/totalPieces*100).toFixed(1)+'%' : '0%';
  const bestCreator = Object.entries(byCreator).sort((a,b)=>b[1].revenue/b[1].spend-a[1].revenue/a[1].spend)[0]?.[0] || '-';
  const bestProduct = Object.entries(byProduct).sort((a,b)=>b[1].revenue-a[1].revenue)[0]?.[0] || '-';

  // ── TITLE ──
  ws.getRange('A1').setValue('📊 HTW Content KPI Dashboard')
    .setFontFamily('Arial').setFontSize(20).setFontWeight('bold').setFontColor(C.DARK_PURPLE);
  ws.getRange('A1:N1').merge().setBackground(C.LIGHT_PURPLE);
  ws.setRowHeight(1, 52);

  ws.getRange('A2').setValue(`📅 Feb 2026  |  Meta + TikTok  |  อัปเดต: ${new Date().toLocaleDateString('th-TH')}  |  Total Content Pieces: ${totalPieces}`)
    .setFontFamily('Arial').setFontSize(9).setFontColor(C.DARK_GRAY).setFontStyle('italic');
  ws.getRange('A2:N2').merge().setBackground(C.MID_GRAY);
  ws.setRowHeight(2, 20);

  // ── KPI CARDS (row 4-7) ──
  const cards = [
    {label:'💰 Total Spend',   value:`฿${totalSpend.toLocaleString('th-TH',{minimumFractionDigits:0,maximumFractionDigits:0})}`,       sub:'ยอดรวมค่าโฆษณา', bg:C.LIGHT_BLUE,    hdr:C.DARK_BLUE},
    {label:'📈 Total Revenue', value:`฿${totalRev.toLocaleString('th-TH',{minimumFractionDigits:0,maximumFractionDigits:0})}`,         sub:'รายได้รวม',       bg:C.LIGHT_GREEN,   hdr:C.DARK_GREEN},
    {label:'🎯 ROAS รวม',      value:totalROAS.toFixed(2)+'x',   sub:`เกณฑ์ ≥ ${SETTINGS.TARGET_ROAS}x`, bg:totalROAS>=SETTINGS.TARGET_ROAS?C.PASS_GREEN:C.FAIL_RED, hdr:C.DARK_PURPLE},
    {label:'✅ ผ่านเกณฑ์',     value:passPct,                    sub:`${totalPass} / ${totalPieces} pieces`, bg:C.PASS_GREEN,  hdr:C.DARK_GREEN},
    {label:'🏆 Best Creator',  value:bestCreator,                sub:'ROAS สูงสุด',  bg:C.RIW_COLOR,     hdr:C.DARK_PURPLE},
    {label:'📚 Best Product',  value:bestProduct.length>20?bestProduct.substring(0,18)+'..':bestProduct, sub:'Revenue สูงสุด', bg:C.LIGHT_ORANGE, hdr:C.ORANGE},
  ];

  ws.setRowHeight(4, 28); ws.setRowHeight(5, 52); ws.setRowHeight(6, 20); ws.setRowHeight(7, 10);

  cards.forEach((card,i) => {
    const col = (i*2)+1;
    // Header
    ws.getRange(4,col,1,2).merge().setValue(card.label)
      .setBackground(card.hdr).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Value
    ws.getRange(5,col,1,2).merge().setValue(card.value)
      .setBackground(card.bg).setFontColor(card.hdr)
      .setFontFamily('Arial').setFontSize(20).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Sub-label
    ws.getRange(6,col,1,2).merge().setValue(card.sub)
      .setBackground(card.bg).setFontColor(C.DARK_GRAY)
      .setFontFamily('Arial').setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // ── CREATOR TABLE (row 9) ──
  ws.setRowHeight(8, 16);
  _sectionHeader(ws, 9, '👤 KPI รายบุคคล (Creator Performance)');
  ws.getRange(9,1,1,14).merge();
  ws.setRowHeight(9, 28);

  const crHdrs = ['Creator','Spend (฿)','Revenue (฿)','ROAS','CPM avg','CPA avg','Content Pieces','✅ ผ่าน','❌ ไม่ผ่าน','% ผ่าน'];
  crHdrs.forEach((h,i) => {
    ws.getRange(10,i+1).setValue(h)
      .setBackground(C.DARK_PURPLE).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(10, 28);

  const CR_COLORS = {ริว:C.RIW_COLOR, 'มิ้น':C.MIN_COLOR, หมิว:C.MIUW_COLOR};
  let r=11;
  ['ริว','มิ้น','หมิว'].forEach(cr => {
    const s = byCreator[cr]; if(!s) return;
    const roas  = s.spend>0 ? s.revenue/s.spend : 0;
    // Get CPM/CPA from KPI sheet data
    const bg    = CR_COLORS[cr]||C.WHITE;
    const pct   = s.count>0?(s.pass/s.count*100).toFixed(1)+'%':'0%';
    const fail  = s.count - s.pass - (s.count-s.pass-Math.max(0,s.count-s.pass));
    const vals  = [cr, s.spend, s.revenue, roas.toFixed(2)+'x', '-', '-', s.count, s.pass, s.count-s.pass, pct];
    vals.forEach((v,i) => {
      const cell = ws.getRange(r,i+1);
      cell.setValue(v).setFontFamily('Arial').setFontSize(9)
        .setBackground(i===3?(roas>=SETTINGS.TARGET_ROAS?C.PASS_GREEN:C.FAIL_RED):bg)
        .setHorizontalAlignment(i===0?'left':'center').setVerticalAlignment('middle')
        .setFontWeight(i===0?'bold':'normal');
      if(i===1||i===2) cell.setNumberFormat('#,##0');
    });
    ws.setRowHeight(r, 24); r++;
  });

  // ── PLATFORM TABLE (row r+1) ──
  r+=2;
  _sectionHeader(ws, r, '📱 KPI รายแพลตฟอร์ม');
  ws.getRange(r,1,1,14).merge(); ws.setRowHeight(r,28); r++;
  const plHdrs = ['Platform','Spend (฿)','Revenue (฿)','ROAS','Content Pieces','% ของทั้งหมด'];
  plHdrs.forEach((h,i) => {
    ws.getRange(r,i+1).setValue(h)
      .setBackground(C.DARK_GREEN).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(r, 26); r++;
  const PL_BG = {Meta:C.LIGHT_GREEN, TikTok:C.LIGHT_BLUE};
  ['Meta','TikTok'].forEach(pl => {
    const s=byPlatform[pl]; if(!s) return;
    const roas=s.spend>0?s.revenue/s.spend:0;
    const pct =totalPieces>0?(s.count/totalPieces*100).toFixed(1)+'%':'0%';
    const vals=[pl,s.spend,s.revenue,roas.toFixed(2)+'x',s.count,pct];
    vals.forEach((v,i) => {
      const cell=ws.getRange(r,i+1);
      cell.setValue(v).setFontFamily('Arial').setFontSize(9)
        .setBackground(i===3?(roas>=SETTINGS.TARGET_ROAS?C.PASS_GREEN:C.FAIL_RED):PL_BG[pl]||C.WHITE)
        .setHorizontalAlignment(i===0?'left':'center').setVerticalAlignment('middle');
      if(i===1||i===2) cell.setNumberFormat('#,##0');
    });
    ws.setRowHeight(r, 22); r++;
  });

  // ── PRODUCT TABLE ──
  r+=2;
  _sectionHeader(ws, r, '📚 KPI รายสินค้า (Top Products by Revenue)');
  ws.getRange(r,1,1,14).merge(); ws.setRowHeight(r,28); r++;
  const prHdrs=['Product','Spend (฿)','Revenue (฿)','ROAS','Pieces'];
  prHdrs.forEach((h,i) => {
    ws.getRange(r,i+1).setValue(h)
      .setBackground(C.ORANGE).setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(r, 26); r++;
  Object.entries(byProduct).sort((a,b)=>b[1].revenue-a[1].revenue).slice(0,8).forEach(([prod,s],idx) => {
    const roas=s.spend>0?s.revenue/s.spend:0;
    const vals=[prod,s.spend,s.revenue,roas.toFixed(2)+'x',s.count];
    const bg=idx%2===0?C.LIGHT_ORANGE:'#FFFFFF';
    vals.forEach((v,i) => {
      const cell=ws.getRange(r,i+1);
      cell.setValue(v).setFontFamily('Arial').setFontSize(9)
        .setBackground(i===3?(roas>=SETTINGS.TARGET_ROAS?C.PASS_GREEN:C.FAIL_RED):bg)
        .setHorizontalAlignment(i===0?'left':'center').setVerticalAlignment('middle');
      if(i===1||i===2) cell.setNumberFormat('#,##0');
    });
    ws.setRowHeight(r, 20); r++;
  });

  // ── CHARTS ──
  // Charts are built separately via buildChartsOnly() to avoid timeout
  // _buildCharts(ss, ws, rows, colMap, byCreator, byPlatform, byProduct);

  // Column widths
  const dashWidths=[130,110,110,90,90,90,90,90,90,90,90,110,90,90];
  dashWidths.forEach((w,i) => ws.setColumnWidth(i+1, w));
  ws.setFrozenRows(2);
}


// ═══════════════════════════════════════════════════════════════════
// 📈 BUILD CHARTS
// ═══════════════════════════════════════════════════════════════════
function _buildCharts(ss, dashWs, rows, colMap, byCreator, byPlatform, byProduct) {
  // Remove existing charts
  dashWs.getCharts().forEach(c => dashWs.removeChart(c));

  const kpiWs = ss.getSheetByName(SN.KPI);
  if (!kpiWs) return;

  // ── Chart 1: ROAS by Creator (Bar) ──
  const creatorNames = Object.keys(byCreator);
  const roasVals     = creatorNames.map(cr => {
    const s=byCreator[cr]; return s.spend>0?parseFloat((s.revenue/s.spend).toFixed(2)):0;
  });

  // Build temp data range in KPI sheet for charts
  const chartDataRow = kpiWs.getLastRow() + 3;
  kpiWs.getRange(chartDataRow, 1).setValue('Creator');
  kpiWs.getRange(chartDataRow, 2).setValue('ROAS');
  kpiWs.getRange(chartDataRow, 3).setValue('Revenue (฿)');
  kpiWs.getRange(chartDataRow, 4).setValue('Spend (฿)');
  const crOrder = ['ริว','มิ้น','หมิว'];
  crOrder.forEach((cr,i) => {
    const s=byCreator[cr]||{spend:0,revenue:0};
    const roas=s.spend>0?parseFloat((s.revenue/s.spend).toFixed(2)):0;
    kpiWs.getRange(chartDataRow+1+i, 1).setValue(cr);
    kpiWs.getRange(chartDataRow+1+i, 2).setValue(roas);
    kpiWs.getRange(chartDataRow+1+i, 3).setValue(Math.round(s.revenue));
    kpiWs.getRange(chartDataRow+1+i, 4).setValue(Math.round(s.spend));
  });

  const roasRange  = kpiWs.getRange(chartDataRow, 1, crOrder.length+1, 2);
  const revRange   = kpiWs.getRange(chartDataRow, 1, crOrder.length+1, 4);

  // Chart 1: ROAS by Creator
  const chart1 = dashWs.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(roasRange)
    .setOption('title', '🎯 ROAS by Creator')
    .setOption('titleTextStyle', {fontSize:12, bold:true, color:'#2D1B69'})
    .setOption('hAxis', {title:'ROAS', minValue:0, baseline:{color:'#2D1B69',width:2}})
    .setOption('vAxis', {title:'Creator'})
    .setOption('colors', ['#6B3FA0'])
    .setOption('legend', {position:'none'})
    .setOption('backgroundColor', '#F8F4FF')
    .setPosition(9, 1, 5, 5)
    .setNumRows(5)
    .setNumColumns(6)
    .build();
  dashWs.insertChart(chart1);

  // Chart 2: Revenue by Platform (Pie)
  const platDataRow = chartDataRow + 5;
  kpiWs.getRange(platDataRow, 1).setValue('Platform');
  kpiWs.getRange(platDataRow, 2).setValue('Revenue');
  let pi = 1;
  ['Meta','TikTok'].forEach(pl => {
    const s=byPlatform[pl]; if(!s) return;
    kpiWs.getRange(platDataRow+pi, 1).setValue(pl);
    kpiWs.getRange(platDataRow+pi, 2).setValue(Math.round(s.revenue));
    pi++;
  });
  const platRange = kpiWs.getRange(platDataRow, 1, pi+1, 2);
  const chart2 = dashWs.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(platRange)
    .setOption('title', '📱 Revenue by Platform')
    .setOption('titleTextStyle', {fontSize:12, bold:true, color:'#1B5E20'})
    .setOption('colors', ['#1B5E20','#0D47A1'])
    .setOption('backgroundColor', '#E8F5E9')
    .setOption('legend', {position:'right'})
    .setOption('pieSliceText', 'percentage')
    .setPosition(9, 7, 5, 5)
    .setNumRows(5).setNumColumns(6)
    .build();
  dashWs.insertChart(chart2);

  // Chart 3: Revenue by Creator (Column)
  const chart3 = dashWs.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(kpiWs.getRange(chartDataRow, 1, crOrder.length+1, 1))
    .addRange(kpiWs.getRange(chartDataRow, 3, crOrder.length+1, 2))
    .setOption('title', '💰 Spend vs Revenue by Creator')
    .setOption('titleTextStyle', {fontSize:12, bold:true, color:'#1B5E20'})
    .setOption('vAxis', {title:'฿ บาท', format:'#,##0'})
    .setOption('hAxis', {title:'Creator'})
    .setOption('colors', ['#FF7043','#2E7D32'])
    .setOption('backgroundColor', '#F1F8E9')
    .setOption('isStacked', false)
    .setPosition(20, 1, 5, 5)
    .setNumRows(5).setNumColumns(12)
    .build();
  dashWs.insertChart(chart3);
}


// ═══════════════════════════════════════════════════════════════════
// 🎨 CONDITIONAL FORMATTING on Ads_Performance
// ═══════════════════════════════════════════════════════════════════
function _addConditionalFormatting(ws) {
  const lastRow = Math.max(ws.getLastRow(), 100);

  // Find ROAS column (col K = 11 typically)
  // Find hit_roas column (last column)
  // Apply green/red based on ROAS value
  const roasCol   = 11;
  const hitCol    = 19;
  const roasRange = ws.getRange(4, roasCol, lastRow-3, 1);
  const hitRange  = ws.getRange(4, hitCol,  lastRow-3, 1);

  // Clear existing
  ws.clearConditionalFormatRules();
  const rules = [];

  // ROAS ≥ 2.5 → green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(SETTINGS.TARGET_ROAS)
    .setBackground(C.PASS_GREEN).setRanges([roasRange]).build());

  // ROAS > 0 < 2.5 → red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.01, SETTINGS.TARGET_ROAS - 0.001)
    .setBackground(C.FAIL_RED).setRanges([roasRange]).build());

  // hit_roas text
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('ผ่าน')
    .setBackground(C.PASS_GREEN).setRanges([hitRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('ไม่ผ่าน')
    .setBackground(C.FAIL_RED).setRanges([hitRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('ไม่มียอดขาย')
    .setBackground(C.WAIT_YELLOW).setRanges([hitRange]).build());

  ws.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════════════
// 🔧 UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════
function _setCell(ws, row, col, value, opts={}) {
  const cell = ws.getRange(row, col);
  cell.setValue(value).setFontFamily('Arial');
  if (opts.bold)    cell.setFontWeight('bold');
  if (opts.italic)  cell.setFontStyle('italic');
  if (opts.size)    cell.setFontSize(opts.size);
  if (opts.color)   cell.setFontColor(opts.color);
  if (opts.bg)      cell.setBackground(opts.bg);
  if (opts.align)   cell.setHorizontalAlignment(opts.align);
  if (opts.wrap)    cell.setWrap(true);
  return cell;
}

function _sectionHeader(ws, row, text) {
  const cell = ws.getRange(row, 1);
  cell.setValue(text)
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground(C.DARK_PURPLE)
    .setVerticalAlignment('middle').setPaddingTop !== undefined
    ? cell.setPaddingTop(4) : null;
  ws.setRowHeight(row, 28);
}

function _cleanDefaultSheets(ss) {
  ['Sheet1','แผ่น1','Untitled'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) try { ss.deleteSheet(sh); } catch(e) {}
  });
}


// ═══════════════════════════════════════════════════════════════════
// 📊 BUILD CHARTS ONLY — Run แยกต่างหากเพื่อไม่ให้ Timeout
// ═══════════════════════════════════════════════════════════════════
function buildChartsOnly() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const ads = _findSheet(ss, SN.ADS);
  if (!ads || ads.getLastRow() < 2) {
    ui.alert('⚠️ ยังไม่มีข้อมูลใน Ads_Performance\nกรุณา Import CSV และรัน "สร้าง KPI & Charts" ก่อน');
    return;
  }
  try {
    const data = ads.getDataRange().getValues();
    let headerRow = -1, colMap = {};
    for (let r = 0; r < Math.min(5, data.length); r++) {
      const row = data[r].map(v => String(v).toLowerCase().trim());
      if (row.includes('spend_thb')) { headerRow = r; row.forEach((h,i) => colMap[h]=i); break; }
    }
    if (headerRow < 0) throw new Error('ไม่พบ Header row');
    const rows = data.slice(headerRow + 1).filter(r => r[colMap['creator']] !== '');
    const num  = (row,col) => { const v=row[colMap[col]??-1]; return (v===''||v==null)?0:Number(v)||0; };
    const str  = (row,col) => String(row[colMap[col]??-1]||'').trim();
    const byCreator={}, byPlatform={}, byProduct={};
    rows.forEach(row => {
      const cr=str(row,'creator'), pl=str(row,'platform'), prod=str(row,'product')||'Other';
      if(!byCreator[cr])  byCreator[cr]  = {spend:0,revenue:0};
      if(!byPlatform[pl]) byPlatform[pl] = {spend:0,revenue:0};
      if(!byProduct[prod])byProduct[prod]= {spend:0,revenue:0};
      byCreator[cr].spend  +=num(row,'spend_thb'); byCreator[cr].revenue  +=num(row,'revenue_thb');
      byPlatform[pl].spend +=num(row,'spend_thb'); byPlatform[pl].revenue +=num(row,'revenue_thb');
      byProduct[prod].spend+=num(row,'spend_thb'); byProduct[prod].revenue+=num(row,'revenue_thb');
    });
    const dashWs = _findSheet(ss, SN.DASH);
    if (!dashWs) throw new Error('ไม่พบ Sheet Dashboard — กรุณารัน buildKPIAndCharts ก่อน');
    _buildCharts(ss, dashWs, rows, colMap, byCreator, byPlatform, byProduct);
    ss.setActiveSheet(dashWs);
    ui.alert('✅ สร้างกราฟสำเร็จ! ดูผลลัพธ์ที่ Sheet 📊 Dashboard');
  } catch(e) {
    ui.alert('❌ Error: ' + e.message);
  }
}


// ═══════════════════════════════════════════════════════════════════
// 📌 MENU — จะขึ้นอัตโนมัติเมื่อเปิด Google Sheets
// ═══════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 HTW Dashboard')
    .addItem('🚀 ตั้งค่าครั้งแรก (Setup)', 'setupDashboard')
    .addSeparator()
    .addItem('📈 สร้าง KPI & Dashboard', 'buildKPIAndCharts')
    .addItem('📊 เพิ่มกราฟ (ทำหลัง KPI)', 'buildChartsOnly')
    .addItem('🔄 รีเฟรช Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('ℹ️ วิธีใช้', 'showHelp')
    .addToUi();
}

function showHelp() {
  const html = `
  <b>📊 HTW Content KPI Dashboard — วิธีใช้</b><br><br>
  <b>ครั้งแรก:</b><br>
  1. Run "ตั้งค่าครั้งแรก (Setup)" → สร้าง Sheet ทั้งหมด<br>
  2. ไปที่ Sheet "📥 Ads_Performance"<br>
  3. File → Import → Upload ไฟล์ CSV → Replace current sheet<br>
  4. กลับมากด "สร้าง KPI & Charts"<br><br>
  <b>ทุกครั้งที่มีข้อมูลใหม่:</b><br>
  1. เพิ่มข้อมูลใน Ads_Performance<br>
  2. กด "รีเฟรช Dashboard"<br><br>
  <b>Target ปัจจุบัน:</b><br>
  • ROAS ≥ 2.5 | CPA ≤ ฿150 | CTR ≥ 1.5%<br>
  (แก้ไขได้ที่ Sheet ⚙️ Config)
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(300),
    'ℹ️ วิธีใช้ HTW Dashboard'
  );
}
