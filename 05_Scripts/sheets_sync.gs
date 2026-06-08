/**
 * Google Apps Script — Auto Sync MK13 + ADSM44 → GitHub Dashboard
 * ─────────────────────────────────────────────────────────────────
 * วิธีติดตั้ง:
 *   1. เปิด Google Sheet → Extensions → Apps Script
 *   2. วาง code นี้ทั้งหมด (แทน code เดิม)
 *   3. กด Run → "setupTrigger" เพื่อตั้ง cron ทุกชั่วโมง
 *   4. อนุญาต permission ที่ Google ขอ
 */

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  SHEET_ID:      '1qYdwXuCHDHeHN6a8vU_RVFBYT5MuYq-8K5AMK3cP4DM',
  MK13_TAB:      'MK13',
  ADSM44_TAB:    'ADSM44',
  GITHUB_TOKEN:  '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',
  GITHUB_OWNER:  'OPD-Atiwat',
  GITHUB_REPO:   'htw-kpi-dashboard',
  GITHUB_FILE:   'index.html',
  GITHUB_BRANCH: 'main',
};
// ============================================================

const MONTH_LABEL = {
  '2026-01':'Jan 26','2026-02':'Feb 26','2026-03':'Mar 26',
  '2026-04':'Apr 26','2026-05':'May 26','2026-06':'Jun 26',
  '2026-07':'Jul 26','2026-08':'Aug 26','2026-09':'Sep 26',
  '2026-10':'Oct 26','2026-11':'Nov 26','2026-12':'Dec 26',
  '2025-01':'Jan 25','2025-02':'Feb 25','2025-03':'Mar 25',
  '2025-04':'Apr 25','2025-05':'May 25','2025-06':'Jun 25',
  '2025-07':'Jul 25','2025-08':'Aug 25','2025-09':'Sep 25',
  '2025-10':'Oct 25','2025-11':'Nov 25','2025-12':'Dec 25',
};


// ════════════════════════════════════════════════════════════
// MAIN — รันทุกชั่วโมงโดย trigger
// ════════════════════════════════════════════════════════════
function syncDashboard() {
  try {
    Logger.log('▶ เริ่ม syncDashboard ' + new Date().toLocaleString('th-TH'));

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    const opdData    = readMK13(ss);
    const adsm44Data = readADSM44(ss);

    Logger.log('MK13 rows: '    + opdData.data.length);
    Logger.log('MK13 channels: ' + opdData.channels.join(', '));
    Logger.log('ADSM44 months: ' + Object.keys(adsm44Data).join(', '));

    const { content, sha } = getGitHubFile();

    let updated = content;
    updated = replaceVar(updated, 'OPD_DAILY',     JSON.stringify(opdData));
    updated = replaceVar(updated, 'ADSM44_PCTADS', JSON.stringify(adsm44Data));

    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd');
    updated = updated.replace(/v\d{8}[a-z]*/g, 'v' + today + 'as');

    pushGitHubFile(updated, sha, 'auto: sync MK13+ADSM44 ' + today);
    Logger.log('✅ syncDashboard เสร็จสิ้น');

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    throw e;
  }
}


// ════════════════════════════════════════════════════════════
// READ MK13 — raw order rows → OPD_DAILY format
// ════════════════════════════════════════════════════════════
function readMK13(ss) {
  const sheet = ss.getSheetByName(CONFIG.MK13_TAB);
  if (!sheet) throw new Error('ไม่เจอ tab: ' + CONFIG.MK13_TAB);

  const data   = sheet.getDataRange().getValues();
  const header = data[0].map(function(h){ return String(h).trim(); });

  const dateCol   = findCol(header, ['วันที่', 'Date', 'date']);
  const chCol     = findCol(header, ['Sale Channel']);
  const methodCol = findCol(header, ['Sale Method']);
  const amtCol    = findCol(header, ['ราคาแยกรายการ']);
  const freeCol   = findCol(header, ['แถม?', 'แถม']);

  if (dateCol < 0) throw new Error('MK13: ไม่เจอ column วันที่');
  if (chCol   < 0) throw new Error('MK13: ไม่เจอ column Sale Channel');
  if (amtCol  < 0) throw new Error('MK13: ไม่เจอ column ราคาแยกรายการ');

  const byDate = {};
  const chSet  = {};

  for (var r = 1; r < data.length; r++) {
    const row = data[r];

    // ข้าม row แถม / voucher / discount
    if (freeCol >= 0 && String(row[freeCol]).toUpperCase() === 'TRUE') continue;

    const rawD = row[dateCol];
    if (!rawD) continue;
    const d = formatDate(rawD);
    if (!d) continue;

    const saleCh  = String(row[chCol]     || '').trim();
    const saleMt  = methodCol >= 0 ? String(row[methodCol] || '').trim() : '';
    const channel = mapMK13Channel(saleCh, saleMt);
    if (!channel) continue;

    const amt = parseFloat(row[amtCol]) || 0;
    if (amt <= 0) continue;

    chSet[channel] = true;
    if (!byDate[d]) byDate[d] = { d: d };
    byDate[d][channel] = (byDate[d][channel] || 0) + amt;
  }

  const channels = Object.keys(chSet).sort();
  const rows     = Object.values(byDate).sort(function(a, b){ return a.d.localeCompare(b.d); });

  return { channels: channels, data: rows };
}

// Sale Channel + Sale Method → ชื่อ channel ใน dashboard
function mapMK13Channel(ch, mt) {
  if (ch === 'TikTok') {
    if (mt === 'Affiliate')                         return 'TikTok Affi';
    if (mt === 'Live' || mt === 'TikTokLive')       return 'TikTok Live';
    return 'TikTok';   // TikTokShop
  }
  if (ch === 'Shopee') {
    if (mt === 'Live' || mt === 'ShopeeLive')       return 'Shopee Live';
    return 'Shopee';
  }
  if (ch === 'Facebook') {
    if (mt === 'Salepage' || mt === 'Shopify')      return 'Shopify';
    return 'Facebook';
  }
  if (ch === 'Instagram')                           return 'Instagram';
  if (ch === 'LINE'  || ch === 'Line')              return 'LINE';
  if (ch === 'YouTube')                             return 'YouTube';
  if (ch === 'หน้าร้าน')                           return 'หน้าร้าน';
  if (ch === 'Bookfair')                            return 'Bookfair';
  return null;
}


// ════════════════════════════════════════════════════════════
// READ ADSM44 → ADSM44_PCTADS format
// ════════════════════════════════════════════════════════════
function readADSM44(ss) {
  const sheet = ss.getSheetByName(CONFIG.ADSM44_TAB);
  if (!sheet) throw new Error('ไม่เจอ tab: ' + CONFIG.ADSM44_TAB);

  const data   = sheet.getDataRange().getValues();
  const header = data[0].map(function(h){ return String(h).trim(); });

  Logger.log('ADSM44 header: ' + header.filter(function(h){return h;}).join(' | '));

  const productCol  = findCol(header, ['Product','product','เล่ม','หนังสือ']);
  const monthCol    = findCol(header, ['Month','month','เดือน','Date','วันที่']);

  // Spend columns
  const ttAdsCol   = findCol(header, ['Ads Tiktok Ads']);
  const ttAffCol   = findCol(header, ['Ads Tiktok Aff']);
  const fbMsgCol   = findCol(header, ['Ads FB MSG']);
  const fbSpCol    = findCol(header, ['Ads FB Salepage']);
  const spAdsCol   = findCol(header, ['Ads Shopee']);
  const totalAdCol = findCol(header, ['Ads Cost','_spend','Total Ads']);

  // Revenue columns
  const saleTtCol    = findCol(header, ['Sale Tiktok']);
  const saleTtAffCol = findCol(header, ['Sale Tiktok Aff']);
  const saleFbMsgCol = findCol(header, ['Sale FB MSG']);
  const saleFbSpCol  = findCol(header, ['Sale FB Salepage']);
  const saleSpCol    = findCol(header, ['Sale Shopee']);
  const saleTotalCol = findCol(header, ['Sale Total']);

  if (productCol < 0) {
    Logger.log('⚠️ ADSM44: ไม่เจอ column Product — ข้าม ADSM44');
    return {};
  }

  const result = {};

  for (var r = 1; r < data.length; r++) {
    const row     = data[r];
    const product = String(row[productCol] || '').trim();
    if (!product) continue;

    // หา monthLabel
    var monthLabel = null;
    if (monthCol >= 0 && row[monthCol]) {
      monthLabel = parseMonthLabel(row[monthCol]);
    }
    if (!monthLabel) {
      // fallback: ใช้เดือนปัจจุบัน
      const key = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM');
      monthLabel = MONTH_LABEL[key] || null;
    }
    if (!monthLabel) continue;

    if (!result[monthLabel]) result[monthLabel] = {};

    const ttSpend  = getVal(row, ttAdsCol)  + getVal(row, ttAffCol);
    const fbSpend  = getVal(row, fbMsgCol)  + getVal(row, fbSpCol);
    const spSpend  = getVal(row, spAdsCol);
    const totSpend = totalAdCol >= 0 ? getVal(row, totalAdCol) : (ttSpend + fbSpend + spSpend);

    const ttRev   = getVal(row, saleTtCol)    + getVal(row, saleTtAffCol);
    const fbRev   = getVal(row, saleFbMsgCol) + getVal(row, saleFbSpCol);
    const spRev   = getVal(row, saleSpCol);
    const totRev  = saleTotalCol >= 0 ? getVal(row, saleTotalCol) : (ttRev + fbRev + spRev);

    const ttPct = totSpend > 0 ? ttSpend / totSpend : null;
    const fbPct = totSpend > 0 ? fbSpend / totSpend : null;
    const spPct = totSpend > 0 ? spSpend / totSpend : null;

    result[monthLabel][product] = {
      TikTok:    ttPct,
      Facebook:  fbPct,
      Shopee:    spPct,
      _tt_spend: ttSpend,
      _fb_spend: fbSpend,
      _sp_spend: spSpend,
      _spend:    totSpend,
      _tt_rev:   ttRev,
      _fb_rev:   fbRev,
      _sp_rev:   spRev,
      _rev:      totRev,
    };
  }

  return result;
}

function getVal(row, idx) {
  if (idx < 0 || idx >= row.length) return 0;
  return parseFloat(row[idx]) || 0;
}


// ════════════════════════════════════════════════════════════
// GITHUB HELPERS
// ════════════════════════════════════════════════════════════
function getGitHubFile() {
  const url = 'https://api.github.com/repos/' + CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO
            + '/contents/' + CONFIG.GITHUB_FILE + '?ref=' + CONFIG.GITHUB_BRANCH;

  const res = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'token ' + CONFIG.GITHUB_TOKEN,
      'Accept': 'application/vnd.github.v3+json',
    },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('GitHub GET failed: ' + res.getContentText().substring(0, 300));
  }

  const json    = JSON.parse(res.getContentText());
  const content = Utilities.newBlob(Utilities.base64Decode(json.content)).getDataAsString();
  return { content: content, sha: json.sha };
}

function pushGitHubFile(content, sha, message) {
  const url     = 'https://api.github.com/repos/' + CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO
                + '/contents/' + CONFIG.GITHUB_FILE;
  const encoded = Utilities.base64Encode(Utilities.newBlob(content).getBytes());

  const res = UrlFetchApp.fetch(url, {
    method: 'PUT',
    headers: {
      'Authorization': 'token ' + CONFIG.GITHUB_TOKEN,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify({
      message: message,
      content: encoded,
      sha:     sha,
      branch:  CONFIG.GITHUB_BRANCH,
    }),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200 && res.getResponseCode() !== 201) {
    throw new Error('GitHub PUT failed: ' + res.getContentText().substring(0, 300));
  }
  Logger.log('✅ GitHub push สำเร็จ: ' + message);
}


// ════════════════════════════════════════════════════════════
// REPLACE VARIABLE in index.html
// ════════════════════════════════════════════════════════════
function replaceVar(html, varName, newValue) {
  const patterns = [
    new RegExp('(const\\s+' + varName + '\\s*=\\s*)([\\[\\{][\\s\\S]*?[\\]\\}])(\\s*;)', 'g'),
    new RegExp('(var\\s+'   + varName + '\\s*=\\s*)([\\[\\{][\\s\\S]*?[\\]\\}])(\\s*;)', 'g'),
  ];

  for (var i = 0; i < patterns.length; i++) {
    var replaced = false;
    var result = html.replace(patterns[i], function(_, prefix, _old, suffix) {
      replaced = true;
      return prefix + newValue + suffix;
    });
    if (replaced) {
      Logger.log('✅ แทนที่ ' + varName + ' สำเร็จ');
      return result;
    }
  }

  Logger.log('⚠️ ไม่เจอ ' + varName + ' ใน index.html');
  return html;
}


// ════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════
function findCol(header, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    for (var j = 0; j < header.length; j++) {
      if (header[j] === candidates[i] || header[j].includes(candidates[i])) return j;
    }
  }
  return -1;
}

function formatDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  const s = String(raw).trim();
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m1) return m1[3] + '-' + m1[2].padStart(2,'0') + '-' + m1[1].padStart(2,'0');
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return s.substring(0, 10);
  return null;
}

function parseMonthLabel(raw) {
  if (!raw) return null;
  if (raw instanceof Date) {
    const key = Utilities.formatDate(raw, 'Asia/Bangkok', 'yyyy-MM');
    return MONTH_LABEL[key] || null;
  }
  const s = String(raw).trim();
  if (/^[A-Za-z]{3}\s+\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})/);
  if (m) return MONTH_LABEL[m[1] + '-' + m[2]] || null;
  const m2 = s.match(/^(\d{2})\/(\d{4})/);
  if (m2) return MONTH_LABEL[m2[2] + '-' + m2[1]] || null;
  return null;
}


// ════════════════════════════════════════════════════════════
// SETUP TRIGGER — รันทุก 1 ชั่วโมง
// ════════════════════════════════════════════════════════════
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncDashboard').timeBased().everyHours(1).create();
  Logger.log('✅ Trigger ตั้งค่าแล้ว: syncDashboard ทุก 1 ชั่วโมง');
}


// ════════════════════════════════════════════════════════════
// DEBUG FUNCTIONS
// ════════════════════════════════════════════════════════════
function debugSheetStructure() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  ss.getSheets().forEach(function(s) {
    const name = s.getName();
    const row1 = s.getRange(1, 1, 1, Math.min(s.getLastColumn(), 20)).getValues()[0];
    Logger.log('Tab: "' + name + '" | Columns: ' + row1.filter(function(c){return c;}).join(' | '));
  });
}

function debugADSM44Columns() {
  const ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet  = ss.getSheetByName(CONFIG.ADSM44_TAB);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('ADSM44 total columns: ' + sheet.getLastColumn());
  header.forEach(function(h, i) {
    if (h) Logger.log(i + ': ' + h);
  });
  const data = sheet.getRange(2, 1, 2, sheet.getLastColumn()).getValues();
  data.forEach(function(row, ri) {
    Logger.log('Row ' + (ri+2) + ': ' + row.slice(0,10).join(' | '));
  });
}

function testMK13Read() {
  const ss     = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const result = readMK13(ss);
  Logger.log('Channels: ' + result.channels.join(', '));
  Logger.log('Total dates: ' + result.data.length);
  if (result.data.length > 0) {
    Logger.log('First: ' + JSON.stringify(result.data[0]));
    Logger.log('Last:  ' + JSON.stringify(result.data[result.data.length - 1]));
  }
}
