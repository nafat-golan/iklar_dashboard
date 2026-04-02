// ============================================================
// דשבורד תוכניות עבודה יקל"ר — Google Apps Script API
// ============================================================
// Paste this into: Google Sheet → Extensions → Apps Script
// Deploy as: Web app → Execute as: Me → Access: Anyone
// ============================================================

const PLANS_SHEET = 'plans';
const CONFIG_SHEET = 'config';

// ─── GET Handler ───────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;

  try {
    switch (action) {
      case 'config':
        result = getConfig();
        break;
      case 'weeks':
        result = getWeeks();
        break;
      case 'plans':
        result = getPlans(e.parameter.week || '');
        break;
      case 'validate':
        result = validateKey(e.parameter.key || '');
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── POST Handler ──────────────────────────────────────────
function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    switch (action) {
      case 'submit':
        result = submitPlan(body);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Config ────────────────────────────────────────────────
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { error: 'Missing config sheet' };

  const data = ws.getDataRange().getValues();
  const units = [];
  const efforts = [];

  // Row 0 = headers: unit | key | (ignored)
  // Rows 1+ = unit data
  // We also look for an "efforts" section
  let section = 'units';
  for (let i = 1; i < data.length; i++) {
    const col0 = String(data[i][0] || '').trim();
    const col1 = String(data[i][1] || '').trim();
    const col2 = String(data[i][2] || '').trim();

    if (col0 === '---') {
      section = 'efforts';
      continue;
    }

    if (section === 'units' && col0) {
      units.push({ name: col0 });
    }

    if (section === 'efforts' && col0) {
      efforts.push({ name: col0, desc: col1, icon: col2 });
    }
  }

  return { units, efforts };
}

// ─── Validate Key ──────────────────────────────────────────
function validateKey(key) {
  if (!key) return { valid: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { valid: false };

  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const unitName = String(data[i][0] || '').trim();
    const unitKey = String(data[i][1] || '').trim();
    if (unitKey && unitKey === key) {
      return { valid: true, unit: unitName };
    }
    if (String(data[i][0] || '').trim() === '---') break;
  }

  return { valid: false };
}

// ─── Get Available Weeks ───────────────────────────────────
function getWeeks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(PLANS_SHEET);
  if (!ws || ws.getLastRow() < 2) return { weeks: [] };

  const data = ws.getDataRange().getValues();
  const weekMap = {};

  for (let i = 1; i < data.length; i++) {
    const week = String(data[i][0] || '').trim();
    const unit = String(data[i][1] || '').trim();
    if (!week) continue;

    if (!weekMap[week]) weekMap[week] = new Set();
    weekMap[week].add(unit);
  }

  const weeks = Object.keys(weekMap).map(w => ({
    label: w,
    unitCount: weekMap[w].size
  }));

  // Sort by parsing first date
  weeks.sort((a, b) => {
    const pa = parseDateLabel(a.label);
    const pb = parseDateLabel(b.label);
    return pb - pa; // newest first
  });

  return { weeks };
}

function parseDateLabel(label) {
  // "22.3-28.3" → parse first part "22.3"
  const part = label.split('-')[0].trim();
  const bits = part.split('.');
  if (bits.length >= 2) {
    const day = parseInt(bits[0]) || 1;
    const month = parseInt(bits[1]) || 1;
    const year = new Date().getFullYear();
    return new Date(year, month - 1, day).getTime();
  }
  return 0;
}

// ─── Get Plans for a Week ──────────────────────────────────
function getPlans(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(PLANS_SHEET);
  if (!ws || ws.getLastRow() < 2) return { plans: [] };

  const data = ws.getDataRange().getValues();
  const plans = [];

  for (let i = 1; i < data.length; i++) {
    const w = String(data[i][0] || '').trim();
    if (week && w !== week) continue;

    plans.push({
      week: w,
      unit: String(data[i][1] || '').trim(),
      day: String(data[i][2] || '').trim(),
      rowType: String(data[i][3] || '').trim(),
      rowName: String(data[i][4] || '').trim(),
      content: String(data[i][5] || '').trim(),
      timestamp: String(data[i][6] || '').trim()
    });
  }

  return { plans };
}

// ─── Submit Plan ───────────────────────────────────────────
function submitPlan(body) {
  const { key, week, data: rows } = body;

  // Validate key
  const validation = validateKey(key);
  if (!validation.valid) {
    return { error: 'קוד יחידה לא תקין' };
  }

  const unit = validation.unit;
  if (!week || !rows || !Array.isArray(rows)) {
    return { error: 'Missing week or data' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(PLANS_SHEET);

  // Create sheet if missing
  if (!ws) {
    ws = ss.insertSheet(PLANS_SHEET);
    ws.appendRow(['week', 'unit', 'day', 'row_type', 'row_name', 'content', 'timestamp']);
  }

  // Delete existing rows for this unit+week
  const allData = ws.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][0]).trim() === week && String(allData[i][1]).trim() === unit) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }
  // Delete in reverse order to preserve indices
  for (const r of rowsToDelete) {
    ws.deleteRow(r);
  }

  // Append new rows
  const timestamp = new Date().toISOString();
  const newRows = [];
  for (const row of rows) {
    if (row.content && row.content.trim()) {
      newRows.push([week, unit, row.day, row.rowType, row.rowName, row.content.trim(), timestamp]);
    }
  }

  if (newRows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }

  return { success: true, unit, week, rowCount: newRows.length };
}

// ─── Initial Setup Helper ──────────────────────────────────
// Run this once to create the config + plans sheets
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config sheet
  let cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_SHEET);
  }
  cfg.clear();
  cfg.appendRow(['unit', 'key', 'icon']);
  cfg.appendRow(['עין קיניא', 'ein1', '']);
  cfg.appendRow(["מג'דל שמס", 'majdal2', '']);
  cfg.appendRow(['בוקעתא', 'bokta3', '']);
  cfg.appendRow(['מסעדה', 'masada4', '']);
  cfg.appendRow(['גולן', 'golan5', '']);
  cfg.appendRow(['קצרין', 'katzrin6', '']);
  cfg.appendRow(['מגדל', 'migdal7', '']);
  cfg.appendRow(['טבריה', 'tveria8', '']);
  cfg.appendRow(['עמק הירדן', 'emek9', '']);
  cfg.appendRow(['---', '', '']);
  cfg.appendRow(['חוסן', 'הרצאות מקצועיות לצוותים, סיוע אזרחי, ביקור קשישים, הכשרות ורענון לצוותים, פעילות מתנדבים ברשות', '🛡️']);
  cfg.appendRow(['הסברה', "מלשחיות, שליחת קיט הסברה לחנ\"מ, לשבת, לנסיעה ברכב וכו', סרטון ראש רשות/רב/שייח, עריכת הודעות נצורות, הסברה יעודית לאוכ' מוחלשות, הדברה דלת לדלת: D2D", '📣']);
  cfg.appendRow(['שוטף', 'צופרים, פעילות מס"ר, עדכון תיק מודיעין אוכלוסיה, הע"מ נפתי/ימי העמקה בנפה', '⚙️']);
  cfg.appendRow(['זמן יקר יקל"ר', 'אימון פנים יקל"ר', '🎯']);
  cfg.appendRow(['קשר עם הרשות', 'סיור רשות, הע"מ רשות, הדרכות מכלולים, סע"ר', '🤝']);
  cfg.appendRow(['תיאום בעלי תפקידים מהנפה', 'נמרוד, מירב, מיגון, קה"א נפה, חפ"ק אלפ"א, מפקד נפה, קנ"ר', '📞']);

  // Plans sheet
  let plans = ss.getSheetByName(PLANS_SHEET);
  if (!plans) {
    plans = ss.insertSheet(PLANS_SHEET);
  }
  plans.clear();
  plans.appendRow(['week', 'unit', 'day', 'row_type', 'row_name', 'content', 'timestamp']);

  SpreadsheetApp.flush();
  Logger.log('Setup complete!');
}
