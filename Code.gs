// ============================================================
// דשבורד תוכניות עבודה יקל"ר — Google Apps Script API
// ============================================================
// Paste this into: Google Sheet → Extensions → Apps Script
// Deploy as: Web app → Execute as: Me → Access: Anyone
// ============================================================

const PLANS_SHEET = 'plans';
const CONFIG_SHEET = 'config';

// ─── Read secret from config sheet (after 3rd --- separator) ───
function getSecret(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return '';
  const data = ws.getDataRange().getValues();
  let separatorCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === '---') { separatorCount++; continue; }
    if (separatorCount >= 3) {
      if (String(data[i][0] || '').trim() === name) {
        return String(data[i][1] || '').trim();
      }
    }
  }
  return '';
}

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
      case 'validateDept':
        result = validateDeptKey(e.parameter.key || '');
        break;
      case 'validateGate':
        result = validateGateCode(e.parameter.code || '');
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
      case 'followup':
        result = submitFollowup(body);
        break;
      case 'validateAdmin':
        result = validateAdmin(body);
        break;
      case 'addEffort':
        result = addEffort(body);
        break;
      case 'updateEffort':
        result = updateEffort(body);
        break;
      case 'removeEffort':
        result = removeEffort(body);
        break;
      case 'reorderEfforts':
        result = reorderEfforts(body);
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
  const departments = [];

  let section = 'units';
  for (let i = 1; i < data.length; i++) {
    const col0 = String(data[i][0] || '').trim();
    const col1 = String(data[i][1] || '').trim();
    const col2 = String(data[i][2] || '').trim();

    if (col0 === '---') {
      if (section === 'units') section = 'efforts';
      else if (section === 'efforts') section = 'departments';
      else if (section === 'departments') break; // secrets section — stop
      continue;
    }

    if (section === 'units' && col0) {
      units.push({ name: col0 });
    }

    if (section === 'efforts' && col0) {
      efforts.push({ name: col0, desc: col1, icon: col2 });
    }

    if (section === 'departments' && col0) {
      departments.push({ name: col0 });
    }
  }

  return { units, efforts, departments };
}

// ─── Validate Gate Code ──────────────────────────────────
function validateGateCode(code) {
  if (!code) return { valid: false };
  const gateCode = getSecret('gate_code');
  return { valid: code === gateCode };
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

// ─── Validate Department Key ───────────────────────────────
function validateDeptKey(key) {
  if (!key) return { valid: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { valid: false };

  const data = ws.getDataRange().getValues();
  let separatorCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === '---') { separatorCount++; continue; }
    if (separatorCount >= 2) {
      const deptName = String(data[i][0] || '').trim();
      const deptKey = String(data[i][1] || '').trim();
      if (deptKey && deptKey === key) {
        return { valid: true, dept: deptName };
      }
    }
  }

  return { valid: false };
}

// ─── Get Available Weeks ───────────────────────────────────
function getWeeks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(PLANS_SHEET);
  if (!ws || ws.getLastRow() < 2) return { weeks: [] };

  const data = ws.getDataRange().getValues();

  // Build set of department names
  const cfgData = getConfig();
  const deptNames = new Set((cfgData.departments || []).map(d => d.name));

  const weekMap = {};

  for (let i = 1; i < data.length; i++) {
    const week = String(data[i][0] || '').trim();
    const unit = String(data[i][1] || '').trim();
    if (!week || !unit) continue;

    if (!weekMap[week]) weekMap[week] = { units: new Set(), depts: new Set() };
    if (deptNames.has(unit)) {
      weekMap[week].depts.add(unit);
    } else {
      weekMap[week].units.add(unit);
    }
  }

  const weeks = Object.keys(weekMap).map(w => ({
    label: w,
    unitCount: weekMap[w].units.size,
    deptCount: weekMap[w].depts.size
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
      timestamp: String(data[i][6] || '').trim(),
      status: String(data[i][7] || '').trim(),
      followupNote: String(data[i][8] || '').trim(),
      followupTs: String(data[i][9] || '').trim()
    });
  }

  return { plans };
}

// ─── Submit Plan ───────────────────────────────────────────
function submitPlan(body) {
  const { key, week, data: rows } = body;

  // Validate key — try unit first, then department
  let validation = validateKey(key);
  let unit;
  if (validation.valid) {
    unit = validation.unit;
  } else {
    validation = validateDeptKey(key);
    if (validation.valid) {
      unit = validation.dept;
    } else {
      return { error: 'קוד לא תקין' };
    }
  }
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
      newRows.push([week, unit, row.day, row.rowType, row.rowName, row.content.trim(), timestamp, '', '', '']);
    }
  }

  if (newRows.length > 0) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
  }

  return { success: true, unit, week, rowCount: newRows.length };
}

// ─── Submit Followup ───────────────────────────────────────
function submitFollowup(body) {
  const { key, week, data: items } = body;

  const validation = validateKey(key);
  if (!validation.valid) {
    return { error: 'קוד יחידה לא תקין' };
  }

  const unit = validation.unit;
  if (!week || !items || !Array.isArray(items)) {
    return { error: 'Missing week or data' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(PLANS_SHEET);
  if (!ws) return { error: 'Missing plans sheet' };

  const allData = ws.getDataRange().getValues();
  const followupTs = new Date().toISOString();
  let updated = 0;

  for (const item of items) {
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === week &&
          String(allData[i][1]).trim() === unit &&
          String(allData[i][2]).trim() === item.day &&
          String(allData[i][3]).trim() === item.rowType &&
          String(allData[i][4]).trim() === item.rowName) {
        // Write followup columns directly — H, I, J (columns 8, 9, 10)
        ws.getRange(i + 1, 8, 1, 3).setValues([[item.status || '', item.note || '', followupTs]]);
        updated++;
        break;
      }
    }
  }

  SpreadsheetApp.flush();
  return { success: true, unit, week, updated };
}

// ─── Validate Admin Password ───────────────────────────────
function validateAdmin(body) {
  const password = String(body.password || '').trim();
  const adminPw = getSecret('admin_password');
  return { valid: password === adminPw };
}

// ─── Add Effort ────────────────────────────────────────
function addEffort(body) {
  if (String(body.password || '') !== getSecret('admin_password')) return { error: 'Unauthorized' };

  const name = String(body.name || '').trim();
  const desc = String(body.desc || '').trim();
  const icon = String(body.icon || '📌').trim();
  if (!name) return { error: 'Missing effort name' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { error: 'Missing config sheet' };

  // Check for duplicate name
  const data = ws.getDataRange().getValues();
  let separatorCount = 0;
  let secondSepRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === '---') {
      separatorCount++;
      if (separatorCount === 2) { secondSepRow = i; break; }
      continue;
    }
    if (separatorCount === 1 && String(data[i][0]).trim() === name) {
      return { error: 'Effort already exists: ' + name };
    }
  }

  // Insert before the second --- (or append if no second separator)
  if (secondSepRow >= 0) {
    ws.insertRowBefore(secondSepRow + 1);
    ws.getRange(secondSepRow + 1, 1, 1, 3).setValues([[name, desc, icon]]);
  } else {
    ws.appendRow([name, desc, icon]);
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// ─── Update Effort ─────────────────────────────────────────
function updateEffort(body) {
  if (String(body.password || '') !== getSecret('admin_password')) return { error: 'Unauthorized' };

  const oldName = String(body.oldName || '').trim();
  const newName = String(body.newName || '').trim();
  const desc = String(body.desc || '').trim();
  const icon = String(body.icon || '📌').trim();
  if (!oldName || !newName) return { error: 'Missing effort name' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { error: 'Missing config sheet' };

  const data = ws.getDataRange().getValues();
  let found = false;
  let separatorCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === '---') { separatorCount++; if (separatorCount >= 2) break; continue; }
    if (separatorCount === 1 && String(data[i][0]).trim() === oldName) {
      ws.getRange(i + 1, 1, 1, 3).setValues([[newName, desc, icon]]);
      found = true;
      break;
    }
  }

  if (!found) return { error: 'Effort not found: ' + oldName };

  // If name changed, update all matching rows in plans sheet
  if (oldName !== newName) {
    const plansWs = ss.getSheetByName(PLANS_SHEET);
    if (plansWs && plansWs.getLastRow() >= 2) {
      const plansData = plansWs.getDataRange().getValues();
      for (let i = 1; i < plansData.length; i++) {
        if (String(plansData[i][4]).trim() === oldName) {
          plansWs.getRange(i + 1, 5).setValue(newName);
        }
      }
    }
  }

  SpreadsheetApp.flush();
  return { success: true };
}

// ─── Remove Effort ─────────────────────────────────────────
function removeEffort(body) {
  if (String(body.password || '') !== getSecret('admin_password')) return { error: 'Unauthorized' };

  const name = String(body.name || '').trim();
  if (!name) return { error: 'Missing effort name' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { error: 'Missing config sheet' };

  const data = ws.getDataRange().getValues();
  let separatorCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === '---') { separatorCount++; if (separatorCount >= 2) break; continue; }
    if (separatorCount === 1 && String(data[i][0]).trim() === name) {
      ws.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }

  return { error: 'Effort not found: ' + name };
}

// ─── Reorder Efforts ───────────────────────────────────────
function reorderEfforts(body) {
  if (String(body.password || '') !== getSecret('admin_password')) return { error: 'Unauthorized' };

  const order = body.order; // array of effort names in desired order
  if (!order || !Array.isArray(order)) return { error: 'Missing order array' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CONFIG_SHEET);
  if (!ws) return { error: 'Missing config sheet' };

  const data = ws.getDataRange().getValues();

  // Find the first separator row and collect effort rows (stop at second ---)
  let separatorRow = -1;
  const effortRows = []; // {name, desc, icon}
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === '---') {
      if (separatorRow < 0) { separatorRow = i; }
      else { break; } // stop at second separator
      continue;
    }
    if (separatorRow >= 0 && String(data[i][0]).trim()) {
      effortRows.push({
        name: String(data[i][0]).trim(),
        desc: String(data[i][1] || '').trim(),
        icon: String(data[i][2] || '').trim()
      });
    }
  }

  if (separatorRow < 0) return { error: 'Config format error: no separator row' };

  // Sort effortRows by the provided order
  const sorted = [];
  for (const name of order) {
    const found = effortRows.find(e => e.name === name);
    if (found) sorted.push(found);
  }
  // Append any efforts not in the order array (safety)
  for (const e of effortRows) {
    if (!sorted.find(s => s.name === e.name)) sorted.push(e);
  }

  // Find the second separator (departments) to know where efforts end
  let secondSeparator = -1;
  for (let i = separatorRow + 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === '---') { secondSeparator = i; break; }
  }

  // Clear effort rows and rewrite (preserve departments section)
  const firstEffortRow = separatorRow + 2; // 1-indexed
  const lastEffortRow = secondSeparator >= 0 ? secondSeparator : ws.getLastRow(); // stop before second ---
  if (lastEffortRow >= firstEffortRow) {
    ws.deleteRows(firstEffortRow, lastEffortRow - firstEffortRow);
  }

  // Insert sorted efforts before the departments separator
  const insertAt = separatorRow + 2; // 1-indexed, after first ---
  for (let i = sorted.length - 1; i >= 0; i--) {
    ws.insertRowBefore(insertAt);
    ws.getRange(insertAt, 1, 1, 3).setValues([[sorted[i].name, sorted[i].desc, sorted[i].icon]]);
  }

  SpreadsheetApp.flush();
  return { success: true };
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
  cfg.appendRow(['---', '', '']);
  cfg.appendRow(['אג"ם', 'agam1', '']);
  cfg.appendRow(['תקשוב', 'tikshov2', '']);
  cfg.appendRow(['רפואה', 'refua3', '']);
  cfg.appendRow(['תכנון', 'tichnun4', '']);
  cfg.appendRow(['מלכ"א', 'malka5', '']);
  cfg.appendRow(['משא"ן', 'mashan6', '']);
  cfg.appendRow(['מודיעין', 'modin7', '']);
  cfg.appendRow(['אוכלוסיה', 'ukhlusiya8', '']);
  cfg.appendRow(['---', '', '']);
  cfg.appendRow(['gate_code', 'iklar2026', '']);
  cfg.appendRow(['admin_password', 'admin2026', '']);

  // Plans sheet
  let plans = ss.getSheetByName(PLANS_SHEET);
  if (!plans) {
    plans = ss.insertSheet(PLANS_SHEET);
  }
  plans.clear();
  plans.appendRow(['week', 'unit', 'day', 'row_type', 'row_name', 'content', 'timestamp', 'status', 'followup_note', 'followup_ts']);

  SpreadsheetApp.flush();
  Logger.log('Setup complete!');
}
