function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  var file = page === 'fin' ? 'Fin' : 'Index';
  var template = HtmlService.createTemplateFromFile(file);
  template.appUrl = ScriptApp.getService().getUrl();
  var title = page === 'fin' ? 'Monthly budget' : 'Daily Expense Tracker';
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function budgetMonthParts_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { year: raw.getFullYear(), month: raw.getMonth() };
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    var parsed = new Date(Date.UTC(1899, 11, 30) + Math.round(raw) * 86400000);
    return { year: parsed.getUTCFullYear(), month: parsed.getUTCMonth() };
  }
  if (typeof raw === 'string' && raw.trim() !== '') {
    var text = raw.trim();
    var match = text.match(/^(\d{4})-(\d{1,2})/);
    if (match) return { year: Number(match[1]), month: Number(match[2]) - 1 };
    var testD = new Date(text);
    if (!isNaN(testD.getTime())) return { year: testD.getFullYear(), month: testD.getMonth() };
  }
  return null;
}

function readBudgetTable_(budgetSheet) {
  var rows = [];
  if (!budgetSheet || budgetSheet.getLastRow() < 2) return rows;
  var values = budgetSheet.getRange(2, 1, budgetSheet.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < values.length; i++) {
    var parts = budgetMonthParts_(values[i][0]);
    if (!parts) continue;
    rows.push({
      year: parts.year,
      month: parts.month,
      target: Number(parseFloat(values[i][1]) || 0),
      actual: Number(parseFloat(values[i][2]) || 0),
      extra: Number(parseFloat(values[i][3]) || 0)
    });
  }
  return rows;
}

function addExpense(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    const tz = ss.getSpreadsheetTimeZone() || 'GMT+8';
    const recordTimestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      String(formData.date || ''),
      String(formData.category || ''),
      parseFloat(formData.amount) || 0,
      String(formData.items || ''),
      formData.isSignificant ? true : false,
      recordTimestamp
    ]);

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).sort({ column: 6, ascending: false });
    }

    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function updateExpense(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { status: 'error', message: 'No records found' };

    const tz = ss.getSpreadsheetTimeZone() || 'GMT+8';
    const data = sheet.getRange(2, 6, lastRow - 1, 1).getValues();

    for (let i = 0; i < data.length; i++) {
      let cellVal = data[i][0];
      let formattedTime = cellVal;
      if (cellVal instanceof Date) {
        formattedTime = Utilities.formatDate(cellVal, tz, 'yyyy-MM-dd HH:mm:ss');
      }
      
      if (String(formattedTime) === String(formData.recordTime)) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, 1).setValue(String(formData.date || ''));
        sheet.getRange(rowIndex, 2).setValue(String(formData.category || ''));
        sheet.getRange(rowIndex, 3).setValue(parseFloat(formData.amount) || 0);
        sheet.getRange(rowIndex, 4).setValue(String(formData.items || ''));
        sheet.getRange(rowIndex, 5).setValue(formData.isSignificant ? true : false);
        
        return { status: 'success' };
      }
    }
    
    return { status: 'error', message: 'Record not found' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function getInitialDashboardData(reqYear, reqMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone() || 'GMT+8';
    const now = new Date();

    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const monthShortNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    // 1. Target Year & Month Resolution
    let targetYear = (reqYear !== undefined && reqYear !== null) ? Number(reqYear) : now.getFullYear();
    let targetMonth = (reqMonth !== undefined && reqMonth !== null) ? Number(reqMonth) : now.getMonth();
    let budgetSheet = ss.getSheetByName('Budget');
    let budgetRows = readBudgetTable_(budgetSheet);
    let budgetMatch = null;
    for (let b = 0; b < budgetRows.length; b++) {
      if (budgetRows[b].year === targetYear && budgetRows[b].month === targetMonth) {
        budgetMatch = budgetRows[b];
        break;
      }
    }
    // Target, actual, and extra come from that month's row on the Budget tab.
    let targetBudget = budgetMatch ? budgetMatch.target : 0;
    let actualExpense = budgetMatch ? budgetMatch.actual : 0;
    let extraExpense = budgetMatch ? budgetMatch.extra : 0;

    // 2. Setup Past 10 Days Mini-Card Totals (Includes Today: i = 0 to 9).
    // The first 5 stay on the top row; the older 5 render on the row below.
    const past10DaysMap = {};
    const past10DaysKeys = [];
    for (let i = 0; i <= 9; i++) {
      let d = new Date(now.getTime() - (i * 24 * 60 * 60 * 1000));
      let dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      past10DaysKeys.push(dateStr);
      let label = (i === 0) ? 'Today' : (i === 1 ? 'Yest' : (i + 'd ago'));
      past10DaysMap[dateStr] = {
        label: label,
        dateFormatted: Utilities.formatDate(d, tz, 'MM/dd'),
        total: 0
      };
    }

    // 3. Setup Last 5 Days Entries Targets (Includes Today: i = 0 to 4)
    const last5DaysTarget = [];
    for (let i = 0; i <= 4; i++) {
      let d = new Date(now.getTime() - (i * 24 * 60 * 60 * 1000));
      let dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      let displayLabel = (i === 0) ? 'Today' : (i === 1 ? 'Yesterday' : (i + ' Days Ago'));
      last5DaysTarget.push({
        dateStr: dateStr,
        header: displayLabel + ' (' + Utilities.formatDate(d, tz, 'MM/dd') + ')',
        order: i
      });
    }

    // 4. Setup 3 Months Summary Indices
    const m0 = new Date(targetYear, targetMonth, 1);
    const m1 = new Date(targetYear, targetMonth - 1, 1);
    const m2 = new Date(targetYear, targetMonth - 2, 1);

    const monthLabels = [
      monthShortNames[m0.getMonth()] + ' ' + m0.getFullYear(),
      monthShortNames[m1.getMonth()] + ' ' + m1.getFullYear(),
      monthShortNames[m2.getMonth()] + ' ' + m2.getFullYear()
    ];

    const categories = [
      "Food & Dining", "Entertainment", "Personal", "Shopping",
      "Family", "Transportation", "Drive", "Others"
    ];

    const category3MonthTotals = {};
    categories.forEach(function(cat) { category3MonthTotals[cat] = [0, 0, 0]; });

    const last5DaysRecords = [];
    const significantItems = [];
    const extraItems = [];

    // 5. Read Expense Sheet Data
    let expenseSheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    const lastRow = expenseSheet.getLastRow();

    if (lastRow > 1) {
      const data = expenseSheet.getRange(2, 1, lastRow - 1, 6).getValues();

      for (let r = 0; r < data.length; r++) {
        let row = data[r];
        let rawDate = row[0];
        let category = String(row[1] || '').trim();
        let amount = Number(parseFloat(row[2]) || 0);
        let itemDesc = String(row[3] || '').trim() || '-';
        let isSig = (row[4] === true || String(row[4]).toLowerCase() === 'true');
        let rawRecTime = row[5];

        let expYear = -1;
        let expMonth = -1;
        let dateStr = '';

        if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
          expYear = rawDate.getFullYear();
          expMonth = rawDate.getMonth();
          dateStr = Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd');
        } else if (typeof rawDate === 'string' && rawDate.trim() !== '') {
          dateStr = rawDate.trim().split('T')[0];
          let testD = new Date(rawDate);
          if (!isNaN(testD.getTime())) {
            expYear = testD.getFullYear();
            expMonth = testD.getMonth();
          }
        }

        let formattedRecTime = '';
        if (rawRecTime instanceof Date && !isNaN(rawRecTime.getTime())) {
          formattedRecTime = Utilities.formatDate(rawRecTime, tz, 'yyyy-MM-dd HH:mm:ss');
        } else {
          formattedRecTime = String(rawRecTime || '');
        }

        if (expYear !== -1 && expMonth !== -1) {
          let mIdx = -1;
          if (expYear === m0.getFullYear() && expMonth === m0.getMonth()) mIdx = 0;
          else if (expYear === m1.getFullYear() && expMonth === m1.getMonth()) mIdx = 1;
          else if (expYear === m2.getFullYear() && expMonth === m2.getMonth()) mIdx = 2;

          if (mIdx !== -1) {
            if (category.toLowerCase() !== 'extra') {
              let matched = false;
              for (let c = 0; c < categories.length; c++) {
                if (categories[c].toLowerCase() === category.toLowerCase()) {
                  category3MonthTotals[categories[c]][mIdx] += amount;
                  matched = true;
                  break;
                }
              }
              if (!matched && category !== '') {
                category3MonthTotals["Others"][mIdx] += amount;
              }
            } else {
              extraItems.push({
                date: String(dateStr),
                monthIdx: Number(mIdx),
                item: String(itemDesc),
                amount: Number(amount)
              });
            }

            if (isSig) {
              significantItems.push({
                date: String(dateStr),
                monthIdx: Number(mIdx),
                category: String(category),
                item: String(itemDesc),
                amount: Number(amount)
              });
            }
          }
        }

        if (past10DaysMap[dateStr] && category.toLowerCase() !== 'extra') {
          past10DaysMap[dateStr].total += amount;
        }

        for (let t = 0; t < last5DaysTarget.length; t++) {
          if (last5DaysTarget[t].dateStr === dateStr) {
            last5DaysRecords.push({
              groupOrder: Number(last5DaysTarget[t].order),
              groupHeader: String(last5DaysTarget[t].header),
              date: String(dateStr),
              category: String(category),
              amount: String(amount.toFixed(2)),
              items: String(itemDesc),
              isSignificant: isSig ? true : false,
              recordTime: String(formattedRecTime)
            });
            break;
          }
        }
      }
    }

    last5DaysRecords.sort(function(a, b) {
      if (a.groupOrder !== b.groupOrder) return a.groupOrder - b.groupOrder;
      return (b.recordTime > a.recordTime) ? 1 : -1;
    });

    significantItems.sort(function(a, b) { return (b.date > a.date) ? 1 : -1; });
    extraItems.sort(function(a, b) { return (b.date > a.date) ? 1 : -1; });

    const past10DaysSummary = [];
    for (let k = 0; k < past10DaysKeys.length; k++) {
      let key = past10DaysKeys[k];
      past10DaysSummary.push({
        label: String(past10DaysMap[key].label),
        date: String(past10DaysMap[key].dateFormatted),
        iso: String(key),
        amount: String(past10DaysMap[key].total.toFixed(2))
      });
    }

    const payload = {
      status: 'ok',
      currentYear: targetYear,
      currentMonth: targetMonth,
      spreadsheetUrl: ss.getUrl(),
      budgetStats: {
        targetBudget: String(targetBudget.toFixed(2)),
        actualExpense: String(actualExpense.toFixed(2)),
        extraExpense: String(extraExpense.toFixed(2)),
        budgetFound: budgetMatch ? true : false,
        remaining: String((targetBudget - actualExpense).toFixed(2)),
        monthName: String(monthNames[targetMonth] + ' ' + targetYear),
        past10Days: past10DaysSummary
      },
      templates: (function() {
        try {
          return readTemplates_(ss);
        } catch (templateErr) {
          return [1, 2, 3].map(function(slot) {
            return { slot: slot, name: 'Template ' + slot, items: [] };
          });
        }
      })(),
      last3DaysRecords: last5DaysRecords,
      summary3Month: {
        months: monthLabels,
        categories: categories,
        totals: category3MonthTotals,
        significantItems: significantItems,
        extraItems: extraItems
      }
    };

    return JSON.stringify(payload);
  } catch (err) {
    return JSON.stringify({ status: 'error', error: err.toString() });
  }
}

var TEMPLATE_HEADERS = ['Slot', 'Name', 'ItemId', 'Category', 'Amount', 'Description', 'Significant'];

function defaultTemplateName_(slot) {
  return 'Template ' + slot;
}

function templateSlot_(value) {
  const slot = Number(value);
  if (slot !== 1 && slot !== 2 && slot !== 3) return 0;
  return slot;
}

function getOrCreateTemplatesSheet_(ss) {
  let sheet = ss.getSheetByName('Templates');
  if (!sheet) {
    try {
      sheet = ss.insertSheet('Templates', ss.getSheets().length);
    } catch (err) {
      sheet = ss.getSheetByName('Templates');
      if (!sheet) throw err;
    }
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, TEMPLATE_HEADERS.length).setValues([TEMPLATE_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readTemplates_(ss) {
  const sheet = getOrCreateTemplatesSheet_(ss);
  const templates = [1, 2, 3].map(function(slot) {
    return { slot: slot, name: defaultTemplateName_(slot), items: [] };
  });
  const bySlot = { 1: templates[0], 2: templates[1], 3: templates[2] };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return templates;

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (let i = 0; i < rows.length; i++) {
    const slot = templateSlot_(rows[i][0]);
    if (!slot) continue;
    const bucket = bySlot[slot];
    const name = String(rows[i][1] || '').trim();
    if (name) bucket.name = name;
    const itemId = String(rows[i][2] || '').trim();
    if (!itemId) continue;
    bucket.items.push({
      id: itemId,
      category: String(rows[i][3] || ''),
      amount: Number(parseFloat(rows[i][4]) || 0).toFixed(2),
      items: String(rows[i][5] || ''),
      isSignificant: rows[i][6] === true || String(rows[i][6]).toLowerCase() === 'true'
    });
  }
  return templates;
}

function templateNameForSlot_(sheet, slot) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return defaultTemplateName_(slot);
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (templateSlot_(rows[i][0]) === slot && String(rows[i][1] || '').trim() !== '') {
      return String(rows[i][1]).trim();
    }
  }
  return defaultTemplateName_(slot);
}

function slotHasRow_(sheet, slot) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const rows = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (templateSlot_(rows[i][0]) === slot) return true;
  }
  return false;
}

function ensureTemplatePlaceholder_(sheet, slot, name) {
  if (slotHasRow_(sheet, slot)) return;
  sheet.appendRow([slot, name, '', '', '', '', false]);
}

function renameTemplate(formData) {
  try {
    const slot = templateSlot_(formData && formData.slot);
    const name = String((formData && formData.name) || '').trim().slice(0, 40);
    if (!slot) return { status: 'error', message: 'Pick one of the 3 templates.' };
    if (!name) return { status: 'error', message: 'Enter a template name.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateTemplatesSheet_(ss);
    const lastRow = sheet.getLastRow();
    let updated = 0;
    if (lastRow >= 2) {
      const slots = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < slots.length; i++) {
        if (templateSlot_(slots[i][0]) === slot) {
          sheet.getRange(i + 2, 2).setValue(name);
          updated++;
        }
      }
    }
    if (updated === 0) ensureTemplatePlaceholder_(sheet, slot, name);
    return { status: 'success', templates: readTemplates_(ss) };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function addTemplateExpense(formData) {
  try {
    const slot = templateSlot_(formData && formData.slot);
    if (!slot) return { status: 'error', message: 'Pick one of the 3 templates.' };
    const category = String((formData && formData.category) || '').trim().slice(0, 80);
    const amount = parseFloat(formData && formData.amount);
    const description = String((formData && formData.items) || '').trim().slice(0, 500);
    if (!category) return { status: 'error', message: 'Select a category.' };
    if (!(amount > 0)) return { status: 'error', message: 'Enter an amount greater than $0.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateTemplatesSheet_(ss);
    const name = templateNameForSlot_(sheet, slot);
    sheet.appendRow([
      slot,
      name,
      Utilities.getUuid(),
      category,
      amount,
      description,
      formData.isSignificant ? true : false
    ]);
    return { status: 'success', templates: readTemplates_(ss) };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function removeTemplateExpense(formData) {
  try {
    const slot = templateSlot_(formData && formData.slot);
    const itemId = String((formData && formData.itemId) || '').trim();
    if (!slot || !itemId) return { status: 'error', message: 'Template expense not found.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateTemplatesSheet_(ss);
    const name = templateNameForSlot_(sheet, slot);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'Template expense not found.' };

    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    let removed = false;
    for (let i = rows.length - 1; i >= 0; i--) {
      if (templateSlot_(rows[i][0]) === slot && String(rows[i][2] || '').trim() === itemId) {
        sheet.deleteRow(i + 2);
        removed = true;
        break;
      }
    }
    if (!removed) return { status: 'error', message: 'Template expense not found.' };
    ensureTemplatePlaceholder_(sheet, slot, name);
    return { status: 'success', templates: readTemplates_(ss) };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function applyTemplate(formData) {
  try {
    const slot = templateSlot_(formData && formData.slot);
    const date = String((formData && formData.date) || '').trim();
    if (!slot) return { status: 'error', message: 'Pick one of the 3 templates.' };
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { status: 'error', message: 'Choose a valid date.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templates = readTemplates_(ss);
    let selected = null;
    for (let i = 0; i < templates.length; i++) {
      if (templates[i].slot === slot) selected = templates[i];
    }
    if (!selected || !selected.items.length) {
      return { status: 'error', message: 'This template has no expenses yet.' };
    }

    const tz = ss.getSpreadsheetTimeZone() || 'GMT+8';
    const sheet = ss.getSheetByName('Expenses') || ss.getSheets()[0];
    const base = new Date();
    const rows = [];
    for (let i = 0; i < selected.items.length; i++) {
      const item = selected.items[i];
      const stamp = Utilities.formatDate(new Date(base.getTime() + i), tz, 'yyyy-MM-dd HH:mm:ss.SSS');
      rows.push([
        date,
        String(item.category || ''),
        parseFloat(item.amount) || 0,
        String(item.items || ''),
        item.isSignificant ? true : false,
        stamp
      ]);
    }

    const startRow = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).sort({ column: 6, ascending: false });
    }

    return { status: 'success', added: rows.length, templates: templates };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function getFinDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Fin');
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 9) {
      return { status: 'error', message: 'The Fin sheet is missing. Add the Fin tab to this spreadsheet.' };
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0];
    const months = [];
    for (let c = 9; c < headers.length; c++) {
      const key = String(headers[c] || '').trim();
      if (/^\d{4}-\d{2}$/.test(key)) months.push(key);
    }
    const items = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row[0] && !row[5]) continue;
      const amounts = {};
      for (let c = 9; c < headers.length; c++) {
        const key = String(headers[c] || '').trim();
        if (!/^\d{4}-\d{2}$/.test(key)) continue;
        const raw = row[c];
        if (raw === '' || raw === null) continue;
        const num = Number(raw);
        if (!isNaN(num)) amounts[key] = num;
      }
      items.push({
        id: String(row[0]),
        section: String(row[1] || ''),
        group: String(row[2] || ''),
        fixed: String(row[3] || ''),
        category: String(row[4] || ''),
        item: String(row[5] || ''),
        note: String(row[6] || ''),
        source: String(row[7] || ''),
        calc: String(row[8] || ''),
        amounts: amounts
      });
    }
    return { status: 'ok', months: months, items: items, spreadsheetUrl: ss.getUrl() };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function saveFinAmount(payload) {
  try {
    const id = String((payload && payload.id) || '').trim();
    const month = String((payload && payload.month) || '').trim();
    if (!id || !/^\d{4}-\d{2}$/.test(month)) {
      return { status: 'error', message: 'Choose an item and a month.' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Fin');
    if (!sheet) return { status: 'error', message: 'The Fin sheet is missing.' };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let col = -1;
    for (let c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === month) { col = c + 1; break; }
    }
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 1).getValues();
    let row = -1;
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) { row = i + 2; break; }
    }
    if (row < 0 || col < 0) return { status: 'error', message: 'That cell is not on the Fin sheet.' };
    const calc = String(sheet.getRange(row, 9).getValue() || '');
    if (calc === 'mpf') return { status: 'error', message: 'MPF follows salary and is not typed in.' };
    const raw = payload.amount;
    if (raw === '' || raw === null || raw === undefined) {
      sheet.getRange(row, col).clearContent();
    } else {
      const num = Number(raw);
      if (isNaN(num)) return { status: 'error', message: 'Enter a number.' };
      sheet.getRange(row, col).setValue(num);
    }
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}