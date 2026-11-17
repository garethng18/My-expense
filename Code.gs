function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Daily Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Allows embedding in other web apps
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
    let targetBudget = 0;
    
    let budgetSheet = ss.getSheetByName('Budget');
    if (budgetSheet && budgetSheet.getLastRow() >= 2) {
      try {
        if (reqYear === undefined || reqYear === null) {
          let rawMonthVal = budgetSheet.getRange('A2').getValue();
          if (rawMonthVal instanceof Date && !isNaN(rawMonthVal.getTime())) {
            targetYear = rawMonthVal.getFullYear();
            targetMonth = rawMonthVal.getMonth();
          } else if (typeof rawMonthVal === 'string' && rawMonthVal.trim() !== '') {
            let testD = new Date(rawMonthVal.trim());
            if (!isNaN(testD.getTime())) {
              targetYear = testD.getFullYear();
              targetMonth = testD.getMonth();
            }
          }
        }
        targetBudget = Number(parseFloat(budgetSheet.getRange('B2').getValue()) || 0);
      } catch (e) {
        // Fallback to defaults
      }
    }

    // 2. Setup Past 5 Days Mini-Card Totals (Includes Today: i = 0 to 4)
    const past5DaysMap = {};
    const past5DaysKeys = [];
    for (let i = 0; i <= 4; i++) {
      let d = new Date(now.getTime() - (i * 24 * 60 * 60 * 1000));
      let dateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      past5DaysKeys.push(dateStr);
      let label = (i === 0) ? 'Today' : (i === 1 ? 'Yest' : (i + 'd ago'));
      past5DaysMap[dateStr] = {
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

    let actualExpense = 0;
    let extraExpense = 0;
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
          if (expYear === targetYear && expMonth === targetMonth) {
            if (category.toLowerCase() === 'extra') {
              extraExpense += amount;
            } else {
              actualExpense += amount;
            }
          }

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

        if (past5DaysMap[dateStr] && category.toLowerCase() !== 'extra') {
          past5DaysMap[dateStr].total += amount;
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

    const past5DaysSummary = [];
    for (let k = 0; k < past5DaysKeys.length; k++) {
      let key = past5DaysKeys[k];
      past5DaysSummary.push({
        label: String(past5DaysMap[key].label),
        date: String(past5DaysMap[key].dateFormatted),
        amount: String(past5DaysMap[key].total.toFixed(2))
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
        remaining: String((targetBudget - actualExpense).toFixed(2)),
        monthName: String(monthNames[targetMonth] + ' ' + targetYear),
        past5Days: past5DaysSummary
      },
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