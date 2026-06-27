// Hell Pizza Weighing Tracker — Google Apps Script
// Paste this entire file into Extensions → Apps Script, then Deploy → New deployment → Web app
// Execute as: Me | Who has access: Anyone

function doPost(e) {
  // Prevent concurrent executions stomping on each other
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var payload            = JSON.parse(e.postData.contents);
    var ss                 = SpreadsheetApp.getActiveSpreadsheet();

    // ── Hard reset: delete all staff sheets and blank the Stats sheet ──────────
    if (payload.reset) {
      var allSheets = ss.getSheets();
      for (var ri = 0; ri < allSheets.length; ri++) {
        if (allSheets[ri].getName() !== 'Stats' && ss.getSheets().length > 1) {
          ss.deleteSheet(allSheets[ri]);
        }
      }
      var statsSheet = ss.getSheetByName('Stats');
      if (!statsSheet) statsSheet = ss.insertSheet('Stats');
      writeStatsSheet(statsSheet, [], {}, [], {}, {}, 0, []);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var records            = payload.records || [];
    var allStaff           = (payload.staff  && payload.staff.length  > 0) ? payload.staff  : null;
    var allPizzas          = (payload.pizzas && payload.pizzas.length > 0) ? payload.pizzas : null;
    var allTimeStaffTotals   = payload.allTimeStaffTotals   || {};
    var allTimeTotal         = (typeof payload.allTimeTotal === 'number') ? payload.allTimeTotal : null;
    var allTimePerPizzaTotals = payload.allTimePerPizzaTotals || {};
    var monthlyBreakdown     = payload.monthlyBreakdown || [];

    // Group records by staff → pizza → [dates]
    var byStaff = {};
    var byPizza = {};

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      if (!byStaff[r.staff])          byStaff[r.staff] = {};
      if (!byStaff[r.staff][r.pizza]) byStaff[r.staff][r.pizza] = [];
      byStaff[r.staff][r.pizza].push(r.date);
      byPizza[r.pizza] = (byPizza[r.pizza] || 0) + 1;
    }

    var staffToWrite = allStaff || Object.keys(byStaff);

    // Write every staff sheet (including 0-record members)
    for (var s = 0; s < staffToWrite.length; s++) {
      var name = staffToWrite[s];
      try {
        var sheet = ss.getSheetByName(name);
        if (!sheet) sheet = ss.insertSheet(name);
        // Build merged display: all-time counts + current month's last-weighed dates
        var currentPizzas  = byStaff[name] || {};
        var allTimePizzas  = allTimePerPizzaTotals[name] || {};
        var displayPizzas  = {};
        var allPizzaNames  = {};
        for (var pn in allTimePizzas) allPizzaNames[pn] = true;
        for (var pn in currentPizzas) allPizzaNames[pn] = true;
        for (var pn in allPizzaNames) {
          var atCount   = allTimePizzas[pn] || 0;
          var curDates  = currentPizzas[pn] || [];
          var lastDate  = curDates.length ? curDates.slice().sort().reverse()[0] : null;
          if (atCount > 0) displayPizzas[pn] = { count: atCount, lastDate: lastDate };
        }
        writeStaffSheet(sheet, name, displayPizzas);
      } catch (sheetErr) {
        try {
          var dead = ss.getSheetByName(name);
          if (dead && ss.getSheets().length > 1) ss.deleteSheet(dead);
          var fresh = ss.insertSheet(name);
          writeStaffSheet(fresh, name, {});
        } catch (retryErr) {}
      }
    }

    // Remove default "Sheet1", "Sheet2" tabs
    var sheets = ss.getSheets();
    for (var j = 0; j < sheets.length; j++) {
      if (/^Sheet\d+$/.test(sheets[j].getName()) && ss.getSheets().length > 1) {
        ss.deleteSheet(sheets[j]);
      }
    }

    // Rename "Summary" → "Stats" if it exists
    var oldSummary = ss.getSheetByName('Summary');
    if (oldSummary) oldSummary.setName('Stats');

    // Write Stats sheet
    var statsSheet = ss.getSheetByName('Stats');
    if (!statsSheet) statsSheet = ss.insertSheet('Stats');
    writeStatsSheet(statsSheet, staffToWrite, byStaff, allPizzas || Object.keys(byPizza), byPizza, allTimeStaffTotals, allTimeTotal, monthlyBreakdown);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ── Constants ──────────────────────────────────────────────────────────────

var RED_DARK   = '#B71C1C';
var GREY_DARK  = '#424242';
var GREY_MID   = '#757575';
var ROW_ALT    = '#E8DBC6'; // bone — alternating data rows (slightly darker)
var ROW_BASE   = '#EDE4D3'; // bone — base data rows
var BONE_TOTAL = '#D9CEBC'; // bone dark — total rows
var BORDER     = '#CCCCCC';

// ── Helpers ────────────────────────────────────────────────────────────────

function clearSheet(sheet) {
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear(); // clears content + formats in one call
}

function applyBorder(range) {
  range.setBorder(true, true, true, true, true, true,
    BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function sectionHeader(sheet, row, label, bg) {
  var r = sheet.getRange(row, 1, 1, 2);
  r.merge();
  r.setValue(label);
  r.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(bg);
  r.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 30);
}

function colHeader(sheet, row, values, bg) {
  var r = sheet.getRange(row, 1, 1, values.length);
  r.setValues([values]);
  r.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(bg);
  r.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 28);
}

function descRow(sheet, row, text) {
  var r = sheet.getRange(row, 1, 1, 2);
  r.merge();
  r.setValue(text);
  r.setFontSize(9).setFontColor(RED_DARK).setFontFamily('Roboto').setFontStyle('italic');
  r.setBackground(ROW_BASE).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setRowHeight(row, 20);
}

function dataRow(sheet, row, values, alt) {
  var r = sheet.getRange(row, 1, 1, values.length);
  r.setValues([values]);
  r.setBackground(alt ? ROW_ALT : ROW_BASE);
  r.setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center');
  if (values.length >= 2) sheet.getRange(row, 2).setNumberFormat('0');
  sheet.setRowHeight(row, 26);
  return r;
}

// Extracts "18th June" from any ISO date string — no Date() parsing, no NaN risk
function formatDate(dateStr) {
  if (!dateStr) return '—';
  var s = String(dateStr);
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return '—';
  var day = parseInt(m[3], 10);
  var mon = parseInt(m[2], 10) - 1;
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var mod = day % 100;
  var suffix = (mod >= 11 && mod <= 13) ? 'th' :
               (day % 10 === 1 ? 'st' : day % 10 === 2 ? 'nd' : day % 10 === 3 ? 'rd' : 'th');
  return day + suffix + ' ' + months[mon];
}

// ── Staff sheet ────────────────────────────────────────────────────────────

function writeStaffSheet(sheet, staffName, pizzaData) {
  clearSheet(sheet);

  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 160);

  // Flood sheet with bone so all gaps and uncovered rows are bone-coloured
  sheet.getRange(1, 1, sheet.getMaxRows(), 3).setBackground(ROW_BASE);

  // Title
  var title = sheet.getRange(1, 1, 1, 3);
  title.merge();
  title.setValue(staffName + '  —  Weighing Record');
  title.setFontSize(16).setFontWeight('bold').setFontColor('#FFFFFF').setFontFamily('Montserrat');
  title.setBackground(RED_DARK).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // Description
  var desc = sheet.getRange(2, 1, 1, 3);
  desc.merge();
  desc.setValue('All-time totals. Times Weighed = amount of pizzas weighed since tracking began. Pizzas not displayed are yet to be weighed.');
  desc.setFontSize(9).setFontColor(RED_DARK).setFontFamily('Roboto').setFontStyle('italic');
  desc.setBackground(ROW_BASE).setHorizontalAlignment('center').setVerticalAlignment('middle');
  desc.setWrap(true);
  sheet.setRowHeight(2, 36);

  colHeader(sheet, 3, ['Pizza', 'Times Weighed', 'Last Weighed'], GREY_DARK);
  sheet.getRange(3, 1, 1, 3).setFontFamily('Montserrat');

  // Build rows from { pizzaName: { count, lastDate } } format
  var pizzaKeys = Object.keys(pizzaData);
  var rows = [];
  for (var i = 0; i < pizzaKeys.length; i++) {
    var pizza = pizzaKeys[i];
    var entry = pizzaData[pizza];
    rows.push([pizza, entry.count, formatDate(entry.lastDate)]);
  }
  rows.sort(function(a, b) { return b[1] - a[1]; });

  if (rows.length === 0) {
    var empty = sheet.getRange(4, 1, 1, 3);
    empty.merge();
    empty.setValue('No pizzas weighed yet this month.');
    empty.setBackground(ROW_BASE).setFontColor('#999999').setFontFamily('Roboto');
    empty.setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(4, 40);
    applyBorder(sheet.getRange(2, 1, 3, 3));
    return;
  }

  var total = 0;
  for (var j = 0; j < rows.length; j++) {
    dataRow(sheet, j + 4, rows[j], j % 2 === 0).setFontWeight('bold').setFontFamily('Roboto');
    total += rows[j][1];
  }

  // Total row
  var totalRowNum = rows.length + 4;
  var tr = sheet.getRange(totalRowNum, 1, 1, 3);
  tr.setValues([['TOTAL', total, '']]);
  tr.setFontWeight('bold').setFontSize(12).setFontFamily('Montserrat').setBackground(BONE_TOTAL);
  tr.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(totalRowNum, 2).setNumberFormat('0');
  sheet.setRowHeight(totalRowNum, 32);

  applyBorder(sheet.getRange(2, 1, rows.length + 3, 3));
  sheet.setFrozenRows(3);
}

// ── Stats sheet ────────────────────────────────────────────────────────────

function writeStatsSheet(sheet, allStaff, byStaff, allPizzas, byPizza, allTimeStaffTotals, allTimeTotal, monthlyBreakdown) {
  clearSheet(sheet);
  var STATS_HDR = '#263238'; // Cool dark charcoal for section headers

  sheet.setColumnWidth(1, 270);
  sheet.setColumnWidth(2, 230);

  // Flood entire sheet with bone so spacer rows and any uncovered cells are bone-coloured
  sheet.getRange(1, 1, sheet.getMaxRows(), 2).setBackground(ROW_BASE);

  var row = 1;

  // Title
  var title = sheet.getRange(row, 1, 1, 2);
  title.merge();
  title.setValue('Store Pizza Weighing Stats');
  title.setFontSize(18).setFontWeight('bold').setFontColor('#FFFFFF');
  title.setFontFamily('Montserrat');
  title.setBackground(RED_DARK).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 56);
  row++;

  sheet.setRowHeight(row, 14); row++; // spacer

  // ── This Month ──────────────────────────────────────────────────────────────

  // Monthly staff totals (include 0-record staff)
  var staffTotals = {};
  for (var si = 0; si < allStaff.length; si++) staffTotals[allStaff[si]] = 0;
  var staffKeys = Object.keys(byStaff);
  for (var sk = 0; sk < staffKeys.length; sk++) {
    var sname = staffKeys[sk];
    var pizzas = Object.keys(byStaff[sname]);
    for (var pk = 0; pk < pizzas.length; pk++) {
      staffTotals[sname] = (staffTotals[sname] || 0) + byStaff[sname][pizzas[pk]].length;
    }
  }

  var totalPizzas = 0;
  var staffEntries = [];
  var staffTotalKeys = Object.keys(staffTotals);
  for (var ti = 0; ti < staffTotalKeys.length; ti++) {
    staffEntries.push([staffTotalKeys[ti], staffTotals[staffTotalKeys[ti]]]);
    totalPizzas += staffTotals[staffTotalKeys[ti]];
  }
  staffEntries.sort(function(a, b) { return b[1] - a[1]; });

  // Pizza totals (include 0-weighed pizzas)
  var pizzaTotals = {};
  for (var pi = 0; pi < allPizzas.length; pi++) pizzaTotals[allPizzas[pi]] = 0;
  var bpKeys = Object.keys(byPizza);
  for (var bp = 0; bp < bpKeys.length; bp++) pizzaTotals[bpKeys[bp]] = byPizza[bpKeys[bp]];

  var pizzaEntries = [];
  var ptKeys = Object.keys(pizzaTotals);
  for (var pti = 0; pti < ptKeys.length; pti++) pizzaEntries.push([ptKeys[pti], pizzaTotals[ptKeys[pti]]]);
  pizzaEntries.sort(function(a, b) { return b[1] - a[1]; });

  var weighedPizzas = [];
  for (var wp = 0; wp < pizzaEntries.length; wp++) {
    if (pizzaEntries[wp][1] > 0) weighedPizzas.push(pizzaEntries[wp]);
  }

  var mostPizza   = weighedPizzas.length ? weighedPizzas[0][0] + '  (' + weighedPizzas[0][1] + '×)' : '—';
  var leastPizza  = pizzaEntries.length  ? pizzaEntries[pizzaEntries.length - 1][0] + '  (' + pizzaEntries[pizzaEntries.length - 1][1] + '×)' : '—';
  var topStaff    = staffEntries.length  ? staffEntries[0][0] + '  (' + staffEntries[0][1] + ' pizzas)' : '—';
  var bottomStaff = staffEntries.length  ? staffEntries[staffEntries.length - 1][0] + '  (' + staffEntries[staffEntries.length - 1][1] + ' pizzas)' : '—';

  var now = new Date();
  var monthLabel = 'This Month  —  ' + now.toLocaleString('default', { month: 'long', year: 'numeric' });
  sectionHeader(sheet, row, monthLabel, STATS_HDR);
  sheet.getRange(row, 1, 1, 2).setFontFamily('Montserrat').setFontSize(12);
  row++;
  descRow(sheet, row, 'Quick live stats this month.'); row++;

  var overview = [
    ['Total pizzas weighed',         totalPizzas],
    ['Most commonly weighed pizza',  mostPizza],
    ['Least commonly weighed pizza', leastPizza],
    ['Top performer',                topStaff],
    ['Most improvement needed',      bottomStaff]
  ];
  var overviewStart = row;
  for (var oi = 0; oi < overview.length; oi++) {
    dataRow(sheet, row, overview[oi], oi % 2 === 0).setFontWeight('bold').setFontFamily('Roboto');
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    row++;
  }
  applyBorder(sheet.getRange(overviewStart - 1, 1, overview.length + 1, 2));

  sheet.setRowHeight(row, 16); row++; // spacer

  // ── All Time ─────────────────────────────────────────────────────────────────

  // Build all-time staff entries from payload (falls back to monthly if not sent)
  var atEntries = [];
  var atTotal = (allTimeTotal !== null) ? allTimeTotal : totalPizzas;
  if (allTimeStaffTotals && Object.keys(allTimeStaffTotals).length > 0) {
    // Include all known staff even if 0
    var atSeen = {};
    for (var si2 = 0; si2 < allStaff.length; si2++) {
      var n = allStaff[si2];
      atSeen[n] = true;
      atEntries.push([n, allTimeStaffTotals[n] || 0]);
    }
    var atKeys = Object.keys(allTimeStaffTotals);
    for (var ak = 0; ak < atKeys.length; ak++) {
      if (!atSeen[atKeys[ak]]) atEntries.push([atKeys[ak], allTimeStaffTotals[atKeys[ak]]]);
    }
    atEntries.sort(function(a, b) { return b[1] - a[1]; });
  } else {
    atEntries = staffEntries;
  }

  // All Time counter row
  sectionHeader(sheet, row, 'All Time', STATS_HDR);
  sheet.getRange(row, 1, 1, 2).setFontFamily('Montserrat').setFontSize(12);
  row++;
  descRow(sheet, row, 'Running store total since tracking began.'); row++;

  var atOverview = [
    ['Store total', atTotal]
  ];
  var atOverviewStart = row;
  for (var aoi = 0; aoi < atOverview.length; aoi++) {
    dataRow(sheet, row, atOverview[aoi], aoi % 2 === 0).setFontWeight('bold').setFontFamily('Roboto');
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    row++;
  }
  applyBorder(sheet.getRange(atOverviewStart - 1, 1, atOverview.length + 1, 2));

  sheet.setRowHeight(row, 16); row++; // spacer

  // Staff Breakdown — All Time
  sectionHeader(sheet, row, 'Staff Breakdown — All Time Totals', STATS_HDR);
  sheet.getRange(row, 1, 1, 2).setFontFamily('Montserrat').setFontSize(12);
  row++;
  descRow(sheet, row, 'Total pizzas weighed per staff member since tracking began.'); row++;

  colHeader(sheet, row, ['Staff Member', 'Total Pizzas Weighed'], '#37474F');
  sheet.getRange(row, 1, 1, 2).setFontFamily('Montserrat');
  row++;

  var breakdownStart = row;
  for (var bi = 0; bi < atEntries.length; bi++) {
    dataRow(sheet, row, atEntries[bi], bi % 2 === 0).setFontWeight('bold').setFontFamily('Roboto');
    row++;
  }

  var gt = sheet.getRange(row, 1, 1, 2);
  gt.setValues([['TOTAL', atTotal]]);
  gt.setFontWeight('bold').setFontFamily('Montserrat').setFontSize(12);
  gt.setBackground(BONE_TOTAL);
  gt.setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, 2).setNumberFormat('0');
  sheet.setRowHeight(row, 32);

  applyBorder(sheet.getRange(breakdownStart - 2, 1, atEntries.length + 3, 2));

  // ── Monthly Breakdown ────────────────────────────────────────────────────────
  if (monthlyBreakdown && monthlyBreakdown.length > 0) {
    sheet.setRowHeight(row, 16); row++; // spacer

    sectionHeader(sheet, row, 'Monthly Breakdown', STATS_HDR);
    sheet.getRange(row, 1, 1, 2).setFontFamily('Montserrat').setFontSize(12);
    row++;
    descRow(sheet, row, 'Month-by-month staff totals.'); row++;

    for (var mi = 0; mi < monthlyBreakdown.length; mi++) {
      var mEntry = monthlyBreakdown[mi];
      if (!mEntry.staff || mEntry.staff.length === 0) continue;

      // Month sub-header
      var mhr = sheet.getRange(row, 1, 1, 2);
      mhr.merge();
      mhr.setValue(mEntry.month);
      mhr.setFontWeight('bold').setFontFamily('Montserrat').setFontSize(11);
      mhr.setFontColor('#FFFFFF').setBackground('#455A64');
      mhr.setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.setRowHeight(row, 26);
      var monthBlockStart = row;
      row++;

      // Store total
      var str = sheet.getRange(row, 1, 1, 2);
      str.setValues([['Store total', mEntry.storeTotal]]);
      str.setFontWeight('bold').setFontFamily('Roboto').setFontSize(11);
      str.setBackground(ROW_ALT).setVerticalAlignment('middle').setHorizontalAlignment('center');
      sheet.getRange(row, 2).setNumberFormat('0');
      sheet.setRowHeight(row, 26);
      row++;

      // Staff rows ranked by count
      for (var msi = 0; msi < mEntry.staff.length; msi++) {
        var sr = mEntry.staff[msi];
        dataRow(sheet, row, [sr.name, sr.count], msi % 2 === 0).setFontWeight('bold').setFontFamily('Roboto');
        sheet.getRange(row, 2).setHorizontalAlignment('center');
        row++;
      }

      applyBorder(sheet.getRange(monthBlockStart, 1, mEntry.staff.length + 2, 2));
      sheet.setRowHeight(row, 8); row++; // small gap between months
    }
  }

  sheet.setFrozenRows(1);
}
