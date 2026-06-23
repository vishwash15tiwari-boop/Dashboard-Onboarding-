// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// Code.gs — Entry point, custom menu, orchestrator
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Platform')
    .addItem('🚀 Initialize / Reset Platform', 'runSetupPlatform')
    .addSeparator()
    .addItem('🔄 Refresh Formulas', 'refreshAll')
    .addItem('ℹ️  About', 'showAbout')
    .addToUi();
}

function runSetupPlatform() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Initialize Platform',
    'This will create / rebuild all sheets with headers, formulas, and formatting.\n\n' +
    'Existing sheets with the same names will be replaced.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  setupPlatform(ss);

  ui.alert(
    '✅ Platform Ready',
    'All sheets have been created:\n\n' +
    '• Control Tower — live KPI dashboard\n' +
    '• Settings — team, verticals, stage weights\n' +
    '• Harshita / Vamsi / Naveen / Vishwash — individual trackers\n\n' +
    'All TAT, progress %, and status columns are auto-calculated.'
  );
}

function refreshAll() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ All formulas refreshed.');
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    'Onboarding Operations Platform',
    'Version 1.0\n\n' +
    'Features:\n' +
    '• Auto-calculated TAT Days, TAT Status, % Done, Case Status\n' +
    '• Real-time Control Tower dashboard\n' +
    '• Team performance summary\n' +
    '• Pipeline view by vertical\n' +
    '• Alerts for breaches, pending docs, open tasks\n' +
    '• Monitoring, Third Party & Other Tasks tracking\n\n' +
    'TAT rule: TAT starts only when Doc % Collected = 100%.\n' +
    'If docs are incomplete, TAT is paused.'
  );
}

// ---- Main orchestrator ----
function setupPlatform(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Build Settings first (team sheets reference it for TAT target)
  setupSettings(ss);

  // Build individual team member sheets
  for (var i = 0; i < TEAM.length; i++) {
    setupTeamSheet(ss, TEAM[i]);
  }

  // Build dashboard last (references all team sheets)
  setupDashboard(ss);

  // Arrange sheet tabs in canonical order
  reorderSheets(ss);

  // Land on dashboard
  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.DASHBOARD));
}

function reorderSheets(ss) {
  var order = [
    SHEET_NAMES.DASHBOARD,
    SHEET_NAMES.SETTINGS
  ].concat(SHEET_NAMES.MEMBERS);

  for (var i = 0; i < order.length; i++) {
    var sheet = ss.getSheetByName(order[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
}

// ---- Helper: get or create a sheet (clears content if exists) ----
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ---- Helper: set border on a range ----
function setBorder(range, top, left, bottom, right, vertical, horizontal, color, style) {
  range.setBorder(top, left, bottom, right, vertical, horizontal,
    color || CLR.GRAY_MED,
    style || SpreadsheetApp.BorderStyle.SOLID);
}

// ---- Helper: apply dropdown validation ----
function addDropdown(sheet, row, col, numRows, numCols, values) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .build();
  sheet.getRange(row, col, numRows, numCols).setDataValidation(rule);
}

// ---- Helper: set number format for percentage ----
function setPercentFormat(range) {
  range.setNumberFormat('0%');
}

// ---- Helper: set date format ----
function setDateFormat(range) {
  range.setNumberFormat('dd-mmm-yyyy');
}
