// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// TeamSheet.gs — Individual team member tracker sheet
//
// Sheet layout per member:
//   Row 1     : Name & title header
//   Row 2     : Role | Vertical | Fallback | 3rd Party
//   Row 3-4   : Mini KPI summary (OB / Monitoring / 3rd Party / Other)
//   Row 5     : (blank)
//   Row 6     : ▌ ONBOARDING CASES section header
//   Row 7     : Column headers
//   Rows 8-27 : Data (20 case rows)
//   Row 28-29 : (blank)
//   Row 30    : ▌ POST-ONBOARDING MONITORING section header
//   Row 31    : Column headers
//   Rows 32-46: Data (15 monitoring rows)
//   Row 47-48 : (blank)
//   Row 49    : ▌ THIRD PARTY HANDLING section header
//   Row 50    : Column headers
//   Rows 51-65: Data (15 third-party rows)
//   Row 66-67 : (blank)
//   Row 68    : ▌ OTHER TASKS / EXTRA section header
//   Row 69    : Column headers
//   Rows 70-89: Data (20 task rows)
// ============================================================

function setupTeamSheet(ss, member) {
  var sh = getOrCreateSheet(ss, member.name);

  // ── Column widths (mirroring original Excel widths in pixels) ─
  sh.setColumnWidth(1,  90);   // A  — Case ID / Seller / 3rd Party Name / Task
  sh.setColumnWidth(2,  165);  // B  — Vendor / Party Name / Case Ref / Description
  sh.setColumnWidth(3,  120);  // C  — Vertical / Case
  sh.setColumnWidth(4,  80);   // D  — Date Initiated / Data Coll Y/N / Assigned / Date Assigned
  sh.setColumnWidth(5,  65);   // E  — Doc % / Data Collect Date / Report Rcvd / Target Date
  sh.setColumnWidth(6,  65);   // F  — Doc Valid Y/N / Follow-up Y/N / Report Date / Actual Date
  sh.setColumnWidth(7,  85);   // G  — Doc Valid Date / Follow-up Date / Bill Rcvd / Days Open
  sh.setColumnWidth(8,  100);  // H  — Stage2 Y/N / Outcome Y/N / Bill Valid / Status
  sh.setColumnWidth(9,  80);   // I  — Stage2 Date / Present Date / Bill Date / Remarks
  sh.setColumnWidth(10, 100);  // J  — Stage3 Y/N / Status / Status
  sh.setColumnWidth(11, 80);   // K  — Stage3 Date / Remarks
  sh.setColumnWidth(12, 100);  // L  — Stage4 Y/N
  sh.setColumnWidth(13, 80);   // M  — Stage4 Date
  sh.setColumnWidth(14, 110);  // N  — Decision / MOM
  sh.setColumnWidth(15, 65);   // O  — MOM Y/N
  sh.setColumnWidth(16, 80);   // P  — MOM Date
  sh.setColumnWidth(17, 65);   // Q  — L1 Approval Y/N
  sh.setColumnWidth(18, 80);   // R  — L1 Approval Date
  sh.setColumnWidth(19, 145);  // S  — Doc Status (formula)
  sh.setColumnWidth(20, 60);   // T  — TAT Days (formula)
  sh.setColumnWidth(21, 110);  // U  — TAT Status (formula)
  sh.setColumnWidth(22, 60);   // V  — % Done (formula)
  sh.setColumnWidth(23, 110);  // W  — Case Status (formula)

  // ── Row heights ──────────────────────────────────────────
  sh.setRowHeight(1, 42);
  sh.setRowHeight(2, 32);
  sh.setRowHeight(3, 36);
  sh.setRowHeight(4, 22);
  sh.setRowHeight(6, 26);
  sh.setRowHeight(7, 52);
  for (var dr = 8; dr <= 27; dr++) sh.setRowHeight(dr, 22);
  sh.setRowHeight(30, 26);
  sh.setRowHeight(31, 52);
  for (var mr = 32; mr <= 46; mr++) sh.setRowHeight(mr, 22);
  sh.setRowHeight(49, 26);
  sh.setRowHeight(50, 52);
  for (var tr = 51; tr <= 65; tr++) sh.setRowHeight(tr, 22);
  sh.setRowHeight(68, 26);
  sh.setRowHeight(69, 52);
  for (var or = 70; or <= 89; or++) sh.setRowHeight(or, 22);

  // ────────────────────────────────────────────────────────
  // ROW 1 — Name header
  // ────────────────────────────────────────────────────────
  var r1 = sh.getRange('A1:W1');
  r1.merge();
  r1.setValue('  ' + member.name.toUpperCase() + '  —  Individual Activity Tracker');
  r1.setBackground(CLR.NAVY)
    .setFontColor(CLR.WHITE)
    .setFontSize(13)
    .setFontWeight('bold')
    .setVerticalAlignment('middle');

  // ROW 2 — Role / Vertical / Fallback / 3rd Party
  var r2 = sh.getRange('A2:W2');
  r2.merge();
  r2.setValue(
    '  Role: ' + member.role +
    '   |   Primary Vertical: ' + member.vertical +
    '   |   Fallback: ' + member.fallback +
    '   |   3rd Party: ' + member.thirdParty
  );
  r2.setBackground(CLR.NAVY_MED)
    .setFontColor(CLR.WHITE)
    .setFontSize(9)
    .setVerticalAlignment('middle');

  // ────────────────────────────────────────────────────────
  // ROWS 3–4 — Mini KPI summary bar
  // ────────────────────────────────────────────────────────
  _kpiBar(sh);

  // ────────────────────────────────────────────────────────
  // SECTION 1 — ONBOARDING CASES  (rows 6-27)
  // ────────────────────────────────────────────────────────
  _obSection(sh, member.name);

  // ────────────────────────────────────────────────────────
  // SECTION 2 — POST-ONBOARDING MONITORING  (rows 30-46)
  // ────────────────────────────────────────────────────────
  _monitoringSection(sh);

  // ────────────────────────────────────────────────────────
  // SECTION 3 — THIRD PARTY HANDLING  (rows 49-65)
  // ────────────────────────────────────────────────────────
  _thirdPartySection(sh);

  // ────────────────────────────────────────────────────────
  // SECTION 4 — OTHER TASKS / EXTRA  (rows 68-89)
  // ────────────────────────────────────────────────────────
  _otherTasksSection(sh);

  // ── Freeze rows 1-4 ──────────────────────────────────────
  sh.setFrozenRows(4);

  // ── Tab color (navy to match OB branding) ────────────────
  sh.setTabColor(CLR.NAVY);
}

// ============================================================
// MINI KPI BAR (Rows 3-4)
// ============================================================
function _kpiBar(sh) {
  // Helper to style a single KPI cell (label or number)
  function kpiCell(cell, bg, bold, fontSize) {
    cell.setBackground(bg)
        .setFontColor(CLR.WHITE)
        .setFontWeight(bold ? 'bold' : 'normal')
        .setFontSize(fontSize)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
  }

  // ── Row 3: label + Open count + Done count per section ──

  // ONBOARDING (navy)
  kpiCell(sh.getRange('A3').setValue('ONBOARDING'),  CLR.SECTION_OB, true,  9);
  kpiCell(sh.getRange('B3').setFormula('=COUNTIF(W8:W27,"In Progress")'), CLR.SECTION_OB, true, 13);
  kpiCell(sh.getRange('C3').setFormula('=COUNTIF(W8:W27,"Completed")+COUNTIF(W8:W27,"Rejected")'), CLR.SECTION_OB, true, 13);

  // Separator
  sh.getRange('D3').setValue('').setBackground(CLR.WHITE);

  // MONITORING (teal)
  kpiCell(sh.getRange('E3').setValue('MONITORING'), CLR.SECTION_MON, true, 9);
  kpiCell(sh.getRange('F3').setFormula('=COUNTIF(J32:J46,"Open")+COUNTIF(J32:J46,"In Progress")+COUNTIF(J32:J46,"Started")'), CLR.SECTION_MON, true, 13);
  kpiCell(sh.getRange('G3').setFormula('=COUNTIF(J32:J46,"Closed")'), CLR.SECTION_MON, true, 13);

  // Separator
  sh.getRange('H3').setValue('').setBackground(CLR.WHITE);

  // THIRD PARTY (purple) — J3 references J51:J65 (third party status column)
  kpiCell(sh.getRange('I3').setValue('THIRD PARTY'), CLR.SECTION_TP, true, 9);
  kpiCell(sh.getRange('J3').setFormula('=COUNTIF(J51:J65,"Open")+COUNTIF(J51:J65,"Report Received")+COUNTIF(J51:J65,"Bill Received")'), CLR.SECTION_TP, true, 13);
  kpiCell(sh.getRange('K3').setFormula('=COUNTIF(J51:J65,"Closed")'), CLR.SECTION_TP, true, 13);

  // Separator
  sh.getRange('L3').setValue('').setBackground(CLR.WHITE);

  // OTHER TASKS (deep orange)
  kpiCell(sh.getRange('M3').setValue('OTHER TASKS'), CLR.SECTION_OT, true, 9);
  kpiCell(sh.getRange('N3').setFormula('=COUNTIF(G70:G89,"Open")+COUNTIF(G70:G89,"Overdue")'), CLR.SECTION_OT, true, 13);
  kpiCell(sh.getRange('O3').setFormula('=COUNTIF(G70:G89,"Closed")'), CLR.SECTION_OT, true, 13);

  // Remainder of row 3
  sh.getRange('P3:W3').setBackground(CLR.WHITE);

  // ── Row 4: "Open" / "Done" sub-labels ────────────────────
  function subLabel(cell, bg, text) {
    cell.setValue(text).setBackground(bg).setFontColor(CLR.WHITE)
        .setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  sh.getRange('A4').setValue('').setBackground(CLR.SECTION_OB);
  subLabel(sh.getRange('B4'), CLR.SECTION_OB, 'Open');
  subLabel(sh.getRange('C4'), CLR.SECTION_OB, 'Done');
  sh.getRange('D4').setValue('').setBackground(CLR.WHITE);
  sh.getRange('E4').setValue('').setBackground(CLR.SECTION_MON);
  subLabel(sh.getRange('F4'), CLR.SECTION_MON, 'Open');
  subLabel(sh.getRange('G4'), CLR.SECTION_MON, 'Done');
  sh.getRange('H4').setValue('').setBackground(CLR.WHITE);
  sh.getRange('I4').setValue('').setBackground(CLR.SECTION_TP);
  subLabel(sh.getRange('J4'), CLR.SECTION_TP, 'Open');
  subLabel(sh.getRange('K4'), CLR.SECTION_TP, 'Done');
  sh.getRange('L4').setValue('').setBackground(CLR.WHITE);
  sh.getRange('M4').setValue('').setBackground(CLR.SECTION_OT);
  subLabel(sh.getRange('N4'), CLR.SECTION_OT, 'Open');
  subLabel(sh.getRange('O4'), CLR.SECTION_OT, 'Done');
  sh.getRange('P4:W4').setBackground(CLR.WHITE);
}

// ============================================================
// SECTION 1 — ONBOARDING CASES  (rows 6–27)
// ============================================================
function _obSection(sh, memberName) {
  // Row 6 — section header
  var hdr = sh.getRange('A6:W6');
  hdr.merge();
  hdr.setValue('  ▌ ONBOARDING CASES');
  hdr.setBackground(CLR.SECTION_OB).setFontColor(CLR.WHITE)
     .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // Row 7 — column headers
  var headers = [
    'Case ID',
    'Vendor / Party Name',
    'Vertical',
    'Date\nInitiated',
    'Doc %\nCollected\n(0–1)',
    'Doc\nValidation\n(Y/N)',
    'Doc Valid\nDate',
    'Verif Report (MP/EPR)\nDoc→3rd Party (OMP)\n(Y/N)',
    'Stage 2\nDate',
    'V-KYC & GST (MP/EPR)\nRating Rpt Rcvd (OMP)\n(Y/N)',
    'Stage 3\nDate',
    'Committee (MP/EPR)\nReview Rating Rpt (OMP)\n(Y/N)',
    'Stage 4\nDate',
    'Decision\n(MP/EPR only)',
    'MOM\n(MP/EPR)\n(Y/N)',
    'MOM\nDate',
    'L1\nApproval\n(Y/N)',
    'L1 Approval\nDate',
    'Doc Status\n(auto)',
    'TAT\nDays\n(auto)',
    'TAT Status\n(auto)',
    '% Done\n(auto)',
    'Case\nStatus\n(auto)'
  ];

  for (var c = 0; c < headers.length; c++) {
    var cell = sh.getRange(7, c + 1);
    cell.setValue(headers[c]);
    cell.setBackground(CLR.NAVY).setFontColor(CLR.WHITE)
        .setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  // Mark auto-calculated columns with lighter shade
  sh.getRange('S7:W7').setBackground(CLR.NAVY_MED);

  // Rows 8–27 — data rows with formulas
  for (var r = 8; r <= 27; r++) {
    var isAlt = (r % 2 === 0);
    var rowBg = isAlt ? CLR.OFF_WHITE : CLR.WHITE;

    // Style manual-entry columns A–R
    sh.getRange(r, 1, 1, 18).setBackground(rowBg).setFontSize(9)
      .setVerticalAlignment('middle');

    // Date columns: D, G, I, K, M, P, R
    setDateFormat(sh.getRange(r, 4));
    setDateFormat(sh.getRange(r, 7));
    setDateFormat(sh.getRange(r, 9));
    setDateFormat(sh.getRange(r, 11));
    setDateFormat(sh.getRange(r, 13));
    setDateFormat(sh.getRange(r, 16));
    setDateFormat(sh.getRange(r, 18));

    // E — Doc % Collected (0–1 → show as %)
    sh.getRange(r, 5).setNumberFormat('0%');

    // ── Auto-calculated columns (S, T, U, V, W) ──────────

    // S — Doc Status
    sh.getRange(r, 19).setFormula(
      '=IF(A' + r + '="","",IF(E' + r + '="","No docs yet",IF(E' + r + '<1,"⚠ Docs incomplete — TAT paused","✓ Docs complete")))'
    );

    // T — TAT Days (working days from Doc Valid Date to L1 Approval Date, only when complete)
    sh.getRange(r, 20).setFormula(
      '=IF(AND(E' + r + '>=1,G' + r + '<>"",R' + r + '<>""),NETWORKDAYS(G' + r + ',R' + r + ')-1,"")'
    );

    // U — TAT Status
    sh.getRange(r, 21).setFormula(
      '=IF(A' + r + '="","",IF(E' + r + '="","—",IF(E' + r + '<1,"Pending Docs",' +
      'IF(T' + r + '="","In Progress",IF(T' + r + '<=Settings!$C$38,"✓ Within TAT","⚠ Breach")))))'
    );

    // V — % Done (weighted stage completion via date columns)
    sh.getRange(r, 22).setFormula(
      '=IF(A' + r + '="","",IF(OR(C' + r + '="Marketplace",C' + r + '="EPR & Sustainability"),' +
      'IF(G' + r + '<>"",0.15,0)+IF(I' + r + '<>"",0.3,0)+IF(K' + r + '<>"",0.25,0)' +
      '+IF(M' + r + '<>"",0.2,0)+IF(P' + r + '<>"",0.05,0)+IF(R' + r + '<>"",0.05,0),' +
      'IF(C' + r + '="Open Marketplace",' +
      'IF(G' + r + '<>"",0.15,0)+IF(I' + r + '<>"",0.15,0)+IF(K' + r + '<>"",0.3,0)' +
      '+IF(M' + r + '<>"",0.3,0)+IF(R' + r + '<>"",0.1,0),"")))'
    );
    sh.getRange(r, 22).setNumberFormat('0%');

    // W — Case Status
    sh.getRange(r, 23).setFormula(
      '=IF(A' + r + '="","—",IF(Q' + r + '="Y","Completed",' +
      'IF(N' + r + '="Rejected","Rejected",' +
      'IF(E' + r + '<1,"Docs Pending",IF(D' + r + '<>"","In Progress","Not Started")))))'
    );

    // Style auto-calculated columns S–W
    sh.getRange(r, 19, 1, 5).setBackground(isAlt ? CLR.NAVY_LIGHT : '#f0f3ff')
      .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // Data validation — Y/N columns: F, H, J, L, O, Q
  addDropdown(sh, 8, 6,  20, 1, YN_OPTIONS);  // F — Doc Validation
  addDropdown(sh, 8, 8,  20, 1, YN_OPTIONS);  // H — Stage 2 Y/N
  addDropdown(sh, 8, 10, 20, 1, YN_OPTIONS);  // J — Stage 3 Y/N
  addDropdown(sh, 8, 12, 20, 1, YN_OPTIONS);  // L — Stage 4 Y/N
  addDropdown(sh, 8, 15, 20, 1, YN_OPTIONS);  // O — MOM Y/N
  addDropdown(sh, 8, 17, 20, 1, YN_OPTIONS);  // Q — L1 Approval Y/N
  // Vertical column C
  addDropdown(sh, 8, 3, 20, 1, VERTICAL_OPTIONS);

  // Borders
  setBorder(sh.getRange('A7:W27'), true, true, true, true, false, true);
  setBorder(sh.getRange('S7:W27'), true, true, true, true, false, true, CLR.NAVY_MED, SpreadsheetApp.BorderStyle.SOLID);

  // Conditional formatting — TAT Status column U
  _applyTatConditionalFormat(sh, 'U8:U27');
  // Conditional formatting — Case Status column W
  _applyCaseStatusFormat(sh, 'W8:W27');
}

// ============================================================
// SECTION 2 — POST-ONBOARDING MONITORING  (rows 30–46)
// ============================================================
function _monitoringSection(sh) {
  // Row 30 — section header
  var hdr = sh.getRange('A30:K30');
  hdr.merge();
  hdr.setValue('  ▌ POST-ONBOARDING MONITORING  (Monthly — per seller)');
  hdr.setBackground(CLR.SECTION_MON).setFontColor(CLR.WHITE)
     .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // Row 31 — column headers
  var mHeaders = [
    'Seller / Vendor',
    'Vertical',
    'Month\n(Period)',
    'Data\nCollection\n(Y/N)',
    'Data Collect\nDate',
    'Follow-up\nwith Teams\n(Y/N)',
    'Follow-up\nDate',
    'Outcome\nPresented\n(Y/N)',
    'Present\nDate',
    'Status\n(auto)',
    'Remarks'
  ];
  for (var c = 0; c < mHeaders.length; c++) {
    var cell = sh.getRange(31, c + 1);
    cell.setValue(mHeaders[c]);
    cell.setBackground(CLR.SECTION_MON).setFontColor(CLR.WHITE)
        .setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  // J31 auto-calculated
  sh.getRange('J31').setBackground('#00695c');

  // Rows 32–46
  for (var r = 32; r <= 46; r++) {
    var isAlt = (r % 2 === 0);
    var rowBg = isAlt ? CLR.OFF_WHITE : CLR.WHITE;
    sh.getRange(r, 1, 1, 11).setBackground(rowBg).setFontSize(9).setVerticalAlignment('middle');
    setDateFormat(sh.getRange(r, 5));  // E — Data Collect Date
    setDateFormat(sh.getRange(r, 7));  // G — Follow-up Date
    setDateFormat(sh.getRange(r, 9));  // I — Present Date

    // J — Status (auto)
    sh.getRange(r, 10).setFormula(
      '=IF(A' + r + '="","—",IF(I' + r + '<>"","Closed",' +
      'IF(H' + r + '="Y","In Progress",' +
      'IF(F' + r + '="Y","In Progress",' +
      'IF(D' + r + '="Y","Started","Open")))))'
    );
    sh.getRange(r, 10).setBackground(isAlt ? CLR.TEAL_LIGHT : '#e8f5f2')
      .setFontSize(9).setHorizontalAlignment('center');
  }

  addDropdown(sh, 32, 4, 15, 1, YN_OPTIONS);  // D — Data Coll Y/N
  addDropdown(sh, 32, 6, 15, 1, YN_OPTIONS);  // F — Follow-up Y/N
  addDropdown(sh, 32, 8, 15, 1, YN_OPTIONS);  // H — Outcome Y/N
  addDropdown(sh, 32, 2, 15, 1, VERTICAL_OPTIONS);  // B — Vertical

  setBorder(sh.getRange('A31:K46'), true, true, true, true, false, true);
  _applyMonitoringStatusFormat(sh, 'J32:J46');
}

// ============================================================
// SECTION 3 — THIRD PARTY HANDLING  (rows 49–65)
// ============================================================
function _thirdPartySection(sh) {
  // Row 49 — section header
  var hdr = sh.getRange('A49:K49');
  hdr.merge();
  hdr.setValue('  ▌ THIRD PARTY HANDLING');
  hdr.setBackground(CLR.SECTION_TP).setFontColor(CLR.WHITE)
     .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // Row 50 — column headers
  var tpHeaders = [
    '3rd Party\nName',
    'Case / Bill\nReference',
    'Vertical / Case',
    'Assigned\nDate',
    'Report/Output\nReceived?\n(Y/N)',
    'Report\nDate',
    'Bill\nReceived?\n(Y/N)',
    'Bill\nValidated?\n(Y/N)',
    'Bill\nDate',
    'Status\n(auto)',
    'Remarks'
  ];
  for (var c = 0; c < tpHeaders.length; c++) {
    var cell = sh.getRange(50, c + 1);
    cell.setValue(tpHeaders[c]);
    cell.setBackground(CLR.SECTION_TP).setFontColor(CLR.WHITE)
        .setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  sh.getRange('J50').setBackground('#6a1b9a');

  // Rows 51–65
  for (var r = 51; r <= 65; r++) {
    var isAlt = (r % 2 === 0);
    var rowBg = isAlt ? CLR.OFF_WHITE : CLR.WHITE;
    sh.getRange(r, 1, 1, 11).setBackground(rowBg).setFontSize(9).setVerticalAlignment('middle');
    setDateFormat(sh.getRange(r, 4));  // D — Assigned Date
    setDateFormat(sh.getRange(r, 6));  // F — Report Date
    setDateFormat(sh.getRange(r, 9));  // I — Bill Date

    // J — Status (auto)
    sh.getRange(r, 10).setFormula(
      '=IF(A' + r + '="","—",IF(AND(H' + r + '="Y",I' + r + '<>""),"Closed",' +
      'IF(G' + r + '="Y","Bill Received",' +
      'IF(E' + r + '="Y","Report Received","Open"))))'
    );
    sh.getRange(r, 10).setBackground(isAlt ? CLR.PURPLE_LIGHT : '#f5eeff')
      .setFontSize(9).setHorizontalAlignment('center');
  }

  addDropdown(sh, 51, 5, 15, 1, YN_OPTIONS);  // E — Report Rcvd Y/N
  addDropdown(sh, 51, 7, 15, 1, YN_OPTIONS);  // G — Bill Rcvd Y/N
  addDropdown(sh, 51, 8, 15, 1, YN_OPTIONS);  // H — Bill Validated Y/N

  setBorder(sh.getRange('A50:K65'), true, true, true, true, false, true);
  _applyTpStatusFormat(sh, 'J51:J65');
}

// ============================================================
// SECTION 4 — OTHER TASKS / EXTRA  (rows 68–89)
// ============================================================
function _otherTasksSection(sh) {
  // Row 68 — section header
  var hdr = sh.getRange('A68:H68');
  hdr.merge();
  hdr.setValue('  ▌ OTHER TASKS / EXTRA  (Beyond KRA)');
  hdr.setBackground(CLR.SECTION_OT).setFontColor(CLR.WHITE)
     .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // Row 69 — column headers
  var otHeaders = [
    'Task Description',
    'Category',
    'Date\nAssigned',
    'Target\nCompletion',
    'Actual\nCompletion',
    'Days\nOpen\n(auto)',
    'Status\n(auto)',
    'Remarks'
  ];
  for (var c = 0; c < otHeaders.length; c++) {
    var cell = sh.getRange(69, c + 1);
    cell.setValue(otHeaders[c]);
    cell.setBackground(CLR.SECTION_OT).setFontColor(CLR.WHITE)
        .setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  sh.getRange('F69:G69').setBackground('#8d2000');

  // Rows 70–89
  for (var r = 70; r <= 89; r++) {
    var isAlt = (r % 2 === 0);
    var rowBg = isAlt ? CLR.OFF_WHITE : CLR.WHITE;
    sh.getRange(r, 1, 1, 8).setBackground(rowBg).setFontSize(9).setVerticalAlignment('middle');
    setDateFormat(sh.getRange(r, 3));  // C — Date Assigned
    setDateFormat(sh.getRange(r, 4));  // D — Target Completion
    setDateFormat(sh.getRange(r, 5));  // E — Actual Completion

    // F — Days Open (auto)
    sh.getRange(r, 6).setFormula(
      '=IF(A' + r + '="","",IF(E' + r + '<>"",E' + r + '-C' + r + ',IF(C' + r + '<>"",TODAY()-C' + r + ',"")))'
    );

    // G — Status (auto)
    sh.getRange(r, 7).setFormula(
      '=IF(A' + r + '="","—",IF(E' + r + '<>"","Closed",' +
      'IF(AND(D' + r + '<>"",D' + r + '<TODAY()),"Overdue","Open")))'
    );

    sh.getRange(r, 6).setBackground(isAlt ? CLR.AMBER_LIGHT : '#fff8f5')
      .setFontSize(9).setHorizontalAlignment('center');
    sh.getRange(r, 7).setBackground(isAlt ? CLR.AMBER_LIGHT : '#fff8f5')
      .setFontSize(9).setHorizontalAlignment('center');
  }

  setBorder(sh.getRange('A69:H89'), true, true, true, true, false, true);
  _applyTaskStatusFormat(sh, 'G70:G89');
}

// ============================================================
// Conditional formatting helpers
// ============================================================

function _applyTatConditionalFormat(sh, rangeNotation) {
  var rules = [];
  var range = sh.getRange(rangeNotation);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('✓ Within TAT')
    .setBackground(CLR.GREEN_LIGHT).setFontColor(CLR.GREEN_DARK)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('⚠ Breach')
    .setBackground(CLR.RED_LIGHT).setFontColor(CLR.RED)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending Docs')
    .setBackground(CLR.AMBER_LIGHT).setFontColor(CLR.AMBER)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground(CLR.NAVY_LIGHT).setFontColor(CLR.NAVY_MED)
    .setRanges([range]).build());

  var existing = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(existing.concat(rules));
}

function _applyCaseStatusFormat(sh, rangeNotation) {
  var rules = [];
  var range = sh.getRange(rangeNotation);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground(CLR.GREEN_LIGHT).setFontColor(CLR.GREEN_DARK)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Rejected')
    .setBackground(CLR.RED_LIGHT).setFontColor(CLR.RED)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Docs Pending')
    .setBackground(CLR.AMBER_LIGHT).setFontColor(CLR.AMBER)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Progress')
    .setBackground(CLR.NAVY_LIGHT).setFontColor(CLR.NAVY_MED)
    .setRanges([range]).build());

  var existing = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(existing.concat(rules));
}

function _applyMonitoringStatusFormat(sh, rangeNotation) {
  var rules = [];
  var range = sh.getRange(rangeNotation);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Closed')
    .setBackground(CLR.GREEN_LIGHT).setFontColor(CLR.GREEN_DARK)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Open')
    .setBackground(CLR.AMBER_LIGHT).setFontColor(CLR.AMBER)
    .setRanges([range]).build());

  var existing = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(existing.concat(rules));
}

function _applyTpStatusFormat(sh, rangeNotation) {
  var rules = [];
  var range = sh.getRange(rangeNotation);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Closed')
    .setBackground(CLR.GREEN_LIGHT).setFontColor(CLR.GREEN_DARK)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Open')
    .setBackground(CLR.AMBER_LIGHT).setFontColor(CLR.AMBER)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Bill Received')
    .setBackground(CLR.NAVY_LIGHT).setFontColor(CLR.NAVY_MED)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Report Received')
    .setBackground(CLR.TEAL_LIGHT).setFontColor(CLR.TEAL)
    .setRanges([range]).build());

  var existing = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(existing.concat(rules));
}

function _applyTaskStatusFormat(sh, rangeNotation) {
  var rules = [];
  var range = sh.getRange(rangeNotation);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Closed')
    .setBackground(CLR.GREEN_LIGHT).setFontColor(CLR.GREEN_DARK)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Overdue')
    .setBackground(CLR.RED_LIGHT).setFontColor(CLR.RED)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Open')
    .setBackground(CLR.AMBER_LIGHT).setFontColor(CLR.AMBER)
    .setRanges([range]).build());

  var existing = sh.getConditionalFormatRules();
  sh.setConditionalFormatRules(existing.concat(rules));
}
