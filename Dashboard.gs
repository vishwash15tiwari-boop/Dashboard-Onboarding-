// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// Dashboard.gs — Dashboard sheet
//
// Layout:
//   Row 1     : Platform title (merged B–L)
//   Row 2     : Subtitle / live view note (merged B–L)
//   Row 3     : (blank spacer)
//   Row 4     : KPI card labels  (B-C | D-E | F-G | H-I | J-K | L-M)
//   Row 5     : KPI card values  (same pairing)
//   Row 6-7   : (blank)
//   Row 8     : TEAM PERFORMANCE SUMMARY header
//   Row 9     : Table column headers
//   Rows 10-13: Per-member rows (Harshita, Vamsi, Naveen, Vishwash)
//   Row 14    : TOTAL / TEAM aggregate row
//   Row 15    : (blank)
//   Row 16    : PIPELINE BY VERTICAL header
//   Row 17    : Table column headers
//   Rows 18-20: Per-vertical rows
//   Row 21    : (blank)
//   Row 22    : ALERTS header
//   Rows 23-27: Alert items
// ============================================================

function setupDashboard(ss) {
  var sh = getOrCreateSheet(ss, SHEET_NAMES.DASHBOARD);

  // ── Column widths (matches original Excel Dashboard) ─────
  sh.setColumnWidth(1,  15);   // A — narrow left margin
  sh.setColumnWidth(2,  295);  // B — main label column
  sh.setColumnWidth(3,  88);   // C
  sh.setColumnWidth(4,  88);   // D
  sh.setColumnWidth(5,  135);  // E
  sh.setColumnWidth(6,  88);   // F
  sh.setColumnWidth(7,  105);  // G
  sh.setColumnWidth(8,  88);   // H
  sh.setColumnWidth(9,  88);   // I
  sh.setColumnWidth(10, 88);   // J
  sh.setColumnWidth(11, 75);   // K
  sh.setColumnWidth(12, 75);   // L
  sh.setColumnWidth(13, 75);   // M
  sh.setColumnWidth(14, 15);   // N — narrow right margin

  // ── Row heights ──────────────────────────────────────────
  sh.setRowHeight(1, 46);
  sh.setRowHeight(2, 30);
  sh.setRowHeight(3, 10);
  sh.setRowHeight(4, 50);
  sh.setRowHeight(5, 46);
  sh.setRowHeight(6, 10);
  sh.setRowHeight(7, 6);
  sh.setRowHeight(8, 28);
  sh.setRowHeight(9, 44);
  for (var pr = 10; pr <= 14; pr++) sh.setRowHeight(pr, 28);
  sh.setRowHeight(15, 10);
  sh.setRowHeight(16, 28);
  sh.setRowHeight(17, 40);
  for (var vr = 18; vr <= 20; vr++) sh.setRowHeight(vr, 26);
  sh.setRowHeight(21, 10);
  sh.setRowHeight(22, 28);
  for (var ar = 23; ar <= 27; ar++) sh.setRowHeight(ar, 26);

  // ════════════════════════════════════════════════════════
  // ROWS 1–2 — Platform header
  // ════════════════════════════════════════════════════════
  var titleRange = sh.getRange('B1:M1');
  titleRange.merge();
  titleRange.setValue('ONBOARDING OPERATIONS PLATFORM  —  CONTROL TOWER');
  titleRange.setBackground(CLR.NAVY)
            .setFontColor(CLR.WHITE)
            .setFontSize(15)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('left');
  sh.getRange('B1').setValue('  ONBOARDING OPERATIONS PLATFORM  —  CONTROL TOWER');

  var subtitleRange = sh.getRange('B2:M2');
  subtitleRange.merge();
  subtitleRange.setValue(
    '  Live view  •  All numbers auto-pulled from individual sheets  •  Team of 4  |  Manager: ' + MANAGER_NAME
  );
  subtitleRange.setBackground(CLR.NAVY_MED)
               .setFontColor(CLR.WHITE)
               .setFontSize(9)
               .setVerticalAlignment('middle');

  // Side margins (A col, N col)
  sh.getRange('A1:A27').setBackground(CLR.NAVY);
  sh.getRange('N1:N27').setBackground(CLR.NAVY);

  // ════════════════════════════════════════════════════════
  // ROWS 4–5 — KPI Cards
  // ════════════════════════════════════════════════════════
  var kpiCards = [
    { labelCols: 'B4:C4', valueCols: 'B5:C5', label: 'Total\nOB Cases',             formula: _kpiFormula('COUNTA_CASES'),  bg: CLR.NAVY,    fg: CLR.WHITE,  numFmt: '0'   },
    { labelCols: 'D4:E4', valueCols: 'D5:E5', label: 'OB\nCompleted',               formula: _kpiFormula('OB_COMPLETED'),  bg: CLR.TEAL,    fg: CLR.WHITE,  numFmt: '0'   },
    { labelCols: 'F4:G4', valueCols: 'F5:G5', label: 'OB\nIn Progress',             formula: _kpiFormula('OB_INPROG'),     bg: CLR.NAVY_MED, fg: CLR.WHITE, numFmt: '0'   },
    { labelCols: 'H4:I4', valueCols: 'H5:I5', label: 'Docs Incomplete\n(TAT paused)',formula: _kpiFormula('DOCS_PENDING'), bg: CLR.AMBER,   fg: CLR.WHITE,  numFmt: '0'   },
    { labelCols: 'J4:K4', valueCols: 'J5:K5', label: 'TAT\nBreaches',               formula: _kpiFormula('TAT_BREACH'),   bg: CLR.RED,     fg: CLR.WHITE,  numFmt: '0'   },
    { labelCols: 'L4:M4', valueCols: 'L5:M5', label: 'TAT Target\n(days)',           formula: '=Settings!C38',             bg: CLR.CHARCOAL, fg: CLR.WHITE, numFmt: '0'   }
  ];

  kpiCards.forEach(function(card) {
    // Label row (row 4)
    var lRange = sh.getRange(card.labelCols);
    lRange.merge();
    lRange.setValue(card.label);
    lRange.setBackground(card.bg)
          .setFontColor(card.fg)
          .setFontSize(9)
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Value row (row 5)
    var vRange = sh.getRange(card.valueCols);
    vRange.merge();
    vRange.setFormula(card.formula);
    vRange.setBackground(card.bg)
          .setFontColor(card.fg)
          .setFontSize(24)
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setNumberFormat(card.numFmt);
  });

  // ════════════════════════════════════════════════════════
  // ROW 8 — TEAM PERFORMANCE SUMMARY header
  // ════════════════════════════════════════════════════════
  var perfHdr = sh.getRange('B8:M8');
  perfHdr.merge();
  perfHdr.setValue('  TEAM PERFORMANCE SUMMARY');
  perfHdr.setBackground(CLR.CHARCOAL).setFontColor(CLR.WHITE)
         .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // ROW 9 — Column headers for team performance table
  var perfHeaders = [
    'Team\nMember', 'OB\nOpen', 'OB\nDone',
    'Monitoring\nOpen', 'Pending Docs\n(TAT paused)',
    '3rd Party\nOpen', '3rd Party\nDone',
    'Other\nOpen', 'Other\nDone',
    'Avg TAT\n(days)', '% Within\nTAT', 'TAT\nBreaches'
  ];
  var perfCols = ['B','C','D','E','F','G','H','I','J','K','L','M'];
  for (var pi = 0; pi < perfHeaders.length; pi++) {
    sh.getRange(perfCols[pi] + '9').setValue(perfHeaders[pi]);
    sh.getRange(perfCols[pi] + '9')
      .setBackground(CLR.GRAY_LIGHT).setFontColor(CLR.BLACK)
      .setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }

  // ROWS 10–13 — Per-member data rows
  var members = SHEET_NAMES.MEMBERS;  // ['Harshita', 'Vamsi', 'Naveen', 'Vishwash']
  for (var mi = 0; mi < members.length; mi++) {
    var row = 10 + mi;
    var nm  = members[mi];
    var isAlt = (mi % 2 === 0);
    var rowBg = isAlt ? CLR.WHITE : CLR.OFF_WHITE;

    sh.getRange('B' + row).setValue(nm)
      .setFontWeight('bold').setFontSize(9)
      .setBackground(rowBg).setFontColor(CLR.NAVY);

    sh.getRange('C' + row).setFormula('=COUNTIF(' + nm + '!W8:W27,"In Progress")')
      .setBackground(rowBg);
    sh.getRange('D' + row).setFormula('=COUNTIF(' + nm + '!W8:W27,"Completed")+COUNTIF(' + nm + '!W8:W27,"Rejected")')
      .setBackground(rowBg);
    sh.getRange('E' + row).setFormula(
      '=COUNTIF(' + nm + '!J32:J46,"Open")+COUNTIF(' + nm + '!J32:J46,"In Progress")+COUNTIF(' + nm + '!J32:J46,"Started")'
    ).setBackground(rowBg);
    sh.getRange('F' + row).setFormula('=COUNTIF(' + nm + '!U8:U27,"Pending Docs")')
      .setBackground(rowBg);
    sh.getRange('G' + row).setFormula(
      '=COUNTIF(' + nm + '!J51:J65,"Open")+COUNTIF(' + nm + '!J51:J65,"Report Received")+COUNTIF(' + nm + '!J51:J65,"Bill Received")'
    ).setBackground(rowBg);
    sh.getRange('H' + row).setFormula('=COUNTIF(' + nm + '!J51:J65,"Closed")')
      .setBackground(rowBg);
    sh.getRange('I' + row).setFormula(
      '=COUNTIF(' + nm + '!G70:G89,"Open")+COUNTIF(' + nm + '!G70:G89,"Overdue")'
    ).setBackground(rowBg);
    sh.getRange('J' + row).setFormula('=COUNTIF(' + nm + '!G70:G89,"Closed")')
      .setBackground(rowBg);
    sh.getRange('K' + row).setFormula('=IFERROR(ROUND(AVERAGEIF(' + nm + '!T8:T27,">0"),1),"—")')
      .setBackground(rowBg);
    sh.getRange('L' + row).setFormula(
      '=IFERROR(COUNTIF(' + nm + '!U8:U27,"✓ Within TAT")/' +
      '(COUNTIF(' + nm + '!U8:U27,"✓ Within TAT")+COUNTIF(' + nm + '!U8:U27,"⚠ Breach")),"-")'
    ).setBackground(rowBg).setNumberFormat('0%');
    sh.getRange('M' + row).setFormula('=COUNTIF(' + nm + '!U8:U27,"⚠ Breach")')
      .setBackground(rowBg);

    sh.getRange('B' + row + ':M' + row).setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange('B' + row).setHorizontalAlignment('left');
  }

  // ROW 14 — TOTAL / TEAM aggregate
  sh.getRange('B14').setValue('  TOTAL / TEAM')
    .setFontWeight('bold').setFontSize(9).setFontColor(CLR.WHITE)
    .setBackground(CLR.NAVY).setHorizontalAlignment('left');

  sh.getRange('C14').setFormula('=SUM(C10:C13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('D14').setFormula('=SUM(D10:D13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('E14').setFormula('=SUM(E10:E13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('F14').setFormula('=SUM(F10:F13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('G14').setFormula('=SUM(G10:G13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('H14').setFormula('=SUM(H10:H13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('I14').setFormula('=SUM(I10:I13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);
  sh.getRange('J14').setFormula('=SUM(J10:J13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);

  // Avg TAT — weighted average across all sheets
  sh.getRange('K14').setFormula(
    '=IFERROR(ROUND(' +
    '(IFERROR(SUMIF(Harshita!T8:T27,">0"),0)+IFERROR(SUMIF(Vamsi!T8:T27,">0"),0)' +
    '+IFERROR(SUMIF(Naveen!T8:T27,">0"),0)+IFERROR(SUMIF(Vishwash!T8:T27,">0"),0))/' +
    '(IFERROR(COUNTIF(Harshita!T8:T27,">0"),0)+IFERROR(COUNTIF(Vamsi!T8:T27,">0"),0)' +
    '+IFERROR(COUNTIF(Naveen!T8:T27,">0"),0)+IFERROR(COUNTIF(Vishwash!T8:T27,">0"),0)),1),"—")'
  ).setBackground(CLR.NAVY).setFontColor(CLR.WHITE);

  // % Within TAT (combined)
  sh.getRange('L14').setFormula(
    '=IFERROR(' +
    '(COUNTIF(Harshita!U8:U27,"✓ Within TAT")+COUNTIF(Vamsi!U8:U27,"✓ Within TAT")' +
    '+COUNTIF(Naveen!U8:U27,"✓ Within TAT")+COUNTIF(Vishwash!U8:U27,"✓ Within TAT"))/' +
    '(COUNTIF(Harshita!U8:U27,"✓ Within TAT")+COUNTIF(Vamsi!U8:U27,"✓ Within TAT")' +
    '+COUNTIF(Naveen!U8:U27,"✓ Within TAT")+COUNTIF(Vishwash!U8:U27,"✓ Within TAT")' +
    '+COUNTIF(Harshita!U8:U27,"⚠ Breach")+COUNTIF(Vamsi!U8:U27,"⚠ Breach")' +
    '+COUNTIF(Naveen!U8:U27,"⚠ Breach")+COUNTIF(Vishwash!U8:U27,"⚠ Breach")),"-")'
  ).setBackground(CLR.NAVY).setFontColor(CLR.WHITE).setNumberFormat('0%');

  sh.getRange('M14').setFormula('=SUM(M10:M13)').setBackground(CLR.NAVY).setFontColor(CLR.WHITE);

  sh.getRange('B14:M14').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('B14').setHorizontalAlignment('left');

  // Borders for performance table
  setBorder(sh.getRange('B9:M14'), true, true, true, true, false, true, CLR.GRAY_MED);

  // ════════════════════════════════════════════════════════
  // ROW 16 — PIPELINE BY VERTICAL header
  // ════════════════════════════════════════════════════════
  var pipeHdr = sh.getRange('B16:M16');
  pipeHdr.merge();
  pipeHdr.setValue('  PIPELINE BY VERTICAL');
  pipeHdr.setBackground(CLR.CHARCOAL).setFontColor(CLR.WHITE)
         .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // ROW 17 — Pipeline column headers
  var pipeHeaders = ['Vertical', 'Owner', 'Total Cases', 'In Progress', 'Completed', 'Rejected'];
  var pipeCols    = ['B', 'C', 'D', 'E', 'F', 'G'];
  for (var ki = 0; ki < pipeHeaders.length; ki++) {
    sh.getRange(pipeCols[ki] + '17').setValue(pipeHeaders[ki]);
    sh.getRange(pipeCols[ki] + '17')
      .setBackground(CLR.GRAY_LIGHT).setFontColor(CLR.BLACK)
      .setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  sh.getRange('H17:M17').setBackground(CLR.GRAY_LIGHT);

  // Rows 18–20 — Per-vertical data
  for (var vi = 0; vi < VERTICALS_CONFIG.length; vi++) {
    var vRow    = 18 + vi;
    var vert    = VERTICALS_CONFIG[vi];
    var vName   = vert.name;
    var isAlt   = (vi % 2 === 0);
    var vBg     = isAlt ? CLR.WHITE : CLR.OFF_WHITE;

    sh.getRange('B' + vRow).setValue(vName)
      .setFontWeight('bold').setFontSize(9).setFontColor(CLR.NAVY)
      .setBackground(vBg);
    sh.getRange('C' + vRow).setValue(vert.owner)
      .setFontSize(9).setBackground(vBg).setHorizontalAlignment('center');

    var sheets = SHEET_NAMES.MEMBERS;
    var totalF = sheets.map(function(s){ return 'COUNTIF(' + s + '!C8:C27,"' + vName + '")'; }).join('+');
    var inProgF = sheets.map(function(s){ return 'COUNTIFS(' + s + '!C8:C27,"' + vName + '",' + s + '!W8:W27,"In Progress")'; }).join('+');
    var compF   = sheets.map(function(s){ return 'COUNTIFS(' + s + '!C8:C27,"' + vName + '",' + s + '!W8:W27,"Completed")'; }).join('+');
    var rejF    = sheets.map(function(s){ return 'COUNTIFS(' + s + '!C8:C27,"' + vName + '",' + s + '!W8:W27,"Rejected")'; }).join('+');

    sh.getRange('D' + vRow).setFormula('=' + totalF).setBackground(vBg).setHorizontalAlignment('center').setFontSize(9);
    sh.getRange('E' + vRow).setFormula('=' + inProgF).setBackground(vBg).setHorizontalAlignment('center').setFontSize(9);
    sh.getRange('F' + vRow).setFormula('=' + compF).setBackground(vBg).setHorizontalAlignment('center').setFontSize(9);
    sh.getRange('G' + vRow).setFormula('=' + rejF).setBackground(vBg).setHorizontalAlignment('center').setFontSize(9);
    sh.getRange('H' + vRow + ':M' + vRow).setBackground(vBg);
  }
  setBorder(sh.getRange('B17:G20'), true, true, true, true, false, true, CLR.GRAY_MED);

  // ════════════════════════════════════════════════════════
  // ROW 22 — ALERTS header
  // ════════════════════════════════════════════════════════
  var alertHdr = sh.getRange('B22:M22');
  alertHdr.merge();
  alertHdr.setValue('  ALERTS  (action needed)');
  alertHdr.setBackground(CLR.CHARCOAL).setFontColor(CLR.WHITE)
          .setFontSize(10).setFontWeight('bold').setVerticalAlignment('middle');

  // Alert rows 23–27
  var alerts = [
    {
      icon:  '🔴',
      label: 'Cases currently breaching TAT  (>' + TAT_TARGET_DAYS + ' days)',
      formula: _alertFormula('breach'),
      bg:    CLR.RED_LIGHT,
      fg:    CLR.RED
    },
    {
      icon:  '🟡',
      label: 'Cases with incomplete docs  (TAT paused — docs being fixed)',
      formula: _alertFormula('docs'),
      bg:    CLR.AMBER_LIGHT,
      fg:    CLR.AMBER
    },
    {
      icon:  '🟣',
      label: '3rd Party tasks still open  (all team)',
      formula: _alertFormula('tp'),
      bg:    CLR.PURPLE_LIGHT,
      fg:    CLR.PURPLE
    },
    {
      icon:  '🟠',
      label: 'Other tasks marked Overdue',
      formula: _alertFormula('overdue'),
      bg:    CLR.AMBER_LIGHT,
      fg:    CLR.AMBER
    },
    {
      icon:  '🔵',
      label: 'Monitoring tasks not yet closed',
      formula: _alertFormula('monitoring'),
      bg:    CLR.TEAL_LIGHT,
      fg:    CLR.TEAL
    }
  ];

  for (var ai = 0; ai < alerts.length; ai++) {
    var aRow = 23 + ai;
    var alert = alerts[ai];

    var labelRange = sh.getRange('B' + aRow + ':D' + aRow);
    labelRange.merge();
    labelRange.setValue(alert.icon + '  ' + alert.label);
    labelRange.setBackground(alert.bg).setFontColor(alert.fg)
              .setFontSize(9).setFontWeight('bold')
              .setVerticalAlignment('middle');

    sh.getRange('E' + aRow).setFormula(alert.formula);
    sh.getRange('E' + aRow)
      .setBackground(alert.bg).setFontColor(alert.fg)
      .setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    sh.getRange('F' + aRow + ':M' + aRow).setBackground(alert.bg);
  }
  setBorder(sh.getRange('B23:E27'), true, true, true, true, false, true, CLR.GRAY_MED);

  // ── Freeze rows 1-2 ──────────────────────────────────────
  sh.setFrozenRows(2);

  // ── Tab color ─────────────────────────────────────────────
  sh.setTabColor(CLR.NAVY_DARK);
}

// ── KPI formulas ────────────────────────────────────────────
function _kpiFormula(type) {
  var s = SHEET_NAMES.MEMBERS;
  switch (type) {
    case 'COUNTA_CASES':
      return '=' + s.map(function(n){ return 'COUNTA(' + n + '!A8:A27)'; }).join('+');
    case 'OB_COMPLETED':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!W8:W27,"Completed")'; }).join('+');
    case 'OB_INPROG':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!W8:W27,"In Progress")'; }).join('+');
    case 'DOCS_PENDING':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!U8:U27,"Pending Docs")'; }).join('+');
    case 'TAT_BREACH':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!U8:U27,"⚠ Breach")'; }).join('+');
    default:
      return '';
  }
}

// ── Alert formulas ──────────────────────────────────────────
function _alertFormula(type) {
  var s = SHEET_NAMES.MEMBERS;
  switch (type) {
    case 'breach':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!U8:U27,"⚠ Breach")'; }).join('+');
    case 'docs':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!U8:U27,"Pending Docs")'; }).join('+');
    case 'tp':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!J51:J65,"Open")'; }).join('+');
    case 'overdue':
      return '=' + s.map(function(n){ return 'COUNTIF(' + n + '!G70:G89,"Overdue")'; }).join('+');
    case 'monitoring':
      return '=' + s.map(function(n){
        return 'COUNTIF(' + n + '!J32:J46,"Open")+COUNTIF(' + n + '!J32:J46,"In Progress")+COUNTIF(' + n + '!J32:J46,"Started")';
      }).join('+');
    default:
      return '';
  }
}
