// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// Settings.gs — Settings sheet: team roster, verticals,
//               stage weights, TAT target, business rules
// ============================================================

function setupSettings(ss) {
  var sh = getOrCreateSheet(ss, SHEET_NAMES.SETTINGS);

  // ── Column widths ────────────────────────────────────────
  sh.setColumnWidth(1, 55);   // A — #
  sh.setColumnWidth(2, 270);  // B — Stage / Name
  sh.setColumnWidth(3, 110);  // C — Weight / Value
  sh.setColumnWidth(4, 140);  // D — Role
  sh.setColumnWidth(5, 175);  // E — Primary Vertical
  sh.setColumnWidth(6, 175);  // F — Fallback
  sh.setColumnWidth(7, 160);  // G — 3rd Party Owned

  sh.setRowHeight(1, 44);

  // ── Row 1 — Sheet title ──────────────────────────────────
  var title = sh.getRange('A1:G1');
  title.merge();
  title.setValue('SETTINGS  —  Team Roster, Verticals, Stage Weights & TAT Target');
  title.setBackground(CLR.CHARCOAL)
       .setFontColor(CLR.WHITE)
       .setFontSize(12)
       .setFontWeight('bold')
       .setVerticalAlignment('middle')
       .setHorizontalAlignment('left');
  sh.getRange('A1').setValue('  SETTINGS  —  Team Roster, Verticals, Stage Weights & TAT Target');

  // ═══════════════════════════════════════════════════════════
  // SECTION 1 — TEAM ROSTER  (rows 3–9)
  // ═══════════════════════════════════════════════════════════
  _sectionHeader(sh, 3, 'A', 'G', 'TEAM ROSTER');

  var rosterHeaders = ['Name', 'Role', 'Primary Vertical', 'Fallback', '3rd Party Owned'];
  _tableHeader(sh, 4, [
    { col: 'A', label: 'Name',            width: null },
    { col: 'B', label: 'Role',            width: null },
    { col: 'C', label: 'Primary Vertical',width: null },
    { col: 'D', label: 'Fallback',        width: null },
    { col: 'E', label: '3rd Party Owned', width: null }
  ]);
  sh.getRange('A4').setValue('Name');
  sh.getRange('B4').setValue('Role');
  sh.getRange('C4').setValue('Primary Vertical');
  sh.getRange('D4').setValue('Fallback');
  sh.getRange('E4').setValue('3rd Party Owned');
  _styleHeaderRow(sh.getRange('A4:G4'), CLR.GRAY_LIGHT, CLR.BLACK);

  // Manager row
  var mgrRow = sh.getRange('A5:G5');
  sh.getRange('A5').setValue(MANAGER_NAME);
  sh.getRange('B5').setValue('Manager');
  sh.getRange('C5').setValue('All (oversight)');
  sh.getRange('D5').setValue('—');
  sh.getRange('E5').setValue('—');
  _styleDataRow(mgrRow, CLR.NAVY_LIGHT, CLR.BLACK, true);

  // Team member rows 6–9
  for (var i = 0; i < TEAM.length; i++) {
    var r = 6 + i;
    var m = TEAM[i];
    var row = sh.getRange(r, 1, 1, 7);
    sh.getRange(r, 1).setValue(m.name);
    sh.getRange(r, 2).setValue(m.role);
    sh.getRange(r, 3).setValue(m.vertical);
    sh.getRange(r, 4).setValue(m.fallback);
    sh.getRange(r, 5).setValue(m.thirdParty);
    _styleDataRow(row, i % 2 === 0 ? CLR.WHITE : CLR.OFF_WHITE, CLR.BLACK, false);
  }
  setBorder(sh.getRange('A4:G9'), true, true, true, true, false, true);

  // ═══════════════════════════════════════════════════════════
  // SECTION 2 — VERTICALS  (rows 11–15)
  // ═══════════════════════════════════════════════════════════
  _sectionHeader(sh, 11, 'A', 'G', 'VERTICALS & OWNERSHIP');
  sh.getRange('A12').setValue('Vertical');
  sh.getRange('B12').setValue('Primary Owner');
  sh.getRange('C12').setValue('Fallback Owner');
  _styleHeaderRow(sh.getRange('A12:G12'), CLR.GRAY_LIGHT, CLR.BLACK);

  for (var j = 0; j < VERTICALS_CONFIG.length; j++) {
    var vr = 13 + j;
    var v  = VERTICALS_CONFIG[j];
    sh.getRange(vr, 1).setValue(v.name);
    sh.getRange(vr, 2).setValue(v.owner);
    sh.getRange(vr, 3).setValue(v.fallback);
    _styleDataRow(sh.getRange(vr, 1, 1, 7), j % 2 === 0 ? CLR.WHITE : CLR.OFF_WHITE, CLR.BLACK, false);
  }
  setBorder(sh.getRange('A12:C15'), true, true, true, true, false, true);

  // ═══════════════════════════════════════════════════════════
  // SECTION 3 — STAGE WEIGHTS (Marketplace & EPR)  rows 17–24
  // ═══════════════════════════════════════════════════════════
  _sectionHeader(sh, 17, 'A', 'G', 'STAGE WEIGHTS  —  Marketplace & EPR / Sustainability');

  sh.getRange('A18').setValue('#');
  sh.getRange('B18').setValue('Stage');
  sh.getRange('C18').setValue('Weight %');
  _styleHeaderRow(sh.getRange('A18:C18'), CLR.GRAY_LIGHT, CLR.BLACK);

  for (var s = 0; s < STAGES_MP_EPR.length; s++) {
    var sr = 19 + s;
    var st = STAGES_MP_EPR[s];
    sh.getRange(sr, 1).setValue(st.order);
    sh.getRange(sr, 2).setValue(st.label);
    sh.getRange(sr, 3).setValue(st.weight);
    _styleDataRow(sh.getRange(sr, 1, 1, 3), s % 2 === 0 ? CLR.WHITE : CLR.OFF_WHITE, CLR.BLACK, false);
    sh.getRange(sr, 3).setNumberFormat('0%');
  }
  // Total row (row 25)
  sh.getRange('B25').setValue('Total');
  sh.getRange('C25').setFormula('=SUM(C19:C24)');
  sh.getRange('C25').setNumberFormat('0%');
  _styleDataRow(sh.getRange('A25:C25'), CLR.NAVY_LIGHT, CLR.NAVY, true);
  setBorder(sh.getRange('A18:C25'), true, true, true, true, false, true);

  // ═══════════════════════════════════════════════════════════
  // SECTION 4 — STAGE WEIGHTS (Open Marketplace)  rows 27–33
  // ═══════════════════════════════════════════════════════════
  _sectionHeader(sh, 27, 'A', 'G', 'STAGE WEIGHTS  —  Open Marketplace');

  sh.getRange('A28').setValue('#');
  sh.getRange('B28').setValue('Stage');
  sh.getRange('C28').setValue('Weight %');
  _styleHeaderRow(sh.getRange('A28:C28'), CLR.GRAY_LIGHT, CLR.BLACK);

  for (var t = 0; t < STAGES_OMP.length; t++) {
    var tr = 29 + t;
    var ot = STAGES_OMP[t];
    sh.getRange(tr, 1).setValue(ot.order);
    sh.getRange(tr, 2).setValue(ot.label);
    sh.getRange(tr, 3).setValue(ot.weight);
    _styleDataRow(sh.getRange(tr, 1, 1, 3), t % 2 === 0 ? CLR.WHITE : CLR.OFF_WHITE, CLR.BLACK, false);
    sh.getRange(tr, 3).setNumberFormat('0%');
  }
  // Total row (row 34)
  sh.getRange('B34').setValue('Total');
  sh.getRange('C34').setFormula('=SUM(C29:C33)');
  sh.getRange('C34').setNumberFormat('0%');
  _styleDataRow(sh.getRange('A34:C34'), CLR.NAVY_LIGHT, CLR.NAVY, true);
  setBorder(sh.getRange('A28:C34'), true, true, true, true, false, true);

  // ═══════════════════════════════════════════════════════════
  // SECTION 5 — TAT TARGET  (row 36–38)
  // Row 38, Col C = TAT_TARGET_DAYS  ← referenced by team sheet formulas as Settings!$C$38
  // ═══════════════════════════════════════════════════════════
  _sectionHeader(sh, 36, 'A', 'G', 'TAT TARGET');

  sh.getRange('A37').setValue('Overall TAT Target  (Doc Validation → L1 Approval, working days)');
  sh.getRange('A37').setFontWeight('bold').setFontColor(CLR.BLACK);
  sh.getRange('C37').setValue(TAT_TARGET_DAYS);
  sh.getRange('C37').setFontSize(14).setFontWeight('bold').setFontColor(CLR.NAVY)
    .setHorizontalAlignment('center');
  sh.getRange('A37:G37').setBackground(CLR.AMBER_LIGHT);
  setBorder(sh.getRange('A37:G37'), true, true, true, true, false, false);

  // Row 38 — the actual data cell referenced by team sheets (Settings!$C$38)
  sh.getRange('A38').setValue('TAT Target (days):');
  sh.getRange('A38').setFontColor(CLR.GRAY_DARK);
  sh.getRange('C38').setValue(TAT_TARGET_DAYS);
  sh.getRange('C38').setFontWeight('bold').setFontColor(CLR.NAVY)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════
  // SECTION 6 — BUSINESS RULE  (row 40)
  // ═══════════════════════════════════════════════════════════
  var rule = sh.getRange('A40:G40');
  rule.merge();
  rule.setValue(
    'Rule: TAT only starts when Doc % Collected = 100%.  ' +
    'If docs are incomplete or incorrect, TAT is paused — case stays In Progress with the onboarding team.'
  );
  rule.setBackground(CLR.GRAY_LIGHT)
      .setFontColor(CLR.GRAY_DARK)
      .setFontSize(9)
      .setItalic(true)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.setRowHeight(40, 36);

  // ── Freeze header row ────────────────────────────────────
  sh.setFrozenRows(1);

  // ── Tab color ────────────────────────────────────────────
  sh.setTabColor(CLR.CHARCOAL);
}

// ============================================================
// Private helpers (used only within Settings.gs scope)
// ============================================================

function _sectionHeader(sh, row, fromCol, toCol, label) {
  var rng = sh.getRange(row + ':' + row);
  var fromIdx = _colLetter(fromCol);
  var toIdx   = _colLetter(toCol);
  var mergeRange = sh.getRange(fromCol + row + ':' + toCol + row);
  mergeRange.merge();
  mergeRange.setValue('  ' + label);
  mergeRange.setBackground(CLR.CHARCOAL)
            .setFontColor(CLR.WHITE)
            .setFontSize(10)
            .setFontWeight('bold')
            .setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
}

function _tableHeader(sh, row, columns) {
  // columns is used only for documentation; actual values set by caller
}

function _styleHeaderRow(range, bg, fg) {
  range.setBackground(bg)
       .setFontColor(fg)
       .setFontWeight('bold')
       .setFontSize(9)
       .setVerticalAlignment('middle')
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function _styleDataRow(range, bg, fg, bold) {
  range.setBackground(bg)
       .setFontColor(fg)
       .setFontSize(9)
       .setFontWeight(bold ? 'bold' : 'normal')
       .setVerticalAlignment('middle');
}

function _colLetter(letter) {
  return letter.charCodeAt(0) - 64;
}
