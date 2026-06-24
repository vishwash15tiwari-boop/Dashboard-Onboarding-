// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// BackendData.gs — Server-side data functions for the SPA
// ============================================================

function dateStr(v) {
  if (!v || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  return String(v);
}

function parseInputDate(s) {
  if (!s || s === '') return '';
  var p = String(s).split('-');
  if (p.length === 3 && p[0].length === 4) return new Date(+p[0], +p[1] - 1, +p[2]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : d;
}

// ── READ: All OB cases across all members ──────────────────
function getCasesListData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var cases = [];
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(8, 1, 20, 23).getValues().forEach(function(row, idx) {
      if (!row[0]) return;
      cases.push({
        owner: name, rowIndex: 8 + idx,
        caseId: row[0], vendor: row[1], vertical: row[2], dateInit: dateStr(row[3]),
        docPct: row[4], docStatus: row[18], tatDays: row[19],
        tatStatus: row[20], pctDone: row[21], caseStatus: row[22]
      });
    });
  });
  return cases;
}

// ── READ: All data for one member ─────────────────────────
function getWorkspaceData(memberName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(memberName);

  // Auto-create sheet on first access if it doesn't exist yet
  if (!sh) {
    var cfg = null;
    for (var i = 0; i < TEAM.length; i++) {
      if (TEAM[i].name === memberName) { cfg = TEAM[i]; break; }
    }
    if (!cfg) return null;
    setupTeamSheet(ss, cfg);
    sh = ss.getSheetByName(memberName);
    if (!sh) return null;
  }

  var cases = [];
  sh.getRange(8, 1, 20, 23).getValues().forEach(function(row, idx) {
    if (!row[0]) return;
    cases.push({
      rowIndex: 8 + idx, caseId: row[0], vendor: row[1], vertical: row[2],
      dateInit: dateStr(row[3]), docPct: row[4], docStatus: row[18],
      tatDays: row[19], tatStatus: row[20], pctDone: row[21], caseStatus: row[22]
    });
  });

  var monitoring = [];
  sh.getRange(32, 1, 15, 11).getValues().forEach(function(row, idx) {
    if (!row[0]) return;
    monitoring.push({
      rowIndex: 32 + idx, seller: row[0], vertical: row[1], month: row[2],
      dataCollected: row[3], dataDate: dateStr(row[4]),
      followUp: row[5], followUpDate: dateStr(row[6]),
      outcome: row[7], presentDate: dateStr(row[8]), status: row[9], remarks: row[10]
    });
  });

  var thirdParty = [];
  sh.getRange(51, 1, 15, 11).getValues().forEach(function(row, idx) {
    if (!row[0]) return;
    thirdParty.push({
      rowIndex: 51 + idx, tpName: row[0], caseRef: row[1], vertical: row[2],
      assignedDate: dateStr(row[3]), reportRcvd: row[4], reportDate: dateStr(row[5]),
      billRcvd: row[6], billValid: row[7], billDate: dateStr(row[8]),
      status: row[9], remarks: row[10]
    });
  });

  var tasks = [];
  sh.getRange(70, 1, 20, 8).getValues().forEach(function(row, idx) {
    if (!row[0]) return;
    tasks.push({
      rowIndex: 70 + idx, description: row[0], category: row[1],
      dateAssigned: dateStr(row[2]), targetDate: dateStr(row[3]), actualDate: dateStr(row[4]),
      daysOpen: row[5], status: row[6], remarks: row[7]
    });
  });

  return {
    member: memberName,
    summary: {
      totalCases:     cases.length,
      completed:      cases.filter(function(c) { return c.caseStatus === 'Completed'; }).length,
      inProgress:     cases.filter(function(c) { return c.caseStatus === 'In Progress' || c.caseStatus === 'Docs Pending' || c.caseStatus === 'Not Started'; }).length,
      tatBreaches:    cases.filter(function(c) { return c.tatStatus === '⚠ Breach'; }).length,
      monitoringOpen: monitoring.filter(function(m) { return m.status === 'Open' || m.status === 'In Progress' || m.status === 'Started'; }).length,
      tpOpen:         thirdParty.filter(function(t) { return t.status === 'Open'; }).length,
      tasksDue:       tasks.filter(function(t) { return t.status === 'Overdue'; }).length
    },
    cases: cases, monitoring: monitoring, thirdParty: thirdParty, tasks: tasks
  };
}

// ── READ: Single case + member monitoring + 3P ────────────
function getCase360Data(ownerName, rowIndex) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ownerName);
  if (!sh) return null;
  var row = sh.getRange(rowIndex, 1, 1, 23).getValues()[0];
  var monitoring = [];
  sh.getRange(32, 1, 15, 11).getValues().forEach(function(r) {
    if (!r[0]) return;
    monitoring.push({ seller: r[0], vertical: r[1], month: r[2], dataCollected: r[3], dataDate: dateStr(r[4]), followUp: r[5], followUpDate: dateStr(r[6]), outcome: r[7], presentDate: dateStr(r[8]), status: r[9], remarks: r[10] });
  });
  var thirdParty = [];
  sh.getRange(51, 1, 15, 11).getValues().forEach(function(r) {
    if (!r[0]) return;
    thirdParty.push({ tpName: r[0], caseRef: r[1], vertical: r[2], assignedDate: dateStr(r[3]), reportRcvd: r[4], reportDate: dateStr(r[5]), billRcvd: r[6], billValid: r[7], billDate: dateStr(r[8]), status: r[9], remarks: r[10] });
  });
  return {
    case: {
      owner: ownerName, rowIndex: rowIndex,
      caseId: row[0], vendor: row[1], vertical: row[2], dateInit: dateStr(row[3]),
      docPct: row[4], docValid: row[5], docValidDate: dateStr(row[6]),
      stage2: row[7], stage2Date: dateStr(row[8]), stage3: row[9], stage3Date: dateStr(row[10]),
      stage4: row[11], stage4Date: dateStr(row[12]), decision: row[13],
      mom: row[14], momDate: dateStr(row[15]), l1Approval: row[16], l1Date: dateStr(row[17]),
      docStatus: row[18], tatDays: row[19], tatStatus: row[20], pctDone: row[21], caseStatus: row[22]
    },
    monitoring: monitoring, thirdParty: thirdParty
  };
}

// ── READ: All monitoring records ───────────────────────────
function getMonitoringListData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var records = [];
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(32, 1, 15, 11).getValues().forEach(function(row, idx) {
      if (!row[0]) return;
      records.push({ owner: name, rowIndex: 32 + idx, seller: row[0], vertical: row[1], month: row[2], dataCollected: row[3], dataDate: dateStr(row[4]), followUp: row[5], followUpDate: dateStr(row[6]), outcome: row[7], presentDate: dateStr(row[8]), status: row[9], remarks: row[10] });
    });
  });
  return records;
}

// ── READ: All third-party records ──────────────────────────
function getThirdPartyListData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var records = [];
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(51, 1, 15, 11).getValues().forEach(function(row, idx) {
      if (!row[0]) return;
      records.push({ owner: name, rowIndex: 51 + idx, tpName: row[0], caseRef: row[1], vertical: row[2], assignedDate: dateStr(row[3]), reportRcvd: row[4], reportDate: dateStr(row[5]), billRcvd: row[6], billValid: row[7], billDate: dateStr(row[8]), status: row[9], remarks: row[10] });
    });
  });
  return records;
}

// ── READ: All tasks ────────────────────────────────────────
function getProjectsData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var records = [];
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(70, 1, 20, 8).getValues().forEach(function(row, idx) {
      if (!row[0]) return;
      records.push({ owner: name, rowIndex: 70 + idx, description: row[0], category: row[1], dateAssigned: dateStr(row[2]), targetDate: dateStr(row[3]), actualDate: dateStr(row[4]), daysOpen: row[5], status: row[6], remarks: row[7] });
    });
  });
  return records;
}

// ── READ: Performance analytics ────────────────────────────
function getPerformanceData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = ss.getSheetByName('Settings');
  var tatTarget = settings ? settings.getRange('C38').getValue() : TAT_TARGET_DAYS;
  var memberPerf = [];
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var total = 0, completed = 0, inProgress = 0, breaches = 0, tatList = [];
    sh.getRange(8, 1, 20, 23).getValues().forEach(function(row) {
      if (!row[0]) return;
      total++;
      if (row[22] === 'Completed') { completed++; if (typeof row[19] === 'number' && row[19] > 0) tatList.push(row[19]); }
      if (row[22] === 'In Progress' || row[22] === 'Docs Pending' || row[22] === 'Not Started') inProgress++;
      if (row[20] === '⚠ Breach') breaches++;
    });
    var avgTat = tatList.length ? +(tatList.reduce(function(a, b) { return a + b; }, 0) / tatList.length).toFixed(1) : null;
    var slaPct = tatList.length ? Math.round(tatList.filter(function(d) { return d <= tatTarget; }).length / tatList.length * 100) : null;
    memberPerf.push({ name: name, totalCases: total, completed: completed, inProgress: inProgress, breaches: breaches, avgTat: avgTat, slaPct: slaPct });
  });
  return { tatTarget: tatTarget, members: memberPerf };
}

// ── READ: Settings/config ──────────────────────────────────
function getSettingsData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = ss.getSheetByName('Settings');
  var tatTarget = settings ? settings.getRange('C38').getValue() : TAT_TARGET_DAYS;
  return { tatTarget: tatTarget, team: TEAM, verticals: VERTICALS_CONFIG, stagesMpEpr: STAGES_MP_EPR, stagesOmp: STAGES_OMP };
}

// ── WRITE: Save / update OB case (cols A-R, leaves S-W for formulas) ──
function saveCaseData(ownerName, rowIndex, data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ownerName);
  if (!sh) throw new Error('Sheet not found: ' + ownerName);
  var target = rowIndex;
  if (!target) {
    var col = sh.getRange(8, 1, 20, 1).getValues();
    for (var i = 0; i < col.length; i++) { if (!col[i][0]) { target = 8 + i; break; } }
    if (!target) throw new Error('No empty row available in range 8–27.');
  }
  var docPctVal = (data.docPct !== '' && data.docPct !== undefined) ? parseFloat(data.docPct) / 100 : 0;
  sh.getRange(target, 1, 1, 18).setValues([[
    data.caseId || '', data.vendor || '', data.vertical || '',
    parseInputDate(data.dateInit), docPctVal,
    data.docValid || '', parseInputDate(data.docValidDate),
    data.stage2 || '', parseInputDate(data.stage2Date),
    data.stage3 || '', parseInputDate(data.stage3Date),
    data.stage4 || '', parseInputDate(data.stage4Date),
    data.decision || '', data.mom || '', parseInputDate(data.momDate),
    data.l1Approval || '', parseInputDate(data.l1Date)
  ]]);
  SpreadsheetApp.flush();
  return { rowIndex: target, success: true };
}

// ── WRITE: Save / update monitoring record ─────────────────
function saveMonitoringRecord(ownerName, rowIndex, data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ownerName);
  if (!sh) throw new Error('Sheet not found: ' + ownerName);
  var target = rowIndex;
  if (!target) {
    var col = sh.getRange(32, 1, 15, 1).getValues();
    for (var i = 0; i < col.length; i++) { if (!col[i][0]) { target = 32 + i; break; } }
    if (!target) throw new Error('No empty row in monitoring range 32–46.');
  }
  sh.getRange(target, 1, 1, 9).setValues([[
    data.seller || '', data.vertical || '', data.month || '',
    data.dataCollected || '', parseInputDate(data.dataDate),
    data.followUp || '', parseInputDate(data.followUpDate),
    data.outcome || '', parseInputDate(data.presentDate)
  ]]);
  sh.getRange(target, 11).setValue(data.remarks || '');
  SpreadsheetApp.flush();
  return { rowIndex: target, success: true };
}

// ── WRITE: Save / update third-party record ────────────────
function saveThirdPartyRecord(ownerName, rowIndex, data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ownerName);
  if (!sh) throw new Error('Sheet not found: ' + ownerName);
  var target = rowIndex;
  if (!target) {
    var col = sh.getRange(51, 1, 15, 1).getValues();
    for (var i = 0; i < col.length; i++) { if (!col[i][0]) { target = 51 + i; break; } }
    if (!target) throw new Error('No empty row in 3P range 51–65.');
  }
  sh.getRange(target, 1, 1, 9).setValues([[
    data.tpName || '', data.caseRef || '', data.vertical || '',
    parseInputDate(data.assignedDate), data.reportRcvd || '',
    parseInputDate(data.reportDate), data.billRcvd || '',
    data.billValid || '', parseInputDate(data.billDate)
  ]]);
  sh.getRange(target, 11).setValue(data.remarks || '');
  SpreadsheetApp.flush();
  return { rowIndex: target, success: true };
}

// ── WRITE: Save / update task record ──────────────────────
function saveTaskRecord(ownerName, rowIndex, data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ownerName);
  if (!sh) throw new Error('Sheet not found: ' + ownerName);
  var target = rowIndex;
  if (!target) {
    var col = sh.getRange(70, 1, 20, 1).getValues();
    for (var i = 0; i < col.length; i++) { if (!col[i][0]) { target = 70 + i; break; } }
    if (!target) throw new Error('No empty row in tasks range 70–89.');
  }
  sh.getRange(target, 1, 1, 5).setValues([[
    data.description || '', data.category || '',
    parseInputDate(data.dateAssigned), parseInputDate(data.targetDate), parseInputDate(data.actualDate)
  ]]);
  sh.getRange(target, 8).setValue(data.remarks || '');
  SpreadsheetApp.flush();
  return { rowIndex: target, success: true };
}

// ── READ: Generate next sequential Case ID ─────────────────
function getNextCaseId() {
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var year = new Date().getFullYear();
  var pref = 'OB-' + year + '-';
  var max  = 0;
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.getRange(8, 1, 20, 1).getValues().forEach(function(row) {
      var id = String(row[0] || '');
      if (id.indexOf(pref) === 0) {
        var n = parseInt(id.substring(pref.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  });
  var next   = max + 1;
  var padded = (next < 10 ? '000' : next < 100 ? '00' : next < 1000 ? '0' : '') + next;
  return pref + padded;
}
