// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// WebApp.gs — Serves the HTML dashboard; getDashboardData()
//             is called client-side via google.script.run
//
// Deploy: Apps Script → Deploy → New deployment
//   Type: Web app
//   Execute as: Me
//   Who has access: Anyone
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Onboarding Operations Platform')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  var settings = ss.getSheetByName('Settings');
  var tatTarget = settings ? settings.getRange('C38').getValue() : TAT_TARGET_DAYS;

  var kpi = {
    totalCases:  0,
    completed:   0,
    inProgress:  0,
    docsPending: 0,
    tatBreaches: 0,
    tatTarget:   tatTarget
  };

  var pipeline = {};
  VERTICALS_CONFIG.forEach(function(v) {
    pipeline[v.name] = { name: v.name, owner: v.owner, total: 0, inProgress: 0, completed: 0, rejected: 0 };
  });

  var alerts = { breach: 0, docs: 0, tpOpen: 0, overdue: 0, monitoringOpen: 0 };
  var memberStats = [];

  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;

    // ── OB Cases (rows 8–27, cols A–W) ──────────────────────
    var obRows = sh.getRange(8, 1, 20, 23).getValues();
    var totalCases = 0, completed = 0, inProgress = 0, docsPending = 0, tatBreaches = 0;

    obRows.forEach(function(row) {
      if (!row[0]) return;
      var vertical   = row[2];   // C
      var tatStatus  = row[20];  // U
      var caseStatus = row[22];  // W

      totalCases++;
      kpi.totalCases++;

      if (caseStatus === 'Completed' || caseStatus === 'Rejected') { completed++;   kpi.completed++;   }
      if (caseStatus === 'In Progress' || caseStatus === 'Docs Pending' || caseStatus === 'Not Started') { inProgress++; kpi.inProgress++; }
      if (tatStatus  === 'Pending Docs')  { docsPending++; kpi.docsPending++; alerts.docs++;   }
      if (tatStatus  === '⚠ Breach')     { tatBreaches++; kpi.tatBreaches++; alerts.breach++; }

      if (pipeline[vertical]) {
        pipeline[vertical].total++;
        if (caseStatus === 'In Progress') pipeline[vertical].inProgress++;
        if (caseStatus === 'Completed')   pipeline[vertical].completed++;
        if (caseStatus === 'Rejected')    pipeline[vertical].rejected++;
      }
    });

    // ── Monitoring status col (J = col 10), rows 32–46 ──────
    var monitoringOpen = 0;
    sh.getRange(32, 10, 15, 1).getValues().forEach(function(row) {
      if (row[0] === 'Open' || row[0] === 'In Progress' || row[0] === 'Started') {
        monitoringOpen++;
        alerts.monitoringOpen++;
      }
    });

    // ── Third Party status col (J = col 10), rows 51–65 ────
    var thirdPartyOpen = 0;
    sh.getRange(51, 10, 15, 1).getValues().forEach(function(row) {
      if (row[0] === 'Open') { thirdPartyOpen++; alerts.tpOpen++; }
    });

    // ── Other Tasks status col (G = col 7), rows 70–89 ─────
    var otherTasks = 0;
    sh.getRange(70, 7, 20, 1).getValues().forEach(function(row) {
      if (row[0] && row[0] !== '') otherTasks++;
      if (row[0] === 'Overdue') alerts.overdue++;
    });

    memberStats.push({
      name:           name,
      totalCases:     totalCases,
      completed:      completed,
      inProgress:     inProgress,
      docsPending:    docsPending,
      tatBreaches:    tatBreaches,
      monitoringOpen: monitoringOpen,
      thirdPartyOpen: thirdPartyOpen,
      otherTasks:     otherTasks
    });
  });

  return {
    kpi:         kpi,
    members:     memberStats,
    pipeline:    Object.values(pipeline),
    alerts:      alerts,
    refreshedAt: new Date().toISOString()
  };
}
