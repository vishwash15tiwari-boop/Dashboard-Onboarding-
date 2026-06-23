// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// WebApp.gs — JSON API endpoint for the React dashboard
//
// Deploy: Apps Script → Deploy → New deployment
//   Type: Web app
//   Execute as: Me
//   Who has access: Anyone
// ============================================================

function doGet(e) {
  try {
    var data = getDashboardData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
    var obOpen = 0, obDone = 0, docsPend = 0, breaches = 0, withinTat = 0;
    var tatSum = 0, tatCount = 0;

    obRows.forEach(function(row) {
      if (!row[0]) return;                // A empty → skip
      var vertical   = row[2];            // C
      var tatDays    = row[19];           // T
      var tatStatus  = row[20];           // U
      var caseStatus = row[22];           // W

      kpi.totalCases++;
      if (caseStatus === 'In Progress')                     { obOpen++;   kpi.inProgress++;  }
      if (caseStatus === 'Completed' || caseStatus === 'Rejected') { obDone++; kpi.completed++; }
      if (tatStatus  === 'Pending Docs')                    { docsPend++; kpi.docsPending++; alerts.docs++;    }
      if (tatStatus  === '⚠ Breach')                       { breaches++; kpi.tatBreaches++; alerts.breach++;  }
      if (tatStatus  === '✓ Within TAT')                   { withinTat++; }
      if (typeof tatDays === 'number' && tatDays > 0)       { tatSum += tatDays; tatCount++; }

      if (pipeline[vertical]) {
        pipeline[vertical].total++;
        if (caseStatus === 'In Progress') pipeline[vertical].inProgress++;
        if (caseStatus === 'Completed')   pipeline[vertical].completed++;
        if (caseStatus === 'Rejected')    pipeline[vertical].rejected++;
      }
    });

    // ── Monitoring status col (J = col 10), rows 32–46 ──────
    sh.getRange(32, 10, 15, 1).getValues().forEach(function(row) {
      var s = row[0];
      if (s === 'Open' || s === 'In Progress' || s === 'Started') alerts.monitoringOpen++;
    });

    // ── Third Party status col (J = col 10), rows 51–65 ─────
    sh.getRange(51, 10, 15, 1).getValues().forEach(function(row) {
      if (row[0] === 'Open') alerts.tpOpen++;
    });

    // ── Other Tasks status col (G = col 7), rows 70–89 ──────
    sh.getRange(70, 7, 20, 1).getValues().forEach(function(row) {
      if (row[0] === 'Overdue') alerts.overdue++;
    });

    memberStats.push({
      name:        name,
      obOpen:      obOpen,
      obDone:      obDone,
      docsPending: docsPend,
      tatBreaches: breaches,
      withinTat:   withinTat,
      avgTat:      tatCount > 0 ? Math.round((tatSum / tatCount) * 10) / 10 : null,
      pctWithinTat:(withinTat + breaches) > 0
                     ? Math.round(withinTat / (withinTat + breaches) * 100)
                     : null
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
