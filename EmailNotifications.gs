// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// EmailNotifications.gs — Individual + summary email dispatch
//
// Setup: Apps Script → Triggers → Add Trigger
//   Function: sendDailyNotifications
//   Event: Time-driven → Day timer (e.g. 8am–9am)
//   (or call manually from the menu)
// ============================================================

// ── Main entry point (set as a trigger) ───────────────────
function sendDailyNotifications() {
  var dash = getDashboardData();
  var perf = getPerformanceData();

  // Individual email to each team member
  SHEET_NAMES.MEMBERS.forEach(function(name) {
    var memberStat = (dash.members || []).filter(function(m) { return m.name === name; })[0];
    var memberPerf = (perf.members || []).filter(function(m) { return m.name === name; })[0];
    var email      = _emailFor(name);
    if (!email || !memberStat) return;

    var subject = '[OB Update] ' + name + ' — ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var html = _memberBody(name, memberStat, memberPerf, perf.tatTarget);

    GmailApp.sendEmail(email, subject, _plain(html), {
      htmlBody: html,
      name: 'Onboarding Operations Platform'
    });
  });

  // Summary to admins (Ajay + Vishwash)
  var adminEmails = USERS.filter(function(u) { return u.role === 'admin'; })
                         .map(function(u) { return u.email; });
  if (adminEmails.length) {
    var subject2 = '[OB Summary] Team Dashboard — ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
    var html2 = _summaryBody(dash, perf);
    adminEmails.forEach(function(e) {
      GmailApp.sendEmail(e, subject2, _plain(html2), {
        htmlBody: html2,
        name: 'Onboarding Operations Platform'
      });
    });
  }
}

// ── Test helpers — call manually to preview emails ────────
function testMemberEmail()  {
  var dash = getDashboardData(), perf = getPerformanceData();
  var stat = (dash.members||[])[0], p = (perf.members||[])[0];
  Logger.log(_memberBody(stat ? stat.name : 'Harshita', stat||{}, p||{}, perf.tatTarget||3));
}
function testSummaryEmail() {
  Logger.log(_summaryBody(getDashboardData(), getPerformanceData()));
}

// ── Private: find email for a team member name ────────────
function _emailFor(name) {
  for (var i = 0; i < USERS.length; i++) {
    if (USERS[i].name === name) return USERS[i].email;
  }
  return null;
}

function _plain(html) {
  return html.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}

// ── Email wrapper ─────────────────────────────────────────
function _wrap(content) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
    'body{font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:20px;}' +
    '.w{max-width:620px;margin:0 auto;background:#fff;border-radius:8px;overflow:hidden;' +
    'box-shadow:0 2px 8px rgba(0,0,0,.1);}' +
    '.hdr{background:#1a237e;color:#fff;padding:18px 24px;font-size:15px;font-weight:700;}' +
    '.sub{background:#37474f;color:#fff;padding:8px 24px;font-size:11px;}' +
    '.bdy{padding:20px 24px;}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px;}' +
    'th{background:#eceff1;padding:7px 10px;text-align:left;font-weight:700;font-size:11px;' +
    'border-bottom:2px solid #b0bec5;}' +
    'td{padding:7px 10px;border-bottom:1px solid #f5f5f5;}' +
    '.n{text-align:center;font-weight:700;}' +
    '.r{color:#b71c1c;} .g{color:#1b5e20;} .o{color:#e65100;} .b{color:#0277bd;}' +
    '.sec{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;' +
    'color:#37474f;margin:14px 0 6px;padding-bottom:4px;border-bottom:2px solid #e0e0e0;}' +
    '.ftr{background:#f5f5f5;padding:10px 24px;font-size:10px;color:#90a4ae;text-align:center;}' +
    '</style></head><body><div class="w">' + content + '</div></body></html>';
}

// ── Individual member email ────────────────────────────────
function _memberBody(name, stat, perf, tatTarget) {
  var dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  var slaCls = perf && perf.slaPct !== null ? (perf.slaPct >= 90 ? 'g' : perf.slaPct >= 70 ? 'o' : 'r') : '';

  var body =
    '<div class="hdr">&#128101; ' + name + ' — Your Onboarding Dashboard Update</div>' +
    '<div class="sub">Generated: ' + dt + ' &nbsp;|&nbsp; TAT Target: ' + (tatTarget||3) + ' working days</div>' +
    '<div class="bdy">' +

    '<div class="sec">OB Cases</div>' +
    '<table><tr><th>Metric</th><th class="n">Count</th></tr>' +
    '<tr><td>Total Cases</td><td class="n b">'      + (stat.totalCases  ||0) + '</td></tr>' +
    '<tr><td>Completed</td><td class="n g">'          + (stat.completed   ||0) + '</td></tr>' +
    '<tr><td>In Progress</td><td class="n b">'        + (stat.inProgress  ||0) + '</td></tr>' +
    '<tr><td>Docs Incomplete</td><td class="n o">'   + (stat.docsPending ||0) + '</td></tr>' +
    '<tr><td>TAT Breaches &#9888;</td><td class="n r">' + (stat.tatBreaches||0) + '</td></tr>' +
    '</table>' +

    '<div class="sec">Open Work Items</div>' +
    '<table><tr><th>Section</th><th class="n">Open</th></tr>' +
    '<tr><td>Monitoring Records</td><td class="n o">' + (stat.monitoringOpen  ||0) + '</td></tr>' +
    '<tr><td>Third Party</td><td class="n o">'        + (stat.thirdPartyOpen  ||0) + '</td></tr>' +
    '<tr><td>Tasks (Overdue)</td><td class="n r">'    + (stat.otherTasks      ||0) + '</td></tr>' +
    '</table>';

  if (perf) {
    body +=
      '<div class="sec">Your Performance</div>' +
      '<table><tr><th>Metric</th><th class="n">Value</th></tr>' +
      '<tr><td>Avg TAT (completed cases)</td>' +
        '<td class="n">' + (perf.avgTat !== null && perf.avgTat !== undefined ? perf.avgTat + 'd' : '—') + '</td></tr>' +
      '<tr><td>SLA Compliance</td>' +
        '<td class="n ' + slaCls + '">' + (perf.slaPct !== null && perf.slaPct !== undefined ? perf.slaPct + '%' : '—') + '</td></tr>' +
      '</table>';
  }

  body +=
    '<p style="font-size:11px;color:#546e7a;margin-top:8px;">Log in to the ' +
    '<strong>Onboarding Operations Platform</strong> to view details and take action.</p>' +
    '</div>' +
    '<div class="ftr">Onboarding Operations Platform &bull; Automated notification &bull; Do not reply</div>';

  return _wrap(body);
}

// ── Admin summary email ───────────────────────────────────
function _summaryBody(dash, perf) {
  var dt      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, hh:mm a');
  var kpi     = dash.kpi     || {};
  var members = dash.members || [];
  var alerts  = dash.alerts  || {};

  var mRows = '';
  members.forEach(function(m) {
    mRows +=
      '<tr><td><strong>' + m.name + '</strong></td>' +
      '<td class="n">'    + (m.totalCases    ||0) + '</td>' +
      '<td class="n g">'  + (m.completed     ||0) + '</td>' +
      '<td class="n b">'  + (m.inProgress    ||0) + '</td>' +
      '<td class="n o">'  + (m.docsPending   ||0) + '</td>' +
      '<td class="n r">'  + (m.tatBreaches   ||0) + '</td>' +
      '<td class="n">'    + (m.monitoringOpen ||0) + '</td>' +
      '<td class="n">'    + (m.thirdPartyOpen ||0) + '</td>' +
      '</tr>';
  });

  var body =
    '<div class="hdr">&#128202; Team Summary — Onboarding Operations</div>' +
    '<div class="sub">Generated: ' + dt + ' &nbsp;|&nbsp; TAT Target: ' + (perf.tatTarget||3) + ' working days</div>' +
    '<div class="bdy">' +

    '<div class="sec">Team KPIs</div>' +
    '<table><tr><th>Total Cases</th><th>Completed</th><th>In Progress</th>' +
    '<th>Docs Pending</th><th>TAT Breaches</th></tr>' +
    '<tr>' +
    '<td class="n b">'  + (kpi.totalCases  ||0) + '</td>' +
    '<td class="n g">'  + (kpi.completed   ||0) + '</td>' +
    '<td class="n b">'  + (kpi.inProgress  ||0) + '</td>' +
    '<td class="n o">'  + (kpi.docsPending ||0) + '</td>' +
    '<td class="n r">'  + (kpi.tatBreaches ||0) + '</td>' +
    '</tr></table>' +

    '<div class="sec">Member Breakdown</div>' +
    '<table><tr><th>Member</th><th class="n">Total</th><th class="n">Done</th>' +
    '<th class="n">In Prog</th><th class="n">Docs&#9888;</th><th class="n">TAT&#9888;</th>' +
    '<th class="n">Monitor</th><th class="n">3P</th></tr>' +
    mRows + '</table>' +

    '<div class="sec">Alerts</div>' +
    '<table><tr><th>Alert</th><th class="n">Count</th></tr>' +
    '<tr><td>TAT Breaches</td><td class="n r">'       + (alerts.breach         ||0) + '</td></tr>' +
    '<tr><td>Docs Incomplete</td><td class="n o">'   + (alerts.docs            ||0) + '</td></tr>' +
    '<tr><td>Third Party Open</td><td class="n o">'  + (alerts.tpOpen          ||0) + '</td></tr>' +
    '<tr><td>Overdue Tasks</td><td class="n r">'      + (alerts.overdue         ||0) + '</td></tr>' +
    '<tr><td>Monitoring Open</td><td class="n b">'   + (alerts.monitoringOpen  ||0) + '</td></tr>' +
    '</table>' +

    '<p style="font-size:11px;color:#546e7a;">Log in to the ' +
    '<strong>Onboarding Operations Platform</strong> for full details.</p>' +
    '</div>' +
    '<div class="ftr">Onboarding Operations Platform &bull; Admin Summary &bull; Do not reply</div>';

  return _wrap(body);
}
