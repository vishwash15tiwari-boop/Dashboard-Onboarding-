// ============================================================
// ONBOARDING OPERATIONS PLATFORM
// Config.gs — All constants, team data, stage weights
// Sheet ID: 1dp2zJ-henn9QZok9qxA1r_4yR0UOPuUdgo3BxDGvf8s
// ============================================================

var SPREADSHEET_ID = '1dp2zJ-henn9QZok9qxA1r_4yR0UOPuUdgo3BxDGvf8s';

var CLR = {
  NAVY:           '#1a237e',
  NAVY_MED:       '#283593',
  NAVY_DARK:      '#0d1b6e',
  NAVY_LIGHT:     '#e8eaf6',
  TEAL:           '#004d40',
  TEAL_LIGHT:     '#e0f2f1',
  GREEN_DARK:     '#1b5e20',
  GREEN_LIGHT:    '#e8f5e9',
  AMBER:          '#bf360c',
  AMBER_LIGHT:    '#fbe9e7',
  RED:            '#b71c1c',
  RED_LIGHT:      '#ffebee',
  PURPLE:         '#4a148c',
  PURPLE_LIGHT:   '#f3e5f5',
  WHITE:          '#ffffff',
  OFF_WHITE:      '#f8f9fa',
  GRAY_LIGHT:     '#eceff1',
  GRAY_MED:       '#b0bec5',
  GRAY_DARK:      '#546e7a',
  CHARCOAL:       '#37474f',
  BLACK:          '#212121',
  // Section-specific accent colors
  SECTION_OB:     '#1a237e',  // Navy — Onboarding
  SECTION_OB_LT:  '#c5cae9',
  SECTION_MON:    '#004d40',  // Teal — Monitoring
  SECTION_MON_LT: '#b2dfdb',
  SECTION_TP:     '#4a148c',  // Purple — Third Party
  SECTION_TP_LT:  '#e1bee7',
  SECTION_OT:     '#bf360c',  // Deep Orange — Other Tasks
  SECTION_OT_LT:  '#ffccbc',
};

var MANAGER_NAME = 'Ajay';

// Team member configuration (includes Ajay for personal sheet creation;
// SHEET_NAMES.MEMBERS stays as 4 so Ajay is excluded from team KPIs)
var TEAM = [
  {
    name:       'Ajay',
    role:       'Manager',
    vertical:   'All (oversight)',
    fallback:   '—',
    thirdParty: '—'
  },
  {
    name:       'Harshita',
    role:       'Executive',
    vertical:   'Marketplace',
    fallback:   'EPR & Sustainability',
    thirdParty: 'ONGRID, RUBIX'
  },
  {
    name:       'Vamsi',
    role:       'Senior Executive',
    vertical:   'Open Marketplace',
    fallback:   'Marketplace',
    thirdParty: 'Finoscale'
  },
  {
    name:       'Naveen',
    role:       'Executive',
    vertical:   'EPR & Sustainability',
    fallback:   'Open Marketplace',
    thirdParty: '—'
  },
  {
    name:       'Vishwash',
    role:       'Management Trainee',
    vertical:   'SOPs / MIS / Tracker',
    fallback:   'All verticals',
    thirdParty: '—'
  }
];

// Vertical → ownership mapping
var VERTICALS_CONFIG = [
  { name: 'Marketplace',        owner: 'Harshita', fallback: 'Vamsi'    },
  { name: 'Open Marketplace',   owner: 'Vamsi',    fallback: 'Naveen'   },
  { name: 'EPR & Sustainability', owner: 'Naveen', fallback: 'Harshita' }
];

// Stage weights — Marketplace & EPR/Sustainability (Settings rows 21–26)
var STAGES_MP_EPR = [
  { order: 1, label: 'Document Validation',            weight: 0.15 },
  { order: 2, label: 'Verification Report',            weight: 0.30 },
  { order: 3, label: 'V-KYC & GST Portal Verification',weight: 0.25 },
  { order: 4, label: 'Committee',                      weight: 0.20 },
  { order: 5, label: 'MOM (Minutes of Meeting)',        weight: 0.05 },
  { order: 6, label: 'L1 Approval',                    weight: 0.05 }
];

// Stage weights — Open Marketplace (Settings rows 31–35)
var STAGES_OMP = [
  { order: 1, label: 'Document Validation',                   weight: 0.15 },
  { order: 2, label: 'Doc sent to 3rd Party for Verification', weight: 0.15 },
  { order: 3, label: 'Rating Report Received',                weight: 0.30 },
  { order: 4, label: 'Review the Rating Report',              weight: 0.30 },
  { order: 5, label: 'L1 Approval',                          weight: 0.10 }
];

var TAT_TARGET_DAYS = 3; // working days — stored in Settings!C38

// Dropdown options
var YN_OPTIONS         = ['Y', 'N'];
var VERTICAL_OPTIONS   = ['Marketplace', 'Open Marketplace', 'EPR & Sustainability'];
var MONITORING_STATUS  = ['Open', 'Started', 'In Progress', 'Closed'];
var TP_STATUS          = ['Open', 'Report Received', 'Bill Received', 'Closed'];
var TASK_STATUS        = ['Open', 'Overdue', 'Closed'];

// Sheet names
var SHEET_NAMES = {
  DASHBOARD: 'Control Tower',
  SETTINGS:  'Settings',
  MEMBERS:   ['Harshita', 'Vamsi', 'Naveen', 'Vishwash']
};

// Users — email, password (change as needed), role, linked sheet
// role: 'admin' → full access (Control Tower + Settings + all workspaces)
// role: 'member' → own workspace only
var USERS = [
  { name: 'Ajay',     email: 'ajay.vunyale@recykal.com',         password: 'Ajay2024',     role: 'admin',  sheet: 'Ajay' },
  { name: 'Harshita', email: 'harshita.shukla@recykal.com',      password: 'Harshita2024', role: 'member', sheet: 'Harshita' },
  { name: 'Vamsi',    email: 'vamsi.ayila@recykal.com',           password: 'Vamsi2024',    role: 'member', sheet: 'Vamsi' },
  { name: 'Naveen',   email: 'paravada.naveenranga@recykal.com', password: 'Naveen2024',   role: 'member', sheet: 'Naveen' },
  { name: 'Vishwash', email: 'vishwash.tiwari@recykal.com',       password: 'Vishwash2024', role: 'admin',  sheet: 'Vishwash' }
];
