const ALERT_ITEMS = [
  { key: 'breach',         icon: '🔴', label: 'Cases currently breaching TAT',            bg: '#ffebee', fg: '#b71c1c' },
  { key: 'docs',           icon: '🟡', label: 'Cases with incomplete docs (TAT paused)',   bg: '#fbe9e7', fg: '#bf360c' },
  { key: 'tpOpen',         icon: '🟣', label: '3rd Party tasks still open (all team)',     bg: '#f3e5f5', fg: '#4a148c' },
  { key: 'overdue',        icon: '🟠', label: 'Other tasks marked Overdue',                bg: '#fbe9e7', fg: '#bf360c' },
  { key: 'monitoringOpen', icon: '🔵', label: 'Monitoring tasks not yet closed',           bg: '#e0f2f1', fg: '#004d40' },
];

export default function AlertsPanel({ alerts }) {
  return (
    <div className="alerts-grid">
      {ALERT_ITEMS.map(({ key, icon, label, bg, fg }) => (
        <div key={key} className="alert-row" style={{ background: bg, color: fg }}>
          <span className="alert-icon">{icon}</span>
          <span style={{ flex: 1 }}>{label}</span>
          <span className="alert-count">{alerts[key] ?? 0}</span>
        </div>
      ))}
    </div>
  );
}
