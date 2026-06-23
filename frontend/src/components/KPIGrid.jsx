export default function KPIGrid({ kpi }) {
  const cards = [
    { label: 'Total\nOB Cases',             value: kpi.totalCases,  bg: '#1a237e' },
    { label: 'OB\nCompleted',               value: kpi.completed,   bg: '#004d40' },
    { label: 'OB\nIn Progress',             value: kpi.inProgress,  bg: '#283593' },
    { label: 'Docs Incomplete\n(TAT paused)',value: kpi.docsPending, bg: '#bf360c' },
    { label: 'TAT\nBreaches',               value: kpi.tatBreaches, bg: '#b71c1c' },
    { label: 'TAT Target\n(days)',           value: kpi.tatTarget,   bg: '#37474f' },
  ];

  return (
    <div className="kpi-grid">
      {cards.map((c) => (
        <div key={c.label} className="kpi-card" style={{ background: c.bg, color: '#fff' }}>
          <div className="kpi-label" style={{ whiteSpace: 'pre-line' }}>{c.label}</div>
          <div className="kpi-value">{c.value ?? '—'}</div>
        </div>
      ))}
    </div>
  );
}
