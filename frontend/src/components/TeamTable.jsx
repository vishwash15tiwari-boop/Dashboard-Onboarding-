export default function TeamTable({ members }) {
  const totals = members.reduce(
    (acc, m) => ({
      obOpen:      acc.obOpen      + m.obOpen,
      obDone:      acc.obDone      + m.obDone,
      docsPending: acc.docsPending + m.docsPending,
      tatBreaches: acc.tatBreaches + m.tatBreaches,
      withinTat:   acc.withinTat   + m.withinTat,
    }),
    { obOpen: 0, obDone: 0, docsPending: 0, tatBreaches: 0, withinTat: 0 }
  );

  const teamPct = (totals.withinTat + totals.tatBreaches) > 0
    ? Math.round(totals.withinTat / (totals.withinTat + totals.tatBreaches) * 100)
    : null;

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Team Member</th>
            <th>OB Open</th>
            <th>OB Done</th>
            <th>Docs Pending</th>
            <th>TAT Breaches</th>
            <th>Within TAT</th>
            <th>Avg TAT (days)</th>
            <th>% Within TAT</th>
          </tr>
        </thead>
        <tbody>
          {members.map((m) => (
            <tr key={m.name}>
              <td>{m.name}</td>
              <td>{m.obOpen}</td>
              <td>{m.obDone}</td>
              <td style={m.docsPending > 0 ? { color: '#bf360c', fontWeight: 700 } : {}}>
                {m.docsPending}
              </td>
              <td style={m.tatBreaches > 0 ? { color: '#b71c1c', fontWeight: 700 } : {}}>
                {m.tatBreaches}
              </td>
              <td style={{ color: '#1b5e20' }}>{m.withinTat}</td>
              <td>{m.avgTat != null ? m.avgTat : '—'}</td>
              <td>{m.pctWithinTat != null ? m.pctWithinTat + '%' : '—'}</td>
            </tr>
          ))}
          <tr className="total-row">
            <td>TOTAL / TEAM</td>
            <td>{totals.obOpen}</td>
            <td>{totals.obDone}</td>
            <td>{totals.docsPending}</td>
            <td>{totals.tatBreaches}</td>
            <td>{totals.withinTat}</td>
            <td>—</td>
            <td>{teamPct != null ? teamPct + '%' : '—'}</td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}
