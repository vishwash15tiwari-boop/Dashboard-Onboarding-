export default function PipelineTable({ pipeline }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Vertical</th>
            <th>Owner</th>
            <th>Total Cases</th>
            <th>In Progress</th>
            <th>Completed</th>
            <th>Rejected</th>
          </tr>
        </thead>
        <tbody>
          {pipeline.map((v) => (
            <tr key={v.name}>
              <td>{v.name}</td>
              <td style={{ color: '#546e7a', fontWeight: 500 }}>{v.owner}</td>
              <td>{v.total}</td>
              <td style={{ color: '#283593' }}>{v.inProgress}</td>
              <td style={{ color: '#1b5e20' }}>{v.completed}</td>
              <td style={v.rejected > 0 ? { color: '#b71c1c' } : {}}>{v.rejected}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
