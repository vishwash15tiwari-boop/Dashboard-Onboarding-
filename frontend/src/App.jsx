import { useState, useEffect, useCallback } from 'react';
import KPIGrid       from './components/KPIGrid.jsx';
import TeamTable     from './components/TeamTable.jsx';
import PipelineTable from './components/PipelineTable.jsx';
import AlertsPanel   from './components/AlertsPanel.jsx';

// ─── Replace with your deployed GAS Web App URL ───────────────
const API_URL = 'YOUR_GAS_WEB_APP_URL';
// ─────────────────────────────────────────────────────────────

const AUTO_REFRESH_MS = 60_000; // refresh every 60 seconds

function formatTime(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' });
}

export default function App() {
  const [data,    setData]    = useState(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState(null);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(API_URL);
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const json = await res.json();
      if (json.error) throw new Error(json.error);
      setData(json);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchData();
    const id = setInterval(fetchData, AUTO_REFRESH_MS);
    return () => clearInterval(id);
  }, [fetchData]);

  return (
    <>
      <header className="app-header">
        <h1>ONBOARDING OPERATIONS PLATFORM — CONTROL TOWER</h1>
        <div className="header-meta">
          {data && <span>Updated {formatTime(data.refreshedAt)}</span>}
          <button className="refresh-btn" onClick={fetchData} disabled={loading}>
            {loading ? 'Loading…' : '↻ Refresh'}
          </button>
        </div>
      </header>

      <main className="page">
        {loading && !data && (
          <div className="state-center">
            <div className="spinner" />
            <span>Loading dashboard data…</span>
          </div>
        )}

        {error && (
          <div className="error-box">
            ⚠ Could not load data: {error}
            {API_URL === 'YOUR_GAS_WEB_APP_URL' && (
              <div style={{ marginTop: 8, fontWeight: 600 }}>
                Replace <code>YOUR_GAS_WEB_APP_URL</code> in <code>src/App.jsx</code> with
                your deployed GAS Web App URL.
              </div>
            )}
          </div>
        )}

        {data && (
          <>
            {/* KPI Cards */}
            <div className="section">
              <div className="section-title">📊 LIVE KPIs</div>
              <KPIGrid kpi={data.kpi} />
            </div>

            {/* Team Performance */}
            <div className="section">
              <div className="section-title">👥 TEAM PERFORMANCE SUMMARY</div>
              <TeamTable members={data.members} />
            </div>

            {/* Pipeline by Vertical */}
            <div className="section">
              <div className="section-title">📋 PIPELINE BY VERTICAL</div>
              <PipelineTable pipeline={data.pipeline} />
            </div>

            {/* Alerts */}
            <div className="section">
              <div className="section-title">🚨 ALERTS  (action needed)</div>
              <AlertsPanel alerts={data.alerts} />
            </div>

            <div className="footer">
              Live view  •  Auto-refreshes every 60 s  •  Manager: Ajay  •  Team: Harshita · Vamsi · Naveen · Vishwash
            </div>
          </>
        )}
      </main>
    </>
  );
}
