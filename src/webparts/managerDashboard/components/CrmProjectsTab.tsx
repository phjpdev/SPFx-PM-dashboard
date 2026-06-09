import * as React from 'react';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { IProject } from '../../../shared/models/IProject';

const FF = 'Montserrat,sans-serif';

const C = {
  surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', muted: '#8a97a8', green: '#2a9e2a',
  thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const fmtShortDate = (iso: string): string => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: '2-digit' });
};

const disciplineBadge = (d: string): React.CSSProperties => ({
  display: 'inline-block', padding: '1px 6px', borderRadius: 3, fontSize: 9, fontWeight: 700,
  letterSpacing: '.05em', textTransform: 'uppercase', fontFamily: FF,
  background: d === 'Concrete' ? 'rgba(107,79,200,0.12)' : 'rgba(37,99,235,0.12)',
  color: d === 'Concrete' ? '#6b4fc8' : '#2563eb',
  border: `1px solid ${d === 'Concrete' ? '#6b4fc8' : '#2563eb'}`,
});

const statusStyle = (s: string): React.CSSProperties => {
  const u = (s || '').toUpperCase();
  if (u.includes('ACTIVE')) return { background: '#e6f5e8', color: '#1e6b38' };
  if (u.includes('HOLD')) return { background: '#f0e8ff', color: '#5a2d9e' };
  if (u.includes('RFI')) return { background: '#e3f0ff', color: '#1a5fa8' };
  if (u.includes('COMPLETE')) return { background: '#f0f2f6', color: '#4a5568' };
  return { background: '#f0f2f6', color: '#4a5568' };
};

const badge: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, fontFamily: FF, padding: '3px 8px',
  borderRadius: 3, whiteSpace: 'nowrap', display: 'inline-block',
};

const KpiCard: React.FC<{ label: string; value: string; sub: string; accent: string }> = ({ label, value, sub, accent }) => (
  <div style={{ flex: '1 1 140px', background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: 'hidden', minWidth: 120 }}>
    <div style={{ height: 3, background: accent }} />
    <div style={{ padding: '12px 14px' }}>
      <div style={{ fontFamily: FF, fontSize: 9, fontWeight: 700, letterSpacing: '.08em', color: C.muted, textTransform: 'uppercase' }}>{label}</div>
      <div style={{ fontFamily: FF, fontSize: 22, fontWeight: 700, color: C.text, marginTop: 4 }}>{value}</div>
      <div style={{ fontFamily: FF, fontSize: 11, color: C.muted, marginTop: 2 }}>{sub}</div>
    </div>
  </div>
);

/** CRM Projects tab — shows the same live projects as the main Projects dashboard (3Edge_Projects). */
const CrmProjectsTab: React.FC<{
  spService: SharePointService;
  refreshKey?: number;
}> = ({ spService, refreshKey = 0 }) => {
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [ready, setReady] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [year, setYear] = React.useState(new Date().getFullYear());

  const load = React.useCallback(async (): Promise<void> => {
    try {
      const data = await spService.loadProjects();
      setProjects(data);
    } catch { /* ignore */ }
  }, [spService]);

  React.useEffect(() => {
    let cancelled = false;
    setReady(false);
    void (async () => {
      await load();
      if (!cancelled) setReady(true);
    })();
    return () => { cancelled = true; };
  }, [load, refreshKey]);

  React.useEffect(() => {
    if (!ready) return;
    const iv = setInterval(() => { void load(); }, 12000);
    return () => clearInterval(iv);
  }, [ready, load]);

  const projectYear = (p: IProject): number => {
    if (p.year) return p.year;
    if (p.startDate) {
      const d = new Date(p.startDate + 'T00:00:00');
      if (!isNaN(d.getTime())) return d.getFullYear();
    }
    return new Date().getFullYear();
  };
  const mainProjects = projects.filter(p => !p.isEwo && projectYear(p) === year);
  const active = mainProjects.filter(p => (p.status || '').toLowerCase() === 'active').length;
  const totalHrs = mainProjects.reduce((s, p) => s + (p.hrsUsed || 0), 0);

  const q = search.toLowerCase();
  const filtered = mainProjects.filter(p => {
    if (!q) return true;
    return (
      p.projNum.toLowerCase().includes(q) ||
      (p.quoteNum || '').toLowerCase().includes(q) ||
      (p.name || '').toLowerCase().includes(q) ||
      (p.company || '').toLowerCase().includes(q) ||
      (p.contact || '').toLowerCase().includes(q)
    );
  });

  if (!ready) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading projects…</div>;
  }

  return (
    <>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Total Projects" accent={C.green} value={String(mainProjects.length)} sub={`${year} year`} />
        <KpiCard label="Active" accent="#3b82c4" value={String(active)} sub="in progress" />
        <KpiCard label="Total Hrs Used" accent="#0d9488" value={String(Math.round(totalHrs))} sub="all projects" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, paddingBottom: 12, flexWrap: 'wrap' }}>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search projects…"
          style={{ padding: '7px 12px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, width: 200, outline: 'none', boxSizing: 'border-box' }}
        />
        <select
          value={year}
          onChange={e => setYear(Number(e.target.value))}
          style={{ padding: '7px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text }}
        >
          {[year, year - 1, year - 2].map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        <span style={{ fontFamily: FF, fontSize: 11, color: C.muted, marginLeft: 'auto' }}>
          Same data as main Projects tab — WON quotes create a row here automatically.
        </span>
      </div>

      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 960 }}>
          <thead>
            <tr>
              {['Project #', 'Quote #', 'Name', 'Company', 'Contact', 'Hours', 'Start', 'Status', 'Owner'].map(h => (
                <th key={h} style={{ padding: '9px 10px', textAlign: 'left', fontFamily: FF, fontWeight: 700, fontSize: 10, letterSpacing: '.06em', textTransform: 'uppercase', color: C.sub, background: C.thBg, borderBottom: `2px solid ${C.borderMd}`, whiteSpace: 'nowrap' }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={9} style={{ padding: 48, textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted }}>
                  No projects for {year}. Mark a quote as WON to create one, or use + New Project on the main Projects tab.
                </td>
              </tr>
            ) : filtered.map(p => {
              const hrsLeft = (p.hrsAllowed || 0) - (p.hrsUsed || 0);
              const over = hrsLeft < 0;
              return (
                <tr
                  key={p.id}
                  onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                  onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
                >
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.green, borderBottom: `1px solid ${C.border}` }}>{p.projNum}</td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: '#6c3fbf', borderBottom: `1px solid ${C.border}` }}>
                    {p.quoteNum ? (p.quoteNum.startsWith('QU-') ? p.quoteNum : `QU-${p.quoteNum}`) : '—'}
                  </td>
                  <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                    <div style={{ fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text }}>{p.name || '—'}</div>
                    {p.discipline && <span style={disciplineBadge(p.discipline)}>{p.discipline.toUpperCase()}</span>}
                  </td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{p.company || '—'}</td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{p.contact || '—'}</td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}`, whiteSpace: 'nowrap' }}>
                    {p.hrsUsed || 0} / {p.hrsAllowed || 0}h
                    {p.hrsAllowed > 0 && (
                      <span style={{ marginLeft: 4, fontSize: 10, fontWeight: 700, color: over ? '#c0392b' : C.green }}>
                        {over ? `+${Math.abs(Math.round(hrsLeft))}h over` : `${Math.round(hrsLeft)}h left`}
                      </span>
                    )}
                  </td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{fmtShortDate(p.startDate)}</td>
                  <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                    <span style={{ ...badge, ...statusStyle(p.status) }}>{(p.status || 'Active').toUpperCase()}</span>
                  </td>
                  <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{p.teamLead || '—'}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg }}>
          {filtered.length} of {mainProjects.length} project{mainProjects.length !== 1 ? 's' : ''} ({year})
        </div>
      </div>
    </>
  );
};

export default CrmProjectsTab;
