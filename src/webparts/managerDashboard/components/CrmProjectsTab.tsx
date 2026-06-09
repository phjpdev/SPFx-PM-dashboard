import * as React from 'react';
import { loadProjectsFromSharePoint, saveProjectsToSharePoint } from './crmStorage';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmProject, CrmRfqDiscipline } from './crmTypes';

const FF = 'Montserrat,sans-serif';
const LS_PROJECTS = '3edge-crm-projects';

const C = {
  surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', muted: '#8a97a8', green: '#2a9e2a',
  purple: '#6c3fbf', thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const fmtShortDate = (iso: string): string => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: '2-digit' });
};

const fmtMoney = (n: number): string =>
  '$' + Math.round(n).toLocaleString('en-AU');

const disciplineBadgeStyle = (d: 'Steel' | 'Concrete'): React.CSSProperties => ({
  display: 'inline-block', padding: '1px 6px', borderRadius: 3, fontSize: 9, fontWeight: 700,
  letterSpacing: '.05em', textTransform: 'uppercase', fontFamily: FF,
  background: d === 'Concrete' ? 'rgba(107,79,200,0.12)' : 'rgba(37,99,235,0.12)',
  color: d === 'Concrete' ? '#6b4fc8' : '#2563eb',
  border: `1px solid ${d === 'Concrete' ? '#6b4fc8' : '#2563eb'}`,
});

const DisciplineBadges: React.FC<{ discipline: CrmRfqDiscipline }> = ({ discipline }) => {
  if (discipline === 'Both') {
    return (
      <span style={{ display: 'flex', gap: 3, flexWrap: 'wrap' }}>
        <span style={disciplineBadgeStyle('Steel')}>STEEL</span>
        <span style={disciplineBadgeStyle('Concrete')}>CONCRETE</span>
      </span>
    );
  }
  const kind = discipline === 'Concrete' ? 'Concrete' : 'Steel';
  return <span style={disciplineBadgeStyle(kind)}>{kind.toUpperCase()}</span>;
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

const CrmProjectsTab: React.FC<{
  spService: SharePointService;
  persons: CrmPerson[];
  companies: CrmCompany[];
  seedProjects?: CrmProject[] | null;
  onSeedApplied?: () => void;
}> = ({ spService, persons, companies, seedProjects, onSeedApplied }) => {
  const [projects, setProjects] = React.useState<CrmProject[]>([]);
  const [ready, setReady] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const projectsRef = React.useRef(projects);
  projectsRef.current = projects;
  const saveTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const reload = React.useCallback(async (): Promise<void> => {
    try {
      const data = await loadProjectsFromSharePoint(spService);
      const remoteStr = JSON.stringify(data);
      const localStr = JSON.stringify(projectsRef.current);
      if (remoteStr !== localStr) setProjects(data);
    } catch {
      try {
        const v = localStorage.getItem(LS_PROJECTS);
        if (v) setProjects(JSON.parse(v) as CrmProject[]);
      } catch { /* ignore */ }
    }
  }, [spService]);

  React.useEffect(() => {
    if (seedProjects?.length) {
      setProjects(seedProjects);
      setReady(true);
      onSeedApplied?.();
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const data = await loadProjectsFromSharePoint(spService);
        if (!cancelled) setProjects(data);
      } catch {
        if (!cancelled) {
          try {
            const v = localStorage.getItem(LS_PROJECTS);
            if (v) setProjects(JSON.parse(v) as CrmProject[]);
          } catch { /* ignore */ }
        }
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [spService, seedProjects]); // eslint-disable-line react-hooks/exhaustive-deps

  React.useEffect(() => {
    if (!ready) return;
    const iv = setInterval(() => { void reload(); }, 12000);
    return () => clearInterval(iv);
  }, [ready, reload]);

  React.useEffect(() => {
    if (!ready) return;
    localStorage.setItem(LS_PROJECTS, JSON.stringify(projects));
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      void saveProjectsToSharePoint(spService, projectsRef.current).catch(() => undefined);
    }, 2000);
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    };
  }, [projects, ready, spService]);

  const companyName = (id: string): string => companies.find(c => c.id === id)?.name || '—';
  const personName = (id: string): string => persons.find(p => p.id === id)?.name || '—';

  const year = new Date().getFullYear();
  const yearProjects = projects.filter(p => p.wonDate.startsWith(String(year)));
  const totalValue = yearProjects.reduce((s, p) => s + (p.projectValue || 0), 0);

  const q = search.toLowerCase();
  const filtered = yearProjects.filter(p => {
    if (!q) return true;
    return (
      p.projNum.toLowerCase().includes(q) ||
      p.rfqNum.toLowerCase().includes(q) ||
      p.quoteNum.toLowerCase().includes(q) ||
      p.projectTitle.toLowerCase().includes(q) ||
      companyName(p.organizationId).toLowerCase().includes(q)
    );
  });

  if (!ready) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading projects…</div>;
  }

  return (
    <>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Won Projects" accent={C.green} value={String(yearProjects.length)} sub={`${year} year`} />
        <KpiCard label="Won Value" accent="#0d9488" value={fmtMoney(totalValue)} sub="from CRM quotes" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, paddingBottom: 12 }}>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search projects…"
          style={{ padding: '7px 12px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, width: 200, outline: 'none', boxSizing: 'border-box' }}
        />
      </div>

      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 960 }}>
          <thead>
            <tr>
              {['Project #', 'Quote #', 'RFQ #', 'Project', 'Company', 'Contact', 'Type', 'Won', 'Est Value', 'Hours', 'Owner'].map(h => (
                <th key={h} style={{ padding: '9px 10px', textAlign: 'left', fontFamily: FF, fontWeight: 700, fontSize: 10, letterSpacing: '.06em', textTransform: 'uppercase', color: C.sub, background: C.thBg, borderBottom: `2px solid ${C.borderMd}`, whiteSpace: 'nowrap' }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={11} style={{ padding: 48, textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted }}>
                  {projects.length === 0
                    ? 'No won projects yet — mark a quote as WON to create a project here.'
                    : 'No results match your search.'}
                </td>
              </tr>
            ) : filtered.map(p => (
              <tr
                key={p.id}
                onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
              >
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.green, borderBottom: `1px solid ${C.border}` }}>{p.projNum}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.purple, borderBottom: `1px solid ${C.border}` }}>{p.quoteNum || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: '#0d9488', borderBottom: `1px solid ${C.border}` }}>{p.rfqNum}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{p.projectTitle || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{companyName(p.organizationId)}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{personName(p.personId)}</td>
                <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                  <DisciplineBadges discipline={p.discipline} />
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{fmtShortDate(p.wonDate)}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{p.projectValue ? fmtMoney(p.projectValue) : '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{p.approximateHours || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{p.assignedTo || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg }}>
          {filtered.length} of {yearProjects.length} project{yearProjects.length !== 1 ? 's' : ''} ({year})
        </div>
      </div>
    </>
  );
};

export default CrmProjectsTab;
