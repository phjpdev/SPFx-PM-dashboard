import * as React from 'react';
import { loadRfqsFromSharePoint, saveRfqsToSharePoint } from './crmStorage';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmRfq, CrmRfqDiscipline, CrmRfqStage } from './crmTypes';

const FF = 'Montserrat,sans-serif';
const LS_RFQS = '3edge-crm-rfqs';

const C = {
  bg: '#f7f8fa', surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', muted: '#8a97a8', green: '#2a9e2a',
  greenBg: 'rgba(42,158,42,.09)', greenBd: 'rgba(42,158,42,.35)',
  red: '#c0392b', purple: '#6c3fbf', thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const STAGES: CrmRfqStage[] = ['New Enquiry', 'Under Review', 'Ready to Quote', 'Won', 'Declined'];
const DISCIPLINES: CrmRfqDiscipline[] = ['Steel', 'Concrete', 'Both'];
const SOURCES = ['Email', 'Phone', 'Referral', 'Repeat Client', 'Website', 'Other'];
const OWNERS = ['MK', 'SK', 'DC', 'JP'];

const mi: React.CSSProperties = {
  padding: '8px 10px', background: C.surface, border: `1px solid ${C.borderMd}`,
  borderRadius: 4, color: C.text, fontSize: 12.5, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: C.sub, letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
const todayIso = (): string => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

const loadRfqsLocal = (): CrmRfq[] => {
  try {
    const v = localStorage.getItem(LS_RFQS);
    return v ? (JSON.parse(v) as CrmRfq[]) : [];
  } catch { return []; }
};

const fmtShortDate = (iso: string): string => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: '2-digit' });
};

const fmtMoney = (n: number): string =>
  '$' + Math.round(n).toLocaleString('en-AU');

const isOverdue = (r: CrmRfq): boolean => {
  if (!r.quoteRequiredBy || r.stage === 'Won' || r.stage === 'Declined') return false;
  return r.quoteRequiredBy < todayIso();
};

const nextRfqNum = (rfqs: CrmRfq[]): string => {
  const year = new Date().getFullYear().toString().slice(-2);
  const prefix = `RFQ-${year}-`;
  const nums = rfqs
    .filter(r => r.rfqNum.startsWith(prefix))
    .map(r => parseInt(r.rfqNum.slice(prefix.length), 10))
    .filter(n => !isNaN(n));
  const next = nums.length ? Math.max(...nums) + 1 : 1;
  return `${prefix}${String(next).padStart(3, '0')}`;
};

const emptyRfq = (rfqs: CrmRfq[]): CrmRfq => ({
  id: uid(),
  rfqNum: nextRfqNum(rfqs),
  dateReceived: todayIso(),
  personId: '',
  organizationId: '',
  projectTitle: '',
  projectAddress: '',
  discipline: 'Steel',
  quoteRequiredBy: '',
  projectValue: 0,
  approximateHours: 0,
  createQuoteXero: false,
  relatedRfqId: '',
  notes: '',
  source: 'Email',
  stage: 'New Enquiry',
  assignedTo: 'MK',
});

const disciplineBadgeStyle = (d: 'Steel' | 'Concrete'): React.CSSProperties => ({
  display: 'inline-block',
  padding: '1px 6px',
  borderRadius: 3,
  fontSize: 9,
  fontWeight: 700,
  letterSpacing: '.05em',
  textTransform: 'uppercase',
  fontFamily: FF,
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

const normalizeRfq = (r: CrmRfq): CrmRfq => ({
  ...r,
  approximateHours: typeof r.approximateHours === 'number' ? r.approximateHours : 0,
});

const stageStyle = (s: CrmRfqStage): React.CSSProperties => {
  const map: Record<CrmRfqStage, { bg: string; color: string }> = {
    'New Enquiry':    { bg: '#e3f0ff', color: '#1a5fa8' },
    'Under Review':   { bg: '#fff3e0', color: '#b36a00' },
    'Ready to Quote': { bg: '#f0e8ff', color: '#5a2d9e' },
    'Won':            { bg: '#e6f5e8', color: '#1e6b38' },
    'Declined':       { bg: '#fde8e8', color: '#a82828' },
  };
  const x = map[s];
  return { background: x.bg, color: x.color };
};

const badge: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, fontFamily: FF, padding: '3px 8px',
  borderRadius: 3, whiteSpace: 'nowrap', display: 'inline-block',
};

const actionBtn: React.CSSProperties = {
  padding: '4px 10px', borderRadius: 3, fontFamily: FF, fontWeight: 700,
  fontSize: 10.5, cursor: 'pointer', width: '100%', boxSizing: 'border-box', textAlign: 'center',
};

// ── Modal ─────────────────────────────────────────────────────────
const RfqModal: React.FC<{
  initial: CrmRfq;
  rfqs: CrmRfq[];
  persons: CrmPerson[];
  companies: CrmCompany[];
  onSave: (r: CrmRfq) => void;
  onClose: () => void;
}> = ({ initial, rfqs, persons, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmRfq>(initial);
  const set = <K extends keyof CrmRfq>(k: K, v: CrmRfq[K]): void => setD(p => ({ ...p, [k]: v }));
  const grid2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' };
  const isNew = !rfqs.some(r => r.id === initial.id);

  const personsForCo = d.organizationId
    ? persons.filter(p => p.organizationId === d.organizationId)
    : persons;

  const onPersonChange = (personId: string): void => {
    const person = persons.find(p => p.id === personId);
    const co = person?.organizationId ? companies.find(c => c.id === person.organizationId) : undefined;
    setD(p => ({
      ...p,
      personId,
      organizationId: person?.organizationId || p.organizationId,
      projectAddress: co?.address || p.projectAddress,
    }));
  };

  const onCompanyChange = (organizationId: string): void => {
    const personStillValid = persons.some(p => p.id === d.personId && p.organizationId === organizationId);
    const co = companies.find(c => c.id === organizationId);
    setD(p => ({
      ...p,
      organizationId,
      personId: personStillValid ? p.personId : '',
      projectAddress: co?.address || p.projectAddress,
    }));
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 640, maxHeight: '92vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>{isNew ? 'New RFQ' : `Edit ${d.rfqNum}`}</span>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 20 }}>×</button>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Company</label>
            <select value={d.organizationId} onChange={e => onCompanyChange(e.target.value)} style={mi}>
              <option value="">— Select company —</option>
              {companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Contact Person</label>
            <select value={d.personId} onChange={e => onPersonChange(e.target.value)} style={mi}>
              <option value="">— Select person —</option>
              {personsForCo.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Title</label>
            <input value={d.projectTitle} onChange={e => set('projectTitle', e.target.value)} style={mi} placeholder="Project name" />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Address</label>
            <input value={d.projectAddress} onChange={e => set('projectAddress', e.target.value)} style={mi} placeholder="Project address" />
          </div>
          <div style={grid2}>
            <div>
              <label style={ml}>Discipline</label>
              <select value={d.discipline} onChange={e => set('discipline', e.target.value as CrmRfqDiscipline)} style={mi}>
                {DISCIPLINES.map(x => <option key={x} value={x}>{x === 'Both' ? 'Steel & Concrete' : x}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>Source</label>
              <select value={d.source} onChange={e => set('source', e.target.value)} style={mi}>
                {SOURCES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>Date Received</label>
              <input type="date" value={d.dateReceived} onChange={e => set('dateReceived', e.target.value)} style={mi} />
            </div>
            <div>
              <label style={ml}>RFQ Required By</label>
              <input type="date" value={d.quoteRequiredBy} onChange={e => set('quoteRequiredBy', e.target.value)} style={mi} />
            </div>
            <div>
              <label style={ml}>Project Value ($)</label>
              <input type="number" min={0} step={100} value={d.projectValue || ''} onChange={e => set('projectValue', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Approximate Hours</label>
              <input type="number" min={0} step={1} value={d.approximateHours || ''} onChange={e => set('approximateHours', Number(e.target.value) || 0)} style={mi} placeholder="e.g. 120" />
            </div>
            <div>
              <label style={ml}>Assigned To</label>
              <select value={d.assignedTo} onChange={e => set('assignedTo', e.target.value)} style={mi}>
                {OWNERS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>Stage</label>
              <select value={d.stage} onChange={e => set('stage', e.target.value as CrmRfqStage)} style={mi}>
                {STAGES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div style={{ display: 'flex', alignItems: 'flex-end', paddingBottom: 4 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer' }}>
                <input type="checkbox" checked={d.createQuoteXero} onChange={e => set('createQuoteXero', e.target.checked)} style={{ width: 16, height: 16, accentColor: C.green }} />
                Create quote in Xero
              </label>
            </div>
          </div>
          <div style={{ marginTop: 14 }}>
            <label style={ml}>Related RFQ (optional)</label>
            <select value={d.relatedRfqId} onChange={e => set('relatedRfqId', e.target.value)} style={mi}>
              <option value="">— None —</option>
              {rfqs.filter(r => r.id !== d.id).map(r => (
                <option key={r.id} value={r.id}>{r.rfqNum} — {r.projectTitle || 'Untitled'}</option>
              ))}
            </select>
          </div>
          <div style={{ marginTop: 14 }}>
            <label style={ml}>Notes</label>
            <textarea
              value={d.notes}
              onChange={e => set('notes', e.target.value)}
              rows={4}
              placeholder="Initial scope, special requirements, attachments needed…"
              style={{ ...mi, resize: 'vertical', minHeight: 80 }}
            />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => { if (d.projectTitle.trim()) onSave(d); }}
            style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}
          >
            Save
          </button>
        </div>
      </div>
    </div>
  );
};

// ── KPI card ──────────────────────────────────────────────────────
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

// ── Main tab ──────────────────────────────────────────────────────
const CrmRfqTab: React.FC<{ spService: SharePointService; persons: CrmPerson[]; companies: CrmCompany[] }> = ({ spService, persons, companies }) => {
  const [rfqs, setRfqs] = React.useState<CrmRfq[]>([]);
  const [rfqReady, setRfqReady] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [stageFilter, setStageFilter] = React.useState('all');
  const [modal, setModal] = React.useState<CrmRfq | null>(null);
  const rfqsRef = React.useRef(rfqs);
  rfqsRef.current = rfqs;
  const saveTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  React.useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const data = await loadRfqsFromSharePoint(spService);
        if (!cancelled) setRfqs(data.map(normalizeRfq));
      } catch {
        if (!cancelled) setRfqs(loadRfqsLocal());
      } finally {
        if (!cancelled) setRfqReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [spService]);

  React.useEffect(() => {
    if (!rfqReady) return;
    const iv = setInterval(() => {
      void (async () => {
        try {
          const data = (await loadRfqsFromSharePoint(spService)).map(normalizeRfq);
          const remoteStr = JSON.stringify(data);
          const localStr = JSON.stringify(rfqsRef.current);
          if (remoteStr !== localStr) setRfqs(data);
        } catch { /* ignore */ }
      })();
    }, 12000);
    return () => clearInterval(iv);
  }, [rfqReady, spService]);

  React.useEffect(() => {
    if (!rfqReady) return;
    localStorage.setItem(LS_RFQS, JSON.stringify(rfqs));
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      void saveRfqsToSharePoint(spService, rfqsRef.current).catch(() => undefined);
    }, 2000);
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    };
  }, [rfqs, rfqReady, spService]);

  const companyName = (id: string): string => companies.find(c => c.id === id)?.name || '—';

  const year = new Date().getFullYear();
  const yearRfqs = rfqs.filter(r => r.dateReceived.startsWith(String(year)));

  const stats = React.useMemo(() => {
    const active = yearRfqs.filter(r => r.stage !== 'Won' && r.stage !== 'Declined');
    const won = yearRfqs.filter(r => r.stage === 'Won');
    const declined = yearRfqs.filter(r => r.stage === 'Declined');
    const overdue = yearRfqs.filter(isOverdue);
    const pipeline = active.reduce((s, r) => s + (r.projectValue || 0), 0);
    const winRate = won.length + declined.length > 0
      ? Math.round((won.length / (won.length + declined.length)) * 100)
      : 0;
    return { total: yearRfqs.length, active: active.length, overdue: overdue.length, won: won.length, declined: declined.length, pipeline, winRate };
  }, [yearRfqs]);

  const q = search.toLowerCase();
  const filtered = yearRfqs.filter(r => {
    if (stageFilter !== 'all' && r.stage !== stageFilter) return false;
    if (!q) return true;
    return (
      r.rfqNum.toLowerCase().includes(q) ||
      r.projectTitle.toLowerCase().includes(q) ||
      companyName(r.organizationId).toLowerCase().includes(q) ||
      r.source.toLowerCase().includes(q)
    );
  });

  const saveRfq = (r: CrmRfq): void => {
    const saved = normalizeRfq(r);
    setRfqs(prev => {
      const next = prev.some(x => x.id === saved.id) ? prev.map(x => x.id === saved.id ? saved : x) : [...prev, saved];
      rfqsRef.current = next;
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      void saveRfqsToSharePoint(spService, next).catch(() => undefined);
      return next;
    });
    setModal(null);
  };

  const deleteRfq = (id: string): void => {
    if (confirm('Delete this RFQ?')) setRfqs(prev => prev.filter(x => x.id !== id));
  };

  if (!rfqReady) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading RFQs…</div>;
  }

  return (
    <>
      {/* KPI row */}
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Total RFQs" accent="#3b82c4" value={String(stats.total)} sub={`${year} year`} />
        <KpiCard label="Active" accent={C.green} value={String(stats.active)} sub="in pipeline" />
        <KpiCard label="Overdue" accent={C.red} value={String(stats.overdue)} sub="past quote by date" />
        <KpiCard label="Won" accent="#0d9488" value={String(stats.won)} sub={`${stats.winRate}% win rate`} />
        <KpiCard label="Declined" accent="#9f1239" value={String(stats.declined)} sub="lost" />
        <KpiCard label="Pipeline Value" accent={C.purple} value={fmtMoney(stats.pipeline)} sub="active enquiries" />
      </div>

      {/* Toolbar — search & filter left, + RFQ right */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'nowrap', paddingBottom: 12, width: '100%', boxSizing: 'border-box' }}>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search RFQs…"
          style={{ padding: '7px 12px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, width: 200, flexShrink: 0, outline: 'none', boxSizing: 'border-box' }}
        />
        <select
          value={stageFilter}
          onChange={e => setStageFilter(e.target.value)}
          style={{ padding: '7px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 150, maxWidth: 150, flexShrink: 0, boxSizing: 'border-box' }}
        >
          <option value="all">All stages</option>
          {STAGES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <button
          onClick={() => setModal(emptyRfq(rfqs))}
          style={{ marginLeft: 'auto', padding: '8px 18px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap', flexShrink: 0 }}
        >
          + RFQ
        </button>
      </div>

      {/* Table */}
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 960 }}>
          <thead>
            <tr>
              {['RFQ #', 'Project', 'Company', 'Type', 'Source', 'Received', 'Quote By', 'Est Value', 'Hours', 'Stage', 'Owner', 'Actions'].map(h => (
                <th key={h} style={{ padding: '9px 10px', textAlign: 'left', fontFamily: FF, fontWeight: 700, fontSize: 10, letterSpacing: '.06em', textTransform: 'uppercase', color: C.sub, background: C.thBg, borderBottom: `2px solid ${C.borderMd}`, whiteSpace: 'nowrap' }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={12} style={{ padding: 48, textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted }}>
                  {rfqs.length === 0 ? 'No RFQs yet — click + RFQ to add one.' : 'No results match your search.'}
                </td>
              </tr>
            ) : filtered.map(r => (
              <tr
                key={r.id}
                onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
              >
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: '#0d9488', borderBottom: `1px solid ${C.border}` }}>{r.rfqNum}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{r.projectTitle || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{companyName(r.organizationId)}</td>
                <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                  <DisciplineBadges discipline={r.discipline} />
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{r.source || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{fmtShortDate(r.dateReceived)}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: isOverdue(r) ? C.red : C.sub, borderBottom: `1px solid ${C.border}` }}>
                  {fmtShortDate(r.quoteRequiredBy)}
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{r.projectValue ? fmtMoney(r.projectValue) : '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{r.approximateHours ? String(r.approximateHours) : '—'}</td>
                <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                  <span style={{ ...badge, ...stageStyle(r.stage) }}>{r.stage.toUpperCase()}</span>
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{r.assignedTo || '—'}</td>
                <td style={{ padding: '8px 10px 8px 6px', borderBottom: `1px solid ${C.border}`, verticalAlign: 'middle' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 52 }}>
                    <button
                      onClick={() => setModal({ ...r })}
                      style={{ ...actionBtn, border: 'none', background: C.purple, color: '#fff' }}
                    >
                      Edit
                    </button>
                    <button
                      onClick={() => deleteRfq(r.id)}
                      style={{ ...actionBtn, border: `1px solid ${C.red}`, background: 'transparent', color: C.red }}
                    >
                      Del
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg }}>
          {filtered.length} of {yearRfqs.length} RFQ{yearRfqs.length !== 1 ? 's' : ''} ({year})
        </div>
      </div>

      {modal && (
        <RfqModal
          initial={modal}
          rfqs={rfqs}
          persons={persons}
          companies={companies}
          onSave={saveRfq}
          onClose={() => setModal(null)}
        />
      )}
    </>
  );
};

export default CrmRfqTab;
