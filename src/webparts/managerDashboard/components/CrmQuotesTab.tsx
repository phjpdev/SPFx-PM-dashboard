import * as React from 'react';
import { loadQuotesFromSharePoint, saveQuotesToSharePoint } from './crmStorage';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmQuote, CrmQuoteStatus, CrmRfqDiscipline } from './crmTypes';

const FF = 'Montserrat,sans-serif';
const LS_QUOTES = '3edge-crm-quotes';

const C = {
  bg: '#f7f8fa', surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', muted: '#8a97a8', green: '#2a9e2a',
  red: '#c0392b', purple: '#6c3fbf', thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const QUOTE_STATUSES: CrmQuoteStatus[] = ['Draft', 'Sent', 'Accepted', 'Declined'];

const mi: React.CSSProperties = {
  padding: '8px 10px', background: C.surface, border: `1px solid ${C.borderMd}`,
  borderRadius: 4, color: C.text, fontSize: 12.5, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: C.sub, letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

const fmtShortDate = (iso: string): string => {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: '2-digit' });
};

const fmtMoney = (n: number): string =>
  '$' + Math.round(n).toLocaleString('en-AU');

const loadQuotesLocal = (): CrmQuote[] => {
  try {
    const v = localStorage.getItem(LS_QUOTES);
    return v ? (JSON.parse(v) as CrmQuote[]) : [];
  } catch { return []; }
};

const normalizeQuote = (q: CrmQuote): CrmQuote => ({
  ...q,
  approximateHours: typeof q.approximateHours === 'number' ? q.approximateHours : 0,
  status: QUOTE_STATUSES.includes(q.status) ? q.status : 'Draft',
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

const statusStyle = (s: CrmQuoteStatus): React.CSSProperties => {
  const map: Record<CrmQuoteStatus, { bg: string; color: string }> = {
    Draft:    { bg: '#f0f2f6', color: '#4a5568' },
    Sent:     { bg: '#e3f0ff', color: '#1a5fa8' },
    Accepted: { bg: '#e6f5e8', color: '#1e6b38' },
    Declined: { bg: '#fde8e8', color: '#a82828' },
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

const QuoteModal: React.FC<{
  initial: CrmQuote;
  onSave: (q: CrmQuote) => void;
  onClose: () => void;
}> = ({ initial, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmQuote>(initial);
  const set = <K extends keyof CrmQuote>(k: K, v: CrmQuote[K]): void => setD(p => ({ ...p, [k]: v }));

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 520, maxHeight: '92vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>{d.quoteNum}</span>
          <button onClick={onClose} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 20 }}>×</button>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <div style={{ marginBottom: 12, fontFamily: FF, fontSize: 12, color: C.sub }}>
            From RFQ <strong style={{ color: '#0d9488' }}>{d.rfqNum}</strong> · {d.projectTitle || 'Untitled'}
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Status</label>
            <select value={d.status} onChange={e => set('status', e.target.value as CrmQuoteStatus)} style={mi}>
              {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Value ($)</label>
            <input type="number" min={0} step={100} value={d.projectValue || ''} onChange={e => set('projectValue', Number(e.target.value) || 0)} style={mi} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Approximate Hours</label>
            <input type="number" min={0} step={1} value={d.approximateHours || ''} onChange={e => set('approximateHours', Number(e.target.value) || 0)} style={mi} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer' }}>
              <input type="checkbox" checked={d.createQuoteXero} onChange={e => set('createQuoteXero', e.target.checked)} style={{ width: 16, height: 16, accentColor: C.green }} />
              Create quote in Xero
            </label>
          </div>
          <div>
            <label style={ml}>Notes</label>
            <textarea value={d.notes} onChange={e => set('notes', e.target.value)} rows={4} style={{ ...mi, resize: 'vertical', minHeight: 80 }} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button onClick={() => onSave(d)} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
        </div>
      </div>
    </div>
  );
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

const CrmQuotesTab: React.FC<{
  spService: SharePointService;
  persons: CrmPerson[];
  companies: CrmCompany[];
  seedQuotes?: CrmQuote[] | null;
  onSeedApplied?: () => void;
}> = ({ spService, persons, companies, seedQuotes, onSeedApplied }) => {
  const [quotes, setQuotes] = React.useState<CrmQuote[]>([]);
  const [ready, setReady] = React.useState(false);
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState('all');
  const [modal, setModal] = React.useState<CrmQuote | null>(null);
  const quotesRef = React.useRef(quotes);
  quotesRef.current = quotes;
  const saveTimerRef = React.useRef<ReturnType<typeof setTimeout> | null>(null);

  const reload = React.useCallback(async (): Promise<void> => {
    try {
      const data = (await loadQuotesFromSharePoint(spService)).map(normalizeQuote);
      const remoteStr = JSON.stringify(data);
      const localStr = JSON.stringify(quotesRef.current);
      if (remoteStr !== localStr) setQuotes(data);
    } catch {
      setQuotes(loadQuotesLocal().map(normalizeQuote));
    }
  }, [spService]);

  React.useEffect(() => {
    if (seedQuotes?.length) {
      setQuotes(seedQuotes.map(normalizeQuote));
      setReady(true);
      onSeedApplied?.();
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const data = await loadQuotesFromSharePoint(spService);
        if (!cancelled) setQuotes(data.map(normalizeQuote));
      } catch {
        if (!cancelled) setQuotes(loadQuotesLocal().map(normalizeQuote));
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [spService, seedQuotes]); // eslint-disable-line react-hooks/exhaustive-deps -- onSeedApplied fires once after seed

  React.useEffect(() => {
    if (!ready) return;
    const iv = setInterval(() => { void reload(); }, 12000);
    return () => clearInterval(iv);
  }, [ready, reload]);

  React.useEffect(() => {
    if (!ready) return;
    localStorage.setItem(LS_QUOTES, JSON.stringify(quotes));
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      void saveQuotesToSharePoint(spService, quotesRef.current).catch(() => undefined);
    }, 2000);
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    };
  }, [quotes, ready, spService]);

  const companyName = (id: string): string => companies.find(c => c.id === id)?.name || '—';
  const personName = (id: string): string => persons.find(p => p.id === id)?.name || '—';

  const year = new Date().getFullYear();
  const yearQuotes = quotes.filter(q => q.quotedDate.startsWith(String(year)));

  const stats = React.useMemo(() => {
    const sent = yearQuotes.filter(q => q.status === 'Sent');
    const accepted = yearQuotes.filter(q => q.status === 'Accepted');
    const draft = yearQuotes.filter(q => q.status === 'Draft');
    const value = yearQuotes.reduce((s, q) => s + (q.projectValue || 0), 0);
    return { total: yearQuotes.length, draft: draft.length, sent: sent.length, accepted: accepted.length, value };
  }, [yearQuotes]);

  const q = search.toLowerCase();
  const filtered = yearQuotes.filter(item => {
    if (statusFilter !== 'all' && item.status !== statusFilter) return false;
    if (!q) return true;
    return (
      item.quoteNum.toLowerCase().includes(q) ||
      item.rfqNum.toLowerCase().includes(q) ||
      item.projectTitle.toLowerCase().includes(q) ||
      companyName(item.organizationId).toLowerCase().includes(q)
    );
  });

  const saveQuote = (item: CrmQuote): void => {
    const saved = normalizeQuote(item);
    setQuotes(prev => {
      const next = prev.map(x => x.id === saved.id ? saved : x);
      quotesRef.current = next;
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      void saveQuotesToSharePoint(spService, next).catch(() => undefined);
      return next;
    });
    setModal(null);
  };

  const deleteQuote = (id: string): void => {
    if (confirm('Delete this quote?')) setQuotes(prev => prev.filter(x => x.id !== id));
  };

  if (!ready) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading quotes…</div>;
  }

  return (
    <>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Total Quotes" accent={C.purple} value={String(stats.total)} sub={`${year} year`} />
        <KpiCard label="Draft" accent="#8a97a8" value={String(stats.draft)} sub="not yet sent" />
        <KpiCard label="Sent" accent="#3b82c4" value={String(stats.sent)} sub="awaiting response" />
        <KpiCard label="Accepted" accent={C.green} value={String(stats.accepted)} sub="won" />
        <KpiCard label="Quote Value" accent="#0d9488" value={fmtMoney(stats.value)} sub="all quotes" />
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'nowrap', paddingBottom: 12 }}>
        <input
          value={search}
          onChange={e => setSearch(e.target.value)}
          placeholder="Search quotes…"
          style={{ padding: '7px 12px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, width: 200, flexShrink: 0, outline: 'none', boxSizing: 'border-box' }}
        />
        <select
          value={statusFilter}
          onChange={e => setStatusFilter(e.target.value)}
          style={{ padding: '7px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, width: 150, flexShrink: 0, boxSizing: 'border-box' }}
        >
          <option value="all">All statuses</option>
          {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </div>

      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 960 }}>
          <thead>
            <tr>
              {['Quote #', 'RFQ #', 'Project', 'Company', 'Contact', 'Type', 'Quoted', 'Est Value', 'Hours', 'Status', 'Owner', 'Actions'].map(h => (
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
                  {quotes.length === 0
                    ? 'No quotes yet — move an RFQ at Ready to Quote stage using the Quote button.'
                    : 'No results match your search.'}
                </td>
              </tr>
            ) : filtered.map(item => (
              <tr
                key={item.id}
                onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
              >
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.purple, borderBottom: `1px solid ${C.border}` }}>{item.quoteNum}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: '#0d9488', borderBottom: `1px solid ${C.border}` }}>{item.rfqNum}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{item.projectTitle || '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{companyName(item.organizationId)}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{personName(item.personId)}</td>
                <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                  <DisciplineBadges discipline={item.discipline} />
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{fmtShortDate(item.quotedDate)}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{item.projectValue ? fmtMoney(item.projectValue) : '—'}</td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 600, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{item.approximateHours ? String(item.approximateHours) : '—'}</td>
                <td style={{ padding: '10px', borderBottom: `1px solid ${C.border}` }}>
                  <span style={{ ...badge, ...statusStyle(item.status) }}>{item.status.toUpperCase()}</span>
                </td>
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.sub, borderBottom: `1px solid ${C.border}` }}>{item.assignedTo || '—'}</td>
                <td style={{ padding: '8px 10px 8px 6px', borderBottom: `1px solid ${C.border}`, verticalAlign: 'middle' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 52 }}>
                    <button onClick={() => setModal({ ...item })} style={{ ...actionBtn, border: 'none', background: C.purple, color: '#fff' }}>Edit</button>
                    <button onClick={() => deleteQuote(item.id)} style={{ ...actionBtn, border: `1px solid ${C.red}`, background: 'transparent', color: C.red }}>Del</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg }}>
          {filtered.length} of {yearQuotes.length} quote{yearQuotes.length !== 1 ? 's' : ''} ({year})
        </div>
      </div>

      {modal && (
        <QuoteModal initial={modal} onSave={saveQuote} onClose={() => setModal(null)} />
      )}
    </>
  );
};

export default CrmQuotesTab;
