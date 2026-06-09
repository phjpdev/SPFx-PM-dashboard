import * as React from 'react';
import { getQuoteAttachments, setQuoteAttachments } from './crmAttachmentStore';
import { DocumentUploadSection } from './crmDocumentUpload';
import {
  loadProjectsFromSharePoint, loadQuoteBudget, loadQuotesFromSharePoint,
  saveProjectsToSharePoint, saveQuoteBudget, saveQuotesToSharePoint,
  type CrmQuoteBudget,
} from './crmStorage';
import { nextProjNum, quoteToCrmProject, quoteToIProject } from './crmProjectHelpers';
import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmAttachment, CrmPerson, CrmCompany, CrmProject, CrmQuote, CrmQuoteStatus, CrmRfqDiscipline } from './crmTypes';

const FF = 'Montserrat,sans-serif';
const LS_QUOTES = '3edge-crm-quotes';

const C = {
  bg: '#f7f8fa', surface: '#ffffff', border: '#e2e5ea', borderMd: '#cdd1d9',
  text: '#1a2030', sub: '#4a5568', muted: '#8a97a8', green: '#2a9e2a',
  red: '#c0392b', purple: '#6c3fbf', thBg: '#f0f2f6', rowHover: '#f5f7fb',
};

const QUOTE_STATUSES: CrmQuoteStatus[] = ['Draft', 'Sent', 'Lost'];
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

const normalizeQuote = (q: CrmQuote): CrmQuote => {
  let status: CrmQuoteStatus = q.status;
  const raw = q.status as string;
  if (raw === 'Accepted') status = 'Draft';
  if (raw === 'Declined') status = 'Lost';
  if (!QUOTE_STATUSES.includes(status)) status = 'Draft';
  return {
    ...q,
    approximateHours: typeof q.approximateHours === 'number' ? q.approximateHours : 0,
    rfiAllowed: typeof q.rfiAllowed === 'number' ? q.rfiAllowed : 0,
    engineerDrawingReceived: !!q.engineerDrawingReceived,
    engineerDrawingDate: q.engineerDrawingDate || '',
    revisionVersionEng: q.revisionVersionEng || '',
    architectDrawingReceived: !!q.architectDrawingReceived,
    architectDrawingDate: q.architectDrawingDate || '',
    revisionVersionArch: q.revisionVersionArch || '',
    quoteRequiredBy: q.quoteRequiredBy || '',
    relatedRfqId: q.relatedRfqId || '',
    createQuoteXero: !!q.createQuoteXero,
    status,
  };
};

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

const statusStyle = (s: CrmQuoteStatus): React.CSSProperties => {
  const map: Record<CrmQuoteStatus, { bg: string; color: string }> = {
    Draft: { bg: '#f0f2f6', color: '#4a5568' },
    Sent:  { bg: '#e3f0ff', color: '#1a5fa8' },
    Lost:  { bg: '#fde8e8', color: '#a82828' },
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

const DrawingReceivedField: React.FC<{
  label: string;
  checked: boolean;
  date: string;
  onCheck: (v: boolean) => void;
  onDate: (v: string) => void;
}> = ({ label, checked, date, onCheck, onDate }) => (
  <div>
    <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer', marginBottom: checked ? 6 : 0 }}>
      <input
        type="checkbox"
        checked={checked}
        onChange={e => {
          const v = e.target.checked;
          onCheck(v);
          if (!v) onDate('');
        }}
        style={{ width: 16, height: 16, accentColor: C.green, flexShrink: 0 }}
      />
      <span style={{ ...ml, marginBottom: 0, textTransform: 'none', letterSpacing: 0, fontSize: 12 }}>{label}</span>
    </label>
    {checked && (
      <input type="date" value={date} onChange={e => onDate(e.target.value)} style={mi} />
    )}
  </div>
);

const QuoteModal: React.FC<{
  initial: CrmQuote;
  persons: CrmPerson[];
  companies: CrmCompany[];
  onSave: (q: CrmQuote) => void;
  onClose: () => void;
}> = ({ initial, persons, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmQuote>(initial);
  const [attachments, setAttachments] = React.useState<CrmAttachment[]>(() => getQuoteAttachments(initial.id));
  const set = <K extends keyof CrmQuote>(k: K, v: CrmQuote[K]): void => setD(p => ({ ...p, [k]: v }));
  const grid2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px 18px' };

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

  const canSave = !d.createQuoteXero || d.quoteNum.trim().length > 0;

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 680, maxHeight: '92vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Edit Quote — {d.rfqNum}</span>
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
            <input value={d.projectTitle} onChange={e => set('projectTitle', e.target.value)} style={mi} />
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Project Address</label>
            <input value={d.projectAddress} onChange={e => set('projectAddress', e.target.value)} style={mi} />
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
            <DrawingReceivedField
              label="Engineer drawing received"
              checked={d.engineerDrawingReceived}
              date={d.engineerDrawingDate}
              onCheck={v => set('engineerDrawingReceived', v)}
              onDate={v => set('engineerDrawingDate', v)}
            />
            <DrawingReceivedField
              label="Architect drawing received"
              checked={d.architectDrawingReceived}
              date={d.architectDrawingDate}
              onCheck={v => set('architectDrawingReceived', v)}
              onDate={v => set('architectDrawingDate', v)}
            />
            <div>
              <label style={ml}>Revision version Eng</label>
              <input value={d.revisionVersionEng} onChange={e => set('revisionVersionEng', e.target.value)} style={mi} placeholder="e.g. Rev A" />
            </div>
            <div>
              <label style={ml}>Revision version Arch</label>
              <input value={d.revisionVersionArch} onChange={e => set('revisionVersionArch', e.target.value)} style={mi} placeholder="e.g. Rev A" />
            </div>
            <div>
              <label style={ml}>Project Value ($)</label>
              <input type="number" min={0} step={100} value={d.projectValue || ''} onChange={e => set('projectValue', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Assigned To</label>
              <select value={d.assignedTo} onChange={e => set('assignedTo', e.target.value)} style={mi}>
                {OWNERS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            <div>
              <label style={ml}>RFI allowed</label>
              <input type="number" min={0} step={1} value={d.rfiAllowed || ''} onChange={e => set('rfiAllowed', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Est Hours</label>
              <input type="number" min={0} step={1} value={d.approximateHours || ''} onChange={e => set('approximateHours', Number(e.target.value) || 0)} style={mi} />
            </div>
            <div>
              <label style={ml}>Status</label>
              <select value={d.status} onChange={e => set('status', e.target.value as CrmQuoteStatus)} style={mi}>
                {QUOTE_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
          <div style={{ marginTop: 14, marginBottom: 14 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: FF, fontSize: 12, color: C.text, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={d.createQuoteXero}
                onChange={e => set('createQuoteXero', e.target.checked)}
                style={{ width: 16, height: 16, accentColor: C.green }}
              />
              Create quote in Xero
            </label>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Quote # (Xero)</label>
            <input
              value={d.quoteNum}
              onChange={e => set('quoteNum', e.target.value)}
              style={mi}
              placeholder="QU-0490"
            />
          </div>
          <div>
            <label style={ml}>Notes</label>
            <textarea value={d.notes} onChange={e => set('notes', e.target.value)} rows={4} style={{ ...mi, resize: 'vertical', minHeight: 80 }} />
          </div>
          <DocumentUploadSection attachments={attachments} onChange={setAttachments} />
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button
            onClick={() => {
              if (canSave) {
                setQuoteAttachments(d.id, attachments);
                onSave({ ...d, quoteNum: d.quoteNum.trim() });
              }
            }}
            style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: canSave ? C.green : C.borderMd, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: canSave ? 'pointer' : 'default' }}
          >
            Save
          </button>
        </div>
      </div>
    </div>
  );
};

const KpiCard: React.FC<{
  label: string; value: string; sub: string; accent: string; onClick?: () => void;
}> = ({ label, value, sub, accent, onClick }) => (
  <div
    onClick={onClick}
    style={{
      flex: '1 1 120px', background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6,
      overflow: 'hidden', minWidth: 100, cursor: onClick ? 'pointer' : 'default',
    }}
  >
    <div style={{ height: 3, background: accent }} />
    <div style={{ padding: '12px 14px' }}>
      <div style={{ fontFamily: FF, fontSize: 9, fontWeight: 700, letterSpacing: '.08em', color: C.muted, textTransform: 'uppercase' }}>{label}</div>
      <div style={{ fontFamily: FF, fontSize: 22, fontWeight: 700, color: C.text, marginTop: 4 }}>{value}</div>
      <div style={{ fontFamily: FF, fontSize: 11, color: C.muted, marginTop: 2 }}>{sub}</div>
    </div>
  </div>
);

const BudgetEditModal: React.FC<{
  budget: CrmQuoteBudget;
  onSave: (b: CrmQuoteBudget) => void;
  onClose: () => void;
}> = ({ budget, onSave, onClose }) => {
  const [b, setB] = React.useState(budget);
  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1001, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: C.surface, borderRadius: 8, width: 400, boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
        <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, background: C.thBg }}>
          <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: C.text }}>Quote Budget Targets</span>
        </div>
        <div style={{ padding: '20px 22px' }}>
          <div style={{ marginBottom: 14 }}>
            <label style={ml}>Year budget ($)</label>
            <input type="number" min={0} step={1000} value={b.yearTarget || ''} onChange={e => setB(p => ({ ...p, yearTarget: Number(e.target.value) || 0 }))} style={mi} />
          </div>
          <div>
            <label style={ml}>Month budget ($)</label>
            <input type="number" min={0} step={500} value={b.monthTarget || ''} onChange={e => setB(p => ({ ...p, monthTarget: Number(e.target.value) || 0 }))} style={mi} />
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', padding: '14px 22px', borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
          <button onClick={() => onSave({ ...b, year: new Date().getFullYear() })} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
        </div>
      </div>
    </div>
  );
};

const CrmQuotesTab: React.FC<{
  spService: SharePointService;
  persons: CrmPerson[];
  companies: CrmCompany[];
  seedQuotes?: CrmQuote[] | null;
  onSeedApplied?: () => void;
  onQuoteWon?: (projects: CrmProject[], quotes: CrmQuote[]) => void;
}> = ({ spService, persons, companies, seedQuotes, onSeedApplied, onQuoteWon }) => {
  const [quotes, setQuotes] = React.useState<CrmQuote[]>([]);
  const [wonProjects, setWonProjects] = React.useState<CrmProject[]>([]);
  const [budget, setBudget] = React.useState<CrmQuoteBudget>({ year: new Date().getFullYear(), yearTarget: 0, monthTarget: 0 });
  const [budgetModal, setBudgetModal] = React.useState(false);
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
      const projs = await loadProjectsFromSharePoint(spService);
      const yr = new Date().getFullYear();
      setWonProjects(projs.filter(p => p.wonDate.startsWith(String(yr))));
    } catch {
      setQuotes(loadQuotesLocal().map(normalizeQuote));
    }
  }, [spService]);

  React.useEffect(() => {
    if (seedQuotes?.length) {
      setQuotes(seedQuotes.map(normalizeQuote));
      setReady(true);
      onSeedApplied?.();
      void Promise.all([
        loadProjectsFromSharePoint(spService),
        loadQuoteBudget(spService),
      ]).then(([projs, b]) => {
        const yr = new Date().getFullYear();
        setWonProjects(projs.filter(p => p.wonDate.startsWith(String(yr))));
        setBudget(b);
      });
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const [data, projs, b] = await Promise.all([
          loadQuotesFromSharePoint(spService),
          loadProjectsFromSharePoint(spService),
          loadQuoteBudget(spService),
        ]);
        if (!cancelled) {
          setQuotes(data.map(normalizeQuote));
          const yr = new Date().getFullYear();
          setWonProjects(projs.filter(p => p.wonDate.startsWith(String(yr))));
          setBudget(b);
        }
      } catch {
        if (!cancelled) setQuotes(loadQuotesLocal().map(normalizeQuote));
      } finally {
        if (!cancelled) setReady(true);
      }
    })();
    return () => { cancelled = true; };
  }, [spService, seedQuotes]); // eslint-disable-line react-hooks/exhaustive-deps

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
  const month = new Date().getMonth();
  const yearQuotes = quotes.filter(q => q.quotedDate.startsWith(String(year)));

  const stats = React.useMemo(() => {
    const sent = yearQuotes.filter(q => q.status === 'Sent');
    const lost = yearQuotes.filter(q => q.status === 'Lost');
    const totalHours = yearQuotes.reduce((s, q) => s + (q.approximateHours || 0), 0);
    const yearActual = wonProjects.reduce((s, p) => s + (p.projectValue || 0), 0);
    const monthActual = wonProjects
      .filter(p => new Date(p.wonDate + 'T00:00:00').getMonth() === month)
      .reduce((s, p) => s + (p.projectValue || 0), 0);
    const yearPct = budget.yearTarget > 0 ? Math.round((yearActual / budget.yearTarget) * 100) : null;
    const monthPct = budget.monthTarget > 0 ? Math.round((monthActual / budget.monthTarget) * 100) : null;
    return {
      total: yearQuotes.length,
      sent: sent.length,
      lost: lost.length,
      won: wonProjects.length,
      totalHours,
      yearActual,
      monthActual,
      yearPct,
      monthPct,
    };
  }, [yearQuotes, wonProjects, budget, month]);

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

  const persistQuotes = (next: CrmQuote[]): void => {
    quotesRef.current = next;
    setQuotes(next);
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    void saveQuotesToSharePoint(spService, next).catch(() => undefined);
  };

  const saveQuote = (item: CrmQuote): void => {
    const saved = normalizeQuote(item);
    persistQuotes(quotesRef.current.map(x => x.id === saved.id ? saved : x));
    setModal(null);
  };

  const markLost = (item: CrmQuote): void => {
    if (!confirm(`Mark ${item.rfqNum} as LOST?`)) return;
    persistQuotes(quotesRef.current.map(x => x.id === item.id ? { ...x, status: 'Lost' as CrmQuoteStatus } : x));
  };

  const markWon = (item: CrmQuote): void => {
    if (!confirm(`Mark ${item.rfqNum} as WON and create project ${item.projectTitle || ''}?`)) return;
    void (async () => {
      try {
        const [crmProjects, dashProjects] = await Promise.all([
          loadProjectsFromSharePoint(spService),
          spService.loadProjects().catch(() => []),
        ]);
        const projNum = nextProjNum(crmProjects, dashProjects);
        const iProject = quoteToIProject(item, projNum, persons, companies);
        let spId: number | undefined;
        try {
          spId = await spService.addProject(iProject);
        } catch { /* CRM project still created if SP list unavailable */ }
        const crmProject = quoteToCrmProject(item, projNum, spId);
        const nextProjects = [...crmProjects, crmProject];
        const nextQuotes = quotesRef.current.filter(q => q.id !== item.id);
        persistQuotes(nextQuotes);
        await saveProjectsToSharePoint(spService, nextProjects);
        const yr = new Date().getFullYear();
        setWonProjects(nextProjects.filter(p => p.wonDate.startsWith(String(yr))));
        onQuoteWon?.(nextProjects, nextQuotes);
      } catch {
        alert('Could not create project. Please try again.');
      }
    })();
  };

  const deleteQuote = (id: string): void => {
    if (confirm('Delete this quote?')) persistQuotes(quotesRef.current.filter(x => x.id !== id));
  };

  if (!ready) {
    return <div style={{ padding: 24, fontFamily: FF, fontSize: 13, color: C.muted }}>Loading quotes…</div>;
  }

  return (
    <>
      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', padding: '16px 0 12px 0' }}>
        <KpiCard label="Total Quotes" accent={C.purple} value={String(stats.total)} sub={`${year} year`} />
        <KpiCard label="Sent" accent="#3b82c4" value={String(stats.sent)} sub="awaiting response" />
        <KpiCard label="Won" accent={C.green} value={String(stats.won)} sub="moved to projects" />
        <KpiCard label="Lost" accent={C.red} value={String(stats.lost)} sub="declined" />
        <KpiCard
          label="Quote Budget Year"
          accent="#b36a00"
          value={`${fmtMoney(stats.yearActual)} / ${budget.yearTarget ? fmtMoney(budget.yearTarget) : '—'}`}
          sub={stats.yearPct !== null ? `${stats.yearPct}% of year budget · click to edit` : 'actual vs budget · click to set'}
          onClick={() => setBudgetModal(true)}
        />
        <KpiCard
          label="Quote Budget Month"
          accent="#d97706"
          value={`${fmtMoney(stats.monthActual)} / ${budget.monthTarget ? fmtMoney(budget.monthTarget) : '—'}`}
          sub={stats.monthPct !== null ? `${stats.monthPct}% of month budget · click to edit` : `${new Date().toLocaleString('en-AU', { month: 'short' })} actual vs budget`}
          onClick={() => setBudgetModal(true)}
        />
        <KpiCard label="Hours Est" accent="#0d9488" value={String(stats.totalHours)} sub="total estimated hours" />
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
              {['Quote #', 'RFQ #', 'Project', 'Company', 'Contact', 'Type', 'Quoted', 'Est Value', 'Hours Est', 'Status', 'Owner', 'Actions'].map(h => (
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
                <td style={{ padding: '10px', fontFamily: FF, fontSize: 12, fontWeight: 700, color: C.purple, borderBottom: `1px solid ${C.border}` }}>{item.quoteNum || '—'}</td>
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
                    {item.status !== 'Lost' && (
                      <>
                        <button onClick={() => markWon(item)} style={{ ...actionBtn, border: 'none', background: C.green, color: '#fff' }}>WON</button>
                        <button onClick={() => markLost(item)} style={{ ...actionBtn, border: `1px solid ${C.red}`, background: 'transparent', color: C.red }}>LOST</button>
                      </>
                    )}
                    <button onClick={() => deleteQuote(item.id)} style={{ ...actionBtn, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.muted, fontSize: 9 }}>Del</button>
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
        <QuoteModal
          initial={modal}
          persons={persons}
          companies={companies}
          onSave={saveQuote}
          onClose={() => setModal(null)}
        />
      )}
      {budgetModal && (
        <BudgetEditModal
          budget={budget}
          onSave={b => {
            setBudget(b);
            void saveQuoteBudget(spService, b);
            setBudgetModal(false);
          }}
          onClose={() => setBudgetModal(false)}
        />
      )}
    </>
  );
};

export default CrmQuotesTab;
