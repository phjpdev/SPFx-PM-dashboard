import * as React from 'react';

// ── Types ─────────────────────────────────────────────────────────
interface CrmPhone { value: string; type: string; }
interface CrmEmail { value: string; type: string; }

export interface CrmPerson {
  id: string;
  name: string;
  organizationId: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}

export interface CrmCompany {
  id: string;
  name: string;
  labels: string;
  address: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}

// ── Constants ─────────────────────────────────────────────────────
const PHONE_TYPES  = ['Work', 'Home', 'Mobile', 'Other'];
const EMAIL_TYPES  = ['Work', 'Home', 'Other'];
const FF           = 'Montserrat,sans-serif';
const LS_PERSONS   = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';

// ── Light-mode palette ────────────────────────────────────────────
const C = {
  bg:       '#f7f8fa',
  surface:  '#ffffff',
  border:   '#e2e5ea',
  borderMd: '#cdd1d9',
  text:     '#1a2030',
  sub:      '#4a5568',
  muted:    '#8a97a8',
  green:    '#2a9e2a',
  greenBg:  'rgba(42,158,42,.09)',
  greenBd:  'rgba(42,158,42,.35)',
  purple:   '#6c3fbf',
  red:      '#c0392b',
  thBg:     '#f0f2f6',
  rowHover: '#f5f7fb',
  tag:      '#eef0f5',
  tagText:  '#5a6a80',
};

// ── Helpers ───────────────────────────────────────────────────────
const uid      = (): string    => `${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;
const loadLS   = <T,>(k: string, fb: T): T => { try { const v = localStorage.getItem(k); return v ? (JSON.parse(v) as T) : fb; } catch { return fb; } };
const firstVal = (arr: CrmPhone[]): string => arr.find(x => x.value)?.value || '—';

const emptyPerson  = (): CrmPerson  => ({ id: uid(), name: '', organizationId: '', phones: [{ value: '', type: 'Work' }], emails: [{ value: '', type: 'Work' }] });
const emptyCompany = (): CrmCompany => ({ id: uid(), name: '', labels: '', address: '', phones: [{ value: '', type: 'Work' }], emails: [{ value: '', type: 'Work' }] });

// ── Shared modal input style ──────────────────────────────────────
const mi: React.CSSProperties = {
  padding: '8px 10px', background: C.surface, border: `1px solid ${C.borderMd}`,
  borderRadius: 4, color: C.text, fontSize: 12.5, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const ml: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: C.sub, letterSpacing: '.07em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

// ── Address autocomplete ──────────────────────────────────────────
interface NominatimResult { display_name: string; }

const AddressSearch: React.FC<{ value: string; onChange: (v: string) => void }> = ({ value, onChange }) => {
  const [suggestions, setSuggestions] = React.useState<string[]>([]);
  const [open, setOpen]               = React.useState(false);
  const timer = React.useRef<ReturnType<typeof setTimeout> | null>(null);
  const wrapRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const handleChange = (q: string): void => {
    onChange(q);
    if (timer.current) clearTimeout(timer.current);
    if (q.length < 3) { setSuggestions([]); setOpen(false); return; }
    timer.current = setTimeout(async () => {
      try {
        const res  = await fetch(`https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(q)}&format=json&limit=6`, { headers: { 'Accept-Language': 'en' } });
        const data = await res.json() as NominatimResult[];
        setSuggestions(data.map(d => d.display_name));
        setOpen(true);
      } catch { setSuggestions([]); }
    }, 420);
  };

  const pick = (s: string): void => { onChange(s); setSuggestions([]); setOpen(false); };

  return (
    <div ref={wrapRef} style={{ position: 'relative' }}>
      <input
        value={value}
        onChange={e => handleChange(e.target.value)}
        onFocus={() => suggestions.length > 0 && setOpen(true)}
        placeholder="Start typing an address…"
        style={mi}
      />
      {open && suggestions.length > 0 && (
        <div style={{ position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, background: C.surface, border: `1px solid ${C.borderMd}`, borderRadius: 6, zIndex: 200, boxShadow: '0 6px 20px rgba(0,0,0,.12)', overflow: 'hidden' }}>
          {suggestions.map((s, i) => (
            <div
              key={i}
              onMouseDown={() => pick(s)}
              style={{ padding: '9px 12px', fontSize: 12, fontFamily: FF, color: C.text, cursor: 'pointer', borderBottom: i < suggestions.length - 1 ? `1px solid ${C.border}` : 'none', display: 'flex', alignItems: 'flex-start', gap: 8 }}
              onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = C.rowHover; }}
              onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = 'transparent'; }}
            >
              <svg style={{ flexShrink: 0, marginTop: 1 }} width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={C.muted} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>
              </svg>
              <span style={{ lineHeight: 1.4 }}>{s}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

// ── Modal wrapper (light) ─────────────────────────────────────────
const Modal: React.FC<{ title: string; onClose: () => void; children: React.ReactNode }> = ({ title, onClose, children }) => (
  <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
    <div style={{ background: C.surface, borderRadius: 8, width: 460, maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 12px 40px rgba(0,0,0,.18)', border: `1px solid ${C.border}` }}>
      <div style={{ padding: '14px 20px', borderBottom: `1px solid ${C.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: C.thBg, borderRadius: '8px 8px 0 0' }}>
        <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 13, color: C.text, letterSpacing: '.04em' }}>{title}</span>
        <button onClick={onClose} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 20, lineHeight: 1, padding: 0 }}>×</button>
      </div>
      <div style={{ padding: '20px 22px' }}>{children}</div>
    </div>
  </div>
);

// ── Multi-value field (phone / email) ─────────────────────────────
const MultiField: React.FC<{
  items: CrmPhone[];
  types: string[];
  addLabel: string;
  placeholder: string;
  onChange: (items: CrmPhone[]) => void;
}> = ({ items, types, addLabel, placeholder, onChange }) => (
  <div style={{ marginBottom: 14 }}>
    {items.map((item, i) => (
      <div key={i} style={{ display: 'flex', gap: 6, marginBottom: 6, alignItems: 'center' }}>
        <input value={item.value} onChange={e => { const n = [...items]; n[i] = { ...n[i], value: e.target.value }; onChange(n); }} placeholder={placeholder} style={{ ...mi, flex: 1 }} />
        <select value={item.type} onChange={e => { const n = [...items]; n[i] = { ...n[i], type: e.target.value }; onChange(n); }} style={{ ...mi, width: 90, flex: 'none' }}>
          {types.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        {items.length > 1 && <button onClick={() => onChange(items.filter((_, j) => j !== i))} style={{ background: 'none', border: 'none', color: C.muted, cursor: 'pointer', fontSize: 18, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>×</button>}
      </div>
    ))}
    <button onClick={() => onChange([...items, { value: '', type: types[0] }])} style={{ background: 'none', border: 'none', color: C.green, fontFamily: FF, fontSize: 12, cursor: 'pointer', padding: 0 }}>
      + Add {addLabel}
    </button>
  </div>
);

// ── Modal footer ──────────────────────────────────────────────────
const ModalFooter: React.FC<{ onCancel: () => void; onSave: () => void }> = ({ onCancel, onSave }) => (
  <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 20, paddingTop: 14, borderTop: `1px solid ${C.border}` }}>
    <button onClick={onCancel} style={{ padding: '8px 20px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: 'transparent', color: C.sub, fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
    <button onClick={onSave}   style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
  </div>
);

// ── Person Modal ──────────────────────────────────────────────────
const PersonModal: React.FC<{ initial: CrmPerson; companies: CrmCompany[]; onSave: (p: CrmPerson) => void; onClose: () => void }> = ({ initial, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmPerson>(initial);
  const set = <K extends keyof CrmPerson>(k: K, v: CrmPerson[K]): void => setD(p => ({ ...p, [k]: v }));
  return (
    <Modal title={initial.name ? 'Edit Person' : 'Add Person'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={mi} placeholder="Full name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Organization</label>
        <div style={{ position: 'relative' }}>
          <span style={{ position: 'absolute', left: 9, top: '50%', transform: 'translateY(-50%)', color: C.muted, pointerEvents: 'none' }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-4 0v2"/></svg>
          </span>
          <select value={d.organizationId} onChange={e => set('organizationId', e.target.value)} style={{ ...mi, paddingLeft: 28 }}>
            <option value="">— None —</option>
            {companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </div>
      </div>
      <label style={ml}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" onChange={v => set('phones', v)} />
      <label style={ml}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />
      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(d); }} />
    </Modal>
  );
};

// ── Company Modal ─────────────────────────────────────────────────
const CompanyModal: React.FC<{ initial: CrmCompany; onSave: (c: CrmCompany) => void; onClose: () => void }> = ({ initial, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmCompany>(initial);
  const set = <K extends keyof CrmCompany>(k: K, v: CrmCompany[K]): void => setD(p => ({ ...p, [k]: v }));
  return (
    <Modal title={initial.name ? 'Edit Company' : 'Add Company'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={mi} placeholder="Company name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Labels</label>
        <input value={d.labels} onChange={e => set('labels', e.target.value)} style={mi} placeholder="e.g. Client, Supplier, Partner" />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={ml}>Address</label>
        <AddressSearch value={d.address} onChange={v => set('address', v)} />
      </div>
      <label style={ml}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" onChange={v => set('phones', v)} />
      <label style={ml}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />
      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(d); }} />
    </Modal>
  );
};

// ── Table helpers ─────────────────────────────────────────────────
const Th: React.FC<{ label: string; w?: string | number }> = ({ label, w }) => (
  <th style={{ padding: '9px 12px', textAlign: 'left', fontFamily: FF, fontWeight: 700, fontSize: 10.5, letterSpacing: '.07em', textTransform: 'uppercase', color: C.sub, background: C.thBg, borderBottom: `2px solid ${C.borderMd}`, whiteSpace: 'nowrap', width: w }}>
    {label}
  </th>
);

const Td: React.FC<{ children: React.ReactNode; muted?: boolean; mono?: boolean }> = ({ children, muted, mono }) => (
  <td style={{ padding: '10px 12px', fontFamily: mono ? 'monospace' : FF, fontSize: 12.5, color: muted ? C.muted : C.text, borderBottom: `1px solid ${C.border}`, verticalAlign: 'middle', maxWidth: 260, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
    {children}
  </td>
);

// ── CrmBoard ──────────────────────────────────────────────────────
const CrmBoard: React.FC = () => {
  const [tab, setTab]             = React.useState<'persons' | 'companies'>('persons');
  const [persons, setPersons]     = React.useState<CrmPerson[]>(() => loadLS<CrmPerson[]>(LS_PERSONS, []));
  const [companies, setCompanies] = React.useState<CrmCompany[]>(() => loadLS<CrmCompany[]>(LS_COMPANIES, []));
  const [personModal, setPersonModal]   = React.useState<CrmPerson | null>(null);
  const [companyModal, setCompanyModal] = React.useState<CrmCompany | null>(null);
  const [search, setSearch]             = React.useState('');

  React.useEffect(() => { localStorage.setItem(LS_PERSONS,   JSON.stringify(persons));   }, [persons]);
  React.useEffect(() => { localStorage.setItem(LS_COMPANIES, JSON.stringify(companies)); }, [companies]);

  const savePerson    = (p: CrmPerson):  void => { setPersons(prev  => prev.some(x => x.id === p.id) ? prev.map(x => x.id === p.id ? p : x) : [...prev, p]);   setPersonModal(null); };
  const deletePerson  = (id: string):    void => { if (confirm('Delete this person?'))  setPersons(prev  => prev.filter(x => x.id !== id)); };
  const saveCompany   = (c: CrmCompany): void => { setCompanies(prev => prev.some(x => x.id === c.id) ? prev.map(x => x.id === c.id ? c : x) : [...prev, c]); setCompanyModal(null); };
  const deleteCompany = (id: string):    void => { if (confirm('Delete this company?')) setCompanies(prev => prev.filter(x => x.id !== id)); };
  const companyName   = (id: string):    string => companies.find(c => c.id === id)?.name || '';

  const q = search.toLowerCase();
  const visPersons  = persons.filter(p  => !q || p.name.toLowerCase().includes(q) || companyName(p.organizationId).toLowerCase().includes(q));
  const visCompanies = companies.filter(c => !q || c.name.toLowerCase().includes(q) || c.labels.toLowerCase().includes(q) || c.address.toLowerCase().includes(q));

  const tabBtn = (active: boolean): React.CSSProperties => ({
    fontFamily: FF, fontWeight: 700, fontSize: 11.5, letterSpacing: '.07em', textTransform: 'uppercase',
    padding: '7px 22px', borderRadius: '4px 4px 0 0', cursor: 'pointer', transition: 'all .12s',
    background: active ? C.surface : 'transparent',
    borderTop:    active ? `2px solid ${C.green}` : '2px solid transparent',
    borderLeft:   active ? `1px solid ${C.border}` : '1px solid transparent',
    borderRight:  active ? `1px solid ${C.border}` : '1px solid transparent',
    borderBottom: active ? `1px solid ${C.surface}` : '1px solid transparent',
    color: active ? C.green : C.muted,
    marginBottom: active ? -1 : 0,
  });

  const ActionCell: React.FC<{ onEdit: () => void; onDel: () => void }> = ({ onEdit, onDel }) => (
    <td style={{ padding: '8px 12px', borderBottom: `1px solid ${C.border}`, whiteSpace: 'nowrap', width: 100 }}>
      <div style={{ display: 'flex', gap: 5 }}>
        <button onClick={onEdit} style={{ padding: '3px 11px', borderRadius: 3, border: 'none', background: C.purple, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 10.5, cursor: 'pointer' }}>Edit</button>
        <button onClick={onDel}  style={{ padding: '3px 11px', borderRadius: 3, border: `1px solid ${C.red}`, background: 'transparent', color: C.red, fontFamily: FF, fontWeight: 700, fontSize: 10.5, cursor: 'pointer' }}>Del</button>
      </div>
    </td>
  );

  const emptyRow = (colspan: number, msg: string): React.ReactNode => (
    <tr><td colSpan={colspan} style={{ padding: '40px 0', textAlign: 'center', fontFamily: FF, fontSize: 13, color: C.muted, borderBottom: `1px solid ${C.border}` }}>{msg}</td></tr>
  );

  return (
    <div style={{ background: C.bg, minHeight: 400, borderRadius: 8, padding: '0 0 24px 0' }}>

      {/* ── Toolbar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: 0, padding: '16px 0 0 0' }}>
        <div style={{ display: 'flex', gap: 0, borderBottom: `1px solid ${C.border}`, paddingBottom: 0 }}>
          <button style={tabBtn(tab === 'persons')}   onClick={() => { setTab('persons');   setSearch(''); }}>Persons</button>
          <button style={tabBtn(tab === 'companies')} onClick={() => { setTab('companies'); setSearch(''); }}>Companies</button>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', paddingBottom: 4 }}>
          <input
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder={`Search ${tab}…`}
            style={{ padding: '6px 10px', borderRadius: 4, border: `1px solid ${C.borderMd}`, background: C.surface, fontFamily: FF, fontSize: 12, color: C.text, outline: 'none', width: 200 }}
          />
          {tab === 'persons' ? (
            <button onClick={() => setPersonModal(emptyPerson())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}>+ Person</button>
          ) : (
            <button onClick={() => setCompanyModal(emptyCompany())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: C.green, color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer', whiteSpace: 'nowrap' }}>+ Company</button>
          )}
        </div>
      </div>

      {/* ── Table */}
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderTop: 'none', borderRadius: '0 0 8px 8px', overflowX: 'auto' }}>

        {/* Persons table */}
        {tab === 'persons' && (
          <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 700 }}>
            <thead>
              <tr>
                <Th label="#"            w={40} />
                <Th label="Name"         />
                <Th label="Organization" />
                <Th label="Phone"        />
                <Th label="Email"        />
                <Th label="Actions"      w={100} />
              </tr>
            </thead>
            <tbody>
              {visPersons.length === 0
                ? emptyRow(6, persons.length === 0 ? 'No persons yet — click + Person to add one.' : 'No results match your search.')
                : visPersons.map((p, idx) => (
                  <tr key={p.id}
                    onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                    onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
                  >
                    <Td muted>{idx + 1}</Td>
                    <Td><span style={{ fontWeight: 700 }}>{p.name}</span></Td>
                    <Td>
                      {p.organizationId ? (
                        <span style={{ color: C.green, fontWeight: 600 }}>{companyName(p.organizationId)}</span>
                      ) : <span style={{ color: C.muted }}>—</span>}
                    </Td>
                    <Td muted>{firstVal(p.phones)}</Td>
                    <Td muted>{firstVal(p.emails)}</Td>
                    <ActionCell onEdit={() => setPersonModal({ ...p })} onDel={() => deletePerson(p.id)} />
                  </tr>
                ))
              }
            </tbody>
          </table>
        )}

        {/* Companies table */}
        {tab === 'companies' && (
          <table style={{ width: '100%', borderCollapse: 'collapse', minWidth: 800 }}>
            <thead>
              <tr>
                <Th label="#"       w={40} />
                <Th label="Company" />
                <Th label="Labels"  />
                <Th label="Phone"   />
                <Th label="Email"   />
                <Th label="Address" />
                <Th label="Actions" w={100} />
              </tr>
            </thead>
            <tbody>
              {visCompanies.length === 0
                ? emptyRow(7, companies.length === 0 ? 'No companies yet — click + Company to add one.' : 'No results match your search.')
                : visCompanies.map((c, idx) => {
                  const linked = persons.filter(p => p.organizationId === c.id);
                  return (
                    <tr key={c.id}
                      onMouseEnter={e => { (e.currentTarget as HTMLTableRowElement).style.background = C.rowHover; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLTableRowElement).style.background = 'transparent'; }}
                    >
                      <Td muted>{idx + 1}</Td>
                      <Td>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                          <span style={{ fontWeight: 700 }}>{c.name}</span>
                          {linked.length > 0 && (
                            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                              {linked.map(p => (
                                <span key={p.id} style={{ fontSize: 10, fontFamily: FF, color: C.green, background: C.greenBg, border: `1px solid ${C.greenBd}`, borderRadius: 3, padding: '1px 6px', cursor: 'pointer' }} onClick={() => setTab('persons')}>
                                  {p.name}
                                </span>
                              ))}
                            </div>
                          )}
                        </div>
                      </Td>
                      <Td>
                        <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                          {c.labels
                            ? c.labels.split(',').map(l => l.trim()).filter(Boolean).map(l => (
                                <span key={l} style={{ fontSize: 10, fontFamily: FF, color: C.tagText, background: C.tag, border: `1px solid ${C.border}`, borderRadius: 3, padding: '1px 6px' }}>{l}</span>
                              ))
                            : <span style={{ color: C.muted }}>—</span>
                          }
                        </div>
                      </Td>
                      <Td muted>{firstVal(c.phones)}</Td>
                      <Td muted>{firstVal(c.emails)}</Td>
                      <Td muted><span title={c.address}>{c.address || '—'}</span></Td>
                      <ActionCell onEdit={() => setCompanyModal({ ...c })} onDel={() => deleteCompany(c.id)} />
                    </tr>
                  );
                })
              }
            </tbody>
          </table>
        )}

        {/* Row count footer */}
        <div style={{ padding: '8px 14px', borderTop: `1px solid ${C.border}`, fontFamily: FF, fontSize: 11, color: C.muted, background: C.thBg, borderRadius: '0 0 8px 8px' }}>
          {tab === 'persons'
            ? `${visPersons.length} of ${persons.length} person${persons.length !== 1 ? 's' : ''}`
            : `${visCompanies.length} of ${companies.length} compan${companies.length !== 1 ? 'ies' : 'y'}`
          }
        </div>
      </div>

      {/* ── Modals */}
      {personModal  && <PersonModal  initial={personModal}  companies={companies} onSave={savePerson}  onClose={() => setPersonModal(null)}  />}
      {companyModal && <CompanyModal initial={companyModal}                        onSave={saveCompany} onClose={() => setCompanyModal(null)} />}
    </div>
  );
};

export default CrmBoard;
