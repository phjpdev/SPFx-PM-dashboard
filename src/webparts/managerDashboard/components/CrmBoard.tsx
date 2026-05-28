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

const PHONE_TYPES = ['Work', 'Home', 'Mobile', 'Other'];
const EMAIL_TYPES = ['Work', 'Home', 'Other'];
const FF = 'Montserrat,sans-serif';
const LS_PERSONS  = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;

const emptyPerson  = (): CrmPerson  => ({ id: uid(), name: '', organizationId: '', phones: [{ value: '', type: 'Work' }], emails: [{ value: '', type: 'Work' }] });
const emptyCompany = (): CrmCompany => ({ id: uid(), name: '', labels: '', address: '', phones: [{ value: '', type: 'Work' }], emails: [{ value: '', type: 'Work' }] });

const loadLS = <T,>(key: string, fallback: T): T => {
  try { const v = localStorage.getItem(key); return v ? (JSON.parse(v) as T) : fallback; } catch { return fallback; }
};

// ── Shared styles ─────────────────────────────────────────────────
const inp: React.CSSProperties = {
  padding: '8px 10px', background: '#151a26', border: '1px solid rgba(138,155,176,.25)',
  borderRadius: 4, color: '#d3d1c7', fontSize: 12, fontFamily: FF,
  width: '100%', boxSizing: 'border-box', outline: 'none',
};
const lbl: React.CSSProperties = {
  fontSize: 10, fontWeight: 700, color: '#8a9bb0', letterSpacing: '.08em',
  textTransform: 'uppercase', marginBottom: 4, display: 'block', fontFamily: FF,
};

// ── Modal wrapper ─────────────────────────────────────────────────
const Modal: React.FC<{ title: string; onClose: () => void; children: React.ReactNode }> = ({ title, onClose, children }) => (
  <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.65)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
    <div style={{ background: '#1b2133', borderRadius: 8, width: 440, maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 8px 32px rgba(0,0,0,.55)', border: '1px solid rgba(138,155,176,.18)' }}>
      <div style={{ padding: '14px 18px', borderBottom: '1px solid rgba(138,155,176,.15)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 14, color: '#d3d1c7' }}>{title}</span>
        <button onClick={onClose} style={{ background: 'none', border: 'none', color: '#8a9bb0', cursor: 'pointer', fontSize: 20, lineHeight: 1, padding: 0 }}>×</button>
      </div>
      <div style={{ padding: '18px 20px' }}>{children}</div>
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
        <input
          value={item.value}
          onChange={e => { const n = [...items]; n[i] = { ...n[i], value: e.target.value }; onChange(n); }}
          placeholder={placeholder}
          style={{ ...inp, flex: 1 }}
        />
        <select
          value={item.type}
          onChange={e => { const n = [...items]; n[i] = { ...n[i], type: e.target.value }; onChange(n); }}
          style={{ ...inp, width: 88, flex: 'none' }}
        >
          {types.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        {items.length > 1 && (
          <button onClick={() => onChange(items.filter((_, j) => j !== i))} style={{ background: 'none', border: 'none', color: '#8a9bb0', cursor: 'pointer', fontSize: 18, lineHeight: 1, padding: '0 2px', flexShrink: 0 }}>×</button>
        )}
      </div>
    ))}
    <button
      onClick={() => onChange([...items, { value: '', type: types[0] }])}
      style={{ background: 'none', border: 'none', color: '#3ebd3e', fontFamily: FF, fontSize: 12, cursor: 'pointer', padding: 0 }}
    >
      + Add {addLabel}
    </button>
  </div>
);

// ── Save / Cancel row ─────────────────────────────────────────────
const ModalFooter: React.FC<{ onCancel: () => void; onSave: () => void }> = ({ onCancel, onSave }) => (
  <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end', marginTop: 20, paddingTop: 14, borderTop: '1px solid rgba(138,155,176,.12)' }}>
    <button onClick={onCancel} style={{ padding: '8px 20px', borderRadius: 4, border: '1px solid rgba(138,155,176,.25)', background: 'transparent', color: '#8a9bb0', fontFamily: FF, fontSize: 12, cursor: 'pointer' }}>Cancel</button>
    <button onClick={onSave} style={{ padding: '8px 20px', borderRadius: 4, border: 'none', background: '#2a9e2a', color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>Save</button>
  </div>
);

// ── Person Modal ──────────────────────────────────────────────────
const PersonModal: React.FC<{
  initial: CrmPerson;
  companies: CrmCompany[];
  onSave: (p: CrmPerson) => void;
  onClose: () => void;
}> = ({ initial, companies, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmPerson>(initial);
  const set = <K extends keyof CrmPerson>(k: K, v: CrmPerson[K]): void => setD(p => ({ ...p, [k]: v }));
  const isEdit = !!initial.name;
  return (
    <Modal title={isEdit ? 'Edit Person' : 'Add Person'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={inp} placeholder="Full name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Organization</label>
        <div style={{ position: 'relative' }}>
          <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: '#8a9bb0', pointerEvents: 'none' }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-4 0v2"/><line x1="12" y1="12" x2="12" y2="16"/>
            </svg>
          </span>
          <select value={d.organizationId} onChange={e => set('organizationId', e.target.value)} style={{ ...inp, paddingLeft: 28 }}>
            <option value="">— None —</option>
            {companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </div>
      </div>
      <label style={lbl}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" onChange={v => set('phones', v)} />
      <label style={lbl}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />
      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(d); }} />
    </Modal>
  );
};

// ── Company Modal ─────────────────────────────────────────────────
const CompanyModal: React.FC<{
  initial: CrmCompany;
  onSave: (c: CrmCompany) => void;
  onClose: () => void;
}> = ({ initial, onSave, onClose }) => {
  const [d, setD] = React.useState<CrmCompany>(initial);
  const set = <K extends keyof CrmCompany>(k: K, v: CrmCompany[K]): void => setD(p => ({ ...p, [k]: v }));
  const isEdit = !!initial.name;
  return (
    <Modal title={isEdit ? 'Edit Company' : 'Add Company'} onClose={onClose}>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Name</label>
        <input value={d.name} onChange={e => set('name', e.target.value)} style={inp} placeholder="Company name" autoFocus />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Labels</label>
        <input value={d.labels} onChange={e => set('labels', e.target.value)} style={inp} placeholder="e.g. Client, Supplier, Partner" />
      </div>
      <div style={{ marginBottom: 14 }}>
        <label style={lbl}>Address</label>
        <textarea value={d.address} onChange={e => set('address', e.target.value)} style={{ ...inp, height: 68, resize: 'vertical' }} placeholder="Street address" />
      </div>
      <label style={lbl}>Phone</label>
      <MultiField items={d.phones} types={PHONE_TYPES} addLabel="phone" placeholder="Phone number" onChange={v => set('phones', v)} />
      <label style={lbl}>Email</label>
      <MultiField items={d.emails} types={EMAIL_TYPES} addLabel="email" placeholder="Email address" onChange={v => set('emails', v)} />
      <ModalFooter onCancel={onClose} onSave={() => { if (d.name.trim()) onSave(d); }} />
    </Modal>
  );
};

// ── Action buttons (Edit / Del) ───────────────────────────────────
const ActionBtn: React.FC<{ label: string; variant: 'edit' | 'del'; onClick: () => void }> = ({ label, variant, onClick }) => (
  <button onClick={onClick} style={{
    padding: '4px 12px', borderRadius: 3, fontSize: 11, fontFamily: FF, fontWeight: 700, cursor: 'pointer',
    border: variant === 'del' ? '1px solid #c0392b' : 'none',
    background: variant === 'edit' ? '#6c3fbf' : 'transparent',
    color: variant === 'edit' ? '#fff' : '#c0392b',
  }}>{label}</button>
);

// ── CrmBoard (main export) ────────────────────────────────────────
const CrmBoard: React.FC = () => {
  const [tab, setTab]               = React.useState<'persons' | 'companies'>('persons');
  const [persons, setPersons]       = React.useState<CrmPerson[]>(() => loadLS<CrmPerson[]>(LS_PERSONS, []));
  const [companies, setCompanies]   = React.useState<CrmCompany[]>(() => loadLS<CrmCompany[]>(LS_COMPANIES, []));
  const [personModal, setPersonModal]   = React.useState<CrmPerson | null>(null);
  const [companyModal, setCompanyModal] = React.useState<CrmCompany | null>(null);

  React.useEffect(() => { localStorage.setItem(LS_PERSONS,   JSON.stringify(persons));   }, [persons]);
  React.useEffect(() => { localStorage.setItem(LS_COMPANIES, JSON.stringify(companies)); }, [companies]);

  const savePerson = (p: CrmPerson): void => {
    setPersons(prev => prev.some(x => x.id === p.id) ? prev.map(x => x.id === p.id ? p : x) : [...prev, p]);
    setPersonModal(null);
  };
  const deletePerson = (id: string): void => {
    if (confirm('Delete this person?')) setPersons(prev => prev.filter(x => x.id !== id));
  };

  const saveCompany = (c: CrmCompany): void => {
    setCompanies(prev => prev.some(x => x.id === c.id) ? prev.map(x => x.id === c.id ? c : x) : [...prev, c]);
    setCompanyModal(null);
  };
  const deleteCompany = (id: string): void => {
    if (confirm('Delete this company?')) setCompanies(prev => prev.filter(x => x.id !== id));
  };

  const companyName = (id: string): string => companies.find(c => c.id === id)?.name || '';

  const tabStyle = (active: boolean): React.CSSProperties => ({
    fontFamily: FF, fontWeight: 700, fontSize: 12, letterSpacing: '.08em', textTransform: 'uppercase',
    padding: '7px 22px', borderRadius: 4, cursor: 'pointer',
    background: active ? 'rgba(42,158,42,.14)' : 'transparent',
    border: active ? '1px solid #2a9e2a' : '1px solid transparent',
    color: active ? '#3ebd3e' : '#8a9bb0', transition: 'all .15s',
  });

  const metaLine = (icon: string, text: string): React.ReactNode => (
    <span style={{ fontFamily: FF, fontSize: 11, color: '#8a9bb0', display: 'flex', alignItems: 'center', gap: 4 }}>
      <span style={{ opacity: .7 }}>{icon}</span>{text}
    </span>
  );

  return (
    <div>
      {/* ── Header bar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
        <div style={{ display: 'flex', gap: 4 }}>
          <button style={tabStyle(tab === 'persons')}   onClick={() => setTab('persons')}>Persons</button>
          <button style={tabStyle(tab === 'companies')} onClick={() => setTab('companies')}>Companies</button>
        </div>
        {tab === 'persons' ? (
          <button onClick={() => setPersonModal(emptyPerson())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: '#2a9e2a', color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>
            + Person
          </button>
        ) : (
          <button onClick={() => setCompanyModal(emptyCompany())} style={{ padding: '7px 16px', borderRadius: 4, border: 'none', background: '#2a9e2a', color: '#fff', fontFamily: FF, fontWeight: 700, fontSize: 12, cursor: 'pointer' }}>
            + Company
          </button>
        )}
      </div>

      {/* ── Persons list */}
      {tab === 'persons' && (
        persons.length === 0
          ? <div style={{ padding: '52px 0', textAlign: 'center', color: '#5a6a80', fontFamily: FF, fontSize: 13 }}>No persons yet — click + Person to add one.</div>
          : <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {persons.map(p => (
                <div key={p.id} style={{ background: '#1b2133', borderRadius: 6, border: '1px solid rgba(138,155,176,.15)', padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12 }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 0 }}>
                    <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 13, color: '#d3d1c7' }}>{p.name}</span>
                    {p.organizationId && metaLine('🏢', companyName(p.organizationId))}
                    {p.phones.filter(x => x.value).map((x, i) => <span key={i}>{metaLine('📞', `${x.value} · ${x.type}`)}</span>)}
                    {p.emails.filter(x => x.value).map((x, i) => <span key={i}>{metaLine('✉', `${x.value} · ${x.type}`)}</span>)}
                  </div>
                  <div style={{ display: 'flex', gap: 6, flexShrink: 0, marginTop: 2 }}>
                    <ActionBtn label="Edit" variant="edit" onClick={() => setPersonModal({ ...p })} />
                    <ActionBtn label="Del"  variant="del"  onClick={() => deletePerson(p.id)} />
                  </div>
                </div>
              ))}
            </div>
      )}

      {/* ── Companies list */}
      {tab === 'companies' && (
        companies.length === 0
          ? <div style={{ padding: '52px 0', textAlign: 'center', color: '#5a6a80', fontFamily: FF, fontSize: 13 }}>No companies yet — click + Company to add one.</div>
          : <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              {companies.map(c => {
                const linked = persons.filter(p => p.organizationId === c.id);
                return (
                  <div key={c.id} style={{ background: '#1b2133', borderRadius: 6, border: '1px solid rgba(138,155,176,.15)', padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 12 }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 0 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span style={{ fontFamily: FF, fontWeight: 700, fontSize: 13, color: '#d3d1c7' }}>{c.name}</span>
                        {c.labels && c.labels.split(',').map(l => l.trim()).filter(Boolean).map(l => (
                          <span key={l} style={{ fontFamily: FF, fontSize: 10, fontWeight: 700, color: '#8a9bb0', background: 'rgba(138,155,176,.12)', border: '1px solid rgba(138,155,176,.2)', borderRadius: 3, padding: '1px 6px' }}>{l}</span>
                        ))}
                      </div>
                      {c.address && metaLine('📍', c.address)}
                      {c.phones.filter(x => x.value).map((x, i) => <span key={i}>{metaLine('📞', `${x.value} · ${x.type}`)}</span>)}
                      {c.emails.filter(x => x.value).map((x, i) => <span key={i}>{metaLine('✉', `${x.value} · ${x.type}`)}</span>)}
                      {linked.length > 0 && (
                        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 2 }}>
                          {linked.map(p => (
                            <span key={p.id} onClick={() => { setTab('persons'); }} style={{ fontFamily: FF, fontSize: 10, color: '#3ebd3e', background: 'rgba(42,158,42,.1)', border: '1px solid rgba(42,158,42,.25)', borderRadius: 3, padding: '1px 7px', cursor: 'pointer' }}>
                              {p.name}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                    <div style={{ display: 'flex', gap: 6, flexShrink: 0, marginTop: 2 }}>
                      <ActionBtn label="Edit" variant="edit" onClick={() => setCompanyModal({ ...c })} />
                      <ActionBtn label="Del"  variant="del"  onClick={() => deleteCompany(c.id)} />
                    </div>
                  </div>
                );
              })}
            </div>
      )}

      {/* ── Modals */}
      {personModal  && <PersonModal  initial={personModal}  companies={companies} onSave={savePerson}  onClose={() => setPersonModal(null)}  />}
      {companyModal && <CompanyModal initial={companyModal}                        onSave={saveCompany} onClose={() => setCompanyModal(null)} />}
    </div>
  );
};

export default CrmBoard;
