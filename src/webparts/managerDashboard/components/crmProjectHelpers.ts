import type { IProject } from '../../../shared/models/IProject';
import type { CrmCompany, CrmPerson, CrmProject, CrmQuote } from './crmTypes';

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

const todayIso = (): string => {
  const t = new Date();
  return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`;
};

export const parse3eNum = (projNum: string): number | null => {
  const m = projNum.match(/^3E-(\d+)$/i);
  return m ? parseInt(m[1], 10) : null;
};

/** Next 3E project number — checks CRM projects and main dashboard projects. */
export const nextProjNum = (
  crmProjects: CrmProject[],
  dashboardProjects: { projNum: string }[] = [],
): string => {
  const nums: number[] = [];
  [...crmProjects, ...dashboardProjects].forEach(p => {
    const n = parse3eNum(p.projNum);
    if (n !== null) nums.push(n);
  });
  const next = nums.length ? Math.max(...nums) + 1 : 500;
  return `3E-${next}`;
};

export const quoteCoreFromQuote = (q: CrmQuote): Omit<CrmProject, 'id' | 'projNum' | 'spProjectId' | 'wonDate'> => ({
  quoteId: q.id,
  quoteNum: q.quoteNum,
  rfqId: q.rfqId,
  rfqNum: q.rfqNum,
  dateReceived: q.dateReceived,
  personId: q.personId,
  organizationId: q.organizationId,
  projectTitle: q.projectTitle,
  projectAddress: q.projectAddress,
  discipline: q.discipline,
  quoteRequiredBy: q.quoteRequiredBy,
  projectValue: q.projectValue,
  approximateHours: q.approximateHours,
  engineerDrawingReceived: q.engineerDrawingReceived,
  engineerDrawingDate: q.engineerDrawingDate,
  revisionVersionEng: q.revisionVersionEng,
  architectDrawingReceived: q.architectDrawingReceived,
  architectDrawingDate: q.architectDrawingDate,
  revisionVersionArch: q.revisionVersionArch,
  rfiAllowed: q.rfiAllowed,
  assignedTo: q.assignedTo,
  source: q.source,
  notes: q.notes,
});

export const quoteToCrmProject = (q: CrmQuote, projNum: string, spProjectId?: number): CrmProject => ({
  id: uid(),
  projNum,
  spProjectId,
  wonDate: todayIso(),
  ...quoteCoreFromQuote(q),
});

export const quoteToIProject = (
  q: CrmQuote,
  projNum: string,
  persons: CrmPerson[],
  companies: CrmCompany[],
): IProject => {
  const person = persons.find(p => p.id === q.personId);
  const company = companies.find(c => c.id === q.organizationId);
  const discipline = q.discipline === 'Both' ? 'Steel' : q.discipline;
  return {
    id: projNum,
    projNum,
    name: q.projectTitle,
    discipline,
    status: 'Active',
    year: new Date().getFullYear(),
    hrsAllowed: q.approximateHours || 0,
    hrsUsed: 0,
    rfisAllowed: q.rfiAllowed || 0,
    quoteNum: q.quoteNum || '',
    contact: person?.name || '',
    company: company?.name || '',
    email: person?.emails?.[0]?.value || '',
    mobile: person?.phones?.[0]?.value || '',
    clientNum: '',
    clientp0: '',
    startDate: todayIso(),
    finishDate: '',
    ifaDate: '',
    ifcDate: '',
    detailers: '',
    teamLead: q.assignedTo || '',
    teamMembers: '',
    notes: q.notes || '',
    invoices: [],
    isEwo: false,
    ewoNum: '',
    parentId: null,
  };
};

/** RFQ numbers already used across RFQs, quotes, and won projects. */
export const nextRfqNum = (
  rfqs: { rfqNum: string }[],
  quotes: { rfqNum: string }[] = [],
  projects: { rfqNum: string }[] = [],
): string => {
  const year = new Date().getFullYear().toString().slice(-2);
  const prefix = `RFQ-${year}-`;
  const nums: number[] = [];
  [...rfqs, ...quotes, ...projects].forEach(r => {
    if (r.rfqNum?.startsWith(prefix)) {
      const n = parseInt(r.rfqNum.slice(prefix.length), 10);
      if (!isNaN(n)) nums.push(n);
    }
  });
  const next = nums.length ? Math.max(...nums) + 1 : 1;
  return `${prefix}${String(next).padStart(3, '0')}`;
};
