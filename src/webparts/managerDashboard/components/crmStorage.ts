import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmRfq, CrmQuote, CrmProject } from './crmTypes';
import { normalizeCrmQuote, normalizeCrmRfq } from './crmRfqNormalize';

export interface CrmQuoteBudget {
  year: number;
  yearTarget: number;
  /** @deprecated use monthTargets — kept for legacy saved budgets */
  monthTarget: number;
  /** Jan–Dec budget targets ($), index 0 = January */
  monthTargets?: number[];
}

export interface CrmQuoteBudgetYearData {
  yearTarget: number;
  monthTargets: number[];
}

export interface CrmQuoteBudgetStore {
  byYear: Record<string, CrmQuoteBudgetYearData>;
}

const MONTH_LABELS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export function normalizeQuoteBudget(raw: CrmQuoteBudget | null | undefined): CrmQuoteBudget {
  const year = raw?.year || new Date().getFullYear();
  const yearTarget = raw?.yearTarget || 0;
  const legacy = raw?.monthTarget || 0;
  const monthTargets = raw?.monthTargets?.length === 12
    ? [...raw.monthTargets]
    : Array(12).fill(legacy);
  return { year, yearTarget, monthTarget: monthTargets[new Date().getMonth()] || 0, monthTargets };
}

export function getMonthBudgetTarget(budget: CrmQuoteBudget, monthIndex: number): number {
  const normalized = normalizeQuoteBudget(budget);
  return normalized.monthTargets?.[monthIndex] ?? 0;
}

export function normalizeQuoteBudgetStore(raw: unknown): CrmQuoteBudgetStore {
  if (raw && typeof raw === 'object' && 'byYear' in raw) {
    const store = raw as CrmQuoteBudgetStore;
    return { byYear: { ...store.byYear } };
  }
  const legacy = raw as CrmQuoteBudget | null | undefined;
  if (legacy?.year) {
    const n = normalizeQuoteBudget(legacy);
    return {
      byYear: {
        [String(n.year)]: { yearTarget: n.yearTarget, monthTargets: n.monthTargets || Array(12).fill(0) },
      },
    };
  }
  return { byYear: {} };
}

export function getQuoteBudgetForYear(store: CrmQuoteBudgetStore, year: number): CrmQuoteBudget {
  const data = store.byYear[String(year)];
  if (!data) return normalizeQuoteBudget({ year, yearTarget: 0, monthTarget: 0 });
  return normalizeQuoteBudget({ year, yearTarget: data.yearTarget, monthTarget: 0, monthTargets: data.monthTargets });
}

export function setQuoteBudgetForYear(store: CrmQuoteBudgetStore, budget: CrmQuoteBudget): CrmQuoteBudgetStore {
  const n = normalizeQuoteBudget(budget);
  return {
    byYear: {
      ...store.byYear,
      [String(n.year)]: { yearTarget: n.yearTarget, monthTargets: n.monthTargets || Array(12).fill(0) },
    },
  };
}

export { MONTH_LABELS };

/**
 * Kept for API compatibility with CrmBoard — the revision field is used to
 * detect concurrent saves. With proper SP lists each record has its own
 * version, so we return a simple timestamp-based value.
 */
export interface CrmDelta {
  revision: number;
  updatedAt: string;
  persons: Record<string, CrmPerson>;
  companies: Record<string, CrmCompany>;
  deletedPersonIds: string[];
  deletedCompanyIds: string[];
}

export interface CrmListsSnapshot {
  persons: CrmPerson[];
  companies: CrmCompany[];
}

const CRM_SP_QUOTE_BUDGET = 'crm_quote_budget';

const LS_PERSONS   = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';
const LS_RFQS      = '3edge-crm-rfqs';
const LS_QUOTES    = '3edge-crm-quotes';
const LS_PROJECTS  = '3edge-crm-projects';
const LS_QUOTE_BUDGET = '3edge-crm-quote-budget';

const loadLS = <T,>(k: string, fb: T): T => {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : fb;
  } catch { return fb; }
};

/** Overlay base64 attachments (stored locally only — too large for SP columns). */
const overlayLocalAttachments = (persons: CrmPerson[], local: CrmPerson[]): CrmPerson[] => {
  const attMap = new Map<string, CrmPerson['attachments']>();
  for (const p of local) {
    if (p.attachments?.length) attMap.set(p.id, p.attachments);
  }
  if (!attMap.size) return persons;
  return persons.map(p => {
    const a = attMap.get(p.id);
    return a ? { ...p, attachments: a } : p;
  });
};

/** SP wins on same id; local-only rows not yet synced are kept. */
export function mergeCrmRemoteList<T extends { id: string }>(remote: T[], local: T[]): T[] {
  const map = new Map<string, T>();
  local.forEach(x => map.set(x.id, x));
  remote.forEach(x => map.set(x.id, x));
  return Array.from(map.values());
}

// ── Persons & Companies ────────────────────────────────────────────────────

export async function loadCrmPersonsCompanies(sp: SharePointService): Promise<CrmListsSnapshot> {
  try {
    const [persons, companies] = await Promise.all([
      sp.loadCrmPersons(),
      sp.loadCrmCompanies(),
    ]);

    // Overlay attachments that are kept in browser storage only
    const local = loadLS<CrmPerson[] | null>(LS_PERSONS, null);
    const merged = local?.length ? overlayLocalAttachments(persons, local) : persons;
    try {
      localStorage.setItem(LS_PERSONS, JSON.stringify(merged));
      localStorage.setItem(LS_COMPANIES, JSON.stringify(companies));
    } catch { /* ignore */ }
    return { persons: merged, companies };
  } catch {
    // SP unavailable — fall back to local cache
    const cachedP = loadLS<CrmPerson[] | null>(LS_PERSONS, null);
    const cachedC = loadLS<CrmCompany[] | null>(LS_COMPANIES, null);
    return { persons: cachedP ?? [], companies: cachedC ?? [] };
  }
}

/** Returns a stub CrmDelta to satisfy CrmBoard's revision-tracking API. */
export async function loadCrmDelta(_sp: SharePointService): Promise<CrmDelta | null> {
  return {
    revision: 1,
    updatedAt: new Date().toISOString(),
    persons: {},
    companies: {},
    deletedPersonIds: [],
    deletedCompanyIds: [],
  };
}

export async function saveCrmPersonsCompanies(
  sp: SharePointService,
  persons: CrmPerson[],
  companies: CrmCompany[],
  _prevRevision = 0,
): Promise<{ ok: boolean; revision: number }> {
  try {
    localStorage.setItem(LS_PERSONS, JSON.stringify(persons));
    localStorage.setItem(LS_COMPANIES, JSON.stringify(companies));
  } catch { /* ignore */ }

  try {
    await Promise.all([
      sp.syncCrmPersons(persons),
      sp.syncCrmCompanies(companies),
    ]);
    return { ok: true, revision: Date.now() };
  } catch {
    return { ok: false, revision: 0 };
  }
}

/**
 * One-time import: clear existing SP rows then add each record individually.
 * Sequential single-record saves avoid SP throttling on large lists.
 */
export async function importCrmPersonsCompanies(
  sp: SharePointService,
  persons: CrmPerson[],
  companies: CrmCompany[],
  onProgress: (done: number, total: number) => void,
): Promise<{ persons: CrmPerson[]; companies: CrmCompany[] }> {
  const total = companies.length + persons.length;
  let done = 0;

  await sp.deleteAllCrmCompanies();
  await sp.deleteAllCrmPersons();

  const savedCompanies: CrmCompany[] = [];
  for (const c of companies) {
    try {
      const spId = await sp.addCrmCompany(c);
      savedCompanies.push({ ...c, spId });
    } catch { savedCompanies.push(c); }
    done++;
    onProgress(done, total);
  }

  const savedPersons: CrmPerson[] = [];
  for (const p of persons) {
    try {
      const spId = await sp.addCrmPerson(p);
      savedPersons.push({ ...p, spId });
    } catch { savedPersons.push(p); }
    done++;
    onProgress(done, total);
  }

  try {
    localStorage.setItem(LS_PERSONS, JSON.stringify(savedPersons));
    localStorage.setItem(LS_COMPANIES, JSON.stringify(savedCompanies));
  } catch { /* ignore */ }

  return { persons: savedPersons, companies: savedCompanies };
}

// ── RFQs ───────────────────────────────────────────────────────────────────

export async function loadRfqsRemote(sp: SharePointService): Promise<CrmRfq[] | null> {
  try {
    const rows = await sp.loadCrmRfqs();
    return rows.map(r => normalizeCrmRfq(r as CrmRfq & Record<string, unknown>));
  } catch { return null; }
}

export async function loadRfqsFromSharePoint(sp: SharePointService): Promise<CrmRfq[]> {
  try {
    const rfqs = (await sp.loadCrmRfqs()).map(
      r => normalizeCrmRfq(r as CrmRfq & Record<string, unknown>),
    );
    try { localStorage.setItem(LS_RFQS, JSON.stringify(rfqs)); } catch { /* ignore */ }
    return rfqs;
  } catch {
    return loadLS<CrmRfq[]>(LS_RFQS, []).map(
      r => normalizeCrmRfq(r as CrmRfq & Record<string, unknown>),
    );
  }
}

export async function saveRfqsToSharePoint(sp: SharePointService, rfqs: CrmRfq[]): Promise<void> {
  try { localStorage.setItem(LS_RFQS, JSON.stringify(rfqs)); } catch { /* ignore */ }
  try { await sp.syncCrmRfqs(rfqs); } catch { /* ignore */ }
}

// ── Quotes ─────────────────────────────────────────────────────────────────

export async function loadQuotesRemote(sp: SharePointService): Promise<CrmQuote[] | null> {
  try {
    const rows = await sp.loadCrmQuotes();
    return rows.map(q => normalizeCrmQuote(q as CrmQuote & Record<string, unknown>));
  } catch { return null; }
}

export async function loadQuotesFromSharePoint(sp: SharePointService): Promise<CrmQuote[]> {
  try {
    const quotes = (await sp.loadCrmQuotes()).map(
      q => normalizeCrmQuote(q as CrmQuote & Record<string, unknown>),
    );
    try { localStorage.setItem(LS_QUOTES, JSON.stringify(quotes)); } catch { /* ignore */ }
    return quotes;
  } catch {
    return loadLS<CrmQuote[]>(LS_QUOTES, []).map(
      q => normalizeCrmQuote(q as CrmQuote & Record<string, unknown>),
    );
  }
}

export async function saveQuotesToSharePoint(sp: SharePointService, quotes: CrmQuote[]): Promise<void> {
  try { localStorage.setItem(LS_QUOTES, JSON.stringify(quotes)); } catch { /* ignore */ }
  try { await sp.syncCrmQuotes(quotes); } catch { /* ignore */ }
}

// ── WIP Projects ───────────────────────────────────────────────────────────

export async function loadProjectsFromSharePoint(sp: SharePointService): Promise<CrmProject[]> {
  try {
    const projects = await sp.loadCrmWip();
    try { localStorage.setItem(LS_PROJECTS, JSON.stringify(projects)); } catch { /* ignore */ }
    return projects;
  } catch {
    return loadLS<CrmProject[]>(LS_PROJECTS, []);
  }
}

export async function saveProjectsToSharePoint(sp: SharePointService, projects: CrmProject[]): Promise<void> {
  try { localStorage.setItem(LS_PROJECTS, JSON.stringify(projects)); } catch { /* ignore */ }
  try { await sp.syncCrmWip(projects); } catch { /* ignore */ }
}

// ── Quote Budget ───────────────────────────────────────────────────────────

const defaultQuoteBudgetStore = (): CrmQuoteBudgetStore => ({ byYear: {} });

export async function loadQuoteBudget(sp: SharePointService): Promise<CrmQuoteBudgetStore> {
  const local = loadLS<CrmQuoteBudgetStore | CrmQuoteBudget | null>(LS_QUOTE_BUDGET, null);
  try {
    const json = await sp.getSettingChunks(CRM_SP_QUOTE_BUDGET);
    if (!json) return normalizeQuoteBudgetStore(local ?? defaultQuoteBudgetStore());
    return normalizeQuoteBudgetStore(JSON.parse(json));
  } catch {
    return normalizeQuoteBudgetStore(local ?? defaultQuoteBudgetStore());
  }
}

export async function saveQuoteBudget(sp: SharePointService, store: CrmQuoteBudgetStore): Promise<void> {
  const normalized = normalizeQuoteBudgetStore(store);
  try { localStorage.setItem(LS_QUOTE_BUDGET, JSON.stringify(normalized)); } catch { /* ignore */ }
  try { await sp.setSettingChunks(CRM_SP_QUOTE_BUDGET, JSON.stringify(normalized)); } catch { /* ignore */ }
}

// ── Per-record person / company operations ─────────────────────────────────

export async function upsertPersonToSharePoint(sp: SharePointService, person: CrmPerson): Promise<CrmPerson> {
  try {
    if (person.spId) { await sp.updateCrmPerson(person.spId, person); return person; }
    const spId = await sp.addCrmPerson(person);
    return { ...person, spId };
  } catch { return person; }
}

export async function upsertCompanyToSharePoint(sp: SharePointService, company: CrmCompany): Promise<CrmCompany> {
  try {
    if (company.spId) { await sp.updateCrmCompany(company.spId, company); return company; }
    const spId = await sp.addCrmCompany(company);
    return { ...company, spId };
  } catch { return company; }
}

export async function deletePersonFromSharePoint(sp: SharePointService, person: Pick<CrmPerson, 'spId'>): Promise<void> {
  if (person.spId) {
    try { await sp.deleteCrmPersonById(person.spId); } catch { /* ignore */ }
  }
}

export async function deleteCompanyFromSharePoint(sp: SharePointService, company: Pick<CrmCompany, 'spId'>): Promise<void> {
  if (company.spId) {
    try { await sp.deleteCrmCompanyById(company.spId); } catch { /* ignore */ }
  }
}

// ── Per-action SP operations (fast path — no bulk sync) ────────────────────

export async function upsertRfqToSharePoint(sp: SharePointService, rfq: CrmRfq): Promise<CrmRfq> {
  try {
    if (rfq.spId) { await sp.updateCrmRfq(rfq.spId, rfq); return rfq; }
    const spId = await sp.addCrmRfq(rfq);
    return { ...rfq, spId };
  } catch { return rfq; }
}

export async function deleteRfqFromSharePoint(sp: SharePointService, rfq: Pick<CrmRfq, 'id' | 'spId'>): Promise<void> {
  if (rfq.spId) {
    try { await sp.deleteCrmRfqById(rfq.spId); } catch { /* ignore */ }
  }
  if (rfq.id) {
    try { await sp.deleteAllCrmRfqsByCrmId(rfq.id); } catch { /* ignore */ }
  }
}

export async function upsertQuoteToSharePoint(sp: SharePointService, quote: CrmQuote): Promise<CrmQuote> {
  try {
    if (quote.spId) { await sp.updateCrmQuote(quote.spId, quote); return quote; }
    const spId = await sp.addCrmQuote(quote);
    return { ...quote, spId };
  } catch { return quote; }
}

export async function deleteQuoteFromSharePoint(sp: SharePointService, quote: Pick<CrmQuote, 'id' | 'spId'>): Promise<void> {
  // Delete by known spId first (fast path), then sweep all rows by crmId to
  // remove any older duplicate SP rows that survived from bulk-sync bugs.
  if (quote.spId) {
    try { await sp.deleteCrmQuoteById(quote.spId); } catch { /* ignore */ }
  }
  if (quote.id) {
    try { await sp.deleteAllCrmQuotesByCrmId(quote.id); } catch { /* ignore */ }
  }
}

export async function upsertWipToSharePoint(sp: SharePointService, project: CrmProject): Promise<CrmProject> {
  try {
    if (project.spId) { await sp.updateCrmWipItem(project.spId, project); return project; }
    const spId = await sp.addCrmWipItem(project);
    return { ...project, spId };
  } catch { return project; }
}

