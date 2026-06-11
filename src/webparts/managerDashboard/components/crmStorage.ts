import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmRfq, CrmQuote, CrmProject } from './crmTypes';
import { normalizeCrmQuote, normalizeCrmRfq } from './crmRfqNormalize';
import { cloneStaticCrm } from './crmStaticData';

export interface CrmQuoteBudget {
  year: number;
  yearTarget: number;
  monthTarget: number;
}

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

    if (persons.length === 0 && companies.length === 0) {
      // First run — seed SP lists from bundled static data
      const baseline = cloneStaticCrm();
      try {
        await Promise.all([
          sp.syncCrmPersons(baseline.persons),
          sp.syncCrmCompanies(baseline.companies),
        ]);
      } catch { /* seed failed — use static in-memory */ }
      return baseline;
    }

    // Overlay attachments that are kept in browser storage only
    const local = loadLS<CrmPerson[] | null>(LS_PERSONS, null);
    const merged = local?.length ? overlayLocalAttachments(persons, local) : persons;
    try {
      localStorage.setItem(LS_PERSONS, JSON.stringify(merged));
      localStorage.setItem(LS_COMPANIES, JSON.stringify(companies));
    } catch { /* ignore */ }
    return { persons: merged, companies };
  } catch {
    // SP unavailable — fall back to local cache, then static
    const cachedP = loadLS<CrmPerson[] | null>(LS_PERSONS, null);
    const cachedC = loadLS<CrmCompany[] | null>(LS_COMPANIES, null);
    if (cachedP || cachedC) {
      return { persons: cachedP ?? [], companies: cachedC ?? [] };
    }
    return cloneStaticCrm();
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

const defaultQuoteBudget = (): CrmQuoteBudget => ({
  year: new Date().getFullYear(),
  yearTarget: 0,
  monthTarget: 0,
});

export async function loadQuoteBudget(sp: SharePointService): Promise<CrmQuoteBudget> {
  const local = loadLS<CrmQuoteBudget | null>(LS_QUOTE_BUDGET, null);
  try {
    const json = await sp.getSettingChunks(CRM_SP_QUOTE_BUDGET);
    if (!json) return local ?? defaultQuoteBudget();
    const remote = JSON.parse(json) as CrmQuoteBudget;
    return remote.year ? remote : (local ?? defaultQuoteBudget());
  } catch {
    return local ?? defaultQuoteBudget();
  }
}

export async function saveQuoteBudget(sp: SharePointService, budget: CrmQuoteBudget): Promise<void> {
  try { localStorage.setItem(LS_QUOTE_BUDGET, JSON.stringify(budget)); } catch { /* ignore */ }
  try { await sp.setSettingChunks(CRM_SP_QUOTE_BUDGET, JSON.stringify(budget)); } catch { /* ignore */ }
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
  if (!rfq.spId) return;
  try { await sp.deleteCrmRfqById(rfq.spId); } catch { /* ignore */ }
}

export async function upsertQuoteToSharePoint(sp: SharePointService, quote: CrmQuote): Promise<CrmQuote> {
  try {
    if (quote.spId) { await sp.updateCrmQuote(quote.spId, quote); return quote; }
    const spId = await sp.addCrmQuote(quote);
    return { ...quote, spId };
  } catch { return quote; }
}

export async function deleteQuoteFromSharePoint(sp: SharePointService, quote: Pick<CrmQuote, 'id' | 'spId'>): Promise<void> {
  if (!quote.spId) return;
  try { await sp.deleteCrmQuoteById(quote.spId); } catch { /* ignore */ }
}

export async function upsertWipToSharePoint(sp: SharePointService, project: CrmProject): Promise<CrmProject> {
  try {
    if (project.spId) { await sp.updateCrmWipItem(project.spId, project); return project; }
    const spId = await sp.addCrmWipItem(project);
    return { ...project, spId };
  } catch { return project; }
}

