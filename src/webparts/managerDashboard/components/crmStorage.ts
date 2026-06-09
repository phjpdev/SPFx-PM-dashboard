import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmRfq, CrmQuote, CrmProject } from './crmTypes';
import { cloneStaticCrm } from './crmStaticData';

/** Site-wide CRM delta index (small JSON — one row per changed person/company). */
export const CRM_SP_DELTA_META = 'crm_delta_meta';
/** Legacy monolithic delta — read-only fallback. */
export const CRM_SP_DELTA = 'crm_delta';
export const CRM_SP_RFQS = 'crm_rfqs';
export const CRM_SP_QUOTES = 'crm_quotes';
export const CRM_SP_PROJECTS = 'crm_projects';
export const CRM_SP_QUOTE_BUDGET = 'crm_quote_budget';

export interface CrmQuoteBudget {
  year: number;
  yearTarget: number;
  monthTarget: number;
}

const LS_DELTA = '3edge-crm-delta';
const LS_PERSONS = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';
const LS_RFQS = '3edge-crm-rfqs';
const LS_QUOTES = '3edge-crm-quotes';
const LS_PROJECTS = '3edge-crm-projects';
const LS_QUOTE_BUDGET = '3edge-crm-quote-budget';

export interface CrmDelta {
  revision: number;
  updatedAt: string;
  persons: Record<string, CrmPerson>;
  companies: Record<string, CrmCompany>;
  deletedPersonIds: string[];
  deletedCompanyIds: string[];
}

interface CrmDeltaMeta {
  revision: number;
  updatedAt: string;
  personIds: string[];
  companyIds: string[];
  deletedPersonIds: string[];
  deletedCompanyIds: string[];
}

export interface CrmListsSnapshot {
  persons: CrmPerson[];
  companies: CrmCompany[];
}

const emptyDelta = (): CrmDelta => ({
  revision: 0,
  updatedAt: '',
  persons: {},
  companies: {},
  deletedPersonIds: [],
  deletedCompanyIds: [],
});

const loadLS = <T,>(k: string, fb: T): T => {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : fb;
  } catch {
    return fb;
  }
};

const jsonEq = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);

const spKey = (prefix: string, id: string): string =>
  `${prefix}${id.replace(/[^a-zA-Z0-9_-]/g, '_')}`;

const spKeyPerson = (id: string): string => spKey('crm_p_', id);
const spKeyCompany = (id: string): string => spKey('crm_c_', id);

/** Attachments stay in browser storage only (base64 is too large for 3Edge_Settings). */
const personForSp = (p: CrmPerson): CrmPerson => {
  const { attachments: _a, ...rest } = p;
  return rest;
};

const overlayLocalAttachments = (persons: CrmPerson[], local: CrmPerson[]): CrmPerson[] => {
  const attMap = new Map<string, CrmPerson['attachments']>();
  for (const p of local) {
    if (p.attachments?.length) attMap.set(p.id, p.attachments);
  }
  if (!attMap.size) return persons;
  return persons.map(p => {
    const attachments = attMap.get(p.id);
    return attachments ? { ...p, attachments } : p;
  });
};

/** Merge static baseline + delta overrides. */
export function applyDelta(baseline: CrmListsSnapshot, delta: CrmDelta | null): CrmListsSnapshot {
  if (!delta) {
    return { persons: [...baseline.persons], companies: [...baseline.companies] };
  }
  const deletedP = new Set(delta.deletedPersonIds || []);
  const deletedC = new Set(delta.deletedCompanyIds || []);

  const personMap = new Map<string, CrmPerson>();
  for (const p of baseline.persons) {
    if (!deletedP.has(p.id)) personMap.set(p.id, p);
  }
  for (const p of Object.values(delta.persons || {})) {
    personMap.set(p.id, p);
  }

  const companyMap = new Map<string, CrmCompany>();
  for (const c of baseline.companies) {
    if (!deletedC.has(c.id)) companyMap.set(c.id, c);
  }
  for (const c of Object.values(delta.companies || {})) {
    companyMap.set(c.id, c);
  }

  return { persons: Array.from(personMap.values()), companies: Array.from(companyMap.values()) };
}

/** Only records that differ from the bundled static import. */
export function computeDelta(baseline: CrmListsSnapshot, current: CrmListsSnapshot): CrmDelta {
  const baseP = new Map(baseline.persons.map(p => [p.id, p]));
  const baseC = new Map(baseline.companies.map(c => [c.id, c]));
  const curP = new Map(current.persons.map(p => [p.id, p]));
  const curC = new Map(current.companies.map(c => [c.id, c]));

  const persons: Record<string, CrmPerson> = {};
  const companies: Record<string, CrmCompany> = {};
  const deletedPersonIds: string[] = [];
  const deletedCompanyIds: string[] = [];

  curP.forEach((p, id) => {
    const b = baseP.get(id);
    if (!b || !jsonEq(b, p)) persons[id] = p;
  });
  baseP.forEach((_, id) => {
    if (!curP.has(id)) deletedPersonIds.push(id);
  });

  curC.forEach((c, id) => {
    const b = baseC.get(id);
    if (!b || !jsonEq(b, c)) companies[id] = c;
  });
  baseC.forEach((_, id) => {
    if (!curC.has(id)) deletedCompanyIds.push(id);
  });

  return { revision: 0, updatedAt: new Date().toISOString(), persons, companies, deletedPersonIds, deletedCompanyIds };
}

async function loadLegacyDeltaFromSp(sp: SharePointService): Promise<CrmDelta | null> {
  try {
    const json = await sp.getSettingChunks(CRM_SP_DELTA);
    if (!json) return null;
    const d = JSON.parse(json) as CrmDelta;
    return d && typeof d === 'object' ? d : null;
  } catch {
    return null;
  }
}

async function loadDeltaFromSp(sp: SharePointService): Promise<CrmDelta | null> {
  try {
    const metaJson = await sp.getSettingChunks(CRM_SP_DELTA_META);
    if (!metaJson) return loadLegacyDeltaFromSp(sp);

    const meta = JSON.parse(metaJson) as CrmDeltaMeta;
    const persons: Record<string, CrmPerson> = {};
    const companies: Record<string, CrmCompany> = {};

    for (const id of meta.personIds || []) {
      const j = await sp.getSettingChunks(spKeyPerson(id));
      if (j) persons[id] = JSON.parse(j) as CrmPerson;
    }
    for (const id of meta.companyIds || []) {
      const j = await sp.getSettingChunks(spKeyCompany(id));
      if (j) companies[id] = JSON.parse(j) as CrmCompany;
    }

    return {
      revision: meta.revision || 0,
      updatedAt: meta.updatedAt || '',
      persons,
      companies,
      deletedPersonIds: meta.deletedPersonIds || [],
      deletedCompanyIds: meta.deletedCompanyIds || [],
    };
  } catch {
    return loadLegacyDeltaFromSp(sp);
  }
}

async function saveDeltaToSp(sp: SharePointService, delta: CrmDelta): Promise<void> {
  const meta: CrmDeltaMeta = {
    revision: delta.revision,
    updatedAt: delta.updatedAt,
    personIds: Object.keys(delta.persons),
    companyIds: Object.keys(delta.companies),
    deletedPersonIds: delta.deletedPersonIds,
    deletedCompanyIds: delta.deletedCompanyIds,
  };
  await sp.setSettingChunks(CRM_SP_DELTA_META, JSON.stringify(meta));
  await Promise.all(Object.keys(delta.persons).map(async id => {
    await sp.setSettingChunks(spKeyPerson(id), JSON.stringify(personForSp(delta.persons[id])));
  }));
  await Promise.all(Object.keys(delta.companies).map(async id => {
    await sp.setSettingChunks(spKeyCompany(id), JSON.stringify(delta.companies[id]));
  }));
}

function loadDeltaLocal(): CrmDelta | null {
  return loadLS<CrmDelta | null>(LS_DELTA, null);
}

function pickNewerDelta(a: CrmDelta | null, b: CrmDelta | null): CrmDelta | null {
  if (!a) return b;
  if (!b) return a;
  return (a.revision || 0) >= (b.revision || 0) ? a : b;
}

/** Static baseline + SharePoint/local delta (newest revision wins). */
export async function loadCrmPersonsCompanies(sp: SharePointService): Promise<CrmListsSnapshot> {
  const baseline = cloneStaticCrm();
  const fromSp = await loadDeltaFromSp(sp);
  const fromLocal = loadDeltaLocal();
  const delta = pickNewerDelta(fromSp, fromLocal);
  const merged = applyDelta(baseline, delta);
  const localPersons = loadLS<CrmPerson[] | null>(LS_PERSONS, null);
  if (localPersons?.length) {
    merged.persons = overlayLocalAttachments(merged.persons, localPersons);
  }
  return merged;
}

export async function loadCrmDelta(sp: SharePointService): Promise<CrmDelta | null> {
  return pickNewerDelta(await loadDeltaFromSp(sp), loadDeltaLocal());
}

/** Save only changed records — one small JSON blob per edited person/company. */
export async function saveCrmPersonsCompanies(
  sp: SharePointService,
  persons: CrmPerson[],
  companies: CrmCompany[],
  prevRevision = 0,
): Promise<{ ok: boolean; revision: number }> {
  const baseline = cloneStaticCrm();
  const current = { persons, companies };
  const delta = computeDelta(baseline, current);
  delta.revision = prevRevision + 1;
  delta.updatedAt = new Date().toISOString();

  try {
    localStorage.setItem(LS_PERSONS, JSON.stringify(persons));
    localStorage.setItem(LS_COMPANIES, JSON.stringify(companies));
    localStorage.setItem(LS_DELTA, JSON.stringify(delta));
  } catch { /* ignore */ }

  try {
    await saveDeltaToSp(sp, delta);
    return { ok: true, revision: delta.revision };
  } catch {
    return { ok: false, revision: delta.revision };
  }
}

/** Union remote + local; local wins on same id (handles stale empty SharePoint reads). */
function mergeListById<T extends { id: string }>(remote: T[], local: T[]): T[] {
  const map = new Map<string, T>();
  remote.forEach(x => map.set(x.id, x));
  local.forEach(x => map.set(x.id, x));
  return Array.from(map.values());
}

export async function loadRfqsFromSharePoint(sp: SharePointService): Promise<CrmRfq[]> {
  const local = loadLS<CrmRfq[]>(LS_RFQS, []);
  try {
    const json = await sp.getSettingChunks(CRM_SP_RFQS);
    if (!json) return local;
    const remote = JSON.parse(json) as CrmRfq[];
    return mergeListById(remote, local);
  } catch {
    return local;
  }
}

export async function saveRfqsToSharePoint(sp: SharePointService, rfqs: CrmRfq[]): Promise<void> {
  try {
    localStorage.setItem(LS_RFQS, JSON.stringify(rfqs));
  } catch { /* ignore */ }
  try {
    await sp.setSettingChunks(CRM_SP_RFQS, JSON.stringify(rfqs));
  } catch { /* ignore */ }
}

export async function loadQuotesFromSharePoint(sp: SharePointService): Promise<CrmQuote[]> {
  const local = loadLS<CrmQuote[]>(LS_QUOTES, []);
  try {
    const json = await sp.getSettingChunks(CRM_SP_QUOTES);
    if (!json) return local;
    const remote = JSON.parse(json) as CrmQuote[];
    return mergeListById(remote, local);
  } catch {
    return local;
  }
}

export async function saveQuotesToSharePoint(sp: SharePointService, quotes: CrmQuote[]): Promise<void> {
  try {
    localStorage.setItem(LS_QUOTES, JSON.stringify(quotes));
  } catch { /* ignore */ }
  try {
    await sp.setSettingChunks(CRM_SP_QUOTES, JSON.stringify(quotes));
  } catch { /* ignore */ }
}

export async function loadProjectsFromSharePoint(sp: SharePointService): Promise<CrmProject[]> {
  const local = loadLS<CrmProject[]>(LS_PROJECTS, []);
  try {
    const json = await sp.getSettingChunks(CRM_SP_PROJECTS);
    if (!json) return local;
    const remote = JSON.parse(json) as CrmProject[];
    return mergeListById(remote, local);
  } catch {
    return local;
  }
}

export async function saveProjectsToSharePoint(sp: SharePointService, projects: CrmProject[]): Promise<void> {
  try {
    localStorage.setItem(LS_PROJECTS, JSON.stringify(projects));
  } catch { /* ignore */ }
  try {
    await sp.setSettingChunks(CRM_SP_PROJECTS, JSON.stringify(projects));
  } catch { /* ignore */ }
}

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
  try {
    localStorage.setItem(LS_QUOTE_BUDGET, JSON.stringify(budget));
  } catch { /* ignore */ }
  try {
    await sp.setSettingChunks(CRM_SP_QUOTE_BUDGET, JSON.stringify(budget));
  } catch { /* ignore */ }
}
