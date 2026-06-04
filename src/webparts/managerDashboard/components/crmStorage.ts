import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmPerson, CrmCompany, CrmRfq } from './crmTypes';

export const CRM_SP_COMPANIES = 'crm_companies';
export const CRM_SP_PERSONS   = 'crm_persons';
export const CRM_SP_RFQS      = 'crm_rfqs';

const LS_PERSONS   = '3edge-crm-persons';
const LS_COMPANIES = '3edge-crm-companies';
const LS_RFQS      = '3edge-crm-rfqs';

const loadLS = <T,>(k: string, fb: T): T => {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : fb;
  } catch {
    return fb;
  }
};

export interface CrmStoreSnapshot {
  companies: CrmCompany[];
  persons: CrmPerson[];
  rfqs: CrmRfq[];
}

/** Load CRM data from SharePoint (site-wide). Falls back to this browser's localStorage. */
export async function loadCrmFromSharePoint(sp: SharePointService): Promise<CrmStoreSnapshot> {
  const [coJson, peJson, rfqJson] = await Promise.all([
    sp.getSettingChunks(CRM_SP_COMPANIES),
    sp.getSettingChunks(CRM_SP_PERSONS),
    sp.getSettingChunks(CRM_SP_RFQS),
  ]);

  const local: CrmStoreSnapshot = {
    companies: loadLS<CrmCompany[]>(LS_COMPANIES, []),
    persons: loadLS<CrmPerson[]>(LS_PERSONS, []),
    rfqs: loadLS<CrmRfq[]>(LS_RFQS, []),
  };

  let companies = coJson ? (JSON.parse(coJson) as CrmCompany[]) : [];
  let persons = peJson ? (JSON.parse(peJson) as CrmPerson[]) : [];
  let rfqs = rfqJson ? (JSON.parse(rfqJson) as CrmRfq[]) : [];

  const spEmpty = !coJson && !peJson && !rfqJson;
  const localHasData = local.companies.length > 0 || local.persons.length > 0 || local.rfqs.length > 0;

  if (spEmpty && localHasData) {
    companies = local.companies;
    persons = local.persons;
    rfqs = local.rfqs;
    await saveCrmToSharePoint(sp, { companies, persons, rfqs });
  }

  return { companies, persons, rfqs };
}

/** Persist companies + persons to SharePoint (site-wide). */
export async function saveCompaniesAndPersons(
  sp: SharePointService,
  companies: CrmCompany[],
  persons: CrmPerson[],
): Promise<void> {
  await Promise.all([
    sp.setSettingChunks(CRM_SP_COMPANIES, JSON.stringify(companies)),
    sp.setSettingChunks(CRM_SP_PERSONS, JSON.stringify(persons)),
  ]);
  localStorage.setItem(LS_COMPANIES, JSON.stringify(companies));
  localStorage.setItem(LS_PERSONS, JSON.stringify(persons));
}

/** Persist CRM data to SharePoint so all users/devices see the same records. */
export async function saveCrmToSharePoint(sp: SharePointService, data: CrmStoreSnapshot): Promise<void> {
  await Promise.all([
    sp.setSettingChunks(CRM_SP_COMPANIES, JSON.stringify(data.companies)),
    sp.setSettingChunks(CRM_SP_PERSONS, JSON.stringify(data.persons)),
    sp.setSettingChunks(CRM_SP_RFQS, JSON.stringify(data.rfqs)),
  ]);
  localStorage.setItem(LS_COMPANIES, JSON.stringify(data.companies));
  localStorage.setItem(LS_PERSONS, JSON.stringify(data.persons));
  localStorage.setItem(LS_RFQS, JSON.stringify(data.rfqs));
}

export async function loadRfqsFromSharePoint(sp: SharePointService): Promise<CrmRfq[]> {
  const json = await sp.getSettingChunks(CRM_SP_RFQS);
  if (json) return JSON.parse(json) as CrmRfq[];
  const local = loadLS<CrmRfq[]>(LS_RFQS, []);
  if (local.length > 0) {
    await sp.setSettingChunks(CRM_SP_RFQS, JSON.stringify(local));
  }
  return local;
}

export async function saveRfqsToSharePoint(sp: SharePointService, rfqs: CrmRfq[]): Promise<void> {
  await sp.setSettingChunks(CRM_SP_RFQS, JSON.stringify(rfqs));
  localStorage.setItem(LS_RFQS, JSON.stringify(rfqs));
}
