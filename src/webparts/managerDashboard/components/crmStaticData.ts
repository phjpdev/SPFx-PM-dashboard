import type { CrmCompany, CrmPerson } from './crmTypes';
import staticCompaniesJson from '../data/crm-companies.json';
import staticPersonsJson from '../data/crm-persons.json';

/** Bundled Pipedrive import — same data for every user (no SharePoint required). */
export const CRM_STATIC_COMPANIES: CrmCompany[] = staticCompaniesJson as CrmCompany[];
export const CRM_STATIC_PERSONS: CrmPerson[] = staticPersonsJson as CrmPerson[];

export const cloneStaticCrm = (): { companies: CrmCompany[]; persons: CrmPerson[] } => ({
  companies: JSON.parse(JSON.stringify(CRM_STATIC_COMPANIES)) as CrmCompany[],
  persons: JSON.parse(JSON.stringify(CRM_STATIC_PERSONS)) as CrmPerson[],
});
