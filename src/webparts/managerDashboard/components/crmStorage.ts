import type { SharePointService } from '../../../shared/services/SharePointService';
import type { CrmRfq } from './crmTypes';

export const CRM_SP_RFQS = 'crm_rfqs';

const LS_RFQS = '3edge-crm-rfqs';

const loadLS = <T,>(k: string, fb: T): T => {
  try {
    const v = localStorage.getItem(k);
    return v ? (JSON.parse(v) as T) : fb;
  } catch {
    return fb;
  }
};

export async function loadRfqsFromSharePoint(sp: SharePointService): Promise<CrmRfq[]> {
  try {
    const json = await sp.getSettingChunks(CRM_SP_RFQS);
    if (json) return JSON.parse(json) as CrmRfq[];
  } catch { /* ignore */ }
  return loadLS<CrmRfq[]>(LS_RFQS, []);
}

export async function saveRfqsToSharePoint(sp: SharePointService, rfqs: CrmRfq[]): Promise<void> {
  localStorage.setItem(LS_RFQS, JSON.stringify(rfqs));
  try {
    await sp.setSettingChunks(CRM_SP_RFQS, JSON.stringify(rfqs));
  } catch { /* ignore */ }
}
