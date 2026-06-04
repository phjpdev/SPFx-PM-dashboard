import type { SharePointService } from '../../../shared/services/SharePointService';

const lsKey = (projId: string): string => `3edge_checklist_v1_${projId}`;

/** SharePoint settings base key (site-wide, per project). */
export const checklistSpKey = (projId: string): string =>
  `checklist_v1_${projId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;

export async function loadChecklistJson(sp: SharePointService, projId: string): Promise<string | null> {
  const fromSp = await sp.getSettingChunks(checklistSpKey(projId));
  if (fromSp) return fromSp;
  try {
    return localStorage.getItem(lsKey(projId));
  } catch {
    return null;
  }
}

export async function saveChecklistJson(sp: SharePointService, projId: string, json: string): Promise<void> {
  await sp.setSettingChunks(checklistSpKey(projId), json);
  try {
    localStorage.setItem(lsKey(projId), json);
  } catch { /* ignore quota */ }
}
