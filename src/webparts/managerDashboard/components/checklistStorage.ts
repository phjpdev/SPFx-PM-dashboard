import type { SharePointService } from '../../../shared/services/SharePointService';

export const checklistLocalKey = (projId: string): string => `3edge_checklist_v1_${projId}`;

/** SharePoint settings base key (site-wide, per project). */
export const checklistSpKey = (projId: string): string =>
  `checklist_v1_${projId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;

const checklistUpdatedAt = (json: string | null): number => {
  if (!json) return 0;
  try {
    const d = JSON.parse(json) as { updatedAt?: number };
    return typeof d.updatedAt === 'number' ? d.updatedAt : 0;
  } catch {
    return 0;
  }
};

export async function loadChecklistJson(sp: SharePointService, projId: string): Promise<string | null> {
  let fromLs: string | null = null;
  try {
    fromLs = localStorage.getItem(checklistLocalKey(projId));
  } catch { /* ignore */ }

  let fromSp: string | null = null;
  try {
    fromSp = await sp.getSettingChunks(checklistSpKey(projId));
  } catch { /* ignore */ }

  if (fromSp && fromLs) {
    return checklistUpdatedAt(fromLs) > checklistUpdatedAt(fromSp) ? fromLs : fromSp;
  }
  if (fromSp) return fromSp;
  return fromLs;
}

/** Saves to browser always; SharePoint when permitted. Returns true if site list sync succeeded. */
export async function saveChecklistJson(sp: SharePointService, projId: string, json: string): Promise<boolean> {
  try {
    localStorage.setItem(checklistLocalKey(projId), json);
  } catch { /* ignore quota */ }
  try {
    await sp.setSettingChunks(checklistSpKey(projId), json);
    return true;
  } catch {
    return false;
  }
}
