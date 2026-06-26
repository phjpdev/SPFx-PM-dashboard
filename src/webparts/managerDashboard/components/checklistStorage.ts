import type {
  ChecklistPersisted,
  SharePointService,
} from '../../../shared/services/SharePointService';

export const checklistLocalKey = (projId: string): string => `3edge_checklist_v1_${projId}`;
export const checklistUiLocalKey = (projId: string): string => `3edge_checklist_ui_v1_${projId}`;

/** Legacy settings key — used once to migrate JSON blobs into SharePoint lists. */
export const checklistSpKey = (projId: string): string =>
  `checklist_v1_${projId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;

export interface ChecklistUiPrefs {
  role?: 'detailer' | 'checker' | 'pm';
  currentPhase?: string;
}

const parseLegacyJson = (raw: string | null): ChecklistPersisted | null => {
  if (!raw) return null;
  try {
    const d = JSON.parse(raw) as ChecklistPersisted & { role?: string; currentPhase?: string };
    return {
      items: d.items || {},
      overrides: d.overrides || [],
      projectType: d.projectType === 'concrete' || d.projectType === 'both' ? d.projectType : 'steel',
      updatedAt: typeof d.updatedAt === 'number' ? d.updatedAt : 0,
    };
  } catch {
    return null;
  }
};

/** Extract 3E-xxx project id from a 3Edge_Settings row title. */
export function projIdFromChecklistSettingsTitle(title: string): string | null {
  if (!title.startsWith('checklist_v1_')) return null;
  const rest = title.slice('checklist_v1_'.length);
  const m = rest.match(/^(.+?)(?:__(?:meta|\d+)|_\d+)?$/);
  return m?.[1] || null;
}

/** Reassemble checklist JSON from 3Edge_Settings rows (supports several chunk naming schemes). */
export async function loadLegacyChecklistJson(
  sp: SharePointService,
  projId: string,
): Promise<ChecklistPersisted | null> {
  const baseKey = checklistSpKey(projId);

  const fromChunks = await sp.getSettingChunks(baseKey);
  const parsed = parseLegacyJson(fromChunks);
  if (parsed) return parsed;

  const rows = await sp.listSettingsByTitlePrefix(baseKey);
  if (rows.length === 0) return null;

  const exact = rows.find(r => r.title === baseKey);
  if (exact?.value) {
    const p = parseLegacyJson(exact.value);
    if (p) return p;
  }

  const metaRow = rows.find(r => r.title === `${baseKey}__meta`);
  if (metaRow?.value) {
    try {
      const meta = JSON.parse(metaRow.value) as { chunks: number };
      let out = '';
      for (let i = 0; i < meta.chunks; i++) {
        const part = rows.find(r => r.title === `${baseKey}__${i}`)?.value;
        if (part === undefined) break;
        out += part;
      }
      const p = parseLegacyJson(out);
      if (p) return p;
    } catch { /* try other formats */ }
  }

  const doubleChunks = rows
    .map(r => {
      const m = r.title.match(new RegExp(`^${baseKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}__(\\d+)$`));
      return m ? { idx: Number(m[1]), value: r.value } : null;
    })
    .filter((x): x is { idx: number; value: string } => x !== null)
    .sort((a, b) => a.idx - b.idx);
  if (doubleChunks.length > 0) {
    const p = parseLegacyJson(doubleChunks.map(c => c.value).join(''));
    if (p) return p;
  }

  const singleChunks = rows
    .map(r => {
      const m = r.title.match(new RegExp(`^${baseKey.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}_(\\d+)$`));
      return m ? { idx: Number(m[1]), value: r.value } : null;
    })
    .filter((x): x is { idx: number; value: string } => x !== null)
    .sort((a, b) => a.idx - b.idx);
  if (singleChunks.length > 0) {
    const p = parseLegacyJson(singleChunks.map(c => c.value).join(''));
    if (p) return p;
  }

  return null;
}

export function loadChecklistUiPrefs(projId: string): ChecklistUiPrefs {
  try {
    const raw = localStorage.getItem(checklistUiLocalKey(projId));
    if (!raw) return {};
    return JSON.parse(raw) as ChecklistUiPrefs;
  } catch {
    return {};
  }
}

export function saveChecklistUiPrefs(projId: string, prefs: ChecklistUiPrefs): void {
  try {
    localStorage.setItem(checklistUiLocalKey(projId), JSON.stringify(prefs));
  } catch { /* ignore quota */ }
}

/** Import every checklist_v1_* blob from 3Edge_Settings into the three checklist lists. */
export async function migrateAllChecklistsFromSettings(sp: SharePointService): Promise<number> {
  await sp.ensureChecklistLists();
  const rows = await sp.listSettingsByTitlePrefix('checklist_v1_');
  const projectIds = new Set<string>();
  for (const r of rows) {
    const pid = projIdFromChecklistSettingsTitle(r.title);
    if (pid) projectIds.add(pid);
  }

  let migrated = 0;
  for (const projId of Array.from(projectIds)) {
    try {
      const existing = await sp.loadChecklist(projId);
      if (existing && Object.keys(existing.items).length > 0) continue;

      const legacy = await loadLegacyChecklistJson(sp, projId);
      if (!legacy) continue;
      if (Object.keys(legacy.items).length === 0 && legacy.overrides.length === 0 && !legacy.updatedAt) continue;

      await sp.saveChecklist(projId, legacy);
      migrated++;
    } catch { /* skip project on error */ }
  }
  return migrated;
}

/** Load checklist from SharePoint lists; migrates legacy JSON once if lists are empty. */
export async function loadChecklistData(
  sp: SharePointService,
  projId: string,
): Promise<ChecklistPersisted | null> {
  let fromSp: ChecklistPersisted | null = null;
  try {
    fromSp = await sp.loadChecklist(projId);
  } catch { /* ignore */ }

  if (fromSp && (Object.keys(fromSp.items).length > 0 || fromSp.overrides.length > 0 || fromSp.updatedAt)) {
    return fromSp;
  }

  let legacy = await loadLegacyChecklistJson(sp, projId);
  if (!legacy) {
    try {
      legacy = parseLegacyJson(localStorage.getItem(checklistLocalKey(projId)));
    } catch { /* ignore */ }
  }

  if (legacy && (Object.keys(legacy.items).length > 0 || legacy.overrides.length > 0 || legacy.updatedAt)) {
    try {
      await sp.saveChecklist(projId, legacy);
    } catch { /* migration best-effort */ }
    return legacy;
  }

  return fromSp;
}

/** Saves shared checklist state to SharePoint; UI prefs stay on this device only. */
export async function saveChecklistData(
  sp: SharePointService,
  projId: string,
  data: ChecklistPersisted,
): Promise<'list' | 'settings' | false> {
  try {
    localStorage.setItem(checklistLocalKey(projId), JSON.stringify(data));
  } catch { /* ignore quota */ }

  try {
    return await sp.saveChecklist(projId, data);
  } catch {
    return false;
  }
}
