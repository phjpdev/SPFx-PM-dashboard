import { IRfi } from '../models/IProject';
import { ITeamMember } from '../models/ITask';

export const RFI_DEFAULT_BY_COMPANY = 'Z Edge Design';

export interface RfiSenderDefaults {
  by: string;
  byCompany: string;
}

const LS_PREFIX = '3edge-rfi-sender-';

export function senderStorageKey(userKey: string): string {
  return LS_PREFIX + userKey.toLowerCase().replace(/[^a-z0-9_-]/g, '_');
}

export function loadSenderDefaults(userKey: string): RfiSenderDefaults | null {
  try {
    const raw = localStorage.getItem(senderStorageKey(userKey));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as RfiSenderDefaults;
    if (parsed.by || parsed.byCompany) return parsed;
  } catch { /* ignore */ }
  return null;
}

export function saveSenderDefaults(userKey: string, defaults: RfiSenderDefaults): void {
  try {
    localStorage.setItem(senderStorageKey(userKey), JSON.stringify(defaults));
  } catch { /* ignore */ }
}

export function resolveSenderDefaults(
  userDisplayName: string,
  teamMembers?: ITeamMember[]
): RfiSenderDefaults {
  const userKey = userDisplayName.trim();
  const saved = loadSenderDefaults(userKey);
  if (saved?.by) {
    return { by: saved.by, byCompany: saved.byCompany || RFI_DEFAULT_BY_COMPANY };
  }
  if (teamMembers?.length) {
    const dn = userDisplayName.trim().toLowerCase();
    const match = teamMembers.find(m =>
      m.isActive !== false && (
        m.fullName.trim().toLowerCase() === dn ||
        m.initials.trim().toLowerCase() === dn
      )
    );
    if (match) {
      return { by: match.fullName, byCompany: RFI_DEFAULT_BY_COMPANY };
    }
  }
  return { by: userDisplayName.trim(), byCompany: RFI_DEFAULT_BY_COMPANY };
}

export function applySenderDefaultsToRfi(
  rfi: IRfi,
  userDisplayName: string,
  teamMembers?: ITeamMember[]
): IRfi {
  const sender = resolveSenderDefaults(userDisplayName, teamMembers);
  return {
    ...rfi,
    by: rfi.by || sender.by,
    byCompany: rfi.byCompany || sender.byCompany,
  };
}

/** Look up sender email from TeamMembers by Prepared By name. */
export function resolveSenderEmail(preparedBy: string, teamMembers?: ITeamMember[]): string {
  const name = preparedBy.trim().toLowerCase();
  if (!name || !teamMembers?.length) return '';
  const match = teamMembers.find(m =>
    m.isActive !== false && m.fullName.trim().toLowerCase() === name
  );
  return match?.email || '';
}
