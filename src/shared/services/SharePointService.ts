import { SPHttpClient } from '@microsoft/sp-http';
import { IProject, IRfi } from '../models/IProject';
import { ITask, ITeamMember, ITaskHistory } from '../models/ITask';
import type {
  CrmPerson, CrmCompany, CrmRfq, CrmQuote, CrmProject as CrmWip, CrmEnquiryCore, CrmQuoteStatus,
} from '../models/ICrm';

const LIST_PROJ = '3Edge_Projects';
const LIST_RFI = '3Edge_RFIs';
const LIST_TASKS = 'WeeklyTasks';
const LIST_TEAM = 'TeamMembers';
const LIST_SETTINGS = '3Edge_Settings';
export const LIST_CHECKLIST = '3Edge_Checklist';
export const LIST_CHECKLIST_OVERRIDES = '3Edge_Checklist_Overrides';
export const LIST_CHECKLIST_META = '3Edge_Checklist_Meta';
/** CRM enquiry rows — one SharePoint item per RFQ / quote / WIP record. */
export const LIST_CRM_RFQS = '3Edge_CRM_RFQs';
export const LIST_CRM_QUOTES = '3Edge_CRM_Quotes';
export const LIST_CRM_WIP = '3Edge_CRM_WIP';
const SITE_DATA_FOLDER = '3EdgeDashboardData';
/** Multiline Note field limit in SharePoint lists. */
const SETTING_CHUNK_SIZE = 60000;
/** Safe chunk size when Value is single-line text (~255 chars). */
const SETTING_CHUNK_SIZE_SAFE = 200;

export interface CrmListRow {
  spId: number;
  crmId: string;
  recordNum: string;
  payloadJson: string;
}

export interface CrmListRowInput {
  spId?: number;
  crmId: string;
  title: string;
  recordNum: string;
  payloadJson: string;
}

export class SharePointService {
  private _siteUrl: string;
  private _digest: string = '';
  private _digestTime: number = 0;
  private _webRelUrl: string = '';
  /** After 403/400 on library writes, use 3Edge_Settings only (no repeated failed file API calls). */
  private _siteDataFilesUnavailable = false;
  /** Cached field names per list — avoids hundreds of redundant POST /fields calls. */
  private _listFieldNames = new Map<string, Set<string>>();
  /** One in-flight/completed provisioning run per service instance. */
  private _ensureAllCrmListsTask: Promise<void> | null = null;

  // spHttpClient param kept for API compat but all requests use plain fetch
  constructor(siteUrl: string, _spHttpClient?: SPHttpClient) {
    this._siteUrl = siteUrl;
  }

  private parseDate(val: string | null | undefined): string {
    if (!val) return '';
    const m = val.match(/\/Date\((\d+)\)\//);
    if (m) return new Date(Number(m[1])).toISOString().substring(0, 10);
    return val.substring(0, 10);
  }

  private async getDigest(): Promise<string> {
    // Refresh digest if older than 25 minutes (SharePoint digests expire at 30 min)
    const age = Date.now() - this._digestTime;
    if (this._digest && age < 25 * 60 * 1000) return this._digest;
    this._digest = '';
    const r = await fetch(this._siteUrl + '/_api/contextinfo', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });
    if (r.ok) {
      const data = await r.json();
      this._digest = data.FormDigestValue || '';
      this._digestTime = Date.now();
    }
    return this._digest;
  }

  private async spGet(path: string): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const r = await fetch(this._siteUrl + path, {
      credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    return r.json();
  }

  private async spPost(path: string, body: any, attempt = 0): Promise<any> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      if (r.status === 429 && attempt < 4) {
        const wait = Math.max(Number(r.headers.get('Retry-After') || '10'), 5) * 1000;
        await new Promise<void>(resolve => setTimeout(resolve, wait));
        return this.spPost(path, body, attempt + 1);
      }
      if (r.status === 403 && attempt === 0) {
        this._digest = ''; this._digestTime = 0;
        return this.spPost(path, body, attempt + 1);
      }
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
    const text = await r.text();
    return text ? JSON.parse(text) : {};
  }

  private async spMerge(path: string, body: any, attempt = 0): Promise<void> { // eslint-disable-line @typescript-eslint/no-explicit-any
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(body)
    });
    if (!r.ok) {
      if (r.status === 429 && attempt < 4) {
        const wait = Math.max(Number(r.headers.get('Retry-After') || '10'), 5) * 1000;
        await new Promise<void>(resolve => setTimeout(resolve, wait));
        return this.spMerge(path, body, attempt + 1);
      }
      if (r.status === 403 && attempt === 0) {
        this._digest = ''; this._digestTime = 0;
        return this.spMerge(path, body, attempt + 1);
      }
      let msg = 'HTTP ' + r.status;
      try { const e = await r.json(); const em = e.error?.message; msg = (typeof em === 'object' && em?.value) ? em.value : (em || e.error?.code || msg); } catch (_x) { /* ignore */ }
      if (r.status === 403) this._digest = '';
      throw new Error(msg);
    }
  }

  private async spDelete(path: string): Promise<void> {
    const digest = await this.getDigest();
    const r = await fetch(this._siteUrl + path, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*',
        'X-RequestDigest': digest
      }
    });
    if (!r.ok && r.status !== 404) {
      if (r.status === 403) this._digest = '';
      throw new Error('DELETE HTTP ' + r.status);
    }
  }

  // ── Project CRUD ──────────────────────────────────────────

  public async loadProjects(): Promise<IProject[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items?$top=5000&$orderby=projNum asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.projNum || String(i.Id),
      spId: i.Id,
      projNum: i.projNum || '',
      name: i.name || i.Title || '',
      discipline: i.discipline || '',
      status: i.status || 'Active',
      year: i.year ? Number(i.year) : new Date().getFullYear(),
      hrsAllowed: Number(i.hrsAllowed) || 0,
      hrsUsed: Number(i.hrsUsed) || 0,
      rfisAllowed: Number(i.rfisAllowed) || 0,
      quoteNum: i.quoteNum || '',
      contact: i.contact || '',
      company: i.company || '',
      email: i.email || '',
      mobile: i.mobile || '',
      clientNum: i.clientNum || '',
      clientp0: i.clientp0 || '',
      startDate: this.parseDate(i.startDate),
      finishDate: this.parseDate(i.finishDate),
      ifaDate: this.parseDate(i.ifaDate),
      ifcDate: this.parseDate(i.ifcDate),
      detailers: i.detailers || '',
      teamLead: i.teamLead || '',
      teamMembers: i.teamMembers || '',
      notes: i.notes || '',
      invoices: (() => { try { return (JSON.parse(i.invoicesJson || '[]') as { invNumber?: string; invPct?: number; invProgressClaim?: string; invDate?: string; invPaid?: boolean }[]).map(inv => ({ invNumber: inv.invNumber || '', invPct: inv.invPct ?? 0, invProgressClaim: inv.invProgressClaim || '', invDate: inv.invDate || '', invPaid: !!inv.invPaid })); } catch { return []; } })(),
      isEwo: i.isEwo || false,
      ewoNum: i.ewoNum || '',
      parentId: i.parentId || null,
      deliveredAt: this.parseDate(i.deliveredAt),
    }));
  }

  private pBody(d: IProject): object {
    return {
      Title: d.projNum || '',
      projNum: d.projNum || '',
      name: d.name || '',
      discipline: d.discipline || '',
      status: d.status || 'Active',
      year: Number(d.year) || new Date().getFullYear(),
      hrsAllowed: Number(d.hrsAllowed) || 0,
      hrsUsed: Number(d.hrsUsed) || 0,
      rfisAllowed: Number(d.rfisAllowed) || 0,
      quoteNum: d.quoteNum || '',
      contact: d.contact || '',
      company: d.company || '',
      email: d.email || '',
      mobile: d.mobile || '',
      clientNum: d.clientNum || '',
      clientp0: d.clientp0 || '',
      startDate: d.startDate ? d.startDate : null,
      finishDate: d.finishDate ? d.finishDate : null,
      ifaDate: d.ifaDate ? d.ifaDate : null,
      ifcDate: d.ifcDate ? d.ifcDate : null,
      detailers: d.detailers || '',
      teamLead: d.teamLead || '',
      teamMembers: d.teamMembers || '',
      notes: d.notes || '',
      invoicesJson: JSON.stringify(d.invoices || []),
      isEwo: d.isEwo === true || (d.isEwo as unknown as string) === 'true',
      ewoNum: d.ewoNum || '',
      parentId: d.parentId || null,
      deliveredAt: d.deliveredAt ? d.deliveredAt : null,
    };
  }

  public async addProject(d: IProject): Promise<number> {
    const r = await this.spPost(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items`, this.pBody(d));
    if (!r || !r.Id) throw new Error('No Id returned');
    return r.Id;
  }

  public async updateProject(spId: number, d: IProject): Promise<void> {
    await this.spMerge(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items(${spId})`, this.pBody(d));
  }

  public async deleteProject(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${LIST_PROJ}')/items(${spId})`);
  }

  // ── Email ─────────────────────────────────────────────────

  public async sendEmail(to: string, cc: string, subject: string, body: string): Promise<void> {
    const digest = await this.getDigest();
    const toList = to.split(/[,;]/).map(s => s.trim()).filter(Boolean);
    const ccList = cc ? cc.split(/[,;]/).map(s => s.trim()).filter(Boolean) : [];
    const r = await fetch(this._siteUrl + '/_api/SP.Utilities.Utility.SendEmail', {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      body: JSON.stringify({
        properties: {
          '__metadata': { type: 'SP.Utilities.EmailProperties' },
          'To': { results: toList },
          'CC': { results: ccList },
          'Subject': subject,
          'Body': body
        }
      })
    });
    if (!r.ok) {
      let msg = 'Email failed: HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
  }

  // ── RFI CRUD ──────────────────────────────────────────────

  public async loadRfis(): Promise<IRfi[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_RFI}')/items?$top=1000&$orderby=rfiSeq asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.rfiNum || String(i.Id),
      spId: i.Id,
      rfiNum: i.rfiNum || '',
      rfiSeq: Number(i.rfiSeq) || 0,
      projectId: i.projectId || '',
      projectName: i.projectName || '',
      rfiType: i.rfiType || '',
      status: i.status || 'Open',
      submittedTo: i.submittedTo || '',
      toCompany: i.toCompany || '',
      by: i.by || '',
      byCompany: i.byCompany || '',
      cc: i.cc || '',
      dateIssued: this.parseDate(i.dateIssued),
      dateRequired: this.parseDate(i.dateRequired),
      description: i.description || '',
      attachments: i.attachments || '',
      clientRfi: i.clientRfi || '',
      dateReceived: this.parseDate(i.dateReceived),
      response: i.response || '',
      responseDesc: i.responseDesc || '',
      sentBy: i.sentBy || '',
      sentByCompany: i.sentByCompany || '',
      impacted: i.impacted === true ? 'Yes' : (i.impacted === false ? 'No' : (i.impacted || 'No')),
      ewoRef: i.ewoRef || '',
      ewoCcn: i.ewoCcn || '',
      tracked: i.tracked || false,
      model: Number(i.model) || 0,
      connections: Number(i.connections) || 0,
      checking: Number(i.checking) || 0,
      drawings: Number(i.drawings) || 0,
      admin: Number(i.admin) || 0,
      revision: i.revision || 'A',
      email: i.email || '',
      emailSentDate: this.parseDate(i.emailSentDate)
    }));
  }

  private rBody(d: IRfi): object {
    return {
      Title: d.rfiNum || '',
      rfiNum: d.rfiNum || '',
      rfiSeq: Number(d.rfiSeq) || 0,
      projectId: d.projectId || '',
      projectName: d.projectName || '',
      rfiType: d.rfiType || '',
      status: d.status || 'Open',
      submittedTo: d.submittedTo || '',
      toCompany: d.toCompany || '',
      by: d.by || '',
      byCompany: d.byCompany || '',
      cc: d.cc || '',
      dateIssued: d.dateIssued ? d.dateIssued : null,
      dateRequired: d.dateRequired ? d.dateRequired : null,
      description: d.description || '',
      clientRfi: d.clientRfi || '',
      dateReceived: d.dateReceived ? d.dateReceived : null,
      response: d.response || '',
      responseDesc: d.responseDesc || '',
      sentBy: d.sentBy || '',
      sentByCompany: d.sentByCompany || '',
      impacted: d.impacted === 'Yes',
      ewoRef: d.ewoRef || '',
      ewoCcn: d.ewoCcn || '',
      model: Number(d.model) || 0,
      connections: Number(d.connections) || 0,
      checking: Number(d.checking) || 0,
      drawings: Number(d.drawings) || 0,
      admin: Number(d.admin) || 0,
      revision: d.revision || 'A',
      email: d.email || '',
      emailSentDate: d.emailSentDate ? d.emailSentDate : null
    };
  }

  public async addRfi(d: IRfi): Promise<number> {
    const r = await this.spPost(`/_api/web/lists/getbytitle('${LIST_RFI}')/items`, this.rBody(d));
    if (!r || !r.Id) throw new Error('No Id returned');
    return r.Id;
  }

  public async updateRfi(spId: number, d: IRfi): Promise<void> {
    await this.spMerge(`/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})`, this.rBody(d));
  }

  public async deleteRfi(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${LIST_RFI}')/items(${spId})`);
  }

  // ── Attachments ──────────────────────────────────────────────

  public async getAttachments(spId: number, listName: string = LIST_RFI): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${listName}')/items(${spId})/AttachmentFiles`);
    return d.value || [];
  }

  public async uploadAttachment(spId: number, file: File, listName: string = LIST_RFI): Promise<void> {
    const digest = await this.getDigest();
    const buf = await file.arrayBuffer();
    const url = this._siteUrl +
      `/_api/web/lists/getbytitle('${listName}')/items(${spId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
    const r = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      },
      body: buf
    });
    if (!r.ok) {
      if (r.status === 403) this._digest = '';
      let msg = 'Upload failed: HTTP ' + r.status;
      try { const e = await r.json(); msg = e.error?.message || msg; } catch (_x) { /* ignore */ }
      throw new Error(msg);
    }
  }

  public async deleteAttachment(spId: number, fileName: string, listName: string = LIST_RFI): Promise<void> {
    await this.spDelete(
      `/_api/web/lists/getbytitle('${listName}')/items(${spId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`
    );
  }

  public getProjectListName(): string { return LIST_PROJ; }

  // ── TeamMembers CRUD ───────────────────────────────────────

  public async loadTeamMembers(): Promise<ITeamMember[]> {
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_TEAM}')/items?$top=200&$orderby=Initials asc`);
    return (d.value || []).map((i: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
      id: i.Initials || String(i.Id),
      spId: i.Id,
      initials: i.Initials || '',
      fullName: i.FullName || i.Title || '',
      roleType: i.RoleType || 'Team',
      email: i.Email || '',
      totalHrsPerWeek: Number(i.TotalHrsPerWeek) || 40,
      productionPct: Number(i.ProductionPct) || 0,
      prodHrsPerWeek: Math.round((Number(i.TotalHrsPerWeek) || 40) * (Number(i.ProductionPct) || 0) / 100),
      startDate: this.parseDate(i.StartDate),
      endDate: this.parseDate(i.EndDate),
      isActive: i.IsActive !== false
    }));
  }

  // ── WeeklyTasks CRUD ───────────────────────────────────────

  public async loadTasks(weekStartDate?: string): Promise<ITask[]> {
    let filter = '';
    if (weekStartDate) filter = `&$filter=weekStartDate eq '${weekStartDate}'`;
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_TASKS}')/items?$top=2000&$orderby=day asc${filter}`);
    return (d.value || []).map((i: any) => { // eslint-disable-line @typescript-eslint/no-explicit-any
      let hist: ITaskHistory[] = [];
      try { hist = i.history ? JSON.parse(i.history) : []; } catch (_e) { /* ignore */ }
      return {
        id: String(i.Id),
        spId: i.Id,
        project: i.project || '',
        taskCode: i.taskCode || '',
        description: i.Title || '',
        assignee: i.assignee || '',
        day: Number(i.day) || 0,
        weekStartDate: this.parseDate(i.weekStartDate),
        hoursPlanned: Number(i.hoursPlanned) || 0,
        hoursActual: Number(i.hoursActual) || 0,
        wipPct: Number(i.wipPct) || 0,
        status: i.status || 'not_started',
        priority: i.priority || 'medium',
        completedBy: i.completedBy || '',
        completedAt: i.completedAt || '',
        completionNote: i.completionNote || '',
        reviewedBy: i.reviewedBy || '',
        reviewStatus: i.reviewStatus || '',
        history: hist
      };
    });
  }

  private tBody(d: ITask): object {
    return {
      Title: d.description || '',
      project: d.project || '',
      taskCode: d.taskCode || '',
      assignee: d.assignee || '',
      day: Number(d.day) || 0,
      weekStartDate: d.weekStartDate ? d.weekStartDate : null,
      hoursPlanned: Number(d.hoursPlanned) || 0,
      hoursActual: Number(d.hoursActual) || 0,
      wipPct: Number(d.wipPct) || 0,
      status: d.status || 'not_started',
      priority: d.priority || 'medium',
      completedBy: d.completedBy || '',
      completedAt: d.completedAt ? d.completedAt : null,
      completionNote: d.completionNote || '',
      reviewedBy: d.reviewedBy || '',
      reviewStatus: d.reviewStatus || '',
      history: JSON.stringify(d.history || [])
    };
  }

  public async addTask(d: ITask): Promise<number> {
    const r = await this.spPost(`/_api/web/lists/getbytitle('${LIST_TASKS}')/items`, this.tBody(d));
    if (!r || !r.Id) throw new Error('No Id returned');
    return r.Id;
  }

  public async updateTask(spId: number, d: ITask): Promise<void> {
    await this.spMerge(`/_api/web/lists/getbytitle('${LIST_TASKS}')/items(${spId})`, this.tBody(d));
  }

  public async deleteTask(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${LIST_TASKS}')/items(${spId})`);
  }

  // ── Settings (key-value store) ─────────────────────────────

  private static oDataStr(s: string): string {
    return s.replace(/'/g, "''");
  }

  /** Keep one record per id, preferring the highest spId (most recently created). */
  private static dedupeById<T extends { id: string; spId?: number }>(items: T[]): T[] {
    const seen = new Map<string, T>();
    for (const item of items) {
      const prev = seen.get(item.id);
      if (!prev || (item.spId ?? 0) > (prev.spId ?? 0)) seen.set(item.id, item);
    }
    const result: T[] = [];
    seen.forEach(v => result.push(v));
    return result;
  }

  /** Server-relative path for OData GetFileByServerRelativeUrl / GetFolderByServerRelativeUrl. */
  private static oDataPath(path: string): string {
    return SharePointService.oDataStr(path);
  }

  public async getSetting(key: string): Promise<string | undefined> {
    try {
      const k = SharePointService.oDataStr(key);
      const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_SETTINGS}')/items?$filter=Title eq '${k}'&$top=1`);
      const items = d.value || [];
      return items.length > 0 ? (items[0].Value || '') : undefined;
    } catch { return undefined; }
  }

  /** All rows from 3Edge_Settings (client-side filter — OData startswith is unreliable on some tenants). */
  public async listAllSettings(): Promise<Array<{ title: string; value: string }>> {
    try {
      const d = await this.spGet(
        `/_api/web/lists/getbytitle('${LIST_SETTINGS}')/items?$select=Title,Value&$top=5000`,
      );
      const items = (d.value || []) as Array<{ Title?: string; Value?: string }>;
      return items
        .filter(x => x.Title)
        .map(x => ({ title: x.Title as string, value: x.Value || '' }));
    } catch {
      return [];
    }
  }

  /** List settings rows whose Title starts with prefix (e.g. crm_rfqs_). */
  public async listSettingsByTitlePrefix(prefix: string): Promise<Array<{ title: string; value: string }>> {
    const all = await this.listAllSettings();
    return all.filter(x => x.title.startsWith(prefix));
  }

  public async setSetting(key: string, value: string): Promise<void> {
    try {
      await this.setSettingStrict(key, value);
    } catch { /* fail silently — non-critical */ }
  }

  public async setSettingStrict(key: string, value: string): Promise<void> {
    const k = SharePointService.oDataStr(key);
    const d = await this.spGet(`/_api/web/lists/getbytitle('${LIST_SETTINGS}')/items?$filter=Title eq '${k}'&$top=1`);
    const items = d.value || [];
    if (items.length > 0) {
      await this.spMerge(`/_api/web/lists/getbytitle('${LIST_SETTINGS}')/items(${items[0].Id})`, { Value: value });
    } else {
      await this.spPost(`/_api/web/lists/getbytitle('${LIST_SETTINGS}')/items`, { Title: key, Value: value });
    }
  }

  private async getWebRelUrl(): Promise<string> {
    if (this._webRelUrl) return this._webRelUrl;
    const web = await this.spGet('/_api/web?$select=ServerRelativeUrl');
    this._webRelUrl = web.ServerRelativeUrl || '';
    return this._webRelUrl;
  }

  private async ensureSiteDataFolder(): Promise<string> {
    const webUrl = await this.getWebRelUrl();
    const folderPath = `${webUrl}/${SITE_DATA_FOLDER}`;
    try {
      await this.spPost('/_api/web/folders', {
        __metadata: { type: 'SP.Folder' },
        ServerRelativeUrl: folderPath,
      });
    } catch {
      /* folder may already exist */
    }
    return folderPath;
  }

  private fileNameForKey(baseKey: string): string {
    return `${baseKey.replace(/[^a-zA-Z0-9_-]/g, '_')}.json`;
  }

  /** Read JSON from site document library (supports large CRM/checklist payloads). */
  private async readSiteDataFile(baseKey: string): Promise<string | null> {
    try {
      const webUrl = await this.getWebRelUrl();
      const filePath = `${webUrl}/${SITE_DATA_FOLDER}/${this.fileNameForKey(baseKey)}`;
      const r = await fetch(`${this._siteUrl}/_api/web/GetFileByServerRelativeUrl('${SharePointService.oDataPath(filePath)}')/$value`, {
        credentials: 'include',
      });
      if (r.status === 404) return null;
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.text();
    } catch {
      return null;
    }
  }

  /** Write JSON to site document library. */
  private async writeSiteDataFile(baseKey: string, value: string): Promise<void> {
    const folderPath = await this.ensureSiteDataFolder();
    const fileName = this.fileNameForKey(baseKey);
    const digest = await this.getDigest();
    const r = await fetch(
      `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${SharePointService.oDataPath(folderPath)}')/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`,
      {
        method: 'POST',
        credentials: 'include',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'X-RequestDigest': digest,
        },
        body: value,
      },
    );
    if (!r.ok) {
      let msg = 'HTTP ' + r.status;
      try {
        const e = await r.json();
        const em = e.error?.message;
        msg = (typeof em === 'object' && em?.value) ? em.value : (em || msg);
      } catch { /* ignore */ }
      throw new Error(msg);
    }
  }

  /** Legacy: read chunks from 3Edge_Settings (Value column size limit caused HTTP 500 on large data). */
  private async getSettingChunksLegacy(baseKey: string): Promise<string | null> {
    const metaRaw = await this.getSetting(`${baseKey}__meta`);
    if (!metaRaw) return null;
    try {
      const meta = JSON.parse(metaRaw) as { chunks: number };
      let out = '';
      for (let i = 0; i < meta.chunks; i++) {
        const part = await this.getSetting(`${baseKey}__${i}`);
        if (part === undefined) return null;
        out += part;
      }
      return out;
    } catch {
      return null;
    }
  }

  private async writeSettingChunks(baseKey: string, value: string, chunkSize: number): Promise<void> {
    if (value.length <= chunkSize) {
      await this.setSettingStrict(baseKey, value);
      return;
    }
    const chunks: string[] = [];
    for (let i = 0; i < value.length; i += chunkSize) {
      chunks.push(value.slice(i, i + chunkSize));
    }
    await this.setSettingStrict(`${baseKey}__meta`, JSON.stringify({ chunks: chunks.length }));
    for (let i = 0; i < chunks.length; i++) {
      await this.setSettingStrict(`${baseKey}__${i}`, chunks[i]);
    }
  }

  /** Write chunks to 3Edge_Settings (works when users lack library Contribute). */
  private async setSettingChunksLegacy(baseKey: string, value: string): Promise<void> {
    if (value.length <= SETTING_CHUNK_SIZE_SAFE) {
      try {
        await this.setSettingStrict(baseKey, value);
        return;
      } catch {
        /* Value column may be single-line — fall through to safe chunking */
      }
    }
    try {
      await this.writeSettingChunks(baseKey, value, SETTING_CHUNK_SIZE);
    } catch {
      await this.writeSettingChunks(baseKey, value, SETTING_CHUNK_SIZE_SAFE);
    }
  }

  /** Read large JSON blobs (settings list first when library unavailable; optional file). */
  public async getSettingChunks(baseKey: string): Promise<string | null> {
    if (!this._siteDataFilesUnavailable) {
      const fromFile = await this.readSiteDataFile(baseKey);
      if (fromFile) return fromFile;
    }
    const direct = await this.getSetting(baseKey);
    if (direct) return direct;
    return this.getSettingChunksLegacy(baseKey);
  }

  /** Write JSON for all site users — library file, then 3Edge_Settings fallback. */
  public async setSettingChunks(baseKey: string, value: string): Promise<void> {
    if (!this._siteDataFilesUnavailable) {
      try {
        await this.writeSiteDataFile(baseKey, value);
        return;
      } catch {
        this._siteDataFilesUnavailable = true;
      }
    }
    await this.setSettingChunksLegacy(baseKey, value);
  }

  // ── CRM lists (one row per RFQ / quote / WIP) ─────────────────

  private async listExists(title: string): Promise<boolean> {
    try {
      const t = SharePointService.oDataStr(title);
      await this.spGet(`/_api/web/lists/getbytitle('${t}')?$select=Id`);
      return true;
    } catch {
      return false;
    }
  }

  private async loadListFieldNames(listTitle: string): Promise<Set<string>> {
    const cached = this._listFieldNames.get(listTitle);
    if (cached) return cached;
    const t = SharePointService.oDataStr(listTitle);
    try {
      const d = await this.spGet(
        `/_api/web/lists/getbytitle('${t}')/fields?$select=Title,StaticName,InternalName&$top=500`,
      );
      const names = new Set<string>();
      for (const f of (d.value || []) as Array<{ Title?: string; StaticName?: string; InternalName?: string }>) {
        if (f.Title) names.add(String(f.Title).toLowerCase());
        if (f.StaticName) names.add(String(f.StaticName).toLowerCase());
        if (f.InternalName) names.add(String(f.InternalName).toLowerCase());
      }
      this._listFieldNames.set(listTitle, names);
      return names;
    } catch {
      return new Set<string>();
    }
  }

  private async ensureCrmListField(listTitle: string, fieldTitle: string, kind: number): Promise<void> {
    const names = await this.loadListFieldNames(listTitle);
    const key = fieldTitle.toLowerCase();
    if (names.has(key)) return;
    const t = SharePointService.oDataStr(listTitle);
    try {
      await this.spPost(`/_api/web/lists/getbytitle('${t}')/fields`, {
        FieldTypeKind: kind,
        Title: fieldTitle,
      });
      names.add(key);
    } catch {
      // Leave cache unchanged so a later ensure pass can retry (e.g. after permissions fix).
    }
  }

  /** Create CRM list + crmId / recordNum / payloadJson columns if missing. */
  public async ensureCrmList(listTitle: string, description: string): Promise<void> {
    if (!(await this.listExists(listTitle))) {
      await this.spPost('/_api/web/lists', {
        Title: listTitle,
        BaseTemplate: 100,
        Description: description,
      });
    }
    await this.ensureCrmListField(listTitle, 'crmId', 2);
    await this.ensureCrmListField(listTitle, 'recordNum', 2);
    await this.ensureCrmListField(listTitle, 'payloadJson', 3);
  }

  public async loadCrmListRows(listTitle: string): Promise<CrmListRow[]> {
    const t = SharePointService.oDataStr(listTitle);
    const d = await this.spGet(
      `/_api/web/lists/getbytitle('${t}')/items?$select=Id,Title,crmId,recordNum,payloadJson&$top=5000&$orderby=recordNum asc`,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (d.value || []).map((i: any) => ({
      spId: Number(i.Id),
      crmId: String(i.crmId || ''),
      recordNum: String(i.recordNum || i.Title || ''),
      payloadJson: String(i.payloadJson || ''),
    }));
  }

  private crmListItemBody(row: CrmListRowInput): object {
    return {
      Title: row.title || row.recordNum || row.crmId,
      crmId: row.crmId,
      recordNum: row.recordNum || row.title || '',
      payloadJson: row.payloadJson,
    };
  }

  /** Upsert rows by crmId; delete list items whose crmId is not in rows. */
  public async syncCrmListRows(listTitle: string, rows: CrmListRowInput[]): Promise<void> {
    const t = SharePointService.oDataStr(listTitle);
    const existing = await this.loadCrmListRows(listTitle);
    const byCrmId = new Map(existing.filter(e => e.crmId).map(e => [e.crmId, e]));
    const keep = new Set(rows.map(r => r.crmId));

    for (const row of rows) {
      const body = this.crmListItemBody(row);
      const prev = byCrmId.get(row.crmId);
      const spId = row.spId || prev?.spId;
      if (spId) {
        await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
      } else {
        await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
      }
    }

    for (const ex of existing) {
      if (ex.crmId && !keep.has(ex.crmId)) {
        await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${ex.spId})`);
      }
    }
  }

  // ── CRM proper-column lists ────────────────────────────────────────────────

  public static readonly LIST_CRM_PERSONS = '3Edge_CRM_Persons';
  public static readonly LIST_CRM_COMPANIES = '3Edge_CRM_Companies';
  // LIST_CRM_RFQS / LIST_CRM_QUOTES / LIST_CRM_WIP already exported above

  private parseJson<T>(v: unknown, fallback: T): T {
    if (typeof v !== 'string' || !v) return fallback;
    try { return JSON.parse(v) as T; } catch { return fallback; }
  }

  // ── Shared enquiry-core helpers ────────────────────────────────────────────

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private mapEnquiryCore(i: any): CrmEnquiryCore {
    return {
      dateReceived: this.parseDate(i.dateReceived),
      quoteRequiredBy: this.parseDate(i.quoteRequiredBy),
      personId: i.personId || '',
      organizationId: i.organizationId || '',
      companyAddress: i.companyAddress || '',
      projectTitle: i.projectTitle || '',
      projectAddress: i.projectAddress || '',
      discipline: i.discipline || 'Steel',
      projectValue: Number(i.projectValue) || 0,
      approximateHours: Number(i.approximateHours) || 0,
      engineerDrawingReceived: !!i.engineerDrawingReceived,
      engineerDrawingDate: this.parseDate(i.engineerDrawingDate),
      revisionVersionEng: i.revisionVersionEng || '',
      architectDrawingReceived: !!i.architectDrawingReceived,
      architectDrawingDate: this.parseDate(i.architectDrawingDate),
      revisionVersionArch: i.revisionVersionArch || '',
      rfiAllowed: Number(i.rfiAllowed) || 0,
      assignedTo: i.assignedTo || '',
      source: i.source || '',
      notes: i.notes || '',
    };
  }

  private enquiryCoreBody(d: CrmEnquiryCore): Record<string, unknown> {
    return {
      dateReceived: d.dateReceived || null,
      quoteRequiredBy: d.quoteRequiredBy || null,
      personId: d.personId || '',
      organizationId: d.organizationId || '',
      companyAddress: d.companyAddress || '',
      projectTitle: d.projectTitle || '',
      projectAddress: d.projectAddress || '',
      discipline: d.discipline || '',
      projectValue: Number(d.projectValue) || 0,
      approximateHours: Number(d.approximateHours) || 0,
      engineerDrawingReceived: !!d.engineerDrawingReceived,
      engineerDrawingDate: d.engineerDrawingDate || null,
      revisionVersionEng: d.revisionVersionEng || '',
      architectDrawingReceived: !!d.architectDrawingReceived,
      architectDrawingDate: d.architectDrawingDate || null,
      revisionVersionArch: d.revisionVersionArch || '',
      rfiAllowed: Number(d.rfiAllowed) || 0,
      assignedTo: d.assignedTo || '',
      source: d.source || '',
      notes: d.notes || '',
    };
  }

  // ── CRM Persons ────────────────────────────────────────────────────────────

  private personBody(d: CrmPerson): object {
    return {
      Title: d.name || '',
      crmId: d.id || '',
      organizationId: d.organizationId || '',
      position: d.position || '',
      phonesJson: JSON.stringify(d.phones || []),
      emailsJson: JSON.stringify(d.emails || []),
      activitiesJson: d.activities?.length ? JSON.stringify(d.activities) : null,
    };
  }

  public async loadCrmPersons(): Promise<CrmPerson[]> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    const d = await this.spGet(
      `/_api/web/lists/getbytitle('${t}')/items?$select=Id,Title,crmId,organizationId,position,phonesJson,emailsJson,activitiesJson&$top=2000&$orderby=Title asc`,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = (d.value || []).map((i: any) => ({
      id: String(i.crmId || String(i.Id)),
      spId: Number(i.Id),
      name: i.Title || '',
      organizationId: i.organizationId || '',
      position: i.position || '',
      phones: this.parseJson(i.phonesJson, []),
      emails: this.parseJson(i.emailsJson, []),
      activities: i.activitiesJson ? this.parseJson<CrmPerson['activities']>(i.activitiesJson, undefined) : undefined,
    }));
    return SharePointService.dedupeById(raw);
  }

  public async syncCrmPersons(persons: CrmPerson[]): Promise<CrmPerson[]> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    const existing = await this.loadCrmPersons();
    const byId = new Map(existing.map(e => [e.id, e.spId as number]));
    const keep = new Set(persons.map(p => p.id));
    const result: CrmPerson[] = [];

    // Run 2 upserts concurrently; individual failures skip that item (don't abort the whole sync).
    const BATCH = 2;
    for (let i = 0; i < persons.length; i += BATCH) {
      const batch = persons.slice(i, i + BATCH);
      const batchResult = await Promise.all(batch.map(async p => {
        try {
          const body = this.personBody(p);
          const spId = p.spId ?? byId.get(p.id);
          if (spId) {
            await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
            return { ...p, spId };
          } else {
            const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
            return { ...p, spId: Number(r.Id) };
          }
        } catch { return p; }
      }));
      result.push(...batchResult);
    }

    // Delete rows not in the new list (also batched so the delete phase always runs).
    const toDelete = existing.filter(e => e.id && !keep.has(e.id) && e.spId);
    for (let i = 0; i < toDelete.length; i += BATCH) {
      await Promise.all(toDelete.slice(i, i + BATCH).map(async e => {
        try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${e.spId})`); } catch { /* ignore */ }
      }));
    }

    return result;
  }

  public async addCrmPerson(p: CrmPerson): Promise<number> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, this.personBody(p));
    return Number(r.Id);
  }

  public async updateCrmPerson(spId: number, p: CrmPerson): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, this.personBody(p));
  }

  public async deleteCrmPersonById(spId: number): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${spId})`);
  }

  public async deleteAllCrmPersons(): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_PERSONS);
    const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$select=Id&$top=2000`);
    for (const item of ((d.value ?? []) as { Id: number }[])) {
      try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${item.Id})`); } catch { /* ignore */ }
    }
  }

  // ── CRM Companies ──────────────────────────────────────────────────────────

  private companyBody(d: CrmCompany): object {
    return {
      Title: d.name || '',
      crmId: d.id || '',
      labels: d.labels || '',
      address: d.address || '',
      phonesJson: JSON.stringify(d.phones || []),
      emailsJson: JSON.stringify(d.emails || []),
    };
  }

  public async loadCrmCompanies(): Promise<CrmCompany[]> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    const d = await this.spGet(
      `/_api/web/lists/getbytitle('${t}')/items?$select=Id,Title,crmId,labels,address,phonesJson,emailsJson&$top=2000&$orderby=Title asc`,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (d.value || []).map((i: any) => ({
      id: String(i.crmId || String(i.Id)),
      spId: Number(i.Id),
      name: i.Title || '',
      labels: i.labels || '',
      address: i.address || '',
      phones: this.parseJson(i.phonesJson, []),
      emails: this.parseJson(i.emailsJson, []),
    }));
  }

  public async syncCrmCompanies(companies: CrmCompany[]): Promise<CrmCompany[]> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    const existing = await this.loadCrmCompanies();
    const byId = new Map(existing.map(e => [e.id, e.spId as number]));
    const keep = new Set(companies.map(c => c.id));
    const result: CrmCompany[] = [];

    const BATCH = 2;
    for (let i = 0; i < companies.length; i += BATCH) {
      const batch = companies.slice(i, i + BATCH);
      const batchResult = await Promise.all(batch.map(async c => {
        try {
          const body = this.companyBody(c);
          const spId = c.spId ?? byId.get(c.id);
          if (spId) {
            await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
            return { ...c, spId };
          } else {
            const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
            return { ...c, spId: Number(r.Id) };
          }
        } catch { return c; }
      }));
      result.push(...batchResult);
    }

    const toDelete = existing.filter(e => e.id && !keep.has(e.id) && e.spId);
    for (let i = 0; i < toDelete.length; i += BATCH) {
      await Promise.all(toDelete.slice(i, i + BATCH).map(async e => {
        try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${e.spId})`); } catch { /* ignore */ }
      }));
    }

    return result;
  }

  public async addCrmCompany(c: CrmCompany): Promise<number> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, this.companyBody(c));
    return Number(r.Id);
  }

  public async updateCrmCompany(spId: number, c: CrmCompany): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, this.companyBody(c));
  }

  public async deleteCrmCompanyById(spId: number): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${spId})`);
  }

  public async deleteAllCrmCompanies(): Promise<void> {
    const t = SharePointService.oDataStr(SharePointService.LIST_CRM_COMPANIES);
    const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$select=Id&$top=2000`);
    for (const item of ((d.value ?? []) as { Id: number }[])) {
      try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${item.Id})`); } catch { /* ignore */ }
    }
  }

  // ── CRM RFQs ───────────────────────────────────────────────────────────────

  private rfqBody(d: CrmRfq): object {
    return {
      Title: d.rfqNum || d.id || '',
      crmId: d.id || '',
      rfqNum: d.rfqNum || '',
      stage: d.stage || 'New Enquiry',
      createQuoteXero: !!d.createQuoteXero,
      relatedRfqId: d.relatedRfqId || '',
      ...this.enquiryCoreBody(d),
    };
  }

  public async loadCrmRfqs(): Promise<CrmRfq[]> {
    const t = SharePointService.oDataStr(LIST_CRM_RFQS);
    const d = await this.spGet(
      `/_api/web/lists/getbytitle('${t}')/items?$top=2000&$orderby=rfqNum asc`,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw: CrmRfq[] = (d.value || []).map((i: any) => ({
      ...this.mapEnquiryCore(i),
      id: String(i.crmId || String(i.Id)),
      spId: Number(i.Id),
      rfqNum: i.rfqNum || i.Title || '',
      stage: i.stage || 'New Enquiry',
      createQuoteXero: !!i.createQuoteXero,
      relatedRfqId: i.relatedRfqId || '',
    }));
    const deduped = SharePointService.dedupeById(raw);
    // Secondary dedup by rfqNum — keep the row with the highest spId per number.
    const byNum = new Map<string, CrmRfq>();
    for (const r of deduped) {
      if (!r.rfqNum) continue;
      const prev = byNum.get(r.rfqNum);
      if (!prev || (r.spId ?? 0) > (prev.spId ?? 0)) byNum.set(r.rfqNum, r);
    }
    return deduped.filter(r => !r.rfqNum || byNum.get(r.rfqNum) === r);
  }

  public async syncCrmRfqs(rfqs: CrmRfq[]): Promise<CrmRfq[]> {
    const t = SharePointService.oDataStr(LIST_CRM_RFQS);
    const existing = await this.loadCrmRfqs();
    const byId = new Map<string, number>();
    const extraSpIds: number[] = [];
    for (const e of existing) {
      if (!e.id) continue;
      const prev = byId.get(e.id);
      if (prev !== undefined) {
        extraSpIds.push(Math.min(prev, e.spId as number));
        byId.set(e.id, Math.max(prev, e.spId as number));
      } else { byId.set(e.id, e.spId as number); }
    }
    for (const spId of extraSpIds) {
      await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${spId})`);
    }
    const keep = new Set(rfqs.map(r => r.id));
    const result: CrmRfq[] = [];

    for (const rfq of rfqs) {
      const body = this.rfqBody(rfq);
      const spId = rfq.spId ?? byId.get(rfq.id);
      if (spId) {
        await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
        result.push({ ...rfq, spId });
      } else {
        const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
        result.push({ ...rfq, spId: Number(r.Id) });
      }
    }

    for (const e of existing) {
      if (e.id && !keep.has(e.id) && e.spId) {
        await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${e.spId})`);
      }
    }

    return result;
  }

  private omitNulls(body: Record<string, unknown>): Record<string, unknown> {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(body)) {
      if (v !== null && v !== undefined) out[k] = v;
    }
    return out;
  }

  private quotePayloadBody(quote: CrmQuote): Record<string, unknown> {
    return {
      Title: quote.quoteNum || quote.id || '',
      crmId: quote.id || '',
      recordNum: quote.quoteNum || '',
      payloadJson: JSON.stringify(quote),
    };
  }

  private async filterBodyToListFields(
    listTitle: string,
    body: Record<string, unknown>,
    payload?: object,
  ): Promise<Record<string, unknown>> {
    const names = await this.loadListFieldNames(listTitle);
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(body)) {
      if (v === null || v === undefined) continue;
      if (names.has(k.toLowerCase())) out[k] = v;
    }
    if (!out.Title && body.Title) out.Title = body.Title;
    if (payload && names.has('payloadjson')) {
      out.payloadJson = JSON.stringify(payload);
    }
    return out;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private mapCrmQuoteRow(i: any): CrmQuote {
    const fromCols: CrmQuote = {
      ...this.mapEnquiryCore(i),
      id: String(i.crmId || String(i.Id)),
      spId: Number(i.Id),
      quoteNum: i.quoteNum || i.Title || '',
      rfqId: i.rfqId || '',
      rfqNum: i.rfqNum || '',
      quotedDate: this.parseDate(i.quotedDate),
      createQuoteXero: !!i.createQuoteXero,
      relatedRfqId: i.relatedRfqId || '',
      lostReason: i.lostReason || '',
      status: (i.status || 'Draft') as CrmQuoteStatus,
      lostAt: this.parseDate(i.lostAt),
      archived: !!i.archived,
      archivedAt: this.parseDate(i.archivedAt),
    };
    if (!i.payloadJson) return fromCols;
    const payload = this.parseJson<Partial<CrmQuote>>(i.payloadJson, {});
    if (!payload || typeof payload !== 'object') return fromCols;
    return {
      ...payload,
      ...fromCols,
      id: fromCols.id,
      spId: fromCols.spId,
      quoteNum: fromCols.quoteNum || payload.quoteNum || '',
      rfqId: fromCols.rfqId || payload.rfqId || '',
      rfqNum: fromCols.rfqNum || payload.rfqNum || '',
    };
  }

  private async postCrmQuoteItem(quote: CrmQuote): Promise<number> {
    const t = SharePointService.oDataStr(LIST_CRM_QUOTES);
    const fullBody = this.omitNulls({ ...this.quoteBody(quote) } as Record<string, unknown>);

    try {
      const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, fullBody);
      return Number(r.Id);
    } catch {
      this._listFieldNames.delete(LIST_CRM_QUOTES);
      const names = await this.loadListFieldNames(LIST_CRM_QUOTES);
      if (names.size > 0) {
        const filtered = await this.filterBodyToListFields(LIST_CRM_QUOTES, fullBody, quote);
        if (Object.keys(filtered).length > 0) {
          try {
            const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, filtered);
            return Number(r.Id);
          } catch { /* fall through to payloadJson row */ }
        }
      }
      const r = await this.spPost(
        `/_api/web/lists/getbytitle('${t}')/items`,
        this.quotePayloadBody(quote),
      );
      return Number(r.Id);
    }
  }

  private async mergeCrmQuoteItem(spId: number, quote: CrmQuote): Promise<void> {
    const t = SharePointService.oDataStr(LIST_CRM_QUOTES);
    const fullBody = this.omitNulls({ ...this.quoteBody(quote) } as Record<string, unknown>);

    try {
      await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, fullBody);
    } catch {
      this._listFieldNames.delete(LIST_CRM_QUOTES);
      const filtered = await this.filterBodyToListFields(LIST_CRM_QUOTES, fullBody, quote);
      if (Object.keys(filtered).length > 0) {
        try {
          await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, filtered);
          return;
        } catch { /* fall through */ }
      }
      await this.spMerge(
        `/_api/web/lists/getbytitle('${t}')/items(${spId})`,
        this.quotePayloadBody(quote),
      );
    }
  }

  private async ensureCrmQuotesListFields(): Promise<void> {
    const ef = (name: string, kind: number): Promise<void> =>
      this.ensureCrmListField(LIST_CRM_QUOTES, name, kind);

    const coreText = ['personId', 'organizationId', 'projectTitle', 'discipline',
      'assignedTo', 'source', 'revisionVersionEng', 'revisionVersionArch'];
    const coreNote = ['companyAddress', 'projectAddress', 'notes'];
    const coreDate = ['dateReceived', 'quoteRequiredBy', 'engineerDrawingDate', 'architectDrawingDate'];
    const coreBool = ['engineerDrawingReceived', 'architectDrawingReceived'];
    const coreNum = ['projectValue', 'approximateHours', 'rfiAllowed'];

    await this.ensureCrmList(LIST_CRM_QUOTES, 'CRM Quotes');
    for (const n of ['quoteNum', 'rfqId', 'rfqNum', 'lostReason', 'status', 'relatedRfqId', ...coreText]) await ef(n, 2);
    for (const n of coreNote) await ef(n, 3);
    for (const n of ['quotedDate', 'lostAt', 'archivedAt', ...coreDate]) await ef(n, 4);
    for (const n of ['createQuoteXero', 'archived', ...coreBool]) await ef(n, 8);
    for (const n of coreNum) await ef(n, 9);
  }

  private quoteBody(d: CrmQuote): object {
    return {
      Title: d.quoteNum || d.id || '',
      crmId: d.id || '',
      quoteNum: d.quoteNum || '',
      rfqId: d.rfqId || '',
      rfqNum: d.rfqNum || '',
      quotedDate: d.quotedDate || null,
      createQuoteXero: !!d.createQuoteXero,
      relatedRfqId: d.relatedRfqId || '',
      lostReason: d.lostReason || '',
      status: d.status || 'Draft',
      lostAt: d.lostAt || null,
      archived: !!d.archived,
      archivedAt: d.archivedAt || null,
      ...this.enquiryCoreBody(d),
    };
  }

  public async loadCrmQuotes(): Promise<CrmQuote[]> {
    const t = SharePointService.oDataStr(LIST_CRM_QUOTES);
    let d: { value?: unknown[] };
    try {
      d = await this.spGet(
        `/_api/web/lists/getbytitle('${t}')/items?$top=2000&$orderby=quoteNum asc`,
      );
    } catch {
      d = await this.spGet(
        `/_api/web/lists/getbytitle('${t}')/items?$top=2000&$orderby=Id asc`,
      );
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw: CrmQuote[] = (d.value || []).map((i: any) => this.mapCrmQuoteRow(i));
    return SharePointService.dedupeById(raw);
  }

  public async syncCrmQuotes(quotes: CrmQuote[]): Promise<CrmQuote[]> {
    const t = SharePointService.oDataStr(LIST_CRM_QUOTES);
    const existing = await this.loadCrmQuotes();
    const byId = new Map<string, number>();
    const extraSpIds: number[] = [];
    for (const e of existing) {
      if (!e.id) continue;
      const prev = byId.get(e.id);
      if (prev !== undefined) {
        extraSpIds.push(Math.min(prev, e.spId as number));
        byId.set(e.id, Math.max(prev, e.spId as number));
      } else { byId.set(e.id, e.spId as number); }
    }
    for (const spId of extraSpIds) {
      await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${spId})`);
    }
    const keep = new Set(quotes.map(q => q.id));
    const result: CrmQuote[] = [];

    for (const q of quotes) {
      const body = this.quoteBody(q);
      const spId = q.spId ?? byId.get(q.id);
      if (spId) {
        await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
        result.push({ ...q, spId });
      } else {
        const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
        result.push({ ...q, spId: Number(r.Id) });
      }
    }

    for (const e of existing) {
      if (e.id && !keep.has(e.id) && e.spId) {
        await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${e.spId})`);
      }
    }

    return result;
  }

  // ── CRM WIP (won projects) ─────────────────────────────────────────────────

  private wipBody(d: CrmWip): object {
    return {
      Title: d.projNum || d.id || '',
      crmId: d.id || '',
      projNum: d.projNum || '',
      quoteId: d.quoteId || '',
      quoteNum: d.quoteNum || '',
      rfqId: d.rfqId || '',
      rfqNum: d.rfqNum || '',
      wonDate: d.wonDate || null,
      ...this.enquiryCoreBody(d),
    };
  }

  public async loadCrmWip(): Promise<CrmWip[]> {
    const t = SharePointService.oDataStr(LIST_CRM_WIP);
    const d = await this.spGet(
      `/_api/web/lists/getbytitle('${t}')/items?$top=2000&$orderby=projNum asc`,
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (d.value || []).map((i: any) => ({
      ...this.mapEnquiryCore(i),
      id: String(i.crmId || String(i.Id)),
      spId: Number(i.Id),
      projNum: i.projNum || i.Title || '',
      quoteId: i.quoteId || '',
      quoteNum: i.quoteNum || '',
      rfqId: i.rfqId || '',
      rfqNum: i.rfqNum || '',
      wonDate: this.parseDate(i.wonDate),
    }));
  }

  public async syncCrmWip(projects: CrmWip[]): Promise<CrmWip[]> {
    const t = SharePointService.oDataStr(LIST_CRM_WIP);
    const existing = await this.loadCrmWip();
    const byId = new Map(existing.map(e => [e.id, e.spId as number]));
    const keep = new Set(projects.map(p => p.id));
    const result: CrmWip[] = [];

    for (const p of projects) {
      const body = this.wipBody(p);
      const spId = p.spId ?? byId.get(p.id);
      if (spId) {
        await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, body);
        result.push({ ...p, spId });
      } else {
        const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
        result.push({ ...p, spId: Number(r.Id) });
      }
    }

    for (const e of existing) {
      if (e.id && !keep.has(e.id) && e.spId) {
        await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${e.spId})`);
      }
    }

    return result;
  }

  // ── Individual CRUD (fast path — no bulk load) ─────────────────────────────

  public async addCrmRfq(rfq: CrmRfq): Promise<number> {
    const t = SharePointService.oDataStr(LIST_CRM_RFQS);
    const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, this.rfqBody(rfq));
    return Number(r.Id);
  }
  public async updateCrmRfq(spId: number, rfq: CrmRfq): Promise<void> {
    const t = SharePointService.oDataStr(LIST_CRM_RFQS);
    await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, this.rfqBody(rfq));
  }
  public async deleteCrmRfqById(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${SharePointService.oDataStr(LIST_CRM_RFQS)}')/items(${spId})`);
  }

  /** Delete every SP row for this crmId (handles duplicate rows left by old bulk-sync bugs). */
  public async deleteAllCrmRfqsByCrmId(crmId: string): Promise<void> {
    const t = SharePointService.oDataStr(LIST_CRM_RFQS);
    const safe = crmId.replace(/'/g, "''");
    const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$select=Id&$filter=crmId eq '${safe}'&$top=200`);
    for (const item of ((d.value ?? []) as { Id: number }[])) {
      try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${item.Id})`); } catch { /* ignore */ }
    }
  }

  public async addCrmQuote(quote: CrmQuote): Promise<number> {
    await this.ensureCrmQuotesListReady();
    return this.postCrmQuoteItem(quote);
  }
  public async updateCrmQuote(spId: number, quote: CrmQuote): Promise<void> {
    await this.ensureCrmQuotesListReady();
    await this.mergeCrmQuoteItem(spId, quote);
  }
  public async deleteCrmQuoteById(spId: number): Promise<void> {
    await this.spDelete(`/_api/web/lists/getbytitle('${SharePointService.oDataStr(LIST_CRM_QUOTES)}')/items(${spId})`);
  }

  /** Delete every SP row for this crmId (handles duplicate rows left by old bulk-sync bugs). */
  public async deleteAllCrmQuotesByCrmId(crmId: string): Promise<void> {
    const t = SharePointService.oDataStr(LIST_CRM_QUOTES);
    const safe = crmId.replace(/'/g, "''");
    const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$select=Id&$filter=crmId eq '${safe}'&$top=200`);
    for (const item of ((d.value ?? []) as { Id: number }[])) {
      try { await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${item.Id})`); } catch { /* ignore */ }
    }
  }

  public async addCrmWipItem(project: CrmWip): Promise<number> {
    const t = SharePointService.oDataStr(LIST_CRM_WIP);
    const r = await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, this.wipBody(project));
    return Number(r.Id);
  }
  public async updateCrmWipItem(spId: number, project: CrmWip): Promise<void> {
    const t = SharePointService.oDataStr(LIST_CRM_WIP);
    await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${spId})`, this.wipBody(project));
  }

  // ── CRM list provisioning ──────────────────────────────────────────────────

  /** Create all 5 CRM lists with proper columns if they do not already exist. */
  public ensureAllCrmLists(): Promise<void> {
    if (!this._ensureAllCrmListsTask) {
      this._ensureAllCrmListsTask = this.ensureAllCrmListsInner().catch((err) => {
        this._ensureAllCrmListsTask = null;
        throw err;
      });
    }
    return this._ensureAllCrmListsTask;
  }

  /** Ensure the quotes list exists with columns — called before every quote write. */
  public async ensureCrmQuotesListReady(): Promise<void> {
    this._listFieldNames.delete(LIST_CRM_QUOTES);
    await this.ensureCrmQuotesListFields();
  }

  private async ensureAllCrmListsInner(): Promise<void> {
    this._listFieldNames.clear();
    const ef = (list: string, name: string, kind: number): Promise<void> =>
      this.ensureCrmListField(list, name, kind);

    const coreText = ['personId', 'organizationId', 'projectTitle', 'discipline',
      'assignedTo', 'source', 'revisionVersionEng', 'revisionVersionArch'];
    const coreNote = ['companyAddress', 'projectAddress', 'notes'];
    const coreDate = ['dateReceived', 'quoteRequiredBy', 'engineerDrawingDate', 'architectDrawingDate'];
    const coreBool = ['engineerDrawingReceived', 'architectDrawingReceived'];
    const coreNum  = ['projectValue', 'approximateHours', 'rfiAllowed'];

    await this.ensureCrmList(SharePointService.LIST_CRM_PERSONS, 'CRM contact persons');
    for (const n of ['organizationId', 'position']) await ef(SharePointService.LIST_CRM_PERSONS, n, 2);
    for (const n of ['phonesJson', 'emailsJson', 'activitiesJson']) await ef(SharePointService.LIST_CRM_PERSONS, n, 3);

    await this.ensureCrmList(SharePointService.LIST_CRM_COMPANIES, 'CRM companies / organisations');
    await ef(SharePointService.LIST_CRM_COMPANIES, 'labels', 2);
    for (const n of ['address', 'phonesJson', 'emailsJson']) await ef(SharePointService.LIST_CRM_COMPANIES, n, 3);

    await this.ensureCrmList(LIST_CRM_RFQS, 'CRM Request for Quotes');
    for (const n of ['rfqNum', 'stage', 'relatedRfqId', ...coreText]) await ef(LIST_CRM_RFQS, n, 2);
    for (const n of coreNote) await ef(LIST_CRM_RFQS, n, 3);
    for (const n of coreDate) await ef(LIST_CRM_RFQS, n, 4);
    for (const n of ['createQuoteXero', ...coreBool]) await ef(LIST_CRM_RFQS, n, 8);
    for (const n of coreNum) await ef(LIST_CRM_RFQS, n, 9);

    await this.ensureCrmQuotesListFields();

    await this.ensureCrmList(LIST_CRM_WIP, 'CRM Work In Progress');
    for (const n of ['projNum', 'quoteId', 'quoteNum', 'rfqId', 'rfqNum', ...coreText]) await ef(LIST_CRM_WIP, n, 2);
    for (const n of coreNote) await ef(LIST_CRM_WIP, n, 3);
    for (const n of ['wonDate', ...coreDate]) await ef(LIST_CRM_WIP, n, 4);
    for (const n of coreBool) await ef(LIST_CRM_WIP, n, 8);
    for (const n of coreNum) await ef(LIST_CRM_WIP, n, 9);
  }

  // ── Checklist lists (crmId + payloadJson — same pattern as CRM RFQs) ───────

  private static checklistItemCrmId(projId: string, itemId: string): string {
    return `chk|${projId}|${itemId}`;
  }

  private static checklistSettingsKey(projId: string): string {
    return `checklist_v1_${projId.replace(/[^a-zA-Z0-9_-]/g, '_')}`;
  }

  private static checklistMetaCrmId(projId: string): string {
    return `chkmeta|${projId}`;
  }

  /** True when checklist lists have crmId + payloadJson columns (site owner provisioned). */
  private async checklistListColumnsReady(): Promise<boolean> {
    this._listFieldNames.delete(LIST_CHECKLIST);
    const names = await this.loadListFieldNames(LIST_CHECKLIST);
    return names.has('payloadjson') && names.has('crmid');
  }

  private parseChecklistPersisted(raw: string | null): ChecklistPersisted | null {
    if (!raw) return null;
    try {
      const d = JSON.parse(raw) as ChecklistPersisted;
      return {
        items: d.items || {},
        overrides: d.overrides || [],
        projectType: d.projectType === 'concrete' || d.projectType === 'both' ? d.projectType : 'steel',
        updatedAt: typeof d.updatedAt === 'number' ? d.updatedAt : 0,
      };
    } catch {
      return null;
    }
  }

  private async loadChecklistFromSettings(projId: string): Promise<ChecklistPersisted | null> {
    const raw = await this.getSettingChunks(SharePointService.checklistSettingsKey(projId));
    return this.parseChecklistPersisted(raw);
  }

  private async saveChecklistToSettings(projId: string, data: ChecklistPersisted): Promise<void> {
    const payload: ChecklistPersisted = {
      items: data.items || {},
      overrides: data.overrides || [],
      projectType: data.projectType || 'steel',
      updatedAt: data.updatedAt || Date.now(),
    };
    await this.setSettingChunks(SharePointService.checklistSettingsKey(projId), JSON.stringify(payload));
  }

  private static checklistOverrideCrmId(projId: string, ov: ChecklistOverrideLog): string {
    return `chkov|${projId}|${ov.itemId}|${ov.at}`;
  }

  private static parseChecklistItemCrmId(crmId: string): { projId: string; itemId: string } | null {
    const m = crmId.match(/^chk\|(.+)\|(p\d+s\d+i\d+)$/);
    return m ? { projId: m[1], itemId: m[2] } : null;
  }

  /** Ensures checklist lists exist; returns whether list columns are ready for row sync. */
  public async ensureChecklistLists(): Promise<boolean> {
    await this.ensureCrmList(LIST_CHECKLIST, '3Edge project checklist item states');
    await this.ensureCrmList(LIST_CHECKLIST_OVERRIDES, '3Edge checklist PM override audit log');
    await this.ensureCrmList(LIST_CHECKLIST_META, '3Edge checklist per-project settings');
    return this.checklistListColumnsReady();
  }

  /** Upsert rows for one project only — does not touch other projects' rows. */
  private async syncProjectCrmRows(
    listTitle: string,
    projectFilter: (crmId: string) => boolean,
    rows: CrmListRowInput[],
  ): Promise<void> {
    const t = SharePointService.oDataStr(listTitle);
    const existing = await this.loadCrmListRows(listTitle);
    const projectExisting = existing.filter(e => projectFilter(e.crmId));
    const byCrmId = new Map(projectExisting.map(e => [e.crmId, e]));
    const keep = new Set(rows.map(r => r.crmId));

    for (const row of rows) {
      const body = this.crmListItemBody(row);
      const prev = byCrmId.get(row.crmId);
      if (prev?.spId) {
        await this.spMerge(`/_api/web/lists/getbytitle('${t}')/items(${prev.spId})`, body);
      } else {
        await this.spPost(`/_api/web/lists/getbytitle('${t}')/items`, body);
      }
    }
    for (const ex of projectExisting) {
      if (ex.crmId && !keep.has(ex.crmId)) {
        await this.spDelete(`/_api/web/lists/getbytitle('${t}')/items(${ex.spId})`);
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private mapLegacyChecklistItemRow(i: any): { itemId: string; state: ChecklistItemState } | null {
    if (i.payloadJson) {
      const parsed = this.parseJson<ChecklistItemState>(i.payloadJson, {});
      const fromCrm = SharePointService.parseChecklistItemCrmId(String(i.crmId || ''));
      if (fromCrm) return { itemId: fromCrm.itemId, state: parsed };
    }
    const c2raw = i.c2 ? String(i.c2) : '';
    const c2 = c2raw === 'cleared' || c2raw === 'na' || c2raw === 'incorrect' ? c2raw : null;
    const itemId = String(i.itemId || '');
    if (!itemId) {
      const title = String(i.Title || '');
      const m = title.match(/(?: · | - )(p\d+s\d+i\d+)$/);
      if (!m) return null;
      const state: ChecklistItemState = {};
      if (i.c1 === true) state.c1 = true;
      if (c2) state.c2 = c2 as ChecklistC2Action;
      if (i.c1By) state.c1By = String(i.c1By);
      if (i.c2By) state.c2By = String(i.c2By);
      if (i.c1At) state.c1At = String(i.c1At);
      if (i.c2At) state.c2At = String(i.c2At);
      if (i.pmOverride === true || i.override === true) state.override = true;
      if (i.overrideBy) state.overrideBy = String(i.overrideBy);
      if (i.overrideAt) state.overrideAt = String(i.overrideAt);
      if (i.overrideReason) state.overrideReason = String(i.overrideReason);
      return { itemId: m[1], state };
    }
    const state: ChecklistItemState = {};
    if (i.c1 === true) state.c1 = true;
    if (c2) state.c2 = c2 as ChecklistC2Action;
    if (i.c1By) state.c1By = String(i.c1By);
    if (i.c2By) state.c2By = String(i.c2By);
    if (i.c1At) state.c1At = String(i.c1At);
    if (i.c2At) state.c2At = String(i.c2At);
    if (i.pmOverride === true || i.override === true) state.override = true;
    if (i.overrideBy) state.overrideBy = String(i.overrideBy);
    if (i.overrideAt) state.overrideAt = String(i.overrideAt);
    if (i.overrideReason) state.overrideReason = String(i.overrideReason);
    return { itemId, state };
  }

  private async loadChecklistFromLists(projId: string): Promise<ChecklistPersisted | null> {
    const itemPrefix = `chk|${projId}|`;
    const ovPrefix = `chkov|${projId}|`;
    const metaCrmId = SharePointService.checklistMetaCrmId(projId);

    const [itemRows, ovRows, metaRows] = await Promise.all([
      this.loadCrmListRows(LIST_CHECKLIST),
      this.loadCrmListRows(LIST_CHECKLIST_OVERRIDES),
      this.loadCrmListRows(LIST_CHECKLIST_META),
    ]);

    const items: Record<string, ChecklistItemState> = {};
    for (const row of itemRows) {
      if (!row.crmId.startsWith(itemPrefix)) continue;
      const ids = SharePointService.parseChecklistItemCrmId(row.crmId);
      if (!ids) continue;
      items[ids.itemId] = this.parseJson<ChecklistItemState>(row.payloadJson, {});
    }

    // Legacy rows saved with Title only (no crmId / payloadJson columns)
    try {
      const t = SharePointService.oDataStr(LIST_CHECKLIST);
      const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$top=5000`);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const i of (d.value || []) as any[]) {
        const title = String(i.Title || '');
        if (!title.startsWith(`${projId} · `) && !title.startsWith(`${projId} - `)) continue;
        const legacy = this.mapLegacyChecklistItemRow(i);
        if (legacy && !items[legacy.itemId]) items[legacy.itemId] = legacy.state;
      }
    } catch { /* ignore */ }

    const overrides: ChecklistOverrideLog[] = [];
    for (const row of ovRows) {
      if (row.crmId.startsWith(ovPrefix)) {
        overrides.push(this.parseJson<ChecklistOverrideLog>(row.payloadJson, {
          itemId: '', by: '', at: '', reason: '', itemText: '', taskCode: '', action: 'cleared',
        }));
      }
    }

    let projectType: 'steel' | 'concrete' | 'both' = 'steel';
    let updatedAt = 0;
    const metaRow = metaRows.find(r => r.crmId === metaCrmId);
    if (metaRow?.payloadJson) {
      const meta = this.parseJson<{ projectType?: string; updatedAt?: number }>(metaRow.payloadJson, {});
      if (meta.projectType === 'concrete' || meta.projectType === 'both') projectType = meta.projectType;
      updatedAt = typeof meta.updatedAt === 'number' ? meta.updatedAt : 0;
    } else {
      try {
        const t = SharePointService.oDataStr(LIST_CHECKLIST_META);
        const d = await this.spGet(`/_api/web/lists/getbytitle('${t}')/items?$top=500`);
        const legacyMeta = (d.value || []).find((i: { Title?: string }) => String(i.Title || '') === projId);
        if (legacyMeta) {
          const lm = legacyMeta as { projectType?: string; updatedAt?: number };
          if (lm.projectType === 'concrete' || lm.projectType === 'both') projectType = lm.projectType;
          updatedAt = Number(lm.updatedAt) || 0;
        }
      } catch { /* ignore */ }
    }

    if (Object.keys(items).length === 0 && overrides.length === 0 && !metaRow && updatedAt === 0) {
      return null;
    }

    return { items, overrides, projectType, updatedAt };
  }

  public async loadChecklist(projId: string): Promise<ChecklistPersisted | null> {
    const columnsReady = await this.ensureChecklistLists();

    let fromLists: ChecklistPersisted | null = null;
    if (columnsReady) {
      try {
        fromLists = await this.loadChecklistFromLists(projId);
      } catch { /* fall through to settings */ }
    }

    let fromSettings: ChecklistPersisted | null = null;
    try {
      fromSettings = await this.loadChecklistFromSettings(projId);
    } catch { /* ignore */ }

    if (fromLists && fromSettings) {
      return (fromSettings.updatedAt || 0) >= (fromLists.updatedAt || 0) ? fromSettings : fromLists;
    }
    return fromLists || fromSettings;
  }

  private async saveChecklistToLists(projId: string, data: ChecklistPersisted): Promise<void> {
    const updatedAt = data.updatedAt || Date.now();

    const itemRows: CrmListRowInput[] = Object.entries(data.items || {}).map(([itemId, st]) => ({
      crmId: SharePointService.checklistItemCrmId(projId, itemId),
      title: `${projId} · ${itemId}`,
      recordNum: itemId,
      payloadJson: JSON.stringify(st),
    }));
    await this.syncProjectCrmRows(
      LIST_CHECKLIST,
      crmId => crmId.startsWith(`chk|${projId}|`),
      itemRows,
    );

    const ovRows: CrmListRowInput[] = (data.overrides || []).map((ov, idx) => ({
      crmId: SharePointService.checklistOverrideCrmId(projId, ov),
      title: `${projId} · ${ov.itemId} · ${ov.at}`,
      recordNum: String(idx),
      payloadJson: JSON.stringify(ov),
    }));
    await this.syncProjectCrmRows(
      LIST_CHECKLIST_OVERRIDES,
      crmId => crmId.startsWith(`chkov|${projId}|`),
      ovRows,
    );

    await this.syncProjectCrmRows(
      LIST_CHECKLIST_META,
      crmId => crmId === SharePointService.checklistMetaCrmId(projId) || crmId === projId,
      [{
        crmId: SharePointService.checklistMetaCrmId(projId),
        title: projId,
        recordNum: projId,
        payloadJson: JSON.stringify({ projectType: data.projectType || 'steel', updatedAt }),
      }],
    );
  }

  public async saveChecklist(projId: string, data: ChecklistPersisted): Promise<'list' | 'settings'> {
    const columnsReady = await this.ensureChecklistLists();
    if (columnsReady) {
      try {
        await this.saveChecklistToLists(projId, data);
        return 'list';
      } catch {
        /* list sync failed — fall back to 3Edge_Settings */
      }
    }
    await this.saveChecklistToSettings(projId, data);
    return 'settings';
  }
}

export type ChecklistC2Action = 'cleared' | 'na' | 'incorrect';

export interface ChecklistItemState {
  c1?: boolean;
  c2?: ChecklistC2Action | null;
  c1By?: string;
  c2By?: string;
  c1At?: string;
  c2At?: string;
  override?: boolean;
  overrideBy?: string;
  overrideAt?: string;
  overrideReason?: string;
}

export interface ChecklistOverrideLog {
  itemId: string;
  by: string;
  at: string;
  reason: string;
  itemText: string;
  taskCode: string;
  action: ChecklistC2Action;
}

export interface ChecklistPersisted {
  items: Record<string, ChecklistItemState>;
  overrides: ChecklistOverrideLog[];
  projectType: 'steel' | 'concrete' | 'both';
  updatedAt?: number;
}
