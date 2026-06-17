import type { CrmAttachment } from './crmTypes';

const LS_RFQ_ATT = '3edge-crm-rfq-attachments';
const LS_QUOTE_ATT = '3edge-crm-quote-attachments';

const uid = (): string => `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

const loadMap = (key: string): Record<string, CrmAttachment[]> => {
  try {
    const v = localStorage.getItem(key);
    return v ? (JSON.parse(v) as Record<string, CrmAttachment[]>) : {};
  } catch {
    return {};
  }
};

const saveMap = (key: string, map: Record<string, CrmAttachment[]>): void => {
  try {
    localStorage.setItem(key, JSON.stringify(map));
  } catch { /* ignore quota */ }
};

export const fileToCrmAttachment = (file: File): Promise<CrmAttachment> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({ id: uid(), name: file.name, dataUrl: reader.result as string });
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });

export const getRfqAttachments = (rfqId: string): CrmAttachment[] =>
  loadMap(LS_RFQ_ATT)[rfqId] ?? [];

export const setRfqAttachments = (rfqId: string, attachments: CrmAttachment[]): void => {
  const map = loadMap(LS_RFQ_ATT);
  if (attachments.length) map[rfqId] = attachments;
  else delete map[rfqId];
  saveMap(LS_RFQ_ATT, map);
};

export const getQuoteAttachments = (quoteId: string): CrmAttachment[] =>
  loadMap(LS_QUOTE_ATT)[quoteId] ?? [];

export const quoteHasAttachments = (quoteId: string): boolean =>
  getQuoteAttachments(quoteId).length > 0;

export const getAllQuoteAttachmentIds = (): Set<string> => {
  const map = loadMap(LS_QUOTE_ATT);
  return new Set(Object.keys(map).filter(id => (map[id]?.length ?? 0) > 0));
};

export const setQuoteAttachments = (quoteId: string, attachments: CrmAttachment[]): void => {
  const map = loadMap(LS_QUOTE_ATT);
  if (attachments.length) map[quoteId] = attachments;
  else delete map[quoteId];
  saveMap(LS_QUOTE_ATT, map);
};

export const copyRfqAttachmentsToQuote = (rfqId: string, quoteId: string): void => {
  const src = getRfqAttachments(rfqId);
  if (src.length) setQuoteAttachments(quoteId, src.map(a => ({ ...a, id: uid() })));
};
