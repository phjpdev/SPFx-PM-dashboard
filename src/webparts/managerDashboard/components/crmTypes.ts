export interface CrmPhone { value: string; type: string; cc?: string; }
export interface CrmEmail { value: string; type: string; }

export type CrmActivityType = 'Phone call' | 'Email' | 'Text' | 'In person';

export interface CrmActivity {
  id: string;
  date: string;
  type: CrmActivityType;
  notes: string;
  followUpDate: string;
  done: boolean;
}

export interface CrmAttachment {
  id: string;
  name: string;
  dataUrl: string;
}

export interface CrmPerson {
  id: string;
  name: string;
  organizationId: string;
  position: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
  activities?: CrmActivity[];
  attachments?: CrmAttachment[];
}

export interface CrmCompany {
  id: string;
  name: string;
  labels: string;
  address: string;
  phones: CrmPhone[];
  emails: CrmEmail[];
}

export type CrmRfqDiscipline = 'Steel' | 'Concrete' | 'Both';
export type CrmRfqStage =
  | 'New Enquiry'
  | 'Under Review'
  | 'Ready to Quote'
  | 'Won'
  | 'Declined';

export interface CrmRfq {
  id: string;
  rfqNum: string;
  dateReceived: string;
  personId: string;
  organizationId: string;
  projectTitle: string;
  projectAddress: string;
  discipline: CrmRfqDiscipline;
  quoteRequiredBy: string;
  projectValue: number;
  approximateHours: number;
  engineerDrawingReceived: boolean;
  engineerDrawingDate: string;
  architectDrawingReceived: boolean;
  architectDrawingDate: string;
  rfiAllowed: boolean;
  createQuoteXero: boolean;
  relatedRfqId: string;
  notes: string;
  source: string;
  stage: CrmRfqStage;
  assignedTo: string;
}

export type CrmQuoteStatus = 'Draft' | 'Sent' | 'Accepted' | 'Declined';

/** Quote created from an RFQ at Ready to Quote stage. */
export interface CrmQuote {
  id: string;
  quoteNum: string;
  rfqId: string;
  rfqNum: string;
  quotedDate: string;
  dateReceived: string;
  personId: string;
  organizationId: string;
  projectTitle: string;
  projectAddress: string;
  discipline: CrmRfqDiscipline;
  projectValue: number;
  approximateHours: number;
  assignedTo: string;
  source: string;
  notes: string;
  createQuoteXero: boolean;
  status: CrmQuoteStatus;
}
