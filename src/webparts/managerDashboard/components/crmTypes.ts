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
