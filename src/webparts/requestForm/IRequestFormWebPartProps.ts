import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IRequestFormWebPartProps {
  context: WebPartContext;
  title: string;
  description: string;
  requestTypes: string;
  maxFileSize: string;
  showManagerField: boolean;
  requireManagerApproval: boolean;
} 