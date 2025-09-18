import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHrAdminGridWebPartProps {
  context: WebPartContext;
  title: string;
  description: string;
  showFilters: boolean;
  itemsPerPage: string;
  enableBulkActions: boolean;
} 