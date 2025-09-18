import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDeptGridWebPartProps {
  context: WebPartContext;
  title: string;
  description: string;
  showFilters: boolean;
  itemsPerPage: string;
} 