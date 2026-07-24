import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeProductivityAppProps {
  userDisplayName: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
