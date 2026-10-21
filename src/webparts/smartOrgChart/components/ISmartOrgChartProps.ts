import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISmartOrgChartWebPartProps } from '../SmartOrgChartWebPart';
import { IExternalPeopleReport } from '../../../services/ExternalPeopleCsvService';

export type OrgChartTheme = 'modern' | 'minimal' | 'corporate' | 'dark' | 'custom';

export interface ISmartOrgChartProps extends ISmartOrgChartWebPartProps {
  context: WebPartContext;
  // Lets the property pane show how external people resolved (CSV file problems, unmatched managers)
  onExternalPeopleStatus?: (report: IExternalPeopleReport) => void;
}

export interface IUserSettings {
  alphabetFilterField: 'firstName' | 'lastName';
  cardSize: 'small' | 'medium' | 'large';
  showEmail: boolean;
  showPhone: boolean;
  showDepartment: boolean;
  showOffice: boolean;
  levelsAbove: number;
  compactCards: boolean;
  fontScale: number;
}
