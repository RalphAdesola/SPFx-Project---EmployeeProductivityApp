import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import EmployeeProductivityApp from './components/EmployeeProductivityApp';
import { IEmployeeProductivityAppProps } from './components/IEmployeeProductivityAppProps';

export interface IEmployeeProductivityAppWebPartProps {
}

export default class EmployeeProductivityAppWebPart extends BaseClientSideWebPart<IEmployeeProductivityAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEmployeeProductivityAppProps> = React.createElement(
      EmployeeProductivityApp,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
