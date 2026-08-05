import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import Dashboard from './components/Dashboard';
import { IDhaPortfolioManagerProps } from './IDhaPortfolioManagerProps';

export default class DhaPortfolioManagerWebPart extends BaseClientSideWebPart<IDhaPortfolioManagerProps> {
  private redirectToWebView(): boolean {
    const isWorkbench = window.location.pathname.toLowerCase().indexOf('/workbench.aspx') >= 0;
    const isEmbedded = window.self !== window.top || Boolean(this.context.sdks.microsoftTeams);
    if (this.displayMode !== DisplayMode.Read || Environment.type === EnvironmentType.Local || isWorkbench || isEmbedded) return false;

    const url = new URL(window.location.href);
    const environment = url.searchParams.get('env');
    if (environment && environment.toLowerCase() === 'webview') return false;

    url.searchParams.set('env', 'WebView');
    window.location.replace(url.toString());
    return true;
  }

  public render(): void {
    if (this.redirectToWebView()) return;
    const element = React.createElement(Dashboard, { context: this.context, intakeListName: this.properties.intakeListName, peopleListName: this.properties.peopleListName, ledgerListName: this.properties.ledgerListName });
    ReactDom.render(element, this.domElement);
  }
  protected get dataVersion(): any { return '1.0'; }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: 'Configure the SharePoint lists that power this dashboard.' }, groups: [{ groupName: 'Data sources', groupFields: [
      PropertyPaneTextField('intakeListName', { label: 'DHA Intake list name' }),
      PropertyPaneTextField('peopleListName', { label: 'ResMan People list name' }),
      PropertyPaneTextField('ledgerListName', { label: 'ResMan TLedger list name' })
    ] }] }] };
  }
}
