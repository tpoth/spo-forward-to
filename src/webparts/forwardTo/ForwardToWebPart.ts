import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ForwardToWebPart.module.scss';
import * as strings from 'ForwardToWebPartStrings';

export interface IForwardToWebPartProps {
  forwardToUrl: string;
  forwardingDelay: number;
  forwardingActive: boolean;
}

export default class ForwardToWebPart extends BaseClientSideWebPart<IForwardToWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    if (this.properties.forwardingActive) {
      setTimeout(function () {
        if (this.displayMode != DisplayMode.Edit) {
          window.location.href = escape(this.properties.forwardToUrl);
        }
      }.bind(this), this.properties.forwardingDelay * 1000);
    }

    this.domElement.innerHTML = `
    <section class="${styles.forwardTo} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <h2>Forward To URL Web Part</h2>
        <div>Version 1.0</div>
        <div>The forwarding targets to this location: <strong>${escape(this.properties.forwardToUrl)}</strong></div>
      </div>
    </section>`;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('forwardToUrl', {
                  label: strings.ForwardToUrlFieldLabel
                }),
                PropertyPaneSlider('forwardingDelay', {
                  label: strings.ForwardingDelayFieldLabel,
                  min: 0,
                  max: 10
                }),
                PropertyPaneToggle('forwardingActive', {
                  label: strings.ForwardingActiveFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
