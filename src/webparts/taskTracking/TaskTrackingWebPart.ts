
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TaskTrackingWebPartStrings';
import TaskTracking from './components/TaskTracking';
import { ITaskTrackingProps } from './components/ITaskTrackingProps';
import { taskService } from '../../services/sp-service';

export interface ITaskTrackingWebPartProps {
  description: string;
}

export default class TaskTrackingWebPart extends BaseClientSideWebPart<ITaskTrackingWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _parentTaskId: number | undefined;
  private _childTaskId: number | undefined;
  private _viewTaskId: number | undefined;

  public async onInit(): Promise<void> {
    await super.onInit();
    taskService.init(this.context);

    // Read URL parameters for direct task navigation
    const urlParams = new URLSearchParams(window.location.search);
    const parentTaskIdStr = urlParams.get('ParentTaskID');
    const childTaskIdStr = urlParams.get('ChildTaskID');
    const viewTaskIdStr = urlParams.get('ViewTaskID');

    if (parentTaskIdStr) {
      this._parentTaskId = parseInt(parentTaskIdStr, 10);
      console.log('[TaskTracking] URL Parameter - ParentTaskID:', this._parentTaskId);
    }

    if (childTaskIdStr) {
      this._childTaskId = parseInt(childTaskIdStr, 10);
      console.log('[TaskTracking] URL Parameter - ChildTaskID:', this._childTaskId);
    }

    if (viewTaskIdStr) {
      this._viewTaskId = parseInt(viewTaskIdStr, 10);
      console.log('[TaskTracking] URL Parameter - ViewTaskID:', this._viewTaskId);
    }

    // Original environment message logic
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  public render(): void {
    const element: React.ReactElement<ITaskTrackingProps> = React.createElement(
      TaskTracking,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        parentTaskId: this._parentTaskId,
        childTaskId: this._childTaskId,
        viewTaskId: this._viewTaskId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
