import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import { IDevOpsService } from '../../services/devOpsService/IDevOpsService';
import { DevOpsService } from '../../services/devOpsService/DevOpsService';

import { DevOps } from './components/DevOps';
import { IDevOpsProps } from './components/IDevOpsProps';

import * as strings from 'DevOpsWebPartStrings';

export interface IDevOpsWebPartProps {
  description: string;
}

export default class DevOpsWebPart extends BaseClientSideWebPart<IDevOpsWebPartProps> {

  private _devOpsService : IDevOpsService;

  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      const serviceScope = this.context.serviceScope;
      this._devOpsService = serviceScope.consume(DevOpsService.serviceKey);  
    });
  }

  public render(): void {

    const element: React.ReactElement<IDevOpsProps> = React.createElement(
      DevOps,
      {
        devOpsService: this._devOpsService
      }
    );

    ReactDom.render(element, this.domElement);
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
