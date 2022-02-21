import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyAgendaWebPartStrings';
import MyAgenda from './components/MyAgenda';
import { IMyAgendaProps } from './components/IMyAgendaProps';


export interface IMyAgendaWebPartProps {
  description: string;
}

import { Providers, SharePointProvider } from "@microsoft/mgt-spfx";

export default class MyAgendaWebPart extends BaseClientSideWebPart<IMyAgendaWebPartProps> {

  protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<IMyAgendaProps> = React.createElement(
      MyAgenda,
      {
        description: this.properties.description,
        displayMode: this.displayMode,
        updateProperty: (value: string): void => {
          // store the new description in the description web part property
          this.properties.description = value;
        }
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

              ]
            }
          ]
        }
      ]
    };
  }
}
