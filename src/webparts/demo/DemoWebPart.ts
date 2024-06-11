
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'DemoWebPartStrings';
import Demo from './components/Demo';
import { IDemoProps } from './components/Demo';

import { spfi, SPFI, SPFx } from "@pnp/sp";

export interface IDemoWebPartProps {
  description: string;
}

export default class DemoWebPart extends BaseClientSideWebPart<IDemoWebPartProps> {
  private sp: SPFI;

  public async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<IDemoProps> = React.createElement(
      Demo,
      {
        description: this.properties.description,
        sp: this.sp
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
