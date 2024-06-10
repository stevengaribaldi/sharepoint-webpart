// //sharepoint webpart class
// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import { Version } from '@microsoft/sp-core-library';

// import { getSP } from './pnpjsConfig';
// import {
//   type IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';
// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';

// import * as strings from 'DemoWebPartStrings';
// import Demo from './components/Demo';
// import { IDemoProps } from './components/IDemoProps';

// export interface IDemoWebPartProps {
//   description: string;
// }

// export default class DemoWebPart extends BaseClientSideWebPart<IDemoWebPartProps> {

//   private _isDarkTheme: boolean = false;
//   private _environmentMessage: string = '';


//   public async onInit(): Promise<void> {
//   try {
//     this._environmentMessage = await this._getEnvironmentMessage();

//     await super.onInit();

//     //Initialize our _sp object that we can then use in other packages without having to pass around the context.
//     // Check out pnpjsConfig.ts for an example of a project setup file.
// //this is are sharepoint context
//     getSP(this.context);
//   } catch (error) {
//     console.error("Error during initialization:", error);
//   }
// }


//   public render(): void {
//     const element: React.ReactElement<IDemoProps> = React.createElement(
//       Demo,
//       {
//         description: this.properties.description,
//         isDarkTheme: this._isDarkTheme,
//         environmentMessage: this._environmentMessage,
//         hasTeamsContext: !!this.context.sdks.microsoftTeams,
//         userDisplayName: this.context.pageContext.user.displayName
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }



//   private _getEnvironmentMessage(): Promise<string> {
//     if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
//       return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
//         .then(context => {
//           let environmentMessage: string = '';
//           switch (context.app.host.name) {
//             case 'Office': // running in Office
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
//               break;
//             case 'Outlook': // running in Outlook
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
//               break;
//             case 'Teams': // running in Teams
//             case 'TeamsModern':
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
//               break;
//             default:
//               environmentMessage = strings.UnknownEnvironment;
//           }

//           return environmentMessage;
//         });
//     }

//     return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
//   }

//   protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
//     if (!currentTheme) {
//       return;
//     }

//     this._isDarkTheme = !!currentTheme.isInverted;
//     const {
//       semanticColors
//     } = currentTheme;

//     if (semanticColors) {
//       this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
//       this.domElement.style.setProperty('--link', semanticColors.link || null);
//       this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
//     }

//   }

//   protected onDispose(): void {
//     ReactDom.unmountComponentAtNode(this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }


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
import { IAsyncAwaitPnPJsProps } from './components/Demo';


import { spfi, SPFI, SPFx } from "@pnp/sp";

export interface IDemoWebPartProps {
  description: string;
}

export default class DemoWebPart extends BaseClientSideWebPart<IDemoWebPartProps> {
  private sp: SPFI;

  // // https://github.com/SharePoint/PnP-JS-Core/wiki/Using-sp-pnp-js-in-SharePoint-Framework
  public async onInit(): Promise<void> {
    await super.onInit();

    this.sp = spfi().using(SPFx(this.context));
  }

  public render(): void {
    const element: React.ReactElement<IAsyncAwaitPnPJsProps> = React.createElement(
      Demo,
      {
        description: this.properties.description,
        sp: this.sp
      }
    );

    ReactDom.render(element, this.domElement);
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

