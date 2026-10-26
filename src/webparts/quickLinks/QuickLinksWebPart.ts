import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'; 
 
import QuickLinks from './components/QuickLinks';
import { IQuickLinksProps } from './components/IQuickLinksProps';

export interface IQuickLinksWebPartProps {
  gridListName: string;
  linkListName:string;
  
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(
      QuickLinks,
      {  
        parentListName: this.properties.gridListName,
        childListName: this.properties.linkListName,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

// No onInit at all
 
 

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
            description: "Quick Links Properties "
          },
          groups: [
            { 
              groupFields: [
                PropertyPaneTextField('gridListName', {
                  label: "Parent List Name"
                }),
                PropertyPaneTextField('linkListName', {
                  label: "Child List Name"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
