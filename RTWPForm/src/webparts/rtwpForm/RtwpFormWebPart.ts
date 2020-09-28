import * as React from 'react';
import * as ReactDom from 'react-dom';
import { sp } from "@pnp/sp";  
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import * as strings from 'RtwpFormWebPartStrings';
import RtwpForm from './components/RtwpForm';
import { IRtwpFormProps } from './components/IRtwpFormProps';

export interface IRtwpFormWebPartProps {
  description: string;
  WorkLoc: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}
  
export default class RtwpFormWebPart extends BaseClientSideWebPart <IRtwpFormWebPartProps> {
  private curr_user=sp.web.currentUser;
  private listDropDownOptions: IPropertyPaneDropdownOption[] =[]; 
  public render(): void {
    
    const element: React.ReactElement<IRtwpFormProps> = React.createElement(
     
      RtwpForm,
      {  
        description: this.properties.description,
        context1:this.context,
        WorkLoc:this.properties.WorkLoc
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
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
    sp.setup({
    spfxContext: this.context
    });
    });
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
