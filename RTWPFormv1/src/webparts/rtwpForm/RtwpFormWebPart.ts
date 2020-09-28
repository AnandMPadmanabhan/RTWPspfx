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
  WorkLoc_Column:string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}
export interface ISPCols {
  value: ISPCol[];
}

export interface ISPCol {
  field: string;
}
  
export default class RtwpFormWebPart extends BaseClientSideWebPart <IRtwpFormWebPartProps> {
  private curr_user=sp.web.currentUser;
  private listDropDownOptions: IPropertyPaneDropdownOption[] =[]; 
  private listDropDownOptions1: IPropertyPaneDropdownOption[] =[];
  private listDropDown1Disabled :boolean=true;
  public render(): void {
    
    const element: React.ReactElement<IRtwpFormProps> = React.createElement(
     
      RtwpForm,
      {  
        description: this.properties.description,
        context1:this.context,
        WorkLoc:this.properties.WorkLoc,
        WorkLoc_Column:this.properties.WorkLoc_Column
      }
    );
    console.log("Here111");
    this.GetLists();
    if(this.properties.WorkLoc_Column!=null){
    ReactDom.render(element, this.domElement);
    }
    
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
    private GetLists():void{  
      // REST API to pull the list names  
      let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Id,Title";  
      
      this.LoadLists(listresturl).then((response)=>{  
        // Render the data in the web part  
        this.LoadDropDownValues(response.value);  
      });  
    }  
      
    private LoadLists(listresturl:string): Promise<ISPLists>{  
      // Call to site to get the list names  
      return this.context.spHttpClient.get(listresturl,SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
        return response.json();  
      });  
    }  
      
    private LoadDropDownValues(lists: ISPList[]): void{  
      lists.forEach((list:ISPList)=>{  
        // Loads the drop down values  
        this.listDropDownOptions.push({key:list.Title,text:list.Title});  
      });  
    }   
    private GetColumns():void{  
      // REST API to pull the list names  
      let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('"+this.properties.WorkLoc+"')/Views/getbytitle('All Items')/ViewFields";  
      
      this.LoadColumns(listresturl).then((response)=>{  
        // Render the data in the web part  
        console.log(response.value);
        this.LoadColumnValues(response.value);  
        this.context.propertyPane.refresh();
      });  
    }  
      
    private LoadColumns(listresturl:string): Promise<ISPCols>{  
      // Call to site to get the list names  
      return this.context.spHttpClient.get(listresturl,SPHttpClient.configurations.v1).then((response: SPHttpClientResponse)=>{  
        return response.json();  
      });  
    }  
      
    private LoadColumnValues(lists: ISPCol[]): void{  
       console.log(lists);
      lists.forEach((list:ISPCol)=>{  
        // Loads the drop down values  

        this.listDropDownOptions1.push({key:list.field,text:list.field});  
      });  
    } 
  
protected onPropertyPaneFieldChanged(propertyPath: string,oldValue:any, newValue: any):void{  
   if(propertyPath === "WorkLoc"){   
// Change only when drop down changes 
  
  super.onPropertyPaneFieldChanged(propertyPath,oldValue,newValue);   
// Clears the existing data  
  this.properties.WorkLoc_Column = undefined;   
  this.onPropertyPaneFieldChanged('WorkLoc_Column',this.properties.WorkLoc_Column,newValue);   

// Get/Load new items data  
  this.GetColumns(); 
  this.listDropDown1Disabled=false;  
  }   
  else {   
// Render the property field  
  super.onPropertyPaneFieldChanged(propertyPath,oldValue,newValue);    
  } 
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
                }),
                PropertyPaneDropdown('WorkLoc', {
                  label: strings.WorkFieldLabel,
                  selectedKey:"Select Work Location List",
                  options:this.listDropDownOptions
                }),
                PropertyPaneDropdown('WorkLoc_Column', {
                  label: strings.WorkCFieldLabel,
                  selectedKey:"Select Work Location List",
                  options:this.listDropDownOptions1,
                  disabled:this.listDropDown1Disabled

                })
              ]
            }
          ]
        }
      ]
    };
  }
}
