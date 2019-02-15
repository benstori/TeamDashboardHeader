import { Version } from '@microsoft/sp-core-library';
import { sp, Items, ItemVersion, Web } from "@pnp/sp";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  Environment,
  EnvironmentType
 } from '@microsoft/sp-core-library';


import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

import styles from './TeamDashboardHeaderWebPart.module.scss';
import * as strings from 'TeamDashboardHeaderWebPartStrings';

import MockHttpClient from './MockHttpClient';

export interface ITeamDashboardHeaderWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
 }

 export interface ISPList {
  Title: string; // this is the department name in the List
  Id: string;
  DeptURL:string;
 }

const logo: any = require('./assests/Team.png');
//global vars
var userDept = "";

export default class TeamDashboardHeaderWebPart extends BaseClientSideWebPart<ITeamDashboardHeaderWebPartProps> {

  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      var propValue = "";
      var userDepartment = "";
  
      props.forEach(function(prop) {
        //this call returns key/value pairs so we need to look for the Dept Key
        if(prop.Key == "Department"){
          // set our global var for the users Dept.
          userDept += prop.Value;
        }
      });
      return result;
    }).then((result) =>{
      this._getListData().then((response) =>{
        this._renderList(response.value);
      });
    });
  });
  

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.teamDashboardHeader }">
                <div id="TeamDashboardHeader"/>
      </div>`;
      //this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }

  // main REST Call to the list...passing in the deaprtment into the call to 
  //return a single list item
  public _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/GetByTitle('TeamDashboardSettings')/Items?$filter=Title eq '`+ userDept +`'`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }
   
   //mock up 
   private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

   // this is required to use the SharePoint PnP shorthand REST CALLS
   public onInit():Promise<void> {
    return super.onInit().then (_=> {
      sp.setup({
        spfxContext:this.context
      });
    });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';

    items.forEach((item: ISPList) => {
      html += `
      <table style="width:100%;height:1px;">
        <tr>
          <td style="height:1px;text-align:right;width:12%">
          <a href="${item.DeptURL} target="_blank">
             <img id="TeamImage" class="${styles.headerImage}" src="${logo}" alt="GSVR Logo" /></a>
          </td>
          <td class="width:70%;height:1px;vertical-align:middle;"> 
          <h2 class="${styles.h2}"><a id="teamHeaderLink" href="${item.DeptURL}" target="_blank">Team Dashboard</a></h2>
          </td>
        </tr>
      </table>   
        `;
    });
 
    const listContainer: Element = this.domElement.querySelector('#TeamDashboardHeader');
    listContainer.innerHTML = html;
    

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
