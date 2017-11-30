import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListColumnsAutoFetchWebPart.module.scss';
import * as strings from 'ListColumnsAutoFetchWebPartStrings';

import pnp from "sp-pnp-js";
import { Web } from "sp-pnp-js";
import { List } from 'sp-pnp-js/lib/sharepoint/lists';

export interface IListColumnsAutoFetchWebPartProps {
  WebUrl: string;
  ListTitle: string;
  Column: string;
  description: string;
}

export default class ListColumnsAutoFetchWebPartWebPart extends BaseClientSideWebPart<IListColumnsAutoFetchWebPartProps> {


  private _listDropdownOptions : IPropertyPaneDropdownOption[] = [];
  private _columsDropdownOptions : IPropertyPaneDropdownOption[] = [];

  public render(): void {
    
		let web = new Web(`${this.properties.WebUrl}`);
    web.lists.getByTitle(`${this.properties.ListTitle}`).items.get().then((items: any[]) => {
      this._renderList(items);
    }).catch((err) => {
    });

    // this.domElement.innerHTML = `
    //   <div class="${ styles.listColumnsAutoFetch }">
    //     <div class="${ styles.container }">
    //       <div class="${ styles.row }">
    //         <div class="${ styles.column }">
    //           <span class="${ styles.title }">Welcome to SharePoint!</span>
    //           <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    //           <p class="${ styles.description }">${escape(this.properties.description)}</p>
    //           <a href="https://aka.ms/spfx" class="${ styles.button }">
    //             <span class="${ styles.label }">Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public onInit<T>(): Promise<T> {
    if(this.displayMode === DisplayMode.Edit){
      this._getListTitles();
    }
    return Promise.resolve();
  }

  private _getListTitles(): void {
    this._listDropdownOptions = [];
    this._columsDropdownOptions = [];
		let web = new Web(`${this.properties.WebUrl}`);
    web.lists.get().then((response:any[]) => {
      var listID = [];
      for(var index = 0; index < response.length; index++){
        if(!response[index].BaseType && !response[index].Hidden){
          this._listDropdownOptions.push({
            key : response[index].Id,
            text : response[index].Title
          });
          listID.push(response[index].Id);
        }
      }
      if(listID.indexOf(this.properties.ListTitle) >= 0){
        this._getListColumns();
      }
      else{
        this.context.propertyPane.refresh();
      }
    });
  }

  private _getListColumns(): void {
      let web = new Web(`${this.properties.WebUrl}`);
      web.lists.getById(this.properties.ListTitle).fields.get().then((listColumnDetails:any[]) =>{
        var columnsList = [];
        for(var index = 0; index < listColumnDetails.length; index++){
          if(!listColumnDetails[index].Hidden){
            this._columsDropdownOptions.push({
              key: listColumnDetails[index].InternalName,
              text: listColumnDetails[index].Title
            });
            columnsList.push(listColumnDetails[index].InternalName);
          }
        }
        this.context.propertyPane.refresh();
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void{
    if(propertyPath === "WebUrl"){
      this._listDropdownOptions = [];
      this._columsDropdownOptions = [];
      // this.properties.columns = "";
      // this.properties.List = "";
      this._getListTitles();
    }
    else if(propertyPath === "ListTitle"){
      this._columsDropdownOptions = [];
      // this.properties.columns = "";
      this._getListColumns();
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
                PropertyPaneTextField('WebUrl',{
                  label : strings.WebUrlFieldLabel
                }),
                PropertyPaneDropdown('List', {
                  label: strings.ListTitleFieldLabel,
                  options : this._listDropdownOptions
                }),
                PropertyPaneDropdown('columns', {
                  label: strings.ColumnFieldLabel,
                  options : this._columsDropdownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
  public _renderList(ListItems: any[]): void {
    
    let htmlString : string = `<div clas="${styles.listColumnsAutoFetch}">`;
    htmlString += `<ul class="${styles.listItemsContainer}">`;
    for(var index = 0; index < ListItems.length; index++){
      htmlString += `<li class="${styles.listItem}">${ListItems[index]["Title"]} : ${ListItems[index][this.properties.Column]}</li>`;
    }
    htmlString += `</ul></div>`;
    this.domElement.innerHTML = htmlString;
  }
}
