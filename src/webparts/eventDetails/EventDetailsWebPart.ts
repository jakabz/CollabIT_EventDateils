import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'EventDetailsWebPartStrings';
import EventDetails from './components/EventDetails';
import { IEventDetailsProps } from './components/IEventDetailsProps';

export interface IEventDetailsWebPartProps {
  listItem: string;
  TitleViewBool: boolean;
  LocationViewBool: boolean;
  StartViewBool: boolean;
  EndViewBool: boolean;
  SaveEventButtonViewBool: boolean;
  RegisterButtonViewBool:boolean;
  listsItems: any;
  pageContext: any;
  eventlistid: string;
  registrationlistid: string;
  userID: string;
  registeredItem: any;
  context:any;
}

export interface ISPLists {
  value: ISPItems[];
}

export interface ISPItems {
  Title: string;
  Id: string;
  EventDate: Date;
  EndDate: Date;
  Location: string;
  EntityTypeName: string;
  EventID: string;
  PersonId: string;
}

export default class EventDetailsWebPart extends BaseClientSideWebPart<IEventDetailsWebPartProps> {

  private listsItems = [];
  private _dropdownOptionsItems: IPropertyPaneDropdownOption[] = [];
  private eventlistid = "";
  private registrationlistid = "";

  public render(): void {
    const element: React.ReactElement<IEventDetailsProps > = React.createElement(
      EventDetails,
      {
        listItem: this.properties.listItem,
        TitleViewBool: this.properties.TitleViewBool,
        LocationViewBool: this.properties.LocationViewBool,
        StartViewBool: this.properties.StartViewBool,
        EndViewBool: this.properties.EndViewBool,
        SaveEventButtonViewBool: this.properties.SaveEventButtonViewBool,
        RegisterButtonViewBool: this.properties.RegisterButtonViewBool,
        listsItems: this.listsItems,
        pageContext: this.context.pageContext,
        context: this.context,
        eventlistid: this.eventlistid,
        registrationlistid: this.registrationlistid,
        userID: this.context.pageContext.legacyPageContext.userId,
        registeredItem: this.properties.registeredItem
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public onInit<T>(): Promise<T> {
    this._getlistInfo('Events')
      .then((response) => {
        this.eventlistid = response.Id;
      });
    this._getlistInfo('EventRegistration')
      .then((response) => {
        this.registrationlistid = response.Id;
      });
    this._getListData('Events',false)
      .then((response) => {
        this._dropdownOptionsItems = response.value.map((list: ISPItems) => {
          this.listsItems.push(list);
          return {
            key: list.Id,
            text: list.Title
          };
        });
        this.render();
      });
    this._getListData('EventRegistration', true)
      .then((response) => {
        //console.info(response.value);
        this.properties.registeredItem = response.value[0];
      });
    return Promise.resolve();
  }

  private _getListData(listName: string, query: boolean): Promise<ISPLists> {
    var querySring = '';
    if(query){
      querySring = 'EventID eq '+this.properties.listItem+' and PersonId eq '+this.context.pageContext.legacyPageContext.userId;
    }
    return this.context.spHttpClient  
    .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/Items?$filter=${querySring}`,
        SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getlistInfo(listName: string): Promise<any> {
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')`,
        SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('listItem', {
                  label: strings.listItemFieldLabel,
                  options: this._dropdownOptionsItems
                })
              ]
            },
            {
              groupName: strings.ViewFieldsGroup,
              groupFields: [
                PropertyPaneToggle('TitleViewBool', {
                  label: strings.TitleViewFieldLabel
                }),
                PropertyPaneToggle('StartViewBool', {
                  label: strings.StartViewFieldLabel
                }),
                PropertyPaneToggle('EndViewBool', {
                  label: strings.EndViewFieldLabel
                }),
                PropertyPaneToggle('LocationViewBool', {
                  label: strings.LocationViewFieldLabel
                }),
                PropertyPaneToggle('RegisterButtonViewBool', {
                  label: strings.RegisterViewFieldLabel
                }),
                PropertyPaneToggle('SaveEventButtonViewBool', {
                  label: strings.SaveEventViewFieldLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
