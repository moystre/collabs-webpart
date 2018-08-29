import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import * as strings from 'CollaboratorsWebPartStrings';
import Collaborators from './components/Collaborators';
import { ICollaboratorsProps } from './components/ICollaboratorsProps';

import { TextField } from 'office-ui-fabric-react/lib/TextField';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { createRef } from 'office-ui-fabric-react/lib/Utilities';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface ICollaboratorsWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
  wpEmail: string;
  wpPhone: string;
}

export default class CollaboratorsWebPart extends BaseClientSideWebPart<ICollaboratorsWebPartProps> {
  public fetchedIsp: string[];
  public columns: IColumn[];

  constructor() {
    super();
    this.fetchedIsp = null;
    this.columns = [
      {
        key: 'column1',
        name: 'Title',
        fieldName: 'title',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        ariaLabel: 'Operations for name'
      },
      {
        key: 'wpEmail',
        name: 'Email',
        fieldName: 'wpEmail',
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        ariaLabel: 'Operations for value'
      }]
  }

  protected async onInit(): Promise<void> {
    this.fetchedIsp = await this.getspLists();
  }

  public render(): void {
    const element: React.ReactElement<ICollaboratorsProps> = React.createElement(
      Collaborators,
      {
        description: this.properties.description,
        ispLists: this.fetchedIsp,
        columns: this.columns
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async getspLists(): Promise<string[]> {
    var renderedList: string[];
    if (Environment.type === EnvironmentType.Local) {
      console.log('Local environment');
      return null;
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      try {
        let container = null;
        var list: string[] = [];
        container = await this._getListData();
        container.value.forEach((item: ISPList) => {
          console.log(item);
          list.push(item.Title);
        });
        renderedList = list;
      }
      catch (exception) {
        console.warn(exception);
      }
      return renderedList;
    }
  }

  private _getListData = async (): Promise<ISPLists> => {
    let returnLists: ISPLists = null;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
      /*  + `/_api/web/lists?$filter=Hidden eq false`, */
        + `/_api/web/lists/GetByTitle('Collaborators')/items`,  
        SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw "Could not fetch list data";
      }
      const lists: ISPLists = await response.json();
      returnLists = lists;
    } catch (exception) {
      console.warn(exception);
    }
    return returnLists;
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
