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
}

export default class CollaboratorsWebPart extends BaseClientSideWebPart<ICollaboratorsWebPartProps> {
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public render(): void {
    const element: React.ReactElement<ICollaboratorsProps> = React.createElement(
      Collaborators,
      {
        description: this.properties.description,
        lists: ["1", "2", "1", "2", "1", "2", "1", "2"]
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
    <ul>
      <li>
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  public _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
     console.log('// Local environment')
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
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
