import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyFirstTeamsWebpartWebPart.module.scss';
import * as strings from 'MyFirstTeamsWebpartWebPartStrings';

export interface IMyFirstTeamsWebpartWebPartProps {
  description: string;
}

export default class MyFirstTeamsWebpartWebPart extends BaseClientSideWebPart<IMyFirstTeamsWebpartWebPartProps> {
  private _msTeamsContext: microsoftTeams.Context;

  protected onInit() : Promise<any> {
    let promise = Promise.resolve();

    if(this.context.microsoftTeams) {
      promise = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext((ctx) => {
          this._msTeamsContext = ctx;
          resolve();
        });
      });
    }

    return promise;
  }

  public render(): void {

    let title;
    if(this._msTeamsContext) {
      title = `From Teams with ❤: ${this._msTeamsContext.teamName}`;
    }
    else {
      title = `From SharePoint with ❤: ${this.context.pageContext.web.title}`;
    }

    this.domElement.innerHTML = `
      <div class="${ styles.myFirstTeamsWebpart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${title}</span>
            </div>
          </div>
        </div>
      </div>`;
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
