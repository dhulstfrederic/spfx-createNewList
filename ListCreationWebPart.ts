import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions} from '@microsoft/sp-http';

import styles from './ListCreationWebPart.module.scss';
import * as strings from 'ListCreationWebPartStrings';

export interface IListCreationWebPartProps {
  description: string;
}

export default class ListCreationWebPart extends BaseClientSideWebPart<IListCreationWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.listCreation} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
      <div>
          New list name: <input type='text' id='txtNewListName'/><br/><br/>
          New list description: <input type='text' id='txtNewListDescription'/><br/><br/>
          <input type='button' id='btnCreateNewList' value='Create new list'/><br/><br/>
      </div>
      
    </section>`;

    this.bindElements();
  }

  private bindElements(): void {
    this.domElement.querySelector("#btnCreateNewList").addEventListener('click', ()=>{this.createNewList()});
  }

  private createNewList(): void {
   
   let newListName = (document.getElementById("txtNewListName") as HTMLInputElement).value;
   
   let newListDescription = (document.getElementById("txtNewListDescription") as HTMLInputElement).value;
   

    const listUrl : string  = this.context.pageContext.web.absoluteUrl+"/_api/web/lists/GetByTitle('"+newListName+"')";
    this.context.spHttpClient.get(listUrl,SPHttpClient.configurations.v1)
    .then((response:SPHttpClientResponse) => {
      if (response.status==200) {
        alert("List already exists");
        return;
      }
      if (response.status==404) {
        const url: string = this.context.pageContext.web.absoluteUrl+"/_api/web/lists";
        const listDefinition: any = {
          "Title": newListName,
          "Description": newListDescription,
          "AllowContentTypes": true,
          "BaseTemplate": 100,
          "ContentTypesEnabled": true,
        };
        const SPHttpClientOptions: ISPHttpClientOptions= {
           "body" : JSON.stringify(listDefinition)
        };
        this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,SPHttpClientOptions)
        .then((response:SPHttpClientResponse)=>{
          if (response.status == 201) {
            alert("A new list has been created");
          } else {
            alert("Error message: "+response.status+" - "+ response.statusText);
          }

        });
      } else {
        alert("Error message: "+response.status+" - "+ response.statusText);
      }
    })
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
