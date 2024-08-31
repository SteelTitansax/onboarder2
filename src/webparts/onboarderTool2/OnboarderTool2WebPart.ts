import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './OnboarderTool2WebPart.module.scss';
import * as strings from 'OnboarderTool2WebPartStrings';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';


export interface SPOnboardAuxList{
  value: SPOnboardAuxListItem[];
}

export interface SPOnboardAuxListItem{
  Title : string;
  Name : string;
  Role: string;
  Status: string

}


export interface IOnboarderTool2WebPartProps {
  description: string;
}

export default class OnboarderTool2WebPart extends BaseClientSideWebPart<IOnboarderTool2WebPartProps> {

 

  private _getListData(id : string ) : Promise <SPOnboardAuxList>{

    return this.context.spHttpClient.get(`https://t8656.sharepoint.com/sites/Sharepoint_Interaction/_api/web/lists/getbytitle('Poc_SharepointInteractionAux')/items?$filter=Title eq '${id}'`,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => { return response.json()});
  }

  private _renderList(workerName : string): void {
    this._getListData(workerName).then((response) => {

      let html: string = `<table width=100% style='border: 1px solid'><tr>
            <th>Name</th>
            <th>Role</th>
            <th>Status</th>
          </tr>`;
      response.value.forEach((item: SPOnboardAuxListItem) =>{
        
          html += `
          
          <tr>
              <td style='border: 1px solid'> ${item.Name} </td> 
              <td style='border: 1px solid'> ${item.Role} </td> 
              <td style='border: 1px solid'> ${item.Status} </td> 
          </tr>
        `    

      });

      html += '</table>'
      
      const listDiv = this.domElement.querySelector('#spListDiv');
      if(listDiv){ listDiv.innerHTML = html;}else{console.log("listDiv not found");}
    
    }
    
  )};
  

  public render(): void  {
    this.domElement.innerHTML = `
    <div>
      <p><strong> Onboarding Status Tracker </strong></p>
      <label>Worker Name</label>
      <input type="text" placeholder=" " id="workerName"/>
      <div id="spListDiv" class="${styles.tableContainer}"></div> 
    </div>  
    `
    this._urlPrepopulation();
    const workerName = (document.getElementById("workerName") as HTMLInputElement).value
    this._renderList(workerName);
  }

  private _urlPrepopulation(): void {
    var url = window.location.href;

    // Define regex patterns for each parameter with values (textInput and dropdowns)
    const patterns: { [key: string]: RegExp } = {
      workerName: /[?&]workerName=([^&]+)/,
    };

    // Define the type for the extracted values
    type ExtractedValues = {
      [key: string]: string | null;
    };

    // Function to extract values using regex patterns
    function extractValues(
      url: string,
      patterns: { [key: string]: RegExp }
    ): ExtractedValues {
      const values: ExtractedValues = {};
      for (const key in patterns) {
        const match = url.match(patterns[key]);
        values[key] = match ? match[1] : null;
      }
      return values;
    }

    // Extract and print the values
    const prepopulatedValues: ExtractedValues = extractValues(url, patterns);

    for (var prepopulatedValue in prepopulatedValues) {
      prepopulatedValues[prepopulatedValue] == undefined
        ? (prepopulatedValues[prepopulatedValue] = "Insert worker name")
        : prepopulatedValues[prepopulatedValue];
      (
        document.getElementById(`${prepopulatedValue}`) as HTMLInputElement
      ).value = `${prepopulatedValues[prepopulatedValue]?.replace(
        /%20/g,
        " "
      )}`;
    }

   
}

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
