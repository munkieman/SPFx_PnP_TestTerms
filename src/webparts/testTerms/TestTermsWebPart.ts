import { Version, Log } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  //PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  //PropertyPaneChoiceGroup,
  //IPropertyPaneGroup,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TestTermsWebPart.module.scss';
import * as strings from 'TestTermsWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
//import { spfi,SPFx } from "@pnp/sp/";
//import { graphfi } from '@pnp/graph';
//import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/taxonomy";
//import { ITermInfo } from "@pnp/sp/taxonomy";

let teamSelect : any={};

export interface ITestTermsWebPartProps {
  description: string;

  URL:string;
  tenantURL: any;
  divisionTitle:string;
  siteTitle: string;
  siteName: string;
  siteArray: string[];

  parentTermID: string;
  divisionDC: string;
  division: string;
  teamLabels: string[];
  teamTerm: string;
  teamTermID: string;
  teamName: string;
  termID: string;
  teamNamePrev : string;
  termFlag:boolean;
  teamDataChk: boolean;

  isDCPowerUser:any;
  isTeamPowerUser:any;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id : string;
  Title : string;
  DC_Team: string;
  TermGuid : string;
}

export default class TestTermsWebPart extends BaseClientSideWebPart<ITestTermsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _teamOptions: IPropertyPaneDropdownOption[];

  public async render(): Promise<void> {

    let libraryFlag : boolean = false;

    this.properties.termFlag = false;
    this._teamOptions = [];
    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.tenantURL = this.properties.URL.split('/',5);
    this.properties.siteTitle = this.context.pageContext.web.title;
    this.properties.siteArray = this.properties.siteTitle.split(" - ");
    this.properties.divisionTitle = this.properties.siteTitle.split(" - ")[0];
    this.properties.siteName = this.properties.siteTitle.split(" - ")[1];

    if(this.properties.division !== undefined){
      this.getTeamOptions();
      alert('team options completed');  
    }

    if(this.properties.teamTerm!==undefined){
      this.properties.teamName = this.properties.teamTerm.split(';')[0];
      this.properties.teamTermID = this.properties.teamTerm.split(';')[1];
      this.properties.termFlag = true;
    } 
  
//    if(this.properties.teamNamePrev === undefined){
//      this.properties.teamNamePrev = this.properties.teamName;
//    }


      //if(this.properties.teamTermID !== undefined){
      //} 

    this.domElement.innerHTML = `
    <section class="${styles.testTerms} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h5 class="text-black">DC Data ${this.properties.divisionDC}</h5>
        <h5 class="text-black">Division ${this.properties.division}</h5>
        <h5 class="text-black">Team ${this.properties.teamName}</h5>
        <h5 class="text-black">Term ID ${this.properties.teamTermID}</h5>
        <h5 class="text-black">Team Prev ${this.properties.teamNamePrev}</h5>
        <h5 class="text-black">TeamDataCHK ${this.properties.teamDataChk}</h5>
        <h5 class="text-black">TermFlag ${this.properties.termFlag}</h5>
        <h5 class="text-black">DC Power User ${this.properties.isDCPowerUser}</h5>
        <h5 class="text-black">Team Power User ${this.properties.isTeamPowerUser}</h5>
      </div>
    </section>`;

    if(this.properties.termFlag){
      libraryFlag = await this.checkDataAsync("policies",this.properties.teamName,"");
    }

    console.log("Render Function");
    console.log("division",this.properties.division);
    console.log("division termID",this.properties.parentTermID);
    console.log("teamName",this.properties.teamName);
    console.log("teamTermID",this.properties.teamTermID);
    console.log("teamName Prev",this.properties.teamNamePrev);
    console.log("teamDataCHK",this.properties.teamDataChk);
    console.log("libraryFlag",libraryFlag);
  }

  private async checkDataAsync(library:string,team:string,category:string):Promise<boolean> {
     
    let dcName : string = "";      
    let dataFlag : boolean = false;

    // *** revise this to check all DCs for Team and SharedWith.
    switch(this.properties.division){
      case "Assessments":
        dcName = "asm_dc";
        break;
      case "Central":
        dcName = "cen_dc";
        break;
      case "Connect":
        dcName = "cnn_dc";
        break;
      case "Employability":
        dcName = "emp_dc";
        break;
      case "Health":
        dcName = "hea_dc";
        break;
    }

    await this.checkData(dcName,library,team,category)
      .then((response: any) => {
        console.log("CheckData",response);
        for(let x=0;x<response.value.length;x++){
          const teamID = response.value[x].DC_Team.TermGuid;

          if(response.value.length>0 && teamID === this.properties.teamTermID){          
            dataFlag = true; 
          }            
        }
      });
      return dataFlag;
  }

  private async checkData(dcName:string,library:string,team:string,category:string):Promise<ISPLists> {
    
    let requestUrl = '';
    console.log("checkdata",dcName," ",team);

    //_api/web/lists/GetByTitle('policies')/items?$select=*,TaxCatchAll/ID,TaxCatchAll/Term&$expand=TaxCatchAll&$filter=TaxCatchAll/@odata.id eq '7578f741-bff1-4093-97fc-283a7f330ccb'

    if(category === ''){
        requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$select=*,TaxCatchAll/Term&$expand=TaxCatchAll/Term&$filter=TaxCatchAll/Term eq '${team}'&$top=10`;
    }else{
      requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$select=*,TaxCatchAll/Term&$expand=TaxCatchAll/Term&$filter=TaxCatchAll/Term eq '${category}'&$top=10`;
    }
    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response : SPHttpClientResponse) => {
        return response.json();
      });   
  }

  private async getTeamLabels():Promise<any>{
    const teamGroupID : string = "a66b7b2f-9f5d-4573-b763-542518574351";
    const teamSetID : string = "a9620950-3da0-4ab8-b191-f976f8b27852";

    switch(this.properties.division){
      case 'Assessments':
        this.properties.parentTermID = '3483090f-6789-4737-9e62-da5a85b0644c';
        break;
      case 'Central':
        this.properties.parentTermID = 'f6e49543-f1e4-4baf-8348-740e4ea33285';
        break;
      case 'Connect':
        this.properties.parentTermID = '41bd7d76-804c-4af8-962b-aa65bb01fae9';
        break;
      case 'Employability':
        this.properties.parentTermID = '99b913be-47a5-4c2b-8f98-632b6288bdc1';
        break;
      case 'Health':
        this.properties.parentTermID = 'c589da32-40f2-4087-b399-ffb692a2fb68';
        break;
    }
   
    const url: string = `https://${this.properties.tenantURL[2]}/_api/v2.1/termStore/groups/${teamGroupID}/sets/${teamSetID}/terms/${this.properties.parentTermID}/children?select=id,labels`;

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      });
  }

  private getTeamOptions(): void {
    try{
      if(this.properties.division !== undefined){
        
        this.getTeamLabels().then( (response) => {
          this._teamOptions = [];

          for(let x=0;x<response.value.length;x++){
            //console.log("teamOptions",response);

            const teamName = response.value[x].labels[0].name;
            const teamTermID = response.value[x].id;
    
            this._teamOptions.push(<IPropertyPaneDropdownOption>{
              text: teamName,
              key: teamName + ";" + teamTermID
            })
          }
          //this.onDispose();         
        }).catch(err => {
          console.log('getTeamOptions ERROR:', err.response.data);
        });
        this.properties.teamDataChk = true;        
      }
    }catch(err){
      Log.error('DocumentCentre', new Error('getTeamOptions Error message'), err);
    }
    return;
  }  
  
  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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

    //const page2Obj : IPropertyPaneGroup["groupFields"] = [];
    let divisionDropdown : any={};
    
    divisionDropdown = PropertyPaneDropdown('division',{
      label:"Please choose your Division",
      options:[
        { key : 'Assessments', text : 'Assessments'},
        { key : 'Central', text : 'Central'},
        { key : 'Connect', text : 'Connect'},
        { key : 'Employability', text : 'Employability'},
        { key : 'Health', text : 'Health'},
      ]
    });

    if(this.properties.division !== undefined){
      //this.getTeamOptions();

      if(this.properties.teamDataChk){
        teamSelect =  PropertyPaneDropdown('teamTerm', {
          label:'Please choose your team',
          options: this._teamOptions
        });
      }

    }

    console.log("Property Pane Config");
    console.log("division",this.properties.division);
    console.log("division termID",this.properties.parentTermID);
    console.log("teamName",this.properties.teamName);
    console.log("teamTermID",this.properties.teamTermID);
    console.log("teamName Prev",this.properties.teamNamePrev);
    console.log("teamDataCHK",this.properties.teamDataChk);

/*    
    page2Obj.push(
      PropertyPaneDropdown('teamTerm', {
        label:'Please choose Team',
        options: this._teamOptions
      }), 
    )    
*/

    if(this.properties.teamTerm!==undefined){
      this.properties.teamName = this.properties.teamTerm.split(';')[0];
      this.properties.teamTermID = this.properties.teamTerm.split(';')[1];
      this.properties.termFlag = true;
    } 

    this.onDispose();         

    return {
      pages: [
        {
          header: {
            description: "Page 1 - Document Centre Setup",
          },
          groups: [
            {
              groupName: "Division",
              groupFields: [
                PropertyPaneLabel('',{
                  text: "Document Centre Data"
                }),
                PropertyPaneLabel('',{
                  text: "This option is where the data is collected from. \n\n\n All Document Centres is the default and will check all DC libraries for the selected team files"
                }),
                PropertyPaneDropdown("divisionDC", {
                  label: "Please Choose Document Centre",
                  options: [
                    { key: "All", text: "All Document Centres" },
                    { key: "ASM", text: "Assessments DC Only" },
                    { key: "CEN", text: "Central DC Only" },
                    { key: "CNN", text: "Connect DC Only" },
                    { key: "EMP", text: "Employability DC Only" },
                    { key: "HEA", text: "Health DC Only" },
                  ],
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneToggle('isDCPowerUser', { 
                  key: 'DCPWToggle',
                  label: 'DC Power User?',
                  onText: 'Yes',
                  offText: 'No'                  
                })
              ]
            }
          ]
        },
        { //Page 2
          header: {
            description: "Page 2 - Team Selection"
          },
          groups: [
            {
              groupName: "Please Select Your Division & Team",
              groupFields: [
                PropertyPaneToggle('isTeamPowerUser', { 
                  key : 'TeamPWToggel',
                  label: 'Team Power User?',
                  onText: 'Yes',
                  offText: 'No'                  
                }),
                divisionDropdown,
                teamSelect
              ]
            }
          ]
        },        
      ]
    };
  }
}