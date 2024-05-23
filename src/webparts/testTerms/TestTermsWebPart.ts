import { Version } from '@microsoft/sp-core-library';
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
  PropertyPaneCheckbox,
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
  libraryChk: boolean;
  teamLabels: string[];
  teamTerm: string;
  teamTermID: string;
  teamName: string;
  termID: string;
  teamNamePrev : string;
  termFlag:boolean;
  teamDataChk: boolean;
  teamSelect : any;

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

    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.siteTitle = this.context.pageContext.web.title;
    this.properties.tenantURL = this.properties.URL.split('/',5);
    this.properties.termFlag = false;
    this.properties.libraryChk = false;
    this._teamOptions = [];
    this.properties.teamSelect = {};

    if(this.properties.division !== undefined){
      await this.getTeamOptions().then(()=>{
        console.log("render getTeamOptions");
        console.log("teamOptions length",this._teamOptions.length);  
      });
    }

    if(this.properties.teamTerm!==undefined){
      this.properties.teamName = this.properties.teamTerm.split(';')[0];
      this.properties.teamTermID = this.properties.teamTerm.split(';')[1];
      this.properties.termFlag = true;
    } 
  
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
      alert("term flag="+this.properties.termFlag);
    }

    this.properties.teamNamePrev = this.properties.teamName;

    //console.log("***********************************************");
    //console.log("Render Function");
    //console.log("division",this.properties.division);
    //console.log("division termID",this.properties.parentTermID);
    //console.log("teamName",this.properties.teamName);
    //console.log("teamTermID",this.properties.teamTermID);
    //console.log("teamName Prev",this.properties.teamNamePrev);
    //console.log("teamDataCHK",this.properties.teamDataChk);
  }

  private getTeamLabels():Promise<any>{
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

  private async getTeamOptions(): Promise<void> {
    try{
      if(this.properties.division !== undefined){
        
        this.getTeamLabels().then( (response) => {
          console.log("response",response);

          for(let x=0;x<response.value.length;x++){
            const teamName = response.value[x].labels[0].name;
            const teamTermID = response.value[x].id;
            this._teamOptions.push(<IPropertyPaneDropdownOption>{
              text: teamName,
              key: teamName + ";" + teamTermID
            })
          }
          if(this._teamOptions.length>0){
            //this.properties.teamDataChk = true;
            //if(this.properties.teamDataChk){
              console.log("getTeamOptions length",this._teamOptions.length);  
              this.properties.teamSelect = PropertyPaneDropdown('teamTerm', {
                label:'Please choose your team',
                options: this._teamOptions
              });
            //}
          }

        }).catch(err => {
          console.log('getTeamOptions ERROR:', err);
        });
        console.log("TeamSelect",this.properties.teamSelect);  
      }
    }catch(err){
      console.log('getTeamOptions Error message', err);
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

    let divisionDropdown : any={};
    let libraryDropdown : any={};
    //let teamSelect : any={};
    
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

    if(this.properties.libraryChk){
      libraryDropdown = PropertyPaneDropdown('library',{
        label:"Please choose a library",
        options:[
          { key : 'Policies', text : 'Policies'},
          { key : 'Procedures', text : 'Procedures'},
          { key : 'Guides', text : 'Guides'},
          { key : 'Forms', text : 'Forms'},
          { key : 'General', text : 'General'},
          { key : 'Management', text : 'Management'},
          { key : 'Custom', text : 'Custom'},
        ]
      });
    }

    console.log("***********************************************");
    console.log("Property Pane Config");
    console.log("division",this.properties.division);
    console.log("division termID",this.properties.parentTermID);
    console.log("teamName",this.properties.teamName);
    console.log("teamTermID",this.properties.teamTermID);
    console.log("teamName Prev",this.properties.teamNamePrev);
    console.log("teamDataCHK",this.properties.teamDataChk);

    //if(this.properties.teamDataChk){
    //  teamSelect = PropertyPaneDropdown('teamTerm', {
    //    label:'Please choose your team',
    //    options: this._teamOptions
    //  });
    //}

    if(this.properties.teamTerm!==undefined){
      this.properties.teamName = this.properties.teamTerm.split(';')[0];
      this.properties.teamTermID = this.properties.teamTerm.split(';')[1];
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
                }),
                PropertyPaneCheckbox('libraryChk', { 
                  text: 'Specify which library to get data for?'
                }),
                libraryDropdown
              ]
            }
          ]
        },
        { //Page 2
          header: {
            description: "Page 2 - Division & Team Selection"
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
                this.properties.teamSelect,
              ]
            }
          ]
        },        
      ]
    };
  }
}