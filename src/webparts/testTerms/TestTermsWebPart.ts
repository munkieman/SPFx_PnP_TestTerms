import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  //IPropertyPaneGroup,
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

export interface ITestTermsWebPartProps {
  description: string;
  teamName: string;
  division: string;
  termID: string;
  URL:string;
  tenantURL: any;
  siteTitle: string;
  siteArray: string[];
  teamLabels: string[];
}

export default class TestTermsWebPart extends BaseClientSideWebPart<ITestTermsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _teamOptions: IPropertyPaneDropdownOption[];

  public async render(): Promise<void> {

    this._teamOptions = [];

    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.tenantURL = this.properties.URL.split('/',5);
    this.properties.siteTitle = this.context.pageContext.web.title;
    this.properties.siteArray = this.properties.siteTitle.split(" - ");
    //this.properties.division = this.properties.siteArray[0];

    console.log("Render division",this.properties.division);

    this.domElement.innerHTML = `
    <section class="${styles.testTerms} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
        The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
    </section>`;

    this.getTerms();
  }

  public async getTerms():Promise<void>{

    if(this.properties.division !== undefined){
      //*** search for a term by label
      //const results = await sp.termStore.searchTerm({
      //  label: "ASM Team A",
      //  setId: setID,
      //  parentTermId : this.properties.termID
      //});
      //console.log(results);

      //*** list all the terms that are direct children of a set
      //const setChildren: ITermInfo[] = await sp.termStore.groups.getById(groupID).sets.getById(setID).getAllChildrenAsOrderedTree();
      //console.log("set children",setChildren);

      //*** list all the terms that are direct children of a given term
      //const termChildren: ITermInfo[] = await sp.termStore.sets.getById(setID).getTermById(this.properties.termID).children();    
      //console.log("terms children",termChildren);
    }

    await this.getTeamLabels()
    .then((response) => {
      console.log("render response",response);
      for(let x=0;x<response.value.length;x++){
        const teamName = response.value[x].labels[0].name;

        //console.log(response.value[x].labels[0].name); 
        //this.properties.teamLabels[x]=response.value[x].labels[0].name;
        this._teamOptions.push(<IPropertyPaneDropdownOption>{
          text: teamName,
          key: teamName 
        })
      }
      console.log(this._teamOptions);
    }); 
  }

  private async getTeamLabels():Promise<any>{
    const groupID : string = "a66b7b2f-9f5d-4573-b763-542518574351";
    const setID : string = "f6c88c73-1bc1-4019-973f-b034ea41e08a";
    //const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  

    const url: string = `https://${this.properties.tenantURL[2]}/_api/v2.1/termStore/groups/${groupID}/sets/${setID}/terms/${this.properties.termID}/children?select=id,labels`;
    console.log(url);

    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        }
      });
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
    let teamDropdown : any = PropertyPaneDropdown('teamName',{
      label:"Please Choose a Team",
      options:[{key : "", text: "test"}]});

    switch (this.properties.division) {
      case "assessments":
        this.properties.termID = "2e21f62b-594b-4a88-aa9f-a1b6aa7e1f62"; 
        break;
      case "central":
        this.properties.termID = "11ae0cc5-d395-4176-81a9-22f57f785afd";  
        break;
      case "connect":
        this.properties.termID = "f414e7f0-4e65-4754-a030-5ac7be12180f";  
        break;
      case "employability":
        this.properties.termID = "d4476663-0780-42da-b6a8-7ef92846f9f4";  
        break;
      case "health":
        this.properties.termID = "7c2683bf-64e6-48e2-9ab6-8021be871cb1";  
        break;
    }
    
    //page2Obj.push(PropertyPaneDropdown('team', {
    //    label:'Please choose Team',
    //    options: this._teamOptions
    //  }), 
    //)

    console.log("division",this.properties.division);
    console.log("termID",this.properties.termID);

    teamDropdown = PropertyPaneDropdown('teamName',{
      label:"Please Choose a Team",
      options:this._teamOptions
    }); 

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
                }),
                PropertyPaneDropdown('division',{
                  label:"Please Choose Division",
                  options:[
                    { key : 'assessments', text : 'Assessments'},
                    { key : 'central', text : 'Central'},
                    { key : 'connect', text : 'Connect'},
                    { key : 'employability', text : 'Employability'},
                    { key : 'health', text : 'Health'},
                  ]
                }), 
                teamDropdown               
              ]
            }
          ]
        },
      ]
    };
  }
}