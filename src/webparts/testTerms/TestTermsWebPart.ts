import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TestTermsWebPart.module.scss';
import * as strings from 'TestTermsWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import { spfi,SPFx } from "@pnp/sp/";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/taxonomy";
import { ITermInfo } from "@pnp/sp/taxonomy";

export interface ITestTermsWebPartProps {
  description: string;
  teamName: string;
  division: string;
  termID: string;
}

export default class TestTermsWebPart extends BaseClientSideWebPart<ITestTermsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _teamOptions: IPropertyPaneDropdownOption[];

  public async render(): Promise<void> {

    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    const results = await sp.termStore.searchTerm({
      label: "ASM Team A",
      setId: "f6c88c73-1bc1-4019-973f-b034ea41e08a",
      parentTermId : "2e21f62b-594b-4a88-aa9f-a1b6aa7e1f62"
    });

    console.log(results);

    // list all the terms that are direct children of this set
    //const infos: ITermInfo[] = await sp.termStore.groups.getById("ad680eae-a3ec-4b8e-86b0-e2d2d64808a1").sets.getById("f6c88c73-1bc1-4019-973f-b034ea41e08a").children();
    //console.log("infos",infos);

    // list all the terms available in this term set by term set id
    const TermSet: ITermInfo[] = await sp.termStore.sets.getById("f6c88c73-1bc1-4019-973f-b034ea41e08a").getTermById("2e21f62b-594b-4a88-aa9f-a1b6aa7e1f62").children();    
    for(let x=0;x<=TermSet.length;x++){
      console.log("termset",TermSet[x].labels[0].name);
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
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");
    
    this._teamOptions = [];

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

    const page2Obj : IPropertyPaneGroup["groupFields"] = [];
    //let termID : string;

    switch (this.properties.division) {
      case "Assessments":
        this.properties.termID = "2e21f62b-594b-4a88-aa9f-a1b6aa7e1f62";  
        break;
      case "Central":
        this.properties.termID = "11ae0cc5-d395-4176-81a9-22f57f785afd";  
        break;
      case "Connect":
        this.properties.termID = "f414e7f0-4e65-4754-a030-5ac7be12180f";  
        break;
      case "Employability":
        this.properties.termID = "d4476663-0780-42da-b6a8-7ef92846f9f4";  
        break;
      case "Health":
        this.properties.termID = "7c2683bf-64e6-48e2-9ab6-8021be871cb1";  
        break;
    }

    page2Obj.push(PropertyPaneDropdown('team', {
        label:'Please choose Team',
        options: this._teamOptions
      }), 
    )

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
              groupName: "Sections",
              groupFields: page2Obj
            }
          ]
        }
      ]
    };
  }
}
