import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape, update } from '@microsoft/sp-lodash-subset';

import styles from './HelloPropertyPaneWebPart.module.scss';
import * as strings from 'HelloPropertyPaneWebPartStrings';
import {
  PropertyPaneContinentSelector,
  IPropertyPaneContinentSelectorProps
} from '../../controls/PropertyPaneContinentSelector';
import {
  IPropertyFieldGroupOrPerson,
  // PropertyFieldPeoplePicker,
  // PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
  people: IPropertyFieldGroupOrPerson[];
  expansionOptions: any[]; // eslint-disable-line @typescript-eslint/no-explicit-any
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<IHelloPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloPropertyPane} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <div>Continent where I reside: <strong>${escape(this.properties.myContinent)}</strong></div>
        <div>Number of continents I've visited: <strong>${this.properties.numContinentsVisited}</strong></div>
      </div>
      <div class="selectedPeople"></div>
      <div class="expansionOptions"></div>
    </section>`;

    if (this.properties.people && this.properties.people.length > 0) {
      let peopleList: string = '';
      this.properties.people.forEach((person) => {
        peopleList = peopleList + `<li>${ person.fullName } (${ person.email })</li>`;
      });
    
      this.domElement.getElementsByClassName('selectedPeople')[0].innerHTML = `<ul>${ peopleList }</ul>`;
    }

    if (this.properties.expansionOptions && this.properties.expansionOptions.length > 0) {
      let expansionOptions: string  = '';
      this.properties.expansionOptions.forEach((option) => {
        expansionOptions = expansionOptions + `<li>${ option.Region }: ${ option.Comment } </li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${ expansionOptions }</ul>`;
      }
    }
  }

  protected onInit(): Promise<void> {
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
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
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
                // PropertyPaneTextField('myContinent', {
                //   label: 'Continent where I currently reside',
                //   onGetErrorMessage: this.validateContinents.bind(this)
                // }),
                PropertyPaneDropdown('myContinent', {
                  label: 'Continent where I currently reside',
                  options: [
                    {
                      key: 'Africa',
                      text: 'Africa'
                    },
                    {
                      key: 'Antarctica',
                      text: 'Antarctica'
                    },
                    {
                      key: 'Asia',
                      text: 'Asia'
                    },
                    {
                      key: 'Australia',
                      text: 'Austraila'
                    },
                    {
                      key: 'Europe',
                      text: 'Europe'
                    },
                    {
                      key: 'North America',
                      text: 'North America'
                    },
                    {
                      key: 'South America',
                      text: 'South America'
                    }
                  ],
                  selectedKey: 'Europe'
                }),
                new PropertyPaneContinentSelector('myContinent', <IPropertyPaneContinentSelectorProps>{
                  label: 'Continent where I currently reside',
                  disabled: false,
                  selectedKey: this.properties.myContinent,
                  onPropertyChange: this.onContinentSelectionChange.bind(this),
                }),
                PropertyPaneSlider('numContinentsVisited', {
                  label: 'Number of continents I\'ve visited',
                  min: 1,
                  max: 7,
                  showValue: true
                }),
                // PropertyFieldPeoplePicker('people', {
                //   label: 'Property Pane Field People Picker PnP Reusable Control',
                //   initialData: this.properties.people,
                //   allowDuplicate: false,
                //   principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   context: this.context as any, // eslint-disable-line @typescript-eslint/no-explicit-any
                //   properties: this.properties,
                //   onGetErrorMessage: null,
                //   deferredValidationTime: 0,
                //   key: 'peopleFieldId'
                // })
                PropertyFieldCollectionData('expansionOptions', {
                  key: 'collectionData',
                  label: 'Possible expansion options',
                  panelHeader: 'Possible expansion options',
                  manageBtnLabel: 'Manage expansion options',
                  value: this.properties.expansionOptions,
                  fields: [
                    {
                      id: 'Region',
                      title: 'Region',
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'Northeast', text: 'Northeast' },
                        { key: 'Northwest', text: 'Northwest' },
                        { key: 'Southeast', text: 'Southeast' },
                        { key: 'Southwest', text: 'Southwest' },
                        { key: 'North', text: 'North' },
                        { key: 'South', text: 'South' }
                      ]
                    },
                    {
                      id: 'Comment',
                      title: 'Comment',
                      type: CustomCollectionFieldType.string
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // private validateContinents(textboxValue: string): string {
  //   const validContinentOptions: string[] = ['africa', 'antarctica', 'asia', 'australia', 'europe', 'north america', 'south america'];
  //   const inputToValidate: string = textboxValue.toLowerCase();

  //   return (validContinentOptions.indexOf(inputToValidate) === -1)
  //     ? 'Invalid continent entry; valid options are "Africa", "Antarctica", "Asia", "Australia", "Europe", "North America", and "South America"'
  //     : '';
  // }

  /* eslint-disable @typescript-eslint/no-explicit-any */
  private onContinentSelectionChange(propertyPath: string, newValue: any): void {
    update(this.properties, propertyPath, (): any => {return newValue});
    this.render();
  }
  /* eslint-enable @typescript-eslint/no-explicit-any */
}
