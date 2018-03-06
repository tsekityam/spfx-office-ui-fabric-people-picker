import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, PrimaryButton, Toggle, NormalPeoplePicker, IPersonaProps, autobind, IDropdownOption, ValidationState, assign, IBasePicker, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { people, mru } from './PeoplePickerExampleData';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';

export interface IPeoplePickerExampleState {
  currentPicker?: number | string;
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems?: IPersonaProps[];
}

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  mostRecentlyUsedHeaderText: 'Suggested Contacts',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'People Picker Suggestions available',
  suggestionsContainerAriaLabel: 'Suggested contacts'
};

export default class HelloWorld extends React.Component<IHelloWorldProps, IPeoplePickerExampleState> {
  private _picker: IBasePicker<IPersonaProps>;

  constructor(props: any) {
    super(props);
    const peopleList: IPersonaWithMenu[] = [];
    people.forEach((persona: IPersonaProps) => {
      const target: IPersonaWithMenu = {};

      assign(target, persona);
      peopleList.push(target);
    });

    this.state = {
      currentPicker: 1,
      peopleList: peopleList,
      mostRecentlyUsed: mru,
      currentSelectedItems: []
    };
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    let currentPicker = this._renderNormalPicker();

    return (
      <div>
        {currentPicker}
      </div>
    );
  }

  private _renderNormalPicker() {
    return (
      <NormalPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        getTextFromItem={ this._getTextFromItem }
        pickerSuggestionsProps={ suggestionProps }
        className={ 'ms-PeoplePicker' }
        key={ 'normal' }
      />
    );
  }

  private _getTextFromItem(persona: IPersonaProps): string {
    return persona.primaryText as string;
  }

  @autobind
  private _onItemsChange(items: any[]) {
    this.setState({
      currentSelectedItems: items
    });
  }

  @autobind
  private _onSetFocusButtonClicked() {
    if (this._picker) {
      this._picker.focusInput();
    }
  }

  @autobind
  private _renderFooterText(): JSX.Element {
    return <div>No additional results</div>;
  }

  @autobind
  private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
    if (filterText) {
          //return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
          if (Environment.type === EnvironmentType.Local) {
            // If the running environment is local, load the data from the mock
      let filteredPersonas: IPersonaProps[] = this._filterPersonasByText(filterText);

      filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
      filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
      return this._filterPromise(filteredPersonas);
    } else if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {
      const userRequestUrl: string = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
      let principalType: number = 0;
        principalType += 1;
      const data = {
        'queryParams': {
          'AllowEmailAddresses': true,
          'AllowMultipleEntities': false,
          'AllUrlZones': false,
          'MaximumEntitySuggestions': 10,
          'PrincipalSource': 15,
          // PrincipalType controls the type of entities that are returned in the results.
          // Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
          // These values can be combined (example: 13 is security + SP groups + users)
          'PrincipalType': principalType,
          'QueryString': filterText
        }
      };

      return new Promise<IPersonaProps[]>((resolve, reject) =>
        this.props.spHttpClient.post(userRequestUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json',
              "content-type": "application/json"
            },
            body: JSON.stringify(data)
          })
          .then((response: SPHttpClientResponse) => {
            return response.json();
          })
          .then((response: any): void => {
            let relevantResults: any = JSON.parse(response.value);
            let resultCount: number = relevantResults.length;
            let people = [];
            let persona: IPersonaProps = {};
            if (resultCount > 0) {
              for (var index = 0; index < resultCount; index++) {
                var p = relevantResults[index];
                let account = p.Key.substr(p.Key.lastIndexOf('|') + 1);

                persona.primaryText = p.DisplayText;
                persona.imageUrl = `/_layouts/15/userphoto.aspx?size=S&accountname=${account}`;
                persona.imageShouldFadeIn = true;
                persona.secondaryText = p.EntityData.Title;
                people.push(persona);
              }
            }
            resolve(people);
          }, (error: any): void => {
            // reject(this._peopleList = []);
          })
        );
};
    } else {
      return [];
    }
  }

 private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    return this._convertResultsToPromise(personasToReturn);
  }

  private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
    if (!personas || !personas.length || personas.length === 0) {
      return false;
    }
    return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
  }

  private _filterPersonasByText(filterText: string): IPersonaProps[] {
    return this.state.peopleList.filter(item => this._doesTextStartWith(item.primaryText as string, filterText));
  }

  private _doesTextStartWith(text: string, filterText: string): boolean {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  }

  private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
  }

  private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
    return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
  }

  /**
   * Takes in the picker input and modifies it in whichever way
   * the caller wants, i.e. parsing entries copied from Outlook (sample
   * input: "Aaron Reid <aaron>").
   *
   * @param input The text entered into the picker.
   */
  private _onInputChange(input: string): string {
    const outlookRegEx = /<.*>/g;
    const emailAddress = outlookRegEx.exec(input);

    if (emailAddress && emailAddress[0]) {
      return emailAddress[0].substring(1, emailAddress[0].length - 1);
    }

    return input;
  }
}
