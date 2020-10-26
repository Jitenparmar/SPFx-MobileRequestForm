import * as React from 'react';
import styles from './PnPTaxonomyPicker.module.scss';
import { IPnPTaxonomyPickerProps } from './IPnPTaxonomyPickerProps';
import { IPnPTaxonomyPickerState } from './IPnPTaxonomyPickerState';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { escape } from '@microsoft/sp-lodash-subset';

// Extra Fields
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings, mergeStyleSets } from 'office-ui-fabric-react';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
// @pnp/sp imports  
import { sp, Web } from '@pnp/sp';
import { getGUID } from "@pnp/common";
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import Select from 'react-select';
// Import button component    
import { IButtonProps, DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { autobind } from 'office-ui-fabric-react';


const companyoptions: IDropdownOption[] = [
  { key: 'PITL', text: 'PITL' },
  { key: 'AITL', text: 'AITL' },
  { key: 'GTPL', text: 'GTPL' }
];

const currencyoptions: IDropdownOption[] = [
  { key: 'USD', text: 'USD' },
  { key: 'INR', text: 'INR' },
  { key: 'EUR', text: 'EUR' }
];

const controlClass = mergeStyleSets({
  control: {
    margin: '0 0 15px 0',
    maxWidth: '300px',
  },
});

// Used to add spacing between example checkboxes
const stackTokens = { childrenGap: 10 };

const DayPickerStrings: IDatePickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ],

  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],

  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  closeButtonAriaLabel: 'Close date picker'
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

export default class PnPTaxonomyPicker extends React.Component<IPnPTaxonomyPickerProps, IPnPTaxonomyPickerState> {

  constructor(props: IPnPTaxonomyPickerProps, state: IPnPTaxonomyPickerState) {
    super(props);

    this.state = {
      selectedTerms: [],
      addUsers: [],
      MobileBusinessJustification: '',
      MobileComments: '',
      MobileCostEstimate: '',
      Company: '',
      MobileCostCurrency: '',
      CompleteBy: null,
      termnCond: false,
      message: '',
      showMsg: false
    };
  }

  public render(): React.ReactElement<IPnPTaxonomyPickerProps> {



    return (
      <div className={styles.pnPTaxonomyPicker}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PeoplePicker
                context={this.props.context}
                titleText="Raised By"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                isRequired={true}
                disabled={false}
                ensureUser={true}
                selectedItems={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />

              <Dropdown
                placeholder="Select a Company"
                label="Company"

                options={companyoptions}
                styles={dropdownStyles}
                onChange={(event, value) => this.setState({ Company: value.text })}
              />

              <TaxonomyPicker allowMultipleSelections={false}
                termsetNameOrID="8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f"
                panelTitle="Department or Team"
                label="Department or Team"
                context={this.props.context}
                onChange={this.onTaxPickerChange}
                isTermSetSelectable={false} />

              <DatePicker
                className={controlClass.control}
                strings={DayPickerStrings}
                label="Complete By"
                placeholder="Select a date..."
                ariaLabel="Select a date"
                onSelectDate={this._onSelectDate}
                value={this.state.CompleteBy}
                formatDate={this._onFormatDate}
              />

              <TextField
                label="Mobile Cost Estimate" value={this.state.MobileCostEstimate}
                onChange={(event, value) => this.setState({ MobileCostEstimate: value })}
              />

              <Dropdown
                placeholder="Select an option"
                label="Mobile Cost Currency"
                options={currencyoptions}
                styles={dropdownStyles}
                onChange={(event, value) => this.setState({ MobileCostCurrency: value.text })}
              />
{/* 
              <Select id="SelectUserRole"
                value={this.state.MobileCostCurrency}
                onChange={(event, value) => this.setState({ MobileCostCurrency: value.text })}
                isMulti
                isSearchable
                options={currencyoptions}
              /> */}

              <TextField
                label="Mobile Business Justification" value={this.state.MobileBusinessJustification}
                onChange={(event, value) => this.setState({ MobileBusinessJustification: value })}
                multiline>
              </TextField>

              <TextField
                label="Mobile Comments" value={this.state.MobileComments}
                onChange={(event, value) => this.setState({ MobileComments: value })}
                multiline>
              </TextField>

              <Stack className={styles.row}>
                <Checkbox label="Agree Mobile Policy (Read)" onChange={this._onCheckboxChange.bind(this)} />
              </Stack>

              <div className={styles.row}>
                <label>{this.state.message}</label>
              </div>

              <PrimaryButton style={{ marginTop: 10 }}
                data-automation-id="addSelectedTerms"
                title="Submit Request"
                onClick={this.addSelectedTerms}>
                Submit Request
              </PrimaryButton>

            </div>
          </div>
        </div>
      </div>
    );
  }

  @autobind
  private onTaxPickerChange(terms: IPickerTerms) {
    console.log("Terms", terms);
    this.setState({ selectedTerms: terms });
  }

  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ CompleteBy: date });
    console.log(this.state.CompleteBy + " " + date);
  }

  private _onFormatDate = (date: Date): string => {
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear();
  }

  @autobind
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({ termnCond: (isChecked) ? true : false });
  }

  @autobind
  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);

    let selectedUsers = [];
    for (let item in items) {
      selectedUsers.push(items[item].id);
    }

    this.setState({ addUsers: selectedUsers });
  }

  @autobind
  private addSelectedTerms(): void {
    // Update single value managed metadata field, with first selected term

    sp.web.lists.getByTitle("MobileRequestApproval").items.add({
      Title: getGUID(),
      Cost: this.state.MobileCostEstimate,
      Comments: this.state.MobileComments,
      Justification: this.state.MobileBusinessJustification,
      Company: this.state.Company,
      CurrencyType: this.state.MobileCostCurrency,
      CompleteBy: this.state.CompleteBy,
      Policy: this.state.termnCond,
      Department: {
        __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
        Label: this.state.selectedTerms[0].name,
        TermGuid: this.state.selectedTerms[0].key,
        WssId: -1
      },
      RaisedById: {
        results: this.state.addUsers
      }
    }).then(i => {
      console.log(i);
      this.setState({ message: 'New Mobile Request Submitted successfully!' });
      this.setState({
        selectedTerms: [],
        addUsers: [],
        MobileBusinessJustification: '',
        MobileComments: '',
        MobileCostEstimate: '',
        Company: '',
        MobileCostCurrency: '',
        CompleteBy: null,
        termnCond: false,
      });
    });

    /*
        // Update multi value managed metadata field
        const spfxList = sp.web.lists.getByTitle('SPFx Users');  
    
        // If the name of your taxonomy field is SomeMultiValueTaxonomyField, the name of your note field will be SomeMultiValueTaxonomyField_0
        const multiTermNoteFieldName = 'Terms_0';
    
        let termsString: string = '';
        this.state.selectedTerms.forEach(term => {
          termsString += `-1;#${term.name}|${term.key};#`;
        });
    
        spfxList.getListItemEntityTypeFullName()
          .then((entityTypeFullName) => {
            spfxList.fields.getByTitle(multiTermNoteFieldName).get()
              .then((taxNoteField) => {
                const multiTermNoteField = taxNoteField.InternalName;
                const updateObject = {
                  Title: 'Item title', 
                };
                updateObject[multiTermNoteField] = termsString;
    
                spfxList.items.add(updateObject, entityTypeFullName)
                  .then((updateResult) => {
                      console.dir(updateResult);
                  })
                  .catch((updateError) => {
                      console.dir(updateError);
                  });
              });
          });
          */
  }
}
