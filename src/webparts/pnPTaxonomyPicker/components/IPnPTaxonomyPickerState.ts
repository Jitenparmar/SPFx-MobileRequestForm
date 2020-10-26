import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";

export interface IPnPTaxonomyPickerState {
    selectedTerms: IPickerTerms;
    addUsers: number[];
    MobileBusinessJustification:string;
    MobileComments:string;
    MobileCostEstimate:string;
    Company:string;
    MobileCostCurrency:string;
    CompleteBy:any|null;
    termnCond:boolean;
    message:string;
    showMsg:boolean;
}
