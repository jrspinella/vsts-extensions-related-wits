import "../css/SettingsPanel.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Label } from "OfficeFabric/Label";
import { Dropdown } from "OfficeFabric/components/Dropdown/Dropdown";
import { IDropdownOption, IDropdownProps } from "OfficeFabric/components/Dropdown/Dropdown.Props";
import { TagPicker, ITag } from 'OfficeFabric/components/pickers/TagPicker/TagPicker';
import { IPickerItemProps } from 'OfficeFabric/components/pickers/PickerItem.Props';
import { TagItem } from 'OfficeFabric/components/pickers/TagPicker/TagItem';
import { autobind } from "OfficeFabric/Utilities";
import { Button, ButtonType } from "OfficeFabric/Button";
import { TextField } from "OfficeFabric/TextField";

import { Loading } from "./Loading";
import { UserPreferenceModel, Constants } from "./Models";
import { isInteger } from "./Helpers";
import { UserPreferences } from "./UserPreferences";

import { WorkItemField } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

interface ISettingsPanelProps {
    settings: UserPreferenceModel;
    onSave: (userPreferenceModel: UserPreferenceModel) => void
}

interface ISettingsPanelState {
    loading: boolean;
    sortField?: WorkItemField;
    queryFields?: WorkItemField[];
    sortableFields?: WorkItemField[];
    queryableFields?: WorkItemField[];   
    top: string; 
}

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
    constructor(props: ISettingsPanelProps, context?: any) {
        super(props, context);

        this.state = {
            loading: true,
            top: props.settings.top.toString()
        };        
    }

    public componentDidMount() {
        this.initialize();
    }

    public async initialize(): Promise<void> {        
        const workItemFormService = await WorkItemFormService.getService();
        const fields = await workItemFormService.getFields();

        const sortableFields = fields.filter(field => 
            Utils_Array.contains(Constants.SortableFieldTypes, field.type) 
            && !Utils_Array.contains(Constants.ExcludedFields, field.referenceName)).sort(this._fieldNameComparer);

        const queryableFields = fields.filter(field => 
            (Utils_Array.contains(Constants.QueryableFieldTypes, field.type) || Utils_String.equals(field.referenceName, "System.Tags", true))
            && !Utils_Array.contains(Constants.ExcludedFields, field.referenceName));

        const sortField = Utils_Array.first(sortableFields, field => Utils_String.equals(field.referenceName, this.props.settings.sortByField, true)) ||
                          Utils_Array.first(sortableFields, field => Utils_String.equals(field.referenceName, Constants.DEFAULT_SORT_BY_FIELD, true));

        let queryFields = this.props.settings.fields.map(fName => Utils_Array.first(queryableFields, field => Utils_String.equals(field.referenceName, fName, true)));
        queryFields = queryFields.filter(f => f != null);

        this.setState({...this.state, loading: false, sortableFields: sortableFields, queryableFields: queryableFields, sortField: sortField, queryFields: queryFields});
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return (
                <div className="settings-panel">
                    <Loading />
                </div>
            );            
        }
        
        let sortableFieldOptions: IDropdownOption[] = this.state.sortableFields.map((field: WorkItemField, index: number) => {
            return {
                key: field.referenceName,
                index: index,
                text: field.name,
                selected: Utils_String.equals(this.state.sortField.referenceName, field.referenceName, true)
            }
        });

        return (
            <div className="settings-panel">
                <div className="settings-controls">
                    <TextField label='Maximum number of workitems to retrieve' 
                        className="top"
                        required={true} 
                        value={`${this.state.top}`} 
                        onChanged={(newValue: string) => this._updateTop(newValue)} 
                        onGetErrorMessage={this._getTopError} />

                    <Dropdown label="Sort by field"                     
                        className="sort-field-dropdown"
                        onRenderList={this._onRenderCallout} 
                        options={sortableFieldOptions} 
                        onChanged={this._updateSortField} /> 
                    
                    <div className="tagpicker-container">
                        <Label>Look for the following fields</Label>
                        <TagPicker
                            className="tagpicker"
                            defaultSelectedItems={this.state.queryFields.map(f => this._getFieldTag(f))}
                            onResolveSuggestions={this._onFieldFilterChanged}
                            getTextFromItem={item => item.name}
                            onChange={this._updateQueryFields}
                            pickerSuggestionsProps={
                                {
                                    suggestionsHeaderText: 'Suggested Fields',
                                    noResultsFoundText: 'No fields Found'
                                }
                            }
                        />
                    </div>
                </div>

                <Button className="save-button" disabled={!this._isSettingsDirty() || !this._isSettingsValid()} buttonType={ButtonType.primary} onClick={this._onSaveClick}>
                    Save
                </Button>
            </div>
        );    
    }

    @autobind
    private _getTopError(value: string): string {
        if (value == null || value.trim() === "") {
            return "A value is required";
        }
        if (!isInteger(value)) {
            return "Enter a positive integer value";
        }
        if (parseInt(value) > 500) {
            return "For better performance, please enter a value less than 500"
        }
        return "";
    }

    private _isSettingsDirty(): boolean {
        return this.props.settings.top.toString() !== this.state.top
            || !Utils_String.equals(this.props.settings.sortByField, this.state.sortField.referenceName, true)
            || !Utils_Array.arrayEquals(this.props.settings.fields, this.state.queryFields, (item1: string, item2: WorkItemField) => Utils_String.equals(item1, item2.referenceName, true))
    }

    private _isSettingsValid(): boolean {
        return isInteger(this.state.top) && parseInt(this.state.top) > 0 && parseInt(this.state.top) <= 500;
    }

    @autobind
    private async _onSaveClick(): Promise<void> {
        if (!this._isSettingsValid()) {
            return;
        }

        let userPreferenceModel: UserPreferenceModel = {
            sortByField: this.state.sortField.referenceName,
            fields: this.state.queryFields.map(f => f.referenceName),
            top: parseInt(this.state.top)
        }

        const workItemFormService = await WorkItemFormService.getService();
        const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;

        await UserPreferences.writeUserSetting(workItemType, userPreferenceModel);
        this.props.onSave(userPreferenceModel);
    }
        
    @autobind
    private _getFieldTag(field: WorkItemField): ITag {
        return {
            key: field.referenceName,
            name: field.name
        }
    }

    @autobind
    private _onRenderCallout(props?: IDropdownProps, defaultRender?: (props?: IDropdownProps) => JSX.Element): JSX.Element {
        return (
            <div className="callout-container">
                {defaultRender(props)}
            </div>
        );
    }

    @autobind
    private _updateSortField(option: IDropdownOption) {
        const sortField = Utils_Array.first(this.state.sortableFields, (field: WorkItemField) => Utils_String.equals(field.referenceName, option.key as string, true));
        this.setState({...this.state, sortField: sortField});
    }

    @autobind
    private _updateQueryFields(items: ITag[]) {
        const queryFields = items.map((item: ITag) => Utils_Array.first(this.state.queryableFields, (field: WorkItemField) => Utils_String.equals(field.referenceName, item.key, true)));
        this.setState({...this.state, queryFields: queryFields});
    }

    @autobind
    private _updateTop(top: string) {
        this.setState({...this.state, top: top});      
    }

    @autobind
    private _onFieldFilterChanged(filterText: string, tagList: ITag[]): ITag[] {
        return filterText
            ? this.state.queryableFields.filter(field => field.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 
                && Utils_Array.findIndex(tagList, (tag: ITag) => Utils_String.equals(tag.key, field.referenceName, true)) === -1).map(field => {
                    return { key: field.referenceName, name: field.name};
                }) 
            : [];
    }

    private _fieldNameComparer(a: WorkItemField, b: WorkItemField): number {
        let aUpper = a.name.toUpperCase();
        let bUpper = b.name.toUpperCase();

        if (aUpper < bUpper) { return -1 }
        if (aUpper > bUpper) { return 1 }
        return 0;
    }
}