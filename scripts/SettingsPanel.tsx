import "../css/SettingsPanel.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Label } from "OfficeFabric/Label";
import { Dropdown } from "OfficeFabric/components/Dropdown/Dropdown";
import { IDropdownOption, IDropdownProps } from "OfficeFabric/components/Dropdown/Dropdown.Props";
import { TagPicker, ITag } from 'OfficeFabric/components/pickers/TagPicker/TagPicker';
import { autobind } from "OfficeFabric/Utilities";

import { Loading } from "./Loading";
import { UserPreferenceModel, Constants } from "./Models";

import { WorkItemField } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

interface ISettingsPanelProps {
    settings: UserPreferenceModel;
}

interface ISettingsPanelState {
    loading: boolean;
    sortField?: WorkItemField;
    queryFields?: WorkItemField[];
    sortableFields?: WorkItemField[];
    queryableFields?: WorkItemField[];    
}

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
    constructor(props: ISettingsPanelProps, context?: any) {
        super(props, context);

        this.state = {
            loading: true
        };

        this.initialize();     
    }

    public async initialize(): Promise<void> {
        const workItemFormService = await WorkItemFormService.getService();
        const fields = await workItemFormService.getFields();

        const sortableFields = fields.filter((field: WorkItemField) => 
            Utils_Array.contains(Constants.SortableFieldTypes, field.type) && !Utils_Array.contains(Constants.ExcludedFields, field.referenceName)).sort(this._fieldNameComparer);

        const queryableFields = fields.filter((field: WorkItemField) => 
            (Utils_Array.contains(Constants.QueryableFieldTypes, field.type) || Utils_String.equals(field.referenceName, "System.Tags", true))
            && !Utils_Array.contains(Constants.ExcludedFields, field.referenceName));

        const sortField = Utils_Array.first(sortableFields, (field: WorkItemField) => Utils_String.equals(field.referenceName, this.props.settings.sortByField, true));
        const queryFields = this.props.settings.fields.map((fName: string) => Utils_Array.first(queryableFields, (field: WorkItemField) => Utils_String.equals(field.referenceName, fName, true)));

        this.setState({...this.state, loading: true, sortableFields: sortableFields, queryableFields: queryableFields, sortField: sortField, queryFields: queryFields});
    }

    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
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
                <Dropdown label="Sort by field" 
                    onRenderList={this._onRenderCallout} 
                    required={true} 
                    options={sortableFieldOptions} 
                    onChanged={this._updateSortField} /> 

                <TagPicker
                    defaultSelectedItems={model.manualFields.map(f => this._getFieldTag(f))}
                    onResolveSuggestions={this._onFieldFilterChanged}
                    getTextFromItem={item => item.name}
                    onChange={items => this._item.updateManualFields(items.map(item => item.key))}
                    pickerSuggestionsProps={
                        {
                            suggestionsHeaderText: 'Suggested Fields',
                            noResultsFoundText: 'No fields Found'
                        }
                    }
                />
            </div>
        );    
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

    private _fieldNameComparer(a: WorkItemField, b: WorkItemField): number {
        let aUpper = a.name.toUpperCase();
        let bUpper = b.name.toUpperCase();

        if (aUpper < bUpper) { return -1 }
        if (aUpper > bUpper) { return 1 }
        return 0;
    }

}