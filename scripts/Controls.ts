import VSS_Utils_Core = require("VSS/Utils/Core");
import Q = require("q");
import {Control, BaseControl} from "VSS/Controls";
import Utils_String = require("VSS/Utils/String");
import Utils_UI = require("VSS/Utils/UI");
import Utils_Array = require("VSS/Utils/Array");
import {StatusIndicator} from "VSS/Controls/StatusIndicator";
import {Combo} from "VSS/Controls/Combos";
import * as WitContracts from "TFS/WorkItemTracking/Contracts";
import {RelatedWitsControlOptions, RelatedFieldsControlOptions, Strings, IdentityReference, Constants} from "scripts/Models";
import {WorkItemTypeColorHelper, StateColorHelper, IdentityHelper, fieldNameComparer} from "scripts/Helpers";

export class RelatedWitsControl extends Control<RelatedWitsControlOptions> {   
    private _container: JQuery;
    private _workItems: WitContracts.WorkItem[];
    private _statusIndicator: StatusIndicator;
    
    constructor(options?: RelatedWitsControlOptions) {
        super(options);
    }

    public initialize(): void {
        super.initialize();

        this._container = $("<div>").addClass("container").appendTo(this.getElement());
        this.refresh(this._options.workItems || []);
    }

    public refresh(workItems: WitContracts.WorkItem[]): void {
        this._container.empty();
        this._workItems = workItems;

        if (this._workItems && this._workItems.length > 0) {
            $.each(this._workItems, (i: number, workItem: WitContracts.WorkItem) => {
                this._renderWorkItem(workItem);
            });   
        }
        else {
            $("<h3/>").addClass("no-related-workitems").appendTo(this._container).text(Strings.NoWorkItemsFound);
        }
    }

    private _renderWorkItem(workItem: WitContracts.WorkItem): void {
        var workItemType = Utils_String.htmlEncode(workItem.fields["System.WorkItemType"]);
        var title = Utils_String.htmlEncode(workItem.fields["System.Title"] || "");
        var id = workItem.id;
        var assignedTo: IdentityReference = IdentityHelper.parseIdentity(workItem.fields["System.AssignedTo"]);
        var state = Utils_String.htmlEncode(workItem.fields["System.State"] || "");
        var tags = workItem.fields["System.Tags"] || "";
        var identityImageUrl = `${VSS.getWebContext().host.uri}/_api/_common/IdentityImage?id=`;
        if (assignedTo.isIdentity && assignedTo.uniqueName) {
            identityImageUrl = `${VSS.getWebContext().host.uri}/_api/_common/IdentityImage?identifier=${assignedTo.uniqueName}&identifierType=0`;
        }
        var stateColor = StateColorHelper.parseColor(state);

        var rowHtmlString = 
            `<div class='work-item-row' title='${workItemType}: ${title}'>
                <div class='work-item-row-header'>
                    <div class='workitem-color' style='background-color: ${WorkItemTypeColorHelper.parseColor(workItemType)}'></div>
                    <div class='workitem-id'><a>${id}</a></div>
                    <div class='workitem-title'>${title}</div>
                </div>
                <div class='work-item-row-subtitle' title='${workItemType}: ${title}'>
                    <div class='workitem-state'><span style='background-color: ${stateColor}; border-color: ${stateColor}'></span>${state}</div>
                    <div class='workitem-assignedTo'>
                        <img src='${identityImageUrl}' />${Utils_String.htmlEncode(assignedTo.displayName)}
                    </div>
                    <div class='workitem-tags'>${this._createTagItemString(tags)}</div>
                </div>
            </div>`;
            
        var $row = $("<div>").append(rowHtmlString);
        $row.find(".workitem-id").click((e) => {
            this._options.openWorkItem(id, e.ctrlKey);
        });

        this._container.append($row);
    }

    private _createTagItemString(tags: string): string {
        if (tags) {
            let tagArray = tags.split(";");
            let html = "<ul class='tag-list'>";
            $.each(tagArray, (i: number, tag: string) => {
                html += `<li class='tag'>${Utils_String.htmlEncode(tag)}</li>`;
            })
            html += "</ul>";

            return html;
        }
        return "";
    }

    public dispose(): void {
        this._container.empty();
        this._container.remove();
        super.dispose();
    }
}

export class RelatedFieldsControl extends Control<RelatedFieldsControlOptions> {   
    private _container: JQuery;
    private _fieldsListContainer: JQuery;
    private _selectedFields: string[];
    private _originalSelectedFields: string[];
    private _allFields: WitContracts.WorkItemField[];
    private _refNameToFieldMap: IDictionaryStringTo<WitContracts.WorkItemField>;
    private _nameToFieldMap: IDictionaryStringTo<WitContracts.WorkItemField>;
    private _isDirty: boolean;
    private _sortByField: string;
    private _originalSortByField: string;

    constructor(options?: RelatedFieldsControlOptions) {
        super(options);
    }

    public initialize(): void {
        super.initialize();

        this._isDirty = false;
        this._selectedFields = this._options.selectedFields || [];
        this._originalSelectedFields = this._selectedFields.slice();
        this._allFields = this._options.allFields || [];
        this._sortByField = this._originalSortByField = this._options.sortByField;

        this._prepareFieldsMap();

        // Initialize elements
        this._container = $("<div/>").addClass("container").appendTo(this.getElement());
        this._fieldsListContainer = $("<ul class='fields-list'>").appendTo(this._container);
        this._render();
    }

    public dispose(): void {
        this._container.empty();
        this._container.remove();
        super.dispose();
    }

    private _render(): void {
        this._fieldsListContainer.empty();
        this._renderSortByFieldSelector();

        $.each(this._selectedFields, (i: number, fieldRefName: string) => {
            this._renderField(fieldRefName);
        });

        this._renderAddField();
        this._renderSaveButton();
        this._renderRefreshButton();
    }

    private _renderSortByFieldSelector(): void {
        var $itemLabel = $("<li class='fields-list-item sort-by-field-label'>").text(Strings.SortBy).appendTo(this._fieldsListContainer);
        var $item = $("<li class='fields-list-item sort-by-field'>").appendTo(this._fieldsListContainer);
        var comboSource = 
                $.map(this._allFields, (f: WitContracts.WorkItemField) => {
                    if (Utils_Array.contains(Constants.SortableFieldTypes, f.type)) {
                        return f.name;
                    }
                    else {
                        return null;
                    }
                }).sort(fieldNameComparer); 

        var combo = <Combo>BaseControl.createIn(Combo, $item, {
                allowEdit: false,
                mode: "drop",
                type: "list",
                value: this._sortByField,
                maxAutoExpandDropWidth: 200,
                source: comboSource,
                indexChanged: (index: number) => {
                    var field = this._nameToFieldMap[comboSource[index]];
                    if (field) {
                        this._sortByField = field.referenceName;
                        this._ensureIsDirty();
                        this._render();
                    }
                }
            });
    }

    private _renderSaveButton(): void {
        var $item = $("<li class='fields-list-item save-preferences bowtie-icon bowtie-save bowtie-white'>").attr("title", Strings.NeedAtleastOneField).appendTo(this._fieldsListContainer);
        if (this._isDirty && this._selectedFields && this._selectedFields.length > 1) {
            $item.addClass("enabled");
            $item.attr("title", Strings.SavePreferenceTitle);
        }
        $item.click(() => {
            if ($.isFunction(this._options.savePreferences) && this._isDirty && this._selectedFields && this._selectedFields.length > 0) {
                this._options.savePreferences({
                    fields: this._selectedFields,
                    sortByField: this._sortByField
                });
                this._isDirty = false;
                $item.removeClass("enabled");
                this._originalSelectedFields = this._selectedFields.slice();
                this._originalSortByField = this._sortByField;
            }
        });
    }

    private _renderRefreshButton(): void {
        var $item = $("<li class='fields-list-item refresh-list bowtie-icon bowtie-navigate-refresh'>").attr("title", Strings.NeedAtleastOneField).appendTo(this._fieldsListContainer);

        if (this._selectedFields && this._selectedFields.length > 1) {
            // We check for selectedFields length = 1 because we dont let users remove TeamProject field. So there'll be atleast 1 field in the array
            $item.addClass("enabled");
            $item.attr("title", Strings.RefreshList);
        }
        $item.click(() => {
            if ($.isFunction(this._options.refresh) && this._selectedFields && this._selectedFields.length > 0) {
                this._options.refresh(this._selectedFields, this._sortByField);
            }
        });
    }

    private _renderField(fieldRefName: string): void {
        var field = this._refNameToFieldMap[fieldRefName];
        if (field && !Utils_String.equals(field.referenceName, "System.TeamProject", true)) {
            var $item = $("<li class='fields-list-item'>").attr("title", fieldRefName).appendTo(this._fieldsListContainer);
            var rowHtmlString = 
                `<div class='field-list-item-name'>${field.name}</div>
                <span class='remove-item bowtie-icon bowtie-edit-delete' title='${Strings.RemoveItemTitle}'></span>`;

            $item.append(rowHtmlString);
            $item.find(".bowtie-edit-delete").click(() => {
                Utils_Array.remove(this._selectedFields, fieldRefName);                
                this._ensureIsDirty();
                this._render();
            });
        }
    }

    private _renderAddField(): void {
        var $item = $("<li class='fields-list-item add-fields-list-item'>").appendTo(this._fieldsListContainer);

        var comboSource = 
                $.map(this._allFields, (f: WitContracts.WorkItemField) => {
                    if (!Utils_Array.contains(this._selectedFields, f.referenceName) && 
                        (Utils_Array.contains(Constants.QueryableFieldTypes, f.type) || Utils_String.equals(f.referenceName, "System.Tags", true)))
                    {
                        return f.name;
                    }
                    else {
                        return null;
                    }
                }).sort(fieldNameComparer);                

        if (comboSource.length > 0 && this._selectedFields.length <= 10) {
            let combo = <Combo>BaseControl.createIn(Combo, $item, {
                allowEdit: false,
                mode: "drop",
                type: "list",
                maxAutoExpandDropWidth: 200,
                source: comboSource,
                indexChanged: (index: number) => {
                    var field = this._nameToFieldMap[comboSource[index]];
                    if (field) {
                        this._selectedFields.push(field.referenceName);
                        this._ensureIsDirty();
                        this._render();
                    }
                }
            });

            Utils_UI.Watermark($("input", combo.getElement()), { watermarkText: Strings.AddFieldPlaceholder });
        }
    }

    private _ensureIsDirty(): void {
        this._isDirty = !Utils_String.equals(this._sortByField, this._originalSortByField, true) || 
                        !Utils_Array.arrayEquals(this._selectedFields, this._originalSelectedFields, (s: string, t: string) => {
                            return Utils_String.equals(s, t, true);
                        });        
    }

    private _prepareFieldsMap(): void {
        this._refNameToFieldMap = {};
        this._nameToFieldMap = {};
        $.each(this._allFields, (i: number, field: WitContracts.WorkItemField) => {
            this._refNameToFieldMap[field.referenceName] = field;
            this._nameToFieldMap[field.name] = field;
        });
    }
}

