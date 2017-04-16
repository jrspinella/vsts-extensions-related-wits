import "../css/app.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemFieldChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, Wiql, WorkItemQueryResult, WorkItemRelation } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { SearchBox } from "OfficeFabric/SearchBox";

import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { InputError } from "./InputError";
import { UserPreferenceModel } from "./Models";
import { UserPreferences } from "./UserPreferences";
import { WorkItemsViewer, IListItem } from "./WorkItemsViewer";

interface IRelatedWitsState {
    loading: boolean;
    items: IListItem[];
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: UserPreferenceModel;
    settingsPanelOpen?: boolean;
    filterText?: string;
    sortColumn: string;
    sortOrder: string;
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {

    constructor(props: void, context?: any) {
        super(props, context);

        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this._updateState({isWorkItemLoaded: true, isNew: true});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (args: IWorkItemChangedArgs) => {
                this._updateState({isWorkItemLoaded: false, workItems: []});
            },
            onSaved: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onRefreshed: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onReset: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
        } as IWorkItemNotificationListener);

        this.state = {
            loading: true,
            items: [],    
            sortColumn: "System.CreatedDate",
            sortOrder: "desc"        
        } as IRelatedWitsState;
    }

    public render(): JSX.Element {
        if (!this.state.isWorkItemLoaded) {
            return null;
        }
        else if (this.state.isNew) {
            return <MessagePanel message="Please save the workitem to get the list of related work items." messageType={MessageType.Info} />;
        }
        else {
            return (
                <Fabric className="fabric-container">
                    <div className="results-view-menu">
                        <SearchBox className="results-view-searchbox" 
                            value={this.state.filterText || ""}
                            onSearch={(searchText: string) => this._updateFilterText(searchText)} 
                            onChange={(newText: string) => {
                                if (newText.trim() === "") {
                                    this._updateFilterText("");
                                }
                            }} />
                        <CommandBar className="all-view-menu-toolbar" items={this._getMenuItems()} />
                    </div>
                    {this._renderList()}
                </Fabric>
            );
        }        
    }

    private _updateState(updatedStates: any) {
        this.setState({...this.state, ...updatedStates});
    }

    private _updateFilterText(searchText: string): void {
        this._updateState({filterText: searchText});
    }

    private _getMenuItems(): IContextualMenuItem[] {
         return [                      
            {
                key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this._refreshList();
                }
            },
            {
                key: "OpenQuery", name: "Open as query", title: "Open all workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this.state.items.length === 0,
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._openQueryWiql())}`;
                    window.open(url, "_parent");
                }
            },
            {
                key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                disabled: this.state.settings == null, checked: this.state.settingsPanelOpen,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this._updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                }
            }
         ] as IContextualMenuItem[];
    }

    private _openQueryWiql(): string {
        let ids = this.state.items.map((item: IListItem) => item.workItem.id).join(",");

        return `SELECT [System.Id], [System.WorkItemType], [System.Title], [System.State], [System.AssignedTo], [System.AreaPath], [System.Tags]
                 FROM WorkItems 
                 WHERE [System.TeamProject] = @project 
                 AND [System.ID] IN (${ids}) 
                 ORDER BY [System.CreatedDate] DESC`;
    }

    private _renderList(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }        
        else if (!this.state.items || this.state.items.length === 0) {
            return <MessagePanel message="No related work items found." messageType={MessageType.Info} />;
        }
        else {
            return <WorkItemsViewer 
                        items={this._sortAndFilterWorkItems(this.state.items)} 
                        sortColumn={this.state.sortColumn} 
                        sortOrder={this.state.sortOrder} 
                        changeSort={this._changeSort} />;
        }
    }

    @autobind
    private _sortAndFilterWorkItems(workItems: WorkItem[]): WorkItem[] {
        let items = workItems.slice();
        let sortedItems = items.sort((w1: WorkItem, w2: WorkItem) => {
            if (Utils_String.equals(this.state.sortColumn, "ID", true)) {
                return this.state.sortOrder === "desc" ? ((w1.id > w2.id) ? -1 : 1) : ((w1.id > w2.id) ? 1 : -1);
            }
            else if (Utils_String.equals(this.state.sortColumn, "System.CreatedDate", true)) {
                let d1 = new Date(w1.fields["System.CreatedDate"]);
                let d2 = new Date(w2.fields["System.CreatedDate"]);
                return this.state.sortOrder === "desc" ? -1 * Utils_Date.defaultComparer(d1, d2) : Utils_Date.defaultComparer(d1, d2);
            }
            else if (Utils_String.equals(this.state.sortColumn, Constants.ACCEPT_STATUS_CELL_NAME, true)) {
                let v1 = Helpers.isWorkItemAccepted(w1) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(w1) ? Constants.REJECTED_TEXT : "");
                let v2 = Helpers.isWorkItemAccepted(w2) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(w2) ? Constants.REJECTED_TEXT : "");
                return this.state.sortOrder === "desc" ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
            }
            else {
                let v1 = w1.fields[this.state.sortColumn];
                let v2 = w2.fields[this.state.sortColumn];
                return this.state.sortOrder === "desc" ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
            }
        });

        if (!this.state.filterText) {
            return sortedItems;
        }
        else {
            return sortedItems.filter((workItem: WorkItem) => {
                let status = Helpers.isWorkItemAccepted(workItem) ? Constants.ACCEPTED_TEXT : (Helpers.isWorkItemRejected(workItem) ? Constants.REJECTED_TEXT : "");
                const filterText = this.state.filterText;
                return `${workItem.id}` === filterText
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AssignedTo"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.State"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.CreatedBy"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.Title"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AreaPath"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(status, filterText);
            });
        }
    }
    
    @autobind
    private _changeSort(sortColumn: string, sortOrder: string): void {
        this._updateState({sortColumn: sortColumn, sortOrder: sortOrder});
    }

    private async _refreshList(): Promise<void> {
        this._updateState({isWorkItemLoaded: true, isNew: false, loading: true, listItems: []});

        if (!this.state.settings) {
            const workItemFormService = await WorkItemFormService.getService();
            const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
            let settings = await UserPreferences.readUserSetting(workItemType);
            this._updateState({settings: settings});
        }


        this._updateState({loading: false, listItems: []});
    } 

    private async _getWorkItems(fieldsToSeek: string[], sortByField: string): Promise<IListItem[]> {
        let data: string[] = await this._createWiql(fieldsToSeek, sortByField);
        let queryResults = await WitClient.getClient().queryByWiql({ query: data[1] }, data[0], null, false, 20);
        return this._readWorkItemsFromQueryResults(queryResults);
    }

    private async _createWiql(fieldsToSeek: string[], sortByField: string): Promise<string[]> {
        let fieldValuesToRead = fieldsToSeek.concat(["System.ID"]);
        let workItemFormService = await this._ensureWorkItemFormService()
        let fieldValues = await workItemFormService.getFieldValues(fieldValuesToRead);
        let witId = fieldValues["System.ID"]; 
        // Generate fields to retrieve part
        let fieldsToRetrieveString: string = "";
        $.each(Constants.DEFAULT_FIELDS_TO_RETRIEVE, (i: number, fieldRefName: string) => {
            fieldsToRetrieveString = `${fieldsToRetrieveString}[${fieldRefName}],`
        });
        // remove last comma
        fieldsToRetrieveString = fieldsToRetrieveString.substring(0, fieldsToRetrieveString.length - 1);

        // Generate fields to seek part
        let fieldsToSeekString: string = "";
        $.each(fieldsToSeek, (i: number, fieldRefName: string) => {
            let fieldValue = fieldValues[fieldRefName];
            if (Utils_String.equals(fieldRefName, "System.Tags", true) && fieldValue) {
                fieldsToSeekString = fieldsToSeekString + " (";
                let fieldValueStr = fieldValue.toString();
                $.each(fieldValueStr.split(";"), (i: number, v: string) => {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] CONTAINS '${v}' OR`
                });
                if (fieldsToSeekString) {
                    // remove last OR
                    fieldsToSeekString = fieldsToSeekString.substring(0, fieldsToSeekString.length - 3) + ") AND";
                }
            }
            else if (!Utils_String.equals(fieldRefName, "System.TeamProject", true)) {
                if (Utils_String.equals(typeof(fieldValue), "string", true) && fieldValue) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = '${fieldValue}' AND`
                }
                else if (Utils_String.equals(typeof(fieldValue), "number", true) && fieldValue != null) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = ${fieldValue} AND`
                }
                else if (Utils_String.equals(typeof(fieldValue), "boolean", true) && fieldValue != null) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = ${fieldValue} AND`
                }
            }
        });
        if (fieldsToSeekString) {
            // remove last OR
            fieldsToSeekString = fieldsToSeekString.substring(0, fieldsToSeekString.length - 4);
        }
        let fieldsToSeekPredicate = fieldsToSeekString ? `AND ${fieldsToSeekString}` : "";
        let wiql = `SELECT ${fieldsToRetrieveString} FROM workitems where [System.TeamProject] = @project AND [System.ID] <> ${witId} ${fieldsToSeekPredicate} order by [${sortByField}] desc`;

        return [fieldValues["System.TeamProject"] as string, wiql];
    }


    private async _readWorkItemsFromQueryResults(queryResults: WorkItemQueryResult): Promise<IListItem[]> {
        if (queryResults.workItems && queryResults.workItems.length > 0) {
            var workItemIdMap: IDictionaryNumberTo<WorkItemReference> = {};
            var workItemIds = $.map(queryResults.workItems, (elem: WorkItemReference) => {
                return elem.id;
            });
            var fields = $.map(queryResults.columns, (elem: WorkItemFieldReference) => {
                return elem.referenceName;
            });

            $.each(queryResults.workItems, (i:number, w: WorkItemReference) => {
                workItemIdMap[w.id] = w;
            })

            let workItems = await WitClient.getClient().getWorkItems(workItemIds, fields);
            // sort the workitems in the same order as they got retrieved from query
            var sortedWorkItems = workItems.sort((w1: WorkItem, w2: WorkItem) => {
                                        if (workItemIds.indexOf(w1.id) < workItemIds.indexOf(w2.id)) { return -1 }
                                        if (workItemIds.indexOf(w1.id) > workItemIds.indexOf(w2.id)) { return 1 }
                                        return 0;
                                    });

            let workItemFormService = await WorkItemFormService.getService();
            let relations = await workItemFormService.getWorkItemRelations();
            return $.map(sortedWorkItems, (w: WorkItem) => $.extend(w, { 
                url: workItemIdMap[w.id].url,
                isLinked: Utils_Array.arrayContains(w.url, relations, (url: string, relation: WorkItemRelation) => {
                    return Utils_String.equals(relation.url, url, true);
                })
            }));
        }
        else {
            return [];
        }
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}