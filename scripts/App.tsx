import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemType, WorkItemStateColor, Wiql, WorkItemQueryResult, WorkItemRelation, WorkItemRelationType } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { SearchBox } from "OfficeFabric/SearchBox";

import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { UserPreferenceModel, Constants } from "./Models";
import { UserPreferences } from "./UserPreferences";
import { WorkItemsViewer } from "./WorkItemsViewer";
import { SettingsPanel } from "./SettingsPanel";

interface IRelatedWitsState {
    loading: boolean;
    items: WorkItem[];
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: UserPreferenceModel;
    workItemTypeColors?: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;
    settingsPanelOpen?: boolean;
    filterText?: string;
    sortColumn: string;
    sortOrder: string; 
    relationsMap: IDictionaryStringTo<boolean>;
    relationTypes?: WorkItemRelationType[];
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {

    constructor(props: void, context?: any) {
        super(props, context);        

        this.state = {
            loading: true,
            items: [],    
            sortColumn: "System.CreatedDate",
            sortOrder: "desc",
            relationsMap: null
        } as IRelatedWitsState;
    }

    public componentDidMount() {
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
                    <div className="menu-container">
                        <SearchBox 
                            className="searchbox" 
                            value={this.state.filterText || ""}
                            onSearch={(searchText: string) => this._updateFilterText(searchText)} 
                            onChange={(newText: string) => {
                                if (newText.trim() === "") {
                                    this._updateFilterText("");
                                }
                            }} />

                        <CommandBar 
                            className="menu-bar" 
                            items={this._getMenuItems()} 
                            farItems={
                                [
                                    {
                                        key: "resultCount", name: this.state.loading ? "" : `${this.state.items.length} results`, className: "result-count"
                                    }
                                ]
                            } />
                    </div>
                    { 
                        this.state.settingsPanelOpen && 
                        <SettingsPanel 
                            settings={this.state.settings} 
                            onSave={(settings: UserPreferenceModel) => {
                                this._updateState({settings: settings});
                                this._refreshList();
                            }} />
                    }
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
                    window.open(url, "_blank");
                }
            },
            {
                key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                style: this.state.settingsPanelOpen ? {backgroundColor: "#EAEAEA"} : null,
                disabled: this.state.settings == null,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this._updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                }
            }
         ] as IContextualMenuItem[];
    }

    private _openQueryWiql(): string {
        let ids = this.state.items.map((workItem: WorkItem) => workItem.id).join(",");

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
        const filteredItems = this._sortAndFilterWorkItems(this.state.items);

        if (filteredItems.length === 0) {
            return <MessagePanel message="No related work items found." messageType={MessageType.Info} />;
        }
        else {
            return <WorkItemsViewer 
                        items={filteredItems} 
                        sortColumn={this.state.sortColumn} 
                        sortOrder={this.state.sortOrder} 
                        workItemTypeColors={this.state.workItemTypeColors}
                        relationsMap={this.state.relationsMap}
                        relationTypes={this.state.relationTypes}
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
                const filterText = this.state.filterText;
                return `${workItem.id}` === filterText
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AssignedTo"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.State"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.Title"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.WorkItemType"] || "", filterText)
                    || Utils_String.caseInsensitiveContains(workItem.fields["System.AreaPath"] || "", filterText)
            });
        }
    }
    
    @autobind
    private _changeSort(sortColumn: string, sortOrder: string): void {
        this._updateState({sortColumn: sortColumn, sortOrder: sortOrder});
    }

    private async _refreshList(): Promise<void> {
        this._updateState({isWorkItemLoaded: true, isNew: false, loading: true, items: [], relationsMap: null});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        const items = await this._getWorkItems(this.state.settings.fields, this.state.settings.sortByField);
        this._updateState({loading: false, items: items, sortColumn: (this.state.settings && this.state.settings.sortByField) || Constants.DEFAULT_SORT_BY_FIELD, sortOrder: "desc"});

        if (!this.state.workItemTypeColors) {
            this._initializeWorkItemTypeColors();
        }

        if (!this.state.relationTypes) {
            this._initializeWorkItemRelationTypes();
        }

        this._initializeLinksData();
    }

    private async _initializeWorkItemRelationTypes() {
        const workItemFormService = await WorkItemFormService.getService();
        const relationTypes = await workItemFormService.getWorkItemRelationTypes();
        this._updateState({relationTypes: relationTypes});
    }

    private async _initializeLinksData() {
        const workItemFormService = await WorkItemFormService.getService();
        const relations = await workItemFormService.getWorkItemRelations();

        let relationsMap = {};
        relations.forEach(relation => {
            relationsMap[relation.url] = true;
        });

        this._updateState({relationsMap: relationsMap});
    }

    private async _initializeSettings() {
        const workItemFormService = await WorkItemFormService.getService();
        const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
        const settings = await UserPreferences.readUserSetting(workItemType);

        this._updateState({settings: settings});
    }

    private async _initializeWorkItemTypeColors() {
        const workItemFormService = await WorkItemFormService.getService();
        const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
        let workItemTypeColors: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}> = {};

        const workItemTypes = await WitClient.getClient().getWorkItemTypes(project);
        workItemTypes.forEach((wit: WorkItemType) => workItemTypeColors[wit.name] = {
            color: wit.color,
            stateColors: {}
        });

        try {
            await Promise.all(workItemTypes.map(async (wit: WorkItemType) => {
                let stateColors = await WitClient.getClient().getWorkItemTypeStates(project, wit.name);
                stateColors.forEach((stateColor: WorkItemStateColor) => workItemTypeColors[wit.name].stateColors[stateColor.name] = stateColor.color);
            }));
        }
        catch (e) {

        }
        
        this._updateState({workItemTypeColors: workItemTypeColors});
    }

    private async _getWorkItems(fieldsToSeek: string[], sortByField: string): Promise<WorkItem[]> {
        let {project, wiql} = await this._createWiql(fieldsToSeek, sortByField);
        let queryResults = await WitClient.getClient().queryByWiql({ query: wiql }, project, null, false, this.state.settings.top);
        return this._readWorkItemsFromQueryResults(queryResults);
    }

    private async _createWiql(fieldsToSeek: string[], sortByField: string): Promise<{project: string, wiql: string}> {
        const workItemFormService = await WorkItemFormService.getService();
        const fieldValues = await workItemFormService.getFieldValues(fieldsToSeek, true);
        const witId = await workItemFormService.getId();
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
       
        // Generate fields to retrieve part
        const fieldsToRetrieveString = Constants.DEFAULT_FIELDS_TO_RETRIEVE.map(fieldRefName => `[${fieldRefName}]`).join(",");

        // Generate fields to seek part
        const fieldsToSeekString = fieldsToSeek.map(fieldRefName => {
            const fieldValue = fieldValues[fieldRefName];
            if (Utils_String.equals(fieldRefName, "System.Tags", true)) {
                if (fieldValue) {
                    let tagStr = fieldValue.toString().split(";").map(v => {
                        return `[System.Tags] CONTAINS '${v}'`;
                    }).join(" OR ");

                    return `(${tagStr})`;
                }                
            }
            else if (Constants.ExcludedFields.indexOf(fieldRefName) === -1) {
                if (Utils_String.equals(typeof(fieldValue), "string", true) && fieldValue) {
                    return `[${fieldRefName}] = '${fieldValue}'`;
                }
                else {
                    return `[${fieldRefName}] = ${fieldValue}`;
                }
            }

            return null;
        }).filter(e => e != null).join(" AND ");

        let fieldsToSeekPredicate = fieldsToSeekString ? `AND ${fieldsToSeekString}` : "";
        let wiql = `SELECT ${fieldsToRetrieveString} FROM workitems 
                    where [System.TeamProject] = '${project}' AND [System.ID] <> ${witId} 
                    ${fieldsToSeekPredicate} order by [${sortByField}] desc`;

        return {
            project: project,
            wiql: wiql
        };
    }

    private async _readWorkItemsFromQueryResults(queryResult: WorkItemQueryResult): Promise<WorkItem[]> {
        let workItemIds = queryResult.workItems.map(workItem => workItem.id);
        let workItems: WorkItem[];

        if (workItemIds.length > 0) {
            return await WitClient.getClient().getWorkItems(workItemIds);
        }
        else {
            return [];
        }
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}