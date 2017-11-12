import "./App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { RelatedWorkItemsActions } from "../Actions/RelatedWorkItemsActions";
import { getQueryUrl } from "../Helpers";
import { Constants, Settings, WorkItemFieldNames } from "../Models";
import { StoresHub } from "../Stores/StoresHub";
import * as SettingsPanel_Async from "./SettingsPanel";

import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import {
    CheckboxVisibility, ConstrainMode, DetailsListLayoutMode, IColumn
} from "OfficeFabric/DetailsList";
import { Fabric } from "OfficeFabric/Fabric";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Panel, PanelType } from "OfficeFabric/Panel";
import {
    DirectionalHint, TooltipDelay, TooltipHost, TooltipOverflowMode
} from "OfficeFabric/Tooltip";
import { autobind } from "OfficeFabric/Utilities";
import { ISelection, Selection, SelectionMode } from "OfficeFabric/utilities/selection";

import { WorkItem, WorkItemRelation, WorkItemRelationType } from "TFS/WorkItemTracking/Contracts";
import {
    IWorkItemChangedArgs, IWorkItemLoadedArgs, IWorkItemNotificationListener
} from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

import { PickListFilterBarItem } from "VSSUI/Components/PickList";
import { FilterBar, IFilterBar, KeywordFilterBarItem } from "VSSUI/FilterBar";
import { Hub } from "VSSUI/Hub";
import { HubHeader } from "VSSUI/HubHeader";
import { PivotBarItem } from "VSSUI/PivotBar";
import { FILTER_CHANGE_EVENT, IFilterState } from "VSSUI/Utilities/Filter";
import { HubViewOptionKeys, HubViewState, IHubViewState } from "VSSUI/Utilities/HubViewState";
import { VssDetailsList } from "VSSUI/VssDetailsList";
import { VssIconType } from "VSSUI/VssIcon";
import { ZeroData } from "VSSUI/ZeroData";

import { IdentityView } from "VSTS_Extension_Widgets/Components/IdentityView";
import { InfoLabel } from "VSTS_Extension_Widgets/Components/InfoLabel";
import { Loading } from "VSTS_Extension_Widgets/Components/Loading";
import {
    getAsyncLoadedComponent
} from "VSTS_Extension_Widgets/Components/Utilities/AsyncLoadedComponent";
import {
    BaseFluxComponent, IBaseFluxComponentProps, IBaseFluxComponentState
} from "VSTS_Extension_Widgets/Components/Utilities/BaseFluxComponent";
import { StateView } from "VSTS_Extension_Widgets/Components/WorkItemComponents/StateView";
import { TitleView } from "VSTS_Extension_Widgets/Components/WorkItemComponents/TitleView";
import { BaseStore } from "VSTS_Extension_Widgets/Flux/Stores/BaseStore";
import { ExtensionDataManager } from "VSTS_Extension_Widgets/Utilities/ExtensionDataManager";
import { parseUniquefiedIdentityName } from "VSTS_Extension_Widgets/Utilities/Identity";
import { StringUtils } from "VSTS_Extension_Widgets/Utilities/String";
import { openWorkItemDialog } from "VSTS_Extension_Widgets/Utilities/WorkItemGridHelpers";

import { initializeIcons } from "@uifabric/icons";

const AsyncSettingsPanel = getAsyncLoadedComponent(
    ["scripts/SettingsPanel"],
    (m: typeof SettingsPanel_Async) => m.SettingsPanel,
    () => <Loading />);

export interface IRelatedWitsState extends IBaseFluxComponentState {
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    workItems: WorkItem[];
    settings?: Settings;
    settingsPanelOpen?: boolean;
    relationsMap?: IDictionaryStringTo<boolean>;
    relationTypes?: WorkItemRelationType[];
}

export class RelatedWits extends BaseFluxComponent<IBaseFluxComponentProps, IRelatedWitsState> {
    private _hubViewState: IHubViewState;
    private _filterBar: IFilterBar;
    private _selection: ISelection;

    constructor(props: IBaseFluxComponentProps, context?: any) {
        super(props, context);

        this._hubViewState = new HubViewState();
        this._hubViewState.viewOptions.setViewOption(HubViewOptionKeys.fullScreen, true);
        this._selection = new Selection({
            getKey: (item: any) => item.id
        });
    }

    protected initializeState(): void {
        this.state = {
            isWorkItemLoaded: false,
            workItems: null,
            settings: null
        }
    }

    protected getStores(): BaseStore<any, any, any>[] {
        return [StoresHub.relatedWorkItemsStore];
    }
    protected getStoresState(): IRelatedWitsState {
        return {
            workItems: StoresHub.relatedWorkItemsStore.isLoading() ? null : StoresHub.relatedWorkItemsStore.getFilteredItems()
        }
    }
    
    public componentDidMount() {
        super.componentDidMount();
        this._hubViewState.filter.subscribe(this._onFilterChange, FILTER_CHANGE_EVENT);
        document.addEventListener("keydown", this._focusFilterBar, false);
        
        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this.setState({isWorkItemLoaded: true, isNew: true, workItems: null, settingsPanelOpen: false});
                }
                else {
                    this.setState({isWorkItemLoaded: true, isNew: false, settingsPanelOpen: false});
                    this._refreshList();
                }
            },
            onUnloaded: (_args: IWorkItemChangedArgs) => {
                this.setState({isWorkItemLoaded: false, workItems: null, settingsPanelOpen: false});
            },
            onSaved: (_args: IWorkItemChangedArgs) => {
                this.setState({isNew: false, settingsPanelOpen: false});
                this._refreshList();
            },
            onRefreshed: (_args: IWorkItemChangedArgs) => {
                this.setState({settingsPanelOpen: false});
                this._refreshList();
            }
        } as IWorkItemNotificationListener);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();

        this._hubViewState.filter.unsubscribe(this._onFilterChange, FILTER_CHANGE_EVENT);
        document.removeEventListener("keydown", this._focusFilterBar);
        VSS.unregister(VSS.getContribution().id);
        RelatedWorkItemsActions.clean();
    }

    public render(): JSX.Element {
        if (!this.state.isWorkItemLoaded) {
            return null;
        }        
        else {
            return (
                <Fabric className="fabric-container">
                    <Panel
                        isOpen={this.state.settingsPanelOpen}
                        type={PanelType.custom}
                        customWidth="450px"
                        isLightDismiss={true} 
                        onDismiss={() => this.setState({settingsPanelOpen: false})}>

                        <AsyncSettingsPanel
                            settings={this.state.settings} 
                            onSave={(settings: Settings) => {
                                this.setState({settings: settings, settingsPanelOpen: false});
                                this._refreshList();
                            }} />
                    </Panel>
                    <Hub
                        className="related-wits-hub"
                        hideFullScreenToggle={true}
                        hubViewState={this._hubViewState}
                        commands={[
                            { 
                                key: "refresh", 
                                name: "Refresh", 
                                disabled: this.state.workItems == null || this.state.isNew, 
                                important: true, 
                                iconProps: { iconName: "Refresh", iconType: VssIconType.fabric }, 
                                onClick: this._refreshList 
                            },
                            { 
                                key: "settings", 
                                name: "Settings", 
                                disabled: this.state.workItems == null || this.state.isNew, 
                                important: true, 
                                iconProps: { iconName: "Settings", iconType: VssIconType.fabric }, 
                                onClick: () => this.setState({settingsPanelOpen: !(this.state.settingsPanelOpen)}) 
                            }
                        ]}>
                        
                        <HubHeader title="Related work items" />
                        <FilterBar componentRef={this._resolveFilterBar}>
                            <KeywordFilterBarItem filterItemKey={"keyword"} />
                            <PickListFilterBarItem
                                placeholder={"Work Item Type"}
                                filterItemKey={WorkItemFieldNames.WorkItemType}
                                selectionMode={SelectionMode.multiple}
                                getPickListItems={() => { return Object.keys(StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.WorkItemType]) }}
                                getListItem={(wit: string) => {
                                    return {
                                        name: wit,
                                        key: wit
                                    }
                                }}
                                indicators={[
                                    {
                                        getItemIndicator: ((wit: string) => {
                                            if (!wit) {
                                                return null;
                                            }
                                            return { title: "" + StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.WorkItemType][wit] }
                                        })
                                    }
                                ]}
                            />
                            <PickListFilterBarItem
                                placeholder={"State"}
                                filterItemKey={WorkItemFieldNames.State}
                                selectionMode={SelectionMode.multiple}
                                getPickListItems={() => { return Object.keys(StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.State]) }}
                                getListItem={(state: string) => {
                                    return {
                                        name: state,
                                        key: state
                                    }
                                }}
                                indicators={[
                                    {
                                        getItemIndicator: ((state: string) => {
                                            if (!state) {
                                                return null;
                                            }
                                            return { title: "" + StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.State][state] }
                                        })
                                    }
                                ]}
                            />
                            <PickListFilterBarItem
                                placeholder={"Assigned To"}
                                filterItemKey={WorkItemFieldNames.AssignedTo}
                                selectionMode={SelectionMode.multiple}
                                getPickListItems={() => { return Object.keys(StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.AssignedTo]) }}
                                getListItem={(assignedTo: string) => {
                                    const identity = parseUniquefiedIdentityName(assignedTo);
                                    return {
                                        name: identity.displayName,
                                        key: assignedTo,
                                        iconProps: identity.imageUrl ? {
                                            iconType: VssIconType.image,
                                            imageProps: {
                                                src: identity.imageUrl
                                            }
                                        } : null
                                    }
                                }}
                                indicators={[
                                    {
                                        getItemIndicator: ((assignedTo: string) => {
                                            if (!assignedTo) {
                                                return null;
                                            }
                                            return { title: "" + StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.AssignedTo][assignedTo] }
                                        })
                                    }
                                ]}
                            />
                            <PickListFilterBarItem
                                placeholder={"Area Path"}
                                filterItemKey={WorkItemFieldNames.AreaPath}
                                selectionMode={SelectionMode.multiple}
                                getPickListItems={() => { return Object.keys(StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.AreaPath]) }}
                                getListItem={(area: string) => {
                                    return {
                                        name: area.substr(area.lastIndexOf("\\") + 1),
                                        key: area
                                    }
                                }}
                                indicators={[
                                    {
                                        getItemIndicator: ((area: string) => {
                                            if (!area) {
                                                return null;
                                            }
                                            return { title: "" + StoresHub.relatedWorkItemsStore.propertyMap[WorkItemFieldNames.AreaPath][area] }
                                        })
                                    }
                                ]}
                            />
                        </FilterBar>

                        <PivotBarItem 
                            name={"Related Work items"} 
                            itemKey={"list"}
                            viewActions={[
                                {
                                    key: "status",
                                    name: this.state.isNew ? "" : (!this.state.workItems ? "Loading..." : `${this.state.workItems.length} results`),
                                    important: true
                                }]}>
                            {this._renderContent()}
                        </PivotBarItem>
                    </Hub>                        
                </Fabric>
            );
        }        
    }

    private _renderContent(): React.ReactNode {
        if (this.state.isNew) {
            return <MessageBar messageBarType={MessageBarType.info}>Please save the workitem to get the list of related work items.</MessageBar>;
        }
        else if (!this.state.workItems) {
            return <Loading />;
        }    
        else if (this.state.workItems.length === 0) {
            return <ZeroData imagePath="images/nodata.png" imageAltText="" primaryText="No results found" />;
        }
        else {
            return <div className="grid-container">
                <VssDetailsList 
                    items={this.state.workItems}
                    columns={this._getColumns()}
                    selectionPreservedOnEmptyClick={true}
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    onColumnHeaderClick={this._onSortChange}
                    checkboxVisibility={CheckboxVisibility.onHover}
                    selectionMode={SelectionMode.multiple}
                    className="work-item-grid"
                    selection={this._selection}
                    getKey={(item: WorkItem) => item.id.toString()}
                    onItemInvoked={this._onItemInvoked}
                    actionsColumnKey={WorkItemFieldNames.Title}
                    getMenuItems={this._getGridContextMenuItems}
                />
            </div>;
        }
    }

    private _getColumns(): IColumn[] {
        return [
            {
                key: "linked",
                name: "Linked",
                fieldName: "linked",
                minWidth: 60,
                maxWidth: 100,
                isResizable: false,
                isSorted: false,
                isSortedDescending: false,
                onRender: this._renderStatusColumnCell
            },
            {
                key: WorkItemFieldNames.ID,
                fieldName: WorkItemFieldNames.ID,
                name: "ID",
                minWidth: 40,
                maxWidth: 70,
                isResizable: true,
                isSorted: (StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.sortKey) === WorkItemFieldNames.ID,
                isSortedDescending: !!(StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.isSortedDescending),
                onRender: (workItem: WorkItem) => {
                    const id = workItem.id.toString();
                    return <TooltipHost 
                        content={id}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        {id}
                    </TooltipHost>;
                }
            },
            {
                key: WorkItemFieldNames.Title,
                fieldName: WorkItemFieldNames.Title,
                name: "Title",
                minWidth: 300,
                maxWidth: 600,
                isResizable: true,
                isSorted: (StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.sortKey) === WorkItemFieldNames.Title,
                isSortedDescending: !!(StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.isSortedDescending),
                onRender: (workItem: WorkItem) => {
                    const title = workItem.fields[WorkItemFieldNames.Title];
                    return <TooltipHost 
                        content={title}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <TitleView 
                            className="item-grid-cell" 
                            workItemId={workItem.id} 
                            onClick={(e: React.MouseEvent<HTMLElement>) => {
                                this._openWorkItemDialog(e, workItem);
                            }}
                            title={title} 
                            workItemType={workItem.fields[WorkItemFieldNames.WorkItemType]} />
                    </TooltipHost>;
                }
            },
            {
                key: WorkItemFieldNames.State,
                fieldName: WorkItemFieldNames.State,
                name: "State",
                minWidth: 100,
                maxWidth: 200,
                isResizable: true,
                isSorted: (StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.sortKey) === WorkItemFieldNames.State,
                isSortedDescending: !!(StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.isSortedDescending),
                onRender: (workItem: WorkItem) => {
                    const state = workItem.fields[WorkItemFieldNames.State];
                    return <TooltipHost 
                        content={state}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <StateView 
                            className="item-grid-cell" 
                            state={state} 
                            workItemType={workItem.fields[WorkItemFieldNames.WorkItemType]} />
                    </TooltipHost>;
                }
            },
            {
                key: WorkItemFieldNames.AssignedTo,
                fieldName: WorkItemFieldNames.AssignedTo,
                name: "Assigned To",
                minWidth: 150,
                maxWidth: 250,
                isResizable: true,
                isSorted: (StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.sortKey) === WorkItemFieldNames.AssignedTo,
                isSortedDescending: !!(StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.isSortedDescending),
                onRender: (workItem: WorkItem) => {
                    const assignedTo = workItem.fields[WorkItemFieldNames.AssignedTo] || "";
                    return <TooltipHost 
                        content={assignedTo}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        <IdentityView identityDistinctName={assignedTo} />
                    </TooltipHost>;
                }
            },
            {
                key: WorkItemFieldNames.AreaPath,
                fieldName: WorkItemFieldNames.AreaPath,
                name: "Area path",
                minWidth: 250,
                maxWidth: 400,
                isResizable: true,
                isSorted: (StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.sortKey) === WorkItemFieldNames.AreaPath,
                isSortedDescending: !!(StoresHub.relatedWorkItemsStore.sortState && StoresHub.relatedWorkItemsStore.sortState.isSortedDescending),
                onRender: (workItem: WorkItem) => {
                    const area = workItem.fields[WorkItemFieldNames.AreaPath];
                    return <TooltipHost 
                        content={area}
                        delay={TooltipDelay.medium}
                        overflowMode={TooltipOverflowMode.Parent}
                        directionalHint={DirectionalHint.bottomLeftEdge}>
                        {area}
                    </TooltipHost>;
                }
            }
        ];
    }

    @autobind
    private async _onItemInvoked(workItem: WorkItem) {
        this._openWorkItemDialog(null, workItem);
    }

    private async _openWorkItemDialog(e: React.MouseEvent<HTMLElement>, workItem: WorkItem) {
        const updatedWorkItem = await openWorkItemDialog(e, workItem);
        RelatedWorkItemsActions.updateWorkItemInStore(updatedWorkItem);
    }

    @autobind
    private _renderStatusColumnCell(item: WorkItem): JSX.Element {
        if (this.state.relationTypes && this.state.relationsMap) {
            let availableLinks: string[] = [];
            this.state.relationTypes.forEach(r => {
                if (this.state.relationsMap[`${item.url}_${r.referenceName}`]) {
                    availableLinks.push(r.name);
                }
            });

            if (availableLinks.length > 0) {
                return <InfoLabel 
                    className="linked-cell"
                    label="Linked" info={`Linked to this workitem as ${availableLinks.join("; ")}`} />;
            }
            else {
                return <InfoLabel 
                    label="Not linked" 
                    className="unlinked-cell"
                    info="This workitem is not linked to the current work item. You can add a link to this workitem by right clicking on the row" />;
            }
        }
        return null;
    }

    @autobind
    private _getGridContextMenuItems(item: WorkItem): IContextualMenuItem[] {
        let selectedItems = this._selection.getSelection() as WorkItem[];
        if (!selectedItems || selectedItems.length === 0) {
            selectedItems = [item];
        }

        return [
            {
                key: "openinquery", name: "Open selected items in Queries", iconProps: {iconName: "ReplyMirrored"}, 
                onClick: () => {                
                    const url = getQueryUrl(selectedItems, [
                        WorkItemFieldNames.ID,
                        WorkItemFieldNames.Title,
                        WorkItemFieldNames.State,
                        WorkItemFieldNames.AssignedTo,
                        WorkItemFieldNames.AreaPath
                    ]);
                    window.open(url, "_blank");
                }
            },
            {
                key: "add-link", name: "Add Link", title: "Add as a link to the current workitem", iconProps: {iconName: "Link"}, 
                items: this.state.relationTypes.filter(r => r.name != null && r.name.trim() !== "").map(relationType => {
                    return {
                        key: relationType.referenceName,
                        name: relationType.name,
                        onClick: async () => {
                            const workItemFormService = await WorkItemFormService.getService();
                            let workItemRelations = selectedItems.filter(wi => !this.state.relationsMap[`${wi.url}_${relationType.referenceName}`]).map(w => {
                                return {
                                    rel: relationType.referenceName,
                                    attributes: {
                                        isLocked: false
                                    },
                                    url: w.url
                                } as WorkItemRelation;
                            });

                            if (workItemRelations) {
                                workItemFormService.addWorkItemRelations(workItemRelations);
                            }                                
                        }
                    };
                })
            }
        ];
    }

    @autobind
    private async _refreshList(): Promise<void> {
        if (!this.state.settings) {
            await this._initializeSettings();
        }

        if (!this.state.relationTypes) {
            this._initializeWorkItemRelationTypes();
        }

        this._initializeLinksData();

        const query = await this._createQuery(this.state.settings.fields, this.state.settings.sortByField);
        RelatedWorkItemsActions.refresh(query, this.state.settings.top);        
    }

    private async _createQuery(fieldsToSeek: string[], sortByField: string): Promise<{project: string, wiql: string}> {
        const workItemFormService = await WorkItemFormService.getService();
        const fieldValues = await workItemFormService.getFieldValues(fieldsToSeek, true);
        const witId = await workItemFormService.getId();
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
       
        // Generate fields to retrieve part
        const fieldsToRetrieveString = Constants.DEFAULT_FIELDS_TO_RETRIEVE.map(fieldRefName => `[${fieldRefName}]`).join(",");

        // Generate fields to seek part
        const fieldsToSeekString = fieldsToSeek.map(fieldRefName => {
            const fieldValue = fieldValues[fieldRefName] == null ? "" : fieldValues[fieldRefName];
            if (StringUtils.equals(fieldRefName, "System.Tags", true)) {
                if (fieldValue) {
                    let tagStr = fieldValue.toString().split(";").map(v => {
                        return `[System.Tags] CONTAINS '${v}'`;
                    }).join(" OR ");

                    return `(${tagStr})`;
                }                
            }
            else if (Constants.ExcludedFields.indexOf(fieldRefName) === -1) {
                if (fieldValue !== "" && fieldValue != null) {
                    if (StringUtils.equals(typeof(fieldValue), "string", true)) {
                        return `[${fieldRefName}] = '${fieldValue}'`;
                    }
                    else {
                        return `[${fieldRefName}] = ${fieldValue}`;
                    }
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

    private async _initializeSettings() {
        const workItemFormService = await WorkItemFormService.getService();
        const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
        let settings = await ExtensionDataManager.readUserSetting<Settings>(`${Constants.StorageKey}_${project}_${workItemType}`, Constants.DEFAULT_SETTINGS, true);
        if (settings.top == null || settings.top <= 0) {
            settings.top = Constants.DEFAULT_RESULT_SIZE;
        }

        this.setState({settings: settings});
    }

    private async _initializeWorkItemRelationTypes() {
        const workItemFormService = await WorkItemFormService.getService();
        const relationTypes = await workItemFormService.getWorkItemRelationTypes();
        this.setState({relationTypes: relationTypes});
    }

    private async _initializeLinksData() {
        const workItemFormService = await WorkItemFormService.getService();
        const relations = await workItemFormService.getWorkItemRelations();

        let relationsMap = {};
        relations.forEach(relation => {
            relationsMap[`${relation.url}_${relation.rel}`] = true;
        });

        this.setState({relationsMap: relationsMap});
    }

    @autobind
    private _onFilterChange(filterState: IFilterState) {
        RelatedWorkItemsActions.applyFilter(filterState);
    }

    @autobind
    private _onSortChange(_ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
        if (column.key !== "linked") {
            RelatedWorkItemsActions.applySort({
                sortKey: column.key as WorkItemFieldNames,
                isSortedDescending: !column.isSortedDescending
            });
        }        
    }

    @autobind
    private _focusFilterBar(ev: KeyboardEvent) {
        if (this._filterBar && ev.ctrlKey && ev.shiftKey && StringUtils.equals(ev.key, "f", true)) {
            this._filterBar.focus();
        }
    }

    @autobind
    private _resolveFilterBar(filterBar: IFilterBar) {
        this._filterBar = filterBar;
    }
}

export function init() {
    initializeIcons();

    const container = document.getElementById("ext-container");
    ReactDOM.render(<RelatedWits />, container);
}