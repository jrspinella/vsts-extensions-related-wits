import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemRelationType, WorkItemRelation, Wiql, WorkItemQueryResult, WorkItemField} from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { IContextualMenuItem } from "OfficeFabric/ContextualMenu";
import { Panel, PanelType } from "OfficeFabric/Panel";

import { FluxContext } from "VSTS_Extension/Flux/FluxContext";
import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { WorkItemsGrid } from "VSTS_Extension/Components/WorkItemsGrid/WorkItemsGrid";
import { ColumnPosition} from "VSTS_Extension/Components/WorkItemsGrid/WorkItemsGrid.Props";
import { MessagePanel, MessageType } from "VSTS_Extension/Components/Common/MessagePanel";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { Settings, Constants } from "./Models";
import { SettingsPanel } from "./SettingsPanel";

interface IRelatedWitsState {
    areResultsLoaded?: boolean;
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: Settings;
    settingsPanelOpen?: boolean;
    items?: WorkItem[];
    fieldsMap?: IDictionaryStringTo<WorkItemField>;
    relationsMap?: IDictionaryStringTo<boolean>;
    relationTypes?: WorkItemRelationType[];
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {
    private _context: FluxContext;

    constructor(props: void, context?: any) {
        super(props, context);        
        this._context = FluxContext.get();

        this.state = {
            isWorkItemLoaded: false
        } as IRelatedWitsState;
    }

    public componentWillUnmount() {
        this._context.stores.workItemFieldStore.removeChangedListener(this._onStoreChanged);
    }

    public componentDidMount() {
        this._context.stores.workItemFieldStore.addChangedListener(this._onStoreChanged);
        this._context.actionsCreator.initializeWorkItemFields();

        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this._updateState({isWorkItemLoaded: true, isNew: true, areResultsLoaded: false});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (args: IWorkItemChangedArgs) => {
                this._updateState({isWorkItemLoaded: false, items: []});
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
        else if (!this._isDataLoaded()) {
            return <Loading />;
        }        
        else {                    
            return (
                <Fabric className="fabric-container">                    
                    { 
                        this.state.settingsPanelOpen && 
                        <Panel
                            isOpen={true}
                            type={PanelType.smallFixedFar}
                            isLightDismiss={true} 
                            onDismiss={() => this._updateState({settingsPanelOpen: false})}>

                            <SettingsPanel 
                                settings={this.state.settings} 
                                onSave={(settings: Settings) => {
                                    this._updateState({settings: settings, settingsPanelOpen: false});
                                    this._refreshList();
                                }} />
                        
                        </Panel>
                    }
                    <WorkItemsGrid 
                        items={this.state.items}
                        refreshWorkItems={() => this._getWorkItems(this.state.settings.fields, this.state.settings.sortByField)}
                        fieldColumns={Constants.DEFAULT_FIELDS_TO_RETRIEVE.map(fr => this.state.fieldsMap[fr.toLowerCase()]).filter(f => f != null)}
                        contextMenuProps={{
                            extraContextMenuItems: this._getContextMenuItemsCallback()
                        }}
                        commandBarProps={{
                            extraCommandMenuItems:
                                [
                                    {
                                        key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                                        disabled: this.state.settings == null,
                                        onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                            this._updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                                        }
                                    }
                                ]                            
                        }}
                        columnsProps={{
                            extraColumns: [
                                {
                                    key: "Linked",
                                    name: "Linked",
                                    renderer: this.state.relationTypes && this.state.relationsMap ? this._renderStatusColumnCell : null,
                                    position: ColumnPosition.FarLeft
                                }
                            ]
                        }}
                    />
                        
                </Fabric>
            );
        }        
    }

    @autobind
    private _renderStatusColumnCell(item: WorkItem, index?: number): JSX.Element {
        if (this.state.relationTypes && this.state.relationsMap) {
            let availableLinks: string[] = [];
            this.state.relationTypes.forEach(r => {
                if (this.state.relationsMap[`${item.url}_${r.referenceName}`]) {
                    availableLinks.push(r.name);
                }
            });

            if (availableLinks.length > 0) {
                return <InfoLabel label="Linked" info={`Linked to this workitem as ${availableLinks.join("; ")}`} />;
            }
        }
        return null;
    }
    
    private _getContextMenuItemsCallback(): (items: WorkItem[]) => IContextualMenuItem[] {
        let addLinkContextMenuItem: (items: WorkItem[]) => IContextualMenuItem[];

        if (this.state.relationTypes && this.state.relationsMap) {
            addLinkContextMenuItem = (items: WorkItem[]) => { 
                return [{
                    key: "add-link", name: "Add Link", title: "Add as a link to the current workitem", iconProps: {iconName: "Link"}, 
                    items: this.state.relationTypes.filter(r => r.name != null && r.name.trim() !== "").map(relationType => {
                        return {
                            key: relationType.referenceName,
                            name: relationType.name,
                            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                const workItemFormService = await WorkItemFormService.getService();
                                let workItemRelations = items.filter(wi => !this.state.relationsMap[`${wi.url}_${relationType.referenceName}`]).map(w => {
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
                    }),
                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                        
                    }
                }]
            };
        }

        return addLinkContextMenuItem;
    }

    private _updateState(updatedStates: IRelatedWitsState) {
        this.setState({...this.state, ...updatedStates});
    }

    private async _refreshList(): Promise<void> {
        this._updateState({isWorkItemLoaded: true, isNew: false, areResultsLoaded: false});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        if (!this.state.relationTypes) {
            this._initializeWorkItemRelationTypes();
        }

        const items = await this._getWorkItems(this.state.settings.fields, this.state.settings.sortByField);
        this._updateState({areResultsLoaded: true, items: items});

        this._initializeLinksData();
    }

    private async _getWorkItems(fieldsToSeek: string[], sortByField: string): Promise<WorkItem[]> {
        let {project, wiql} = await this._createWiql(fieldsToSeek, sortByField);
        let queryResults = await WitClient.getClient().queryByWiql({ query: wiql }, project, null, false, this.state.settings.top);
        return this._readWorkItemsFromQueryResults(queryResults);
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

    private async _initializeSettings() {
        const workItemFormService = await WorkItemFormService.getService();
        const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
        const project = await workItemFormService.getFieldValue("System.TeamProject") as string;
        let settings = await ExtensionDataManager.readUserSetting<Settings>(`${Constants.StorageKey}_${workItemType}`, Constants.DEFAULT_SETTINGS, true);
        if (settings.top == null || settings.top <= 0) {
            settings.top = Constants.DEFAULT_RESULT_SIZE;
        }

        this._updateState({settings: settings});
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
            relationsMap[`${relation.url}_${relation.rel}`] = true;
        });

        this._updateState({relationsMap: relationsMap});
    }

    @autobind
    private _onStoreChanged() {
         if (!this.state.fieldsMap && this._context.stores.workItemFieldStore.isLoaded()) {
            const fields = this._context.stores.workItemFieldStore.getAll();
            let fieldsMap = {};
            fields.forEach(f => fieldsMap[f.referenceName.toLowerCase()] = f);

            this._updateState({fieldsMap: fieldsMap});
         }
    }

    private _isDataLoaded(): boolean {
        return this.state.areResultsLoaded && this.state.fieldsMap != null;
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}