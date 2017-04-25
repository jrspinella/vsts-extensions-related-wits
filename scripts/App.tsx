import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemType, Wiql, WorkItemQueryResult, WorkItemField} from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { Panel, PanelType } from "OfficeFabric/Panel";

import { Loading } from "VSTS_Extension/components/Loading";
import { MessagePanel, MessageType } from "VSTS_Extension/components/MessagePanel";
import { ExtensionDataManager } from "VSTS_Extension/utilities/ExtensionDataManager";
import { ActionsCreator, ActionsHub } from "./Actions/ActionsCreator";
import { StoresHub } from "./Stores/StoresHub";

import { Settings, Constants } from "./Models";
import { WorkItemsGrid } from "./Components/WorkItemsGrid";
import { SettingsPanel } from "./SettingsPanel";

interface IRelatedWitsState {
    loading?: boolean;
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: Settings;
    settingsPanelOpen?: boolean;
    items?: WorkItem[];
    fieldsMap?: IDictionaryStringTo<WorkItemField>;
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {

    constructor(props: void, context?: any) {
        super(props, context);        

        this.state = {
            loading: true
        } as IRelatedWitsState;
    }

    public componentDidMount() {
        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this._updateState({isWorkItemLoaded: true, isNew: true, loading: false});
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
        else if (this.state.loading) {
            return <Loading />;
        }
        else if (this.state.isNew) {
            return <MessagePanel message="Please save the workitem to get the list of related work items." messageType={MessageType.Info} />;
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
                        extraCommandMenuItems={
                            [
                                {
                                    key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                                    disabled: this.state.settings == null,
                                    onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                        this._updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                                    }
                                }
                            ]
                        }/>
                </Fabric>
            );
        }        
    }

    private _updateState(updatedStates: IRelatedWitsState) {
        this.setState({...this.state, ...updatedStates});
    }

    private async _refreshList(): Promise<void> {
        this._updateState({isWorkItemLoaded: true, isNew: false, loading: true});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        if (!this.state.fieldsMap) {
            await this._initializeFields();
        }

        const items = await this._getWorkItems(this.state.settings.fields, this.state.settings.sortByField);
        this._updateState({loading: false, items: items});

        this._updateState({isWorkItemLoaded: true, isNew: false, loading: false});
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

    private async _initializeFields() {
        const workItemFormService = await WorkItemFormService.getService();
        const fields = await workItemFormService.getFields();
        let fieldsMap = {};
        fields.forEach(f => fieldsMap[f.referenceName.toLowerCase()] = f);

        this._updateState({fieldsMap: fieldsMap});
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}