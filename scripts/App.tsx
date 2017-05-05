import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemRelationType, WorkItemRelation, Wiql} from "TFS/WorkItemTracking/Contracts";
import Utils_String = require("VSS/Utils/String");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { IContextualMenuItem } from "OfficeFabric/ContextualMenu";
import { Panel, PanelType } from "OfficeFabric/Panel";
import { MessageBar, MessageBarType } from 'OfficeFabric/MessageBar';

import { InfoLabel } from "VSTS_Extension/Components/Common/InfoLabel";
import { Loading } from "VSTS_Extension/Components/Common/Loading";
import { BaseComponent, IBaseComponentProps, IBaseComponentState } from "VSTS_Extension/Components/Common/BaseComponent";
import { QueryResultGrid } from "VSTS_Extension/Components/Grids/WorkItemGrid/QueryResultGrid";
import { ColumnPosition} from "VSTS_Extension/Components/Grids/WorkItemGrid/WorkItemGrid.Props";
import { ExtensionDataManager } from "VSTS_Extension/Utilities/ExtensionDataManager";

import { Settings, Constants } from "./Models";
import { SettingsPanel } from "./SettingsPanel";

interface IRelatedWitsState extends IBaseComponentState {
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    query?: {wiql: string, project: string};
    settings?: Settings;
    settingsPanelOpen?: boolean;
    relationsMap?: IDictionaryStringTo<boolean>;
    relationTypes?: WorkItemRelationType[];
}

export class RelatedWits extends BaseComponent<IBaseComponentProps, IRelatedWitsState> {        
    protected initializeState(): void {
        this.state = {
            isWorkItemLoaded: false,
            query: null
        }
    }
    
    protected initialize() {        
        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this.updateState({isWorkItemLoaded: true, isNew: true});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (args: IWorkItemChangedArgs) => {
                this.updateState({isWorkItemLoaded: false});
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
            return <MessageBar messageBarType={MessageBarType.info}>Please save the workitem to get the list of related work items.</MessageBar>;
        }
        else if (!this.state.query) {
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
                            onDismiss={() => this.updateState({settingsPanelOpen: false})}>

                            <SettingsPanel 
                                settings={this.state.settings} 
                                onSave={(settings: Settings) => {
                                    this.updateState({settings: settings, settingsPanelOpen: false});
                                    this._refreshList();
                                }} />
                        
                        </Panel>
                    }
                    <QueryResultGrid
                        className="related-workitems-list"
                        wiql={this.state.query.wiql}
                        project={this.state.query.project}
                        top={this.state.settings.top}
                        extraColumns={
                            [
                                {
                                    position: ColumnPosition.FarLeft,
                                    column: {
                                        key: "Link status",
                                        name: "Linked",
                                        minWidth: 60,
                                        maxWidth: 100,
                                        resizable: false,
                                        onRenderCell: this._renderStatusColumnCell
                                    }
                                }
                            ]
                        }
                        commandBarProps={{
                            menuItems:
                                [
                                    {
                                        key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                                        disabled: this.state.settings == null,
                                        onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                            this.updateState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
                                        }
                                    }
                                ]
                        }}
                        contextMenuProps={{
                            menuItems: this._getContextMenuItemsCallback
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
            else {
                return <InfoLabel label="Not linked" info="This workitem is not linked to the current work item. You can add a link to this workitem by right clicking on the row" />;
            }
        }
        return null;
    }
    
    @autobind
    private _getContextMenuItemsCallback(items: WorkItem[]): IContextualMenuItem[] {
        if (this.state.relationTypes && this.state.relationsMap) {
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
        }

        return [];
    }

    private async _refreshList(): Promise<void> {
        this.updateState({isWorkItemLoaded: true, isNew: false, query: null});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        if (!this.state.relationTypes) {
            this._initializeWorkItemRelationTypes();
        }

        const wiql = await this._createWiql(this.state.settings.fields, this.state.settings.sortByField);
        this.updateState({query: wiql});

        this._initializeLinksData();
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

        this.updateState({settings: settings});
    }

    private async _initializeWorkItemRelationTypes() {
        const workItemFormService = await WorkItemFormService.getService();
        const relationTypes = await workItemFormService.getWorkItemRelationTypes();
        this.updateState({relationTypes: relationTypes});
    }

    private async _initializeLinksData() {
        const workItemFormService = await WorkItemFormService.getService();
        const relations = await workItemFormService.getWorkItemRelations();

        let relationsMap = {};
        relations.forEach(relation => {
            relationsMap[`${relation.url}_${relation.rel}`] = true;
        });

        this.updateState({relationsMap: relationsMap});
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}