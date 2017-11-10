import "./App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Constants, Settings } from "./Models";
import * as SettingsPanel_Async from "./SettingsPanel";

import { IContextualMenuItem } from "OfficeFabric/ContextualMenu";
import { Fabric } from "OfficeFabric/Fabric";
import { MessageBar, MessageBarType } from "OfficeFabric/MessageBar";
import { Panel, PanelType } from "OfficeFabric/Panel";
import { autobind } from "OfficeFabric/Utilities";

import {
    Wiql, WorkItem, WorkItemRelation, WorkItemRelationType
} from "TFS/WorkItemTracking/Contracts";
import {
    IWorkItemChangedArgs, IWorkItemLoadedArgs, IWorkItemNotificationListener
} from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

import { InfoLabel, Loading } from "VSTS_Extension_Widgets/Components";
import {
    BaseFluxComponent, getAsyncLoadedComponent, IBaseFluxComponentProps, IBaseFluxComponentState
} from "VSTS_Extension_Widgets/Components/Utilities";
import { ExtensionDataManager, StringUtils } from "VSTS_Extension_Widgets/Utilities";

const AsyncSettingsPanel = getAsyncLoadedComponent(
    ["scripts/SettingsPanel"],
    (m: typeof SettingsPanel_Async) => m.SettingsPanel,
    () => <Loading />);

interface IRelatedWitsState extends IBaseFluxComponentState {
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    query?: {wiql: string, project: string};
    settings?: Settings;
    settingsPanelOpen?: boolean;
    relationsMap?: IDictionaryStringTo<boolean>;
    relationTypes?: WorkItemRelationType[];
}

export class RelatedWits extends BaseFluxComponent<IBaseFluxComponentProps, IRelatedWitsState> {
    protected initializeState(): void {
        this.state = {
            isWorkItemLoaded: false,
            query: null
        }
    }
    
    public componentDidMount() {
        super.componentDidMount();
        
        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this.setState({isWorkItemLoaded: true, isNew: true});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (_args: IWorkItemChangedArgs) => {
                this.setState({isWorkItemLoaded: false});
            },
            onSaved: (_args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onRefreshed: (_args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onReset: (_args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
        } as IWorkItemNotificationListener);
    }

    public componentWillUnmount() {
        super.componentWillUnmount();
        VSS.unregister(VSS.getContribution().id);
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
                    <Panel
                        isOpen={this.state.settingsPanelOpen}
                        type={PanelType.smallFixedFar}
                        isLightDismiss={true} 
                        onDismiss={() => this.setState({settingsPanelOpen: false})}>

                        <AsyncSettingsPanel
                            settings={this.state.settings} 
                            onSave={(settings: Settings) => {
                                this.setState({settings: settings, settingsPanelOpen: false});
                                this._refreshList();
                            }} />
                    </Panel>
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
                                            this.setState({settingsPanelOpen: !(this.state.settingsPanelOpen)});
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
                return <div className="linked-cell"><InfoLabel label="Linked" info={`Linked to this workitem as ${availableLinks.join("; ")}`} /></div>;
            }
            else {
                return <div className="unlinked-cell"><InfoLabel label="Not linked" info="This workitem is not linked to the current work item. You can add a link to this workitem by right clicking on the row" /></div>;
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
        this.setState({isWorkItemLoaded: true, isNew: false, query: null});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        if (!this.state.relationTypes) {
            this._initializeWorkItemRelationTypes();
        }

        const wiql = await this._createWiql(this.state.settings.fields, this.state.settings.sortByField);
        this.setState({query: wiql});

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
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}