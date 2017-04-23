import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemType, Wiql} from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { Loading } from "VSTS_Extension/components/Loading";
import { MessagePanel, MessageType } from "VSTS_Extension/components/MessagePanel";
import { ExtensionDataManager } from "VSTS_Extension/utilities/ExtensionDataManager";

import { Settings, Constants } from "./Models";
import { WorkItemsGrid } from "./WorkItemsGrid";
import { SettingsPanel } from "./SettingsPanel";

interface IRelatedWitsState {
    loading: boolean;
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: Settings;
    settingsPanelOpen?: boolean;
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {

    constructor(props: void, context?: any) {
        super(props, context);        

        this.state = {
            loading: true,
            relationsMap: null
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
                        <SettingsPanel 
                            settings={this.state.settings} 
                            onSave={(settings: Settings) => {
                                this._updateState({settings: settings});
                                this._refreshList();
                            }} />
                    }
                    <WorkItemsViewer />;
                </Fabric>
            );
        }        
    }

    private _updateState(updatedStates: any) {
        this.setState({...this.state, ...updatedStates});
    }

    private async _refreshList(): Promise<void> {
        this._updateState({isWorkItemLoaded: true, isNew: false, loading: true});

        if (!this.state.settings) {
            await this._initializeSettings();
        }

        this._updateState({isWorkItemLoaded: true, isNew: false, loading: false});
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
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}