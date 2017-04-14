import "../css/settingsPanel.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Loading } from "./Loading";

import { WorkItemField } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

interface ISettingsPanelState {
    loading: boolean;
}

export class SettingsPanel extends React.Component<void, ISettingsPanelState> {
    public render(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }  
        else {
            return (
                <div className="settings-panel">
                    
                </div>
            );
        }        
    }

    private async _getWorkItemFields(): Promise<WorkItemField[]> {
        let workItemFormService = await WorkItemFormService.getService();
        let fields = await workItemFormService.getFields();
        return fields;
    }
}