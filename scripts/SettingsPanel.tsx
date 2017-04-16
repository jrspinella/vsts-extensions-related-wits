import "../css/SettingsPanel.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Loading } from "./Loading";
import { UserPreferenceModel, Constants } from "./Models";

import { WorkItemField } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";

interface ISettingsPanelProps {
    settings: UserPreferenceModel
}

interface ISettingsPanelState {
    loading: boolean;
}

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
    public render(): JSX.Element {
        return (
            <div className="settings-panel">
                
            </div>
        );    
    }

    private async _getWorkItemFields(): Promise<WorkItemField[]> {
        let workItemFormService = await WorkItemFormService.getService();
        let fields = await workItemFormService.getFields();
        return fields;
    }
}