import { Action } from "VSS/Flux/Action";
import { WorkItemField, WorkItemTemplateReference, WorkItemStateColor, WorkItemType, WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";

import { IWorkItemFieldStore } from "../Stores/WorkItemFieldStore";
import { IWorkItemTypeStore } from "../Stores/WorkItemTypeStore";
import { IWorkItemTemplateStore } from "../Stores/WorkItemTemplateStore";
import { IWorkItemTemplateItemStore } from "../Stores/WorkItemTemplateItemStore";
import { IWorkItemColorStore } from "../Stores/WorkItemColorStore";

import * as WitClient from "TFS/WorkItemTracking/RestClient";

export class ActionsHub {    
    public InitializeWorkItemFields = new Action<WorkItemField[]>();
    public InitializeWorkItemTemplates = new Action<WorkItemTemplateReference[]>();
    public InitializeWorkItemTypes = new Action<WorkItemType[]>();
    public InitializeWorkItemColors = new Action<IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>>();

    public WorkItemTemplateItemAdded = new Action<WorkItemTemplate | WorkItemTemplate[]>();
}

export class ActionsCreator {
    constructor(
        private _actionsHub: ActionsHub, 
        private _workItemFieldDataProvider: IWorkItemFieldStore,
        private _workItemTemplateDataProvider: IWorkItemTemplateStore,
        private _workItemTypeDataProvider: IWorkItemTypeStore,
        private _workItemTemplateItemDataProvider: IWorkItemTemplateItemStore,
        private _workItemColorsDataProvider: IWorkItemColorStore) {
    }

    public async initializeWorkItemFields() {
        if (this._workItemFieldDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemFields.invoke(null);
        }
        else {            
            let fields = await WitClient.getClient().getFields(VSS.getWebContext().project.id);
            this._actionsHub.InitializeWorkItemFields.invoke(fields);
        }
    }

    public async initializeWorkItemTemplates() {
        if (this._workItemTemplateDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemTemplates.invoke(null);
        }
        else {            
            let templates = await WitClient.getClient().getTemplates(VSS.getWebContext().project.id, VSS.getWebContext().team.id);
            this._actionsHub.InitializeWorkItemTemplates.invoke(templates);
        }
    }

    public async initializeWorkItemTypes() {
        if (this._workItemTypeDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemTypes.invoke(null);
        }
        else {            
            let workItemTypes = await WitClient.getClient().getWorkItemTypes(VSS.getWebContext().project.id);
            this._actionsHub.InitializeWorkItemTypes.invoke(workItemTypes);
        }
    }
    
    public async initializeWorkItemColors() {
        if (this._workItemColorsDataProvider.isLoaded()) {
            // Do nothing if query hierarchy data is already loaded
            this._actionsHub.InitializeWorkItemColors.invoke(null);
        }
        else {
            let workItemTypeAndStateColors: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}> = {};
            const projectId = VSS.getWebContext().project.id;
            
            await this.initializeWorkItemTypes();
            const workItemTypes = this._workItemTypeDataProvider.getAll();
            workItemTypes.forEach((wit: WorkItemType) => workItemTypeAndStateColors[wit.name] = {
                color: wit.color,
                stateColors: {}
            });

            await Promise.all(workItemTypes.map(async (wit: WorkItemType) => {
                let stateColors = await WitClient.getClient().getWorkItemTypeStates(projectId, wit.name);
                stateColors.forEach((stateColor: WorkItemStateColor) => workItemTypeAndStateColors[wit.name].stateColors[stateColor.name] = stateColor.color);
            }));

            this._actionsHub.InitializeWorkItemColors.invoke(workItemTypeAndStateColors);
        }
    }

    public async ensureTemplateItem(id: string): Promise<boolean> {
        if (!this._workItemTemplateItemDataProvider.itemExists(id)) {
            try {
                let template = await WitClient.getClient().getTemplate(VSS.getWebContext().project.id, VSS.getWebContext().team.id, id)
                if (template) {
                    this._actionsHub.WorkItemTemplateItemAdded.invoke(template);
                    return true;
                }
            }
            catch (e) {
                return false;
            }

            return false;
        }
        else {
            return true;
        }
    }
}
