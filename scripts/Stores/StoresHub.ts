import { WorkItemFieldStore } from "./WorkItemFieldStore";
import { WorkItemTemplateStore } from "./WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "./WorkItemTemplateItemStore";
import { WorkItemTypeStore } from "./WorkItemTypeStore";
import { WorkItemColorStore } from "./WorkItemColorStore";
import { ActionsHub } from "../Actions/ActionsCreator";

export class StoresHub {
    public workItemFieldStore: WorkItemFieldStore;
    public workItemTemplateStore: WorkItemTemplateStore;
    public workItemTemplateItemStore: WorkItemTemplateItemStore;
    public workItemTypeStore: WorkItemTypeStore;
    public workItemColorStore: WorkItemColorStore;

    constructor(actions: ActionsHub) {
        this.workItemFieldStore = new WorkItemFieldStore(actions);
        this.workItemTemplateStore = new WorkItemTemplateStore(actions);
        this.workItemTypeStore = new WorkItemTypeStore(actions);
        this.workItemTemplateItemStore = new WorkItemTemplateItemStore(actions);
        this.workItemColorStore = new WorkItemColorStore(actions);
    }
}
