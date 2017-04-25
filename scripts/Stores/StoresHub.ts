import { WorkItemFieldStore } from "./WorkItemFieldStore";
import { WorkItemTemplateStore } from "./WorkItemTemplateStore";
import { WorkItemTemplateItemStore } from "./WorkItemTemplateItemStore";
import { WorkItemTypeStore } from "./WorkItemTypeStore";
import { ActionsHub } from "../Actions/ActionsCreator";

export class StoresHub {
    public workItemFieldStore: WorkItemFieldStore;
    public workItemTemplateStore: WorkItemTemplateStore;
    public workItemTemplateItemStore: WorkItemTemplateItemStore;
    public workItemTypeStore: WorkItemTypeStore;

    constructor(actions: ActionsHub) {
        this.workItemFieldStore = new WorkItemFieldStore(actions);
        this.workItemTemplateStore = new WorkItemTemplateStore(actions);
        this.workItemTypeStore = new WorkItemTypeStore(actions);
        this.workItemTemplateItemStore = new WorkItemTemplateItemStore(actions);
    }
}
