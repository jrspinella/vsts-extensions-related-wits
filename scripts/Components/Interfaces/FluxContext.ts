import { StoresHub } from "../../Stores/StoresHub";
import { ActionsHub, ActionsCreator } from "../../Actions/ActionsCreator";

export class FluxContext {
    private static _instance: FluxContext;

    actions: ActionsHub;
    stores: StoresHub;
    actionsCreator: ActionsCreator;

    public static get(): FluxContext {
        if (!this._instance) {
            const actionsHub = new ActionsHub();
            const storeHub = new StoresHub(actionsHub);
            const actionsCreator = new ActionsCreator(
                actionsHub,
                storeHub.workItemFieldStore,
                storeHub.workItemTemplateStore,
                storeHub.workItemTypeStore,
                storeHub.workItemTemplateItemStore,
                storeHub.workItemColorStore);
            
            this._instance = {
                actions: actionsHub,
                stores: storeHub,
                actionsCreator: actionsCreator
            };
        }

        return this._instance;
    }
}