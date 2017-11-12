import { RelatedWorkItemsStore } from "./RelatedWorkItemsStore";

import { StoreFactory } from "VSTS_Extension_Widgets/Flux/Stores/BaseStore";

export namespace StoresHub {
    export const relatedWorkItemsStore: RelatedWorkItemsStore = StoreFactory.getInstance<RelatedWorkItemsStore>(RelatedWorkItemsStore);
}