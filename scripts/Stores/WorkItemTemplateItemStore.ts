import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";
import { WorkItemTemplate } from "TFS/WorkItemTracking/Contracts";

import { ActionsHub } from "../Actions/ActionsCreator";

export interface IWorkItemTemplateItemStore {
    itemExists(id: string): boolean;
    getItem(id: string): WorkItemTemplate;
}

export class WorkItemTemplateItemStore extends Store implements IWorkItemTemplateItemStore {
    private _items: WorkItemTemplate[];

    constructor(actions: ActionsHub) {
        super();

        this._items = [];

        actions.WorkItemTemplateItemAdded.addListener((items: WorkItemTemplate | WorkItemTemplate[]) => {
            this._onAdd(items);
        });
    }

    public itemExists(id: string): boolean {
        return this._getById(id) ? true : false;
    }

    public getItem(id: string): WorkItemTemplate {
        return this._getById(id);
    }

    private _getById(id: string): WorkItemTemplate {
         return Utils_Array.first(this._items, (item: WorkItemTemplate) => Utils_String.equals(item.id, id, true));
    }

    private _onAdd(items: WorkItemTemplate | WorkItemTemplate[]): void {
        if (!items) {
            return;
        }

        if (!this._items) {
            this._items = [];
        }

        if (Array.isArray(items)) {
            for (let item of items) {
                this._addItem(item);
            }
        }
        else {
            this._addItem(items);
        }

        this.emitChanged();
    }

    private _addItem(item: WorkItemTemplate): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: WorkItemTemplate) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this._items[existingItemIndex] = item;
        }
        else {
            this._items.push(item);
        }
    }
}