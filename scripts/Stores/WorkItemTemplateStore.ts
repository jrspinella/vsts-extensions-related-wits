import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";
import { WorkItemTemplateReference } from "TFS/WorkItemTracking/Contracts";

import { ActionsHub } from "../Actions/ActionsCreator";

export interface IWorkItemTemplateStore {
    isLoaded(): boolean;
    itemExists(id: string): boolean;
    getItem(id: string): WorkItemTemplateReference;
    getAll(): WorkItemTemplateReference[];
}

export class WorkItemTemplateStore extends Store implements IWorkItemTemplateStore {
    private _items: WorkItemTemplateReference[];

    constructor(actions: ActionsHub) {
        super();

        this._items = null;

        actions.InitializeWorkItemTemplates.addListener((items: WorkItemTemplateReference[]) => {
            if (!items) {
                this.emitChanged();
            }
            this._onAdd(items);
        });
    }

    public isLoaded(): boolean {
        return this._items ? true : false;
    }

    public itemExists(id: string): boolean {
        return this._getById(id) ? true : false;
    }

    public getItem(id: string): WorkItemTemplateReference {
        return this._getById(id);
    }

    public getAll(): WorkItemTemplateReference[] {
        return this._items || [];
    }

    private _getById(id: string): WorkItemTemplateReference {
        if (!this.isLoaded()) {
            return null;
        }

        return Utils_Array.first(this._items, (item: WorkItemTemplateReference) => Utils_String.equals(item.id, id, true));
    }

    private _onAdd(items: WorkItemTemplateReference | WorkItemTemplateReference[]): void {
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

    private _addItem(item: WorkItemTemplateReference): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: WorkItemTemplateReference) => Utils_String.equals(item.id, existingItem.id, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this._items[existingItemIndex] = item;
        }
        else {
            this._items.push(item);
        }
    }
}