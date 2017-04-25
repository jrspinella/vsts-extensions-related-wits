import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";
import { WorkItemType } from "TFS/WorkItemTracking/Contracts";

import { ActionsHub } from "../Actions/ActionsCreator";

export interface IWorkItemTypeStore {
    isLoaded(): boolean;
    itemExists(typeName: string): boolean;
    getItem(typeName: string): WorkItemType;
    getAll(): WorkItemType[];
}

export class WorkItemTypeStore extends Store implements IWorkItemTypeStore {
    private _items: WorkItemType[];

    constructor(actions: ActionsHub) {
        super();

        this._items = null;

        actions.InitializeWorkItemTypes.addListener((items: WorkItemType[]) => {
            if (!items) {
                this.emitChanged();
            }
            this._onAdd(items);
        });
    }

    public isLoaded(): boolean {
        return this._items ? true : false;
    }

    public itemExists(typeName: string): boolean {
        return this._getByTypeName(typeName) ? true : false;
    }

    public getItem(typeName: string): WorkItemType {
        return this._getByTypeName(typeName);
    }

    public getAll(): WorkItemType[] {
        return this._items || [];
    }

    private _getByTypeName(typeName: string): WorkItemType {
        if (!this.isLoaded()) {
            return null;
        }

        return Utils_Array.first(this._items, (item: WorkItemType) => Utils_String.equals(item.name, typeName, true));
    }

    private _onAdd(items: WorkItemType | WorkItemType[]): void {
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

    private _addItem(item: WorkItemType): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: WorkItemType) => Utils_String.equals(item.name, existingItem.name, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this._items[existingItemIndex] = item;
        }
        else {
            this._items.push(item);
        }
    }
}