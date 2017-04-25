import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";
import { WorkItemField } from "TFS/WorkItemTracking/Contracts";

import { ActionsHub } from "../Actions/ActionsCreator";

export interface IWorkItemFieldStore {
    isLoaded(): boolean;
    itemExists(refName: string): boolean;
    getItem(refName: string): WorkItemField;
    getAll(): WorkItemField[];
}

export class WorkItemFieldStore extends Store implements IWorkItemFieldStore {
    private _items: WorkItemField[];

    constructor(actions: ActionsHub) {
        super();

        this._items = null;

        actions.InitializeWorkItemFields.addListener((items: WorkItemField[]) => {
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
        return this._getByRefName(id) ? true : false;
    }

    public getItem(refName: string): WorkItemField {
        return this._getByRefName(refName);
    }

    public getAll(): WorkItemField[] {
        return this._items || [];
    }

    private _getByRefName(refName: string): WorkItemField {
        if (!this.isLoaded()) {
            return null;
        }

        return Utils_Array.first(this._items, (item: WorkItemField) => Utils_String.equals(item.referenceName, refName, true));
    }

    private _onAdd(items: WorkItemField | WorkItemField[]): void {
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

    private _addItem(item: WorkItemField): void {
        let existingItemIndex = Utils_Array.findIndex(this._items, (existingItem: WorkItemField) => Utils_String.equals(item.referenceName, existingItem.referenceName, true));
        if (existingItemIndex != -1) {
            // Overwrite the item data
            this._items[existingItemIndex] = item;
        }
        else {
            this._items.push(item);
        }
    }
}