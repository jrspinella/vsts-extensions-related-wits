import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { Store } from "VSS/Flux/Store";
import { WorkItemField } from "TFS/WorkItemTracking/Contracts";

import { ActionsHub } from "../Actions/ActionsCreator";

export interface IWorkItemColorStore {
    isLoaded(): boolean;
    itemExists(witName: string, stateName?: string): boolean;
    getWorkItemTypeColor(witName: string): string;
    getStateColor(witName: string, state: string): string;
    getAll(): IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;
}

export class WorkItemColorStore extends Store implements IWorkItemColorStore {
    private _items: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;

    constructor(actions: ActionsHub) {
        super();

        this._items = null;

        actions.InitializeWorkItemColors.addListener((items: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>) => {            
            this._items = items;
            this.emitChanged();
        });
    }

    public isLoaded(): boolean {
        return this._items ? true : false;
    }

    public itemExists(witName: string, stateName?: string): boolean {
        if (!this.isLoaded()) {
            return false;
        }

        let workItemType = this._items[witName];
        return workItemType != null && (!stateName || workItemType.stateColors[stateName] != null);
    }

    public getWorkItemTypeColor(witName: string): string {
        if (!this.isLoaded() || !this.itemExists(witName)) {
            return null;
        }

        return this._items[witName].color;
    }

     public getStateColor(witName: string, state: string): string {
        if (!this.isLoaded() || !this.itemExists(witName, state)) {
            return null;
        }

        return this._items[witName].stateColors[state];
    }

    public getAll(): IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}> {
        return this._items || {};
    }
}