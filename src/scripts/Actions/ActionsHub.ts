import { ISortState } from "../Models";

import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { IFilterState } from "VSSUI/Utilities/Filter";

import { Action } from "VSTS_Extension_Widgets/Flux/Actions/Action";

export namespace ActionsHub {
    export const Refresh = new Action<WorkItem[]>();
    export const ApplyFilter = new Action<IFilterState>();
    export const ClearSortAndFilter = new Action<void>();
    export const ApplySort = new Action<ISortState>();
    export const Clean = new Action<void>();
    export const UpdateWorkItemInStore = new Action<WorkItem>();
}