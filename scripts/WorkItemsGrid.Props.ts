import { WorkItem, WorkItemField } from "TFS/WorkItemTracking/Contracts";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";

export interface IWorkItemsGridProps {    
    fieldColumns: WorkItemField[];
    items: WorkItem[];
    refreshWorkItems?: () => Promise<WorkItem[]>;
    sortColumn?: string;
    sortOrder?: string;
    disableSort?: boolean;
    disableColumnResize?: boolean;
    pageSize?: number;
    hideSearchBox?: boolean;
    hideCommandBar?: boolean;
    disableContextMenu?: boolean;
    extraCommandMenuItems?: IContextualMenuItem[];
    farCommandMenuItems?: IContextualMenuItem[];
    extraContextMenuItems?: IContextualMenuItem[];
    onItemInvoked?: (workItem: WorkItem, index: number) => void;
    selectionMode?: SelectionMode;
}

export interface IWorkItemsGridState  {    
    filteredItems?: WorkItem[];
    items?: WorkItem[];
    isContextMenuVisible?: boolean;
    contextMenuTarget?: MouseEvent;
    workItemTypeAndStateColors?: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;
    sortColumn?: string;
    sortOrder?: string;
    filterText?: string;
}

