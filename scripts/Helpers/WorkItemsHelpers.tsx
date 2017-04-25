import * as React from "react";

import { WorkItem, WorkItemField, FieldType } from "TFS/WorkItemTracking/Contracts";
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import Utils_String = require("VSS/Utils/String");

import { IColumn } from "OfficeFabric/components/DetailsList/DetailsList.Props";

import { SortOrder } from "../Components/Interfaces/WorkItemsGrid.Props";
import { IdentityView } from "VSTS_Extension/components/IdentityView";

export interface workItemFieldCellRendererOptions {
    workItemTypeAndStateColors?: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;
}

export function workItemFieldValueComparer(w1: WorkItem, w2: WorkItem, fieldRefName: string, sortOrder: SortOrder): number {
    if (Utils_String.equals(fieldRefName, "System.Id", true)) {
        return sortOrder === SortOrder.DESC ? ((w1.id > w2.id) ? -1 : 1) : ((w1.id > w2.id) ? 1 : -1);
    }            
    else {
        let v1 = w1.fields[fieldRefName];
        let v2 = w2.fields[fieldRefName];
        return sortOrder === SortOrder.DESC ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
    }
}

export function workItemFieldCellRenderer(item: WorkItem, index: number, column: IColumn, extraData?: workItemFieldCellRendererOptions): JSX.Element {
    let text: string;
    switch (column.fieldName.toLowerCase()) {
        case "system.id":  
            text = item.id.toString();
            break;
        case "system.title":
            let witColor = extraData && extraData.workItemTypeAndStateColors && 
                        extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]] && 
                        extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].color;
            return (
                <div className={"title-cell"} title={item.fields[column.fieldName]}>
                    <span 
                        className="overflow-ellipsis" 
                        onClick={(e) => openWorkItemDialog(e, item)}
                        style={{borderColor: witColor ? "#" + witColor : "#000"}}>

                        {item.fields[column.fieldName]}
                    </span>
                </div>
            );
        case "system.state":
            return (
                <div className={"state-cell"} title={item.fields[column.fieldName]}>
                    {
                        extraData &&
                        extraData.workItemTypeAndStateColors &&
                        extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]] &&
                        extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors &&
                        extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]] &&
                        <span 
                            className="work-item-type-state-color" 
                            style={{
                                backgroundColor: "#" +extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]],
                                borderColor: "#" + extraData.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]]
                            }} />
                    }
                    <span className="overflow-ellipsis">{item.fields[column.fieldName]}</span>
                </div>
            );
        case "system.assignedto":  // check isidentity flag
            return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;                      
        default:
            text = item.fields[column.fieldName];  
            break;          
    }

    return <div className="overflow-ellipsis" title={text}>{text}</div>;
}

export function getColumnSize(field: WorkItemField): {minWidth: number, maxWidth: number} {
    if (Utils_String.equals(field.referenceName, "System.Id", true)) {
        return { minWidth: 40, maxWidth: 70}
    }
    else if (Utils_String.equals(field.referenceName, "System.WorkItemType", true)) {
        return { minWidth: 80, maxWidth: 100}
    }
    else if (Utils_String.equals(field.referenceName, "System.Title", true)) {
        return { minWidth: 150, maxWidth: 300}
    }
    else if (Utils_String.equals(field.referenceName, "System.State", true)) {
        return { minWidth: 70, maxWidth: 120}
    }
    else if (field.type === FieldType.TreePath) {
        return { minWidth: 150, maxWidth: 350}
    }
    else if (field.type === FieldType.Boolean) {
        return { minWidth: 40, maxWidth: 40}
    }
    else if (field.type === FieldType.DateTime) {
        return { minWidth: 80, maxWidth: 150}
    }
    else if (field.type === FieldType.Double ||
        field.type === FieldType.Integer ||
        field.type === FieldType.PicklistDouble ||
        field.type === FieldType.PicklistInteger) {
        return { minWidth: 50, maxWidth: 100}
    }
    else {
        return { minWidth: 100, maxWidth: 250}
    }
}

export async function openWorkItemDialog(e: React.MouseEvent<HTMLElement>, item: WorkItem): Promise<void> {
    let newTab = e ? e.ctrlKey : false;
    let workItemNavSvc = await WorkItemFormNavigationService.getService();
    workItemNavSvc.openWorkItem(item.id, newTab);
}