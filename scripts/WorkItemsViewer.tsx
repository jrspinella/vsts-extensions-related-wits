import "../css/WorkItemsViewer.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IconButton } from "OfficeFabric/Button";
import { DetailsList } from "OfficeFabric/DetailsList";
import { DetailsListLayoutMode, IColumn, CheckboxVisibility, ConstrainMode } from "OfficeFabric/components/DetailsList/DetailsList.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { Selection } from "OfficeFabric/utilities/selection/Selection";
import { autobind } from "OfficeFabric/Utilities";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { ContextualMenu } from "OfficeFabric/ContextualMenu";

import Utils_String = require("VSS/Utils/String");
import { WorkItemFormNavigationService, WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemRelationType, WorkItemRelation } from "TFS/WorkItemTracking/Contracts";

import { IdentityView } from "./IdentityView";

interface IWorkItemsViewerProps {
    items: WorkItem[];
    sortColumn: string;
    sortOrder: string;
    workItemTypeColors?: IDictionaryStringTo<{color: string, stateColors: IDictionaryStringTo<string>}>;
    relationsMap: IDictionaryStringTo<boolean>;
    changeSort: (sortColumn: string, sortOrder: string) => void;
    relationTypes: WorkItemRelationType[];
}

interface IWorkItemsViewerState  {
    isContextMenuVisible?: boolean;
    contextMenuTarget?: MouseEvent;
}

export class WorkItemsViewer extends React.Component<IWorkItemsViewerProps, IWorkItemsViewerState> {
    private _selection: Selection;

    constructor(props: IWorkItemsViewerProps, context: any) {
        super(props, context);

        this._selection = new Selection();
        this.state = {
            isContextMenuVisible: false,
            contextMenuTarget: null    
        };
    }

    public render(): JSX.Element {
        return (
            <div className="workitem-list-container">
                <DetailsList 
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    selectionMode={SelectionMode.multiple}
                    isHeaderVisible={true}
                    checkboxVisibility={CheckboxVisibility.onHover}
                    columns={this._getColumns()}
                    onRenderItemColumn={this._onRenderCell}
                    items={this.props.items}
                    className="workitem-list"
                    onItemInvoked={(item: WorkItem, index: number) => {
                        this._openWorkItemDialog(null, item);
                    }}
                    selection={ this._selection }
                    onItemContextMenu={this._showContextMenu}
                    onColumnHeaderClick={this._onColumnHeaderClick}
                />

                { this.state.isContextMenuVisible && (
                    <ContextualMenu
                        className="context-menu"
                        items={this._getContextMenuItems()}
                        target={this.state.contextMenuTarget}
                        shouldFocusOnMount={ true }
                        onDismiss={this._hideContextMenu}
                    />
                )}
            </div>
        );      
    }

    private async _openWorkItemDialog(e: React.MouseEvent<HTMLElement>, item: WorkItem) {
        let newTab = e ? e.ctrlKey : false;
        let workItemNavSvc = await WorkItemFormNavigationService.getService();
        workItemNavSvc.openWorkItem(item.id, newTab);
    }

    @autobind
    private _onColumnHeaderClick(ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
        if (column.fieldName !== "Actions") {
            this.props.changeSort(column.fieldName, column.isSortedDescending ? "asc" : "desc");
        }
    }

    @autobind
    private _showContextMenu(item?: WorkItem, index?: number, e?: MouseEvent) {
        if (!this._selection.isIndexSelected(index)) {
            // if not already selected, unselect every other row and select this one
            this._selection.setAllSelected(false);
            this._selection.setIndexSelected(index, true, true);
        }        
        this.setState({...this.state, contextMenuTarget: e, isContextMenuVisible: true});
    }

    @autobind
    private _hideContextMenu(e?: any) {
        this.setState({...this.state, contextMenuTarget: null, isContextMenuVisible: false});
    }

    private _getColumns(): IColumn[] {
        return [
            {
                fieldName: "Actions",
                key: "Actions",
                name: "",
                minWidth: 50,
                maxWidth: 50,
                isResizable: false,
                isSorted: false,
                isSortedDescending: false
            },
            {
                fieldName: "ID",
                key: "ID",
                name:"ID",
                minWidth: 40,
                maxWidth: 70,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "ID", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.WorkItemType",
                key: "WorkItemType",
                name:"Work Item Type",
                minWidth: 80,
                maxWidth: 100,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.WorkItemType", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.Title",
                key: "Title",
                name:"Title",
                minWidth: 150,
                maxWidth: 300,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.Title", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.State",
                key: "State",
                name:"State",
                minWidth: 70,
                maxWidth: 120,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.State", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.AssignedTo",
                key: "AssignedTo",
                name:"Assigned To",
                minWidth: 100,
                maxWidth: 250,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.AssignedTo", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            },
            {
                fieldName: "System.AreaPath",
                key: "AreaPath",
                name:"Area Path",
                minWidth: 150,
                maxWidth: 350,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.AreaPath", true),
                isSortedDescending: Utils_String.equals(this.props.sortOrder, "desc", true)
            }
        ];
    }    

    @autobind
    private _onRenderCell(item: WorkItem, index?: number, column?: IColumn): React.ReactNode {
        let text: string;
        switch (column.fieldName) {            
            case "ID":
                text = `${item.id}`;
                break;
            case "System.Title":
                let witColor = this.props.workItemTypeColors && 
                            this.props.workItemTypeColors[item.fields["System.WorkItemType"]] && 
                            this.props.workItemTypeColors[item.fields["System.WorkItemType"]].color;
                return (
                    <div className="title-cell" title={item.fields[column.fieldName]}>
                        <span 
                            className="overflow-ellipsis" 
                            onClick={(e) => this._openWorkItemDialog(e, item)}
                            style={{borderColor: witColor ? "#" + witColor : "#000"}}>

                            {item.fields[column.fieldName]}
                        </span>
                    </div>
                );
            case "System.State":
                return (
                    <div className="state-cell" title={item.fields[column.fieldName]}>
                        {
                            this.props.workItemTypeColors &&
                            this.props.workItemTypeColors[item.fields["System.WorkItemType"]] &&
                            this.props.workItemTypeColors[item.fields["System.WorkItemType"]].stateColors &&
                            this.props.workItemTypeColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]] &&
                            <span 
                                className="work-item-type-state-color" 
                                style={{
                                    backgroundColor: "#" + this.props.workItemTypeColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]],
                                    borderColor: "#" + this.props.workItemTypeColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]]
                                }} />
                        }
                        <span className="overflow-ellipsis">{item.fields[column.fieldName]}</span>
                    </div>
                );
            case "System.AssignedTo":
                return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;
            case "Actions":
                let isLinked: boolean = false;
                if (this.props.relationsMap && this.props.relationTypes && !this.props.relationsMap[item.url]) {
                    return (
                        <div className="link-cell">
                            <IconButton 
                                iconProps={{iconName: "Link"}}
                                className="workitem-link-button"
                                title="Add link"
                                menuProps={{
                                    className: "callout-container",
                                    items: this.props.relationTypes.filter(r => r.name != null && r.name.trim() !== "").map(relationType => {
                                        return {
                                            key: relationType.referenceName,
                                            name: relationType.name,
                                            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                                                const workItemFormService = await WorkItemFormService.getService();
                                                let workItemRelation = {
                                                    rel: relationType.referenceName,
                                                    attributes: {
                                                        isLocked: false
                                                    },
                                                    url: item.url
                                                } as WorkItemRelation;
                                                workItemFormService.addWorkItemRelations([workItemRelation]);
                                            }
                                        };
                                    })                                    
                                }}
                            />
                        </div>
                    );
                }
                else {
                    return null;
                }                
            default:
                text = item.fields[column.fieldName];  
                break;          
        }

        return <div className="overflow-ellipsis" title={text}>{text}</div>;
    }    

    private _getContextMenuItems(): IContextualMenuItem[] {
        return [            
            {
                key: "OpenQuery", name: "Open as query", title: "Open selected workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this._selection.getSelectedCount() == 0,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getSelectedWorkItemsWiql())}`;
                    window.open(url, "_blank");
                }
            }
        ];
    }

    private _getSelectedWorkItemsWiql(): string {
        let selectedWorkItems = this._selection.getSelection() as WorkItem[];
        let ids = selectedWorkItems.map((w:WorkItem) => w.id).join(",");

        return `SELECT [System.Id], [System.WorkItemType], [System.Title], [System.State], [System.AssignedTo], [System.AreaPath], [System.Tags]
                 FROM WorkItems 
                 WHERE [System.TeamProject] = @project 
                 AND [System.ID] IN (${ids}) 
                 ORDER BY [System.CreatedDate] DESC`;
    }
}