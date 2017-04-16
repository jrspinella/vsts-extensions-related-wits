import "../css/settingsPanel.scss";

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

import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import { WorkItem } from "TFS/WorkItemTracking/Contracts";

import { IdentityView } from "./IdentityView";

interface IWorkItemsViewerProps {
    items: IListItem[];
    sortColumn: string;
    sortOrder: string;
    changeSort: (sortColumn: string, sortOrder: string) => void;
}

interface IWorkItemsViewerState  {
    isContextMenuVisible?: boolean;
    contextMenuTarget?: MouseEvent;
}

export interface IListItem {
    workItem: WorkItem;
    isLinked: boolean;    
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

    public componentWillReceiveProps(nextProps: IWorkItemsViewerProps) {
        
    }

    public render(): JSX.Element {
        return (
            <div className="results-view-contents">
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
                minWidth: 40,
                maxWidth: 40,
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
                minWidth: 100,
                maxWidth: 150,
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
            },
            {
                fieldName: "System.Tags",
                key: "Tags",
                name:"Tags",
                minWidth: 150,
                maxWidth: 350,
                isResizable: true,
                isSorted: Utils_String.equals(this.props.sortColumn, "System.Tags", true),
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
                return <span className="title-cell overflow-ellipsis" onClick={(e) => this._openWorkItemDialog(e, item)} title={item.fields[column.fieldName]}>{item.fields[column.fieldName]}</span>
            case "System.CreatedDate":
                text = Utils_Date.friendly(new Date(item.fields["System.CreatedDate"]));
                break;
            case "System.CreatedBy":
            case "System.AssignedTo":
                return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;
            case "Actions":                
                return (
                    <div className="workitem-row-actions-cell">
                        <IconButton icon="Link" className="workitem-link-button" title="Add a link" />                        
                    </div>
                );
            default:
                text = item.fields[column.fieldName];  
                break;          
        }

        return <div className="overflow-ellipsis" title={text}>{text}</div>;
    }    


    private _getContextMenuItems(): IContextualMenuItem[] {
        return [
            {
                key: "Delete", name: "Delete", title: "Delete selected workitems", iconProps: {iconName: "Delete"}, 
                disabled: this._selection.getSelectedCount() == 0,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    
                }
            },
            {
                key: "OpenQuery", name: "Open as query", title: "Open selected workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this._selection.getSelectedCount() == 0,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    let url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getSelectedWorkItemsWiql())}`;
                    window.open(url, "_parent");
                }
            }
        ];
    }

    private _getSelectedWorkItemsWiql(): string {
        let selectedWorkItems = this._selection.getSelection() as WorkItem[];
        let ids = selectedWorkItems.map((w:WorkItem) => w.id).join(",");

        return `SELECT [System.Id], [System.Title], [System.CreatedBy], [System.CreatedDate], [System.State], [System.AssignedTo], [System.AreaPath]
                 FROM WorkItems 
                 WHERE [System.TeamProject] = @project 
                 AND [System.ID] IN (${ids}) 
                 ORDER BY [System.CreatedDate] DESC`;
    }
}