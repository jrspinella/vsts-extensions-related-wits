import "../css/WorkItemsViewer.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import {IWorkItemsGridProps, IWorkItemsGridState} from "./WorkItemsGrid.Props";

import { DetailsList } from "OfficeFabric/DetailsList";
import { DetailsListLayoutMode, IColumn, CheckboxVisibility, ConstrainMode } from "OfficeFabric/components/DetailsList/DetailsList.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { Selection } from "OfficeFabric/utilities/selection/Selection";
import { autobind } from "OfficeFabric/Utilities";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { ContextualMenu } from "OfficeFabric/ContextualMenu";
import { CommandBar } from "OfficeFabric/CommandBar";
import { SearchBox } from "OfficeFabric/SearchBox";

import { IdentityView } from "VSTS_Extension/components/IdentityView";
import { MessagePanel, MessageType } from "VSTS_Extension/components/MessagePanel";

import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemField, FieldType } from "TFS/WorkItemTracking/Contracts";

function getColumnSize(field: WorkItemField): {minWidth: number, maxWidth: number} {
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

export class WorkItemsGrid extends React.Component<IWorkItemsGridProps, IWorkItemsGridState> {
    private _selection: Selection;

    constructor(props: IWorkItemsGridProps, context: any) {
        super(props, context);
        this._selection = new Selection();
        this._initializeState();
    }

    private _initializeState() {
        this.state = {
            filteredItems: this.props.items || [],
            items: this.props.items || [],
            sortColumn: this.props.sortColumn,
            sortOrder: this.props.sortOrder,
            filterText: "",
            workItemTypeAndStateColors: {}
        };
    }

    public componentDidMount() {
        
    }

    public componentWillUnmount() {
        
    }

    public render(): JSX.Element {
        return (
            <div className={this._getClassName()}>
                {this._renderCommandBar()}
                {this._renderWorkItemGrid()}
                {this.state.isContextMenuVisible && (
                    <ContextualMenu
                        className={this._getClassName("context-menu")}
                        items={this._getContextMenuItems()}
                        target={this.state.contextMenuTarget}
                        shouldFocusOnMount={ true }
                        onDismiss={this._hideContextMenu}
                    />
                )}
            </div>
        );
    }

    private _renderWorkItemGrid(): JSX.Element {
        if (this.state.filteredItems.length === 0) {
            return <MessagePanel message="No results" messageType={MessageType.Info} />;
        }
        else {
            const selectionMode = this.props.selectionMode || SelectionMode.multiple;
            return <DetailsList 
                    layoutMode={DetailsListLayoutMode.justified}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    selectionMode={selectionMode}
                    isHeaderVisible={true}
                    checkboxVisibility={selectionMode === SelectionMode.none ? CheckboxVisibility.hidden : CheckboxVisibility.onHover}
                    columns={this._getColumns()}
                    onRenderItemColumn={this._onRenderCell}
                    items={this.state.filteredItems}
                    className={this._getClassName("grid")}
                    onItemInvoked={(item: WorkItem, index: number) => {
                        if (this.props.onItemInvoked) {
                            this.props.onItemInvoked(item, index);
                        }
                        else {
                            this._openWorkItemDialog(null, item);
                        }                        
                    }}
                    selection={ this._selection }
                    onItemContextMenu={this._showContextMenu}
                    onColumnHeaderClick={this._onColumnHeaderClick}
                />;
        }
    }

    private _getColumns(): IColumn[] {
        return this.props.fieldColumns.map(f => {
            const columnSize = getColumnSize(f);
            return {
                fieldName: f.referenceName,
                key: f.referenceName,
                name: f.name,
                minWidth: columnSize.minWidth,
                maxWidth: columnSize.maxWidth,
                isResizable: !this.props.disableColumnResize,
                isSorted: this.state.sortColumn && Utils_String.equals(this.state.sortColumn, "ID", true),
                isSortedDescending: this.state.sortOrder && Utils_String.equals(this.state.sortOrder, "desc", true)
            }
        });        
    }

    @autobind
    private _onColumnHeaderClick(ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
        if (!this.props.disableSort) {
            this._updateState({sortColumn: column.fieldName, sortOrder: column.isSortedDescending ? "asc" : "desc"});
        }
    }

    @autobind
    private _onRenderCell(item: WorkItem, index?: number, column?: IColumn): React.ReactNode {
        let text: string;
        switch (column.fieldName.toLowerCase()) {            
            case "system.title":
                let witColor = this.state.workItemTypeAndStateColors && 
                            this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]] && 
                            this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].color;
                return (
                    <div className={this._getClassName("title-cell")} title={item.fields[column.fieldName]}>
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
                    <div className={this._getClassName("state-cell")} title={item.fields[column.fieldName]}>
                        {
                            this.state.workItemTypeAndStateColors &&
                            this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]] &&
                            this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors &&
                            this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]] &&
                            <span 
                                className="work-item-type-state-color" 
                                style={{
                                    backgroundColor: "#" + this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]],
                                    borderColor: "#" + this.state.workItemTypeAndStateColors[item.fields["System.WorkItemType"]].stateColors[item.fields["System.State"]]
                                }} />
                        }
                        <span className="overflow-ellipsis">{item.fields[column.fieldName]}</span>
                    </div>
                );
            case "System.AssignedTo":  // check isidentity flag
                return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;                      
            default:
                text = item.fields[column.fieldName];  
                break;          
        }

        return <div className="overflow-ellipsis" title={text}>{text}</div>;
    }    
    
    private _renderCommandBar(): JSX.Element {
        return (
            <div className={this._getClassName("menu-bar-container")}>
                {this.props.hideSearchBox && (
                    <SearchBox 
                        className={this._getClassName("searchbox")}
                        value={this.state.filterText || ""}
                        onSearch={(filterText: string) => this._updateFilterText(filterText)} 
                        onChange={(newText: string) => {
                            if (newText.trim() === "") {
                                this._updateFilterText("");
                            }
                        }} />
                )}

                {this.props.hideCommandBar && (
                    <CommandBar 
                        className={this._getClassName("menu-bar")}
                        items={this._getCommandMenuItems()} 
                        farItems={
                            [
                                {
                                    key: "resultCount", 
                                    name: `${this.state.filteredItems.length} results`, 
                                    className: this._getClassName("result-count")
                                }
                            ]
                        } />
                )}
            </div>
        );
    }

    private _getCommandMenuItems(): IContextualMenuItem[] {
        let menuItems: IContextualMenuItem[];
        
        if (this.props.refreshWorkItems) {
            menuItems.push({
                key: "refresh", name: "Refresh", title: "Refresh workitems", iconProps: {iconName: "Refresh"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    const workItems = await this.props.refreshWorkItems();
                    this._updateState({items: workItems, filteredItems: workItems});
                }
            });
        }

        menuItems.push({
            key: "OpenQuery", name: "Open as query", title: "Open all workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
            disabled: !this.state.filteredItems || this.state.filteredItems.length === 0,
            onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                const url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getWiql())}`;
                window.open(url, "_blank");
            }
        });
        
        if (this.props.extraCommandMenuItems && this.props.extraCommandMenuItems.length > 0) {
            menuItems = menuItems.concat(this.props.extraCommandMenuItems);
        }

        return menuItems;
    }

    private _getContextMenuItems(): IContextualMenuItem[] {
        let contextMenuItems: IContextualMenuItem[] = [            
            {
                key: "OpenQuery", name: "Open as query", title: "Open selected workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this._selection.getSelectedCount() == 0,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    const selectedWorkItems = this._selection.getSelection() as WorkItem[];
                    const url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getWiql(selectedWorkItems))}`;
                    window.open(url, "_blank");
                }
            }
        ];

        if (this.props.extraContextMenuItems && this.props.extraContextMenuItems.length > 0) {
            contextMenuItems = contextMenuItems.concat(this.props.extraContextMenuItems);
        }

        return contextMenuItems;
    }

    private _getWiql(workItems?: WorkItem[]): string {
        return "";
    }

    private _updateFilterText(filterText: string): void {
        this._updateState({filterText: filterText});
    }

    @autobind
    private _showContextMenu(item?: WorkItem, index?: number, e?: MouseEvent) {
        if (!this._selection.isIndexSelected(index)) {
            // if not already selected, unselect every other row and select this one
            this._selection.setAllSelected(false);
            this._selection.setIndexSelected(index, true, true);
        }        
        this._updateState({contextMenuTarget: e, isContextMenuVisible: true});
    }

    @autobind
    private _hideContextMenu(e?: any) {
        this._updateState({contextMenuTarget: null, isContextMenuVisible: false});
    }

    private async _openWorkItemDialog(e: React.MouseEvent<HTMLElement>, item: WorkItem): Promise<void> {
        let newTab = e ? e.ctrlKey : false;
        let workItemNavSvc = await WorkItemFormNavigationService.getService();
        workItemNavSvc.openWorkItem(item.id, newTab);
    }

    private _updateState(updatedStates: IWorkItemsGridState) {
        this.setState({...this.state, ...updatedStates});
    }

    private _getClassName(className?: string): string {
        if (className) {
            return `work-item-viewer-${className}`;
        }
        else {
            return "work-item-viewer";
        }        
    }
}