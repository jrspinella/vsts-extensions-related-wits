import "../../css/WorkItemsGrid.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import {IWorkItemsGridProps, IWorkItemsGridState, SortOrder} from "./Interfaces/WorkItemsGrid.Props";

import { DetailsList } from "OfficeFabric/DetailsList";
import { DetailsListLayoutMode, IColumn, CheckboxVisibility, ConstrainMode } from "OfficeFabric/components/DetailsList/DetailsList.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { Selection } from "OfficeFabric/utilities/selection/Selection";
import { autobind } from "OfficeFabric/Utilities";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { ContextualMenu } from "OfficeFabric/ContextualMenu";
import { CommandBar } from "OfficeFabric/CommandBar";
import { SearchBox } from "OfficeFabric/SearchBox";

import { Loading } from "VSTS_Extension/components/Loading";
import { IdentityView } from "VSTS_Extension/components/IdentityView";
import { MessagePanel, MessageType } from "VSTS_Extension/components/MessagePanel";
import { FluxContext } from "./Interfaces/FluxContext";

import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import Utils_Core = require("VSS/Utils/Core");
import { WorkItemFormNavigationService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemField, FieldType, WorkItemType, WorkItemStateColor } from "TFS/WorkItemTracking/Contracts";

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
    private _searchTimeout: any
    private _context: FluxContext;

    constructor(props: IWorkItemsGridProps, context: any) {
        super(props, context);
        this._selection = new Selection();
        this._context = FluxContext.get();
        this._initializeState();
    }

    private _initializeState() {
        this.state = {
            filteredItems: this._sortAndFilterWorkItems(this.props.items || [], this.props.sortColumn, this.props.sortOrder, ""),
            items: this.props.items || [],
            sortColumn: this.props.sortColumn,
            sortOrder: this.props.sortOrder,
            filterText: ""
        };
    }

    public componentDidMount() {
        this._context.stores.workItemColorStore.addChangedListener(this._onStoreChanged);
        this._context.actionsCreator.initializeWorkItemColors();
    }

    public componentWillUnmount() {
        this._context.stores.workItemColorStore.removeChangedListener(this._onStoreChanged);
    }    

    public componentWillReceiveProps(nextProps: Readonly<IWorkItemsGridProps>, nextContext: any): void {

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

    private _renderCommandBar(): JSX.Element {
        return (
            <div className={this._getClassName("menu-bar-container")}>
                {!this.props.hideSearchBox && (
                    <SearchBox 
                        className={this._getClassName("searchbox")}
                        value={this.state.filterText || ""}
                        onSearch={this._updateFilterText}
                        onChange={this._updateFilterText} />
                )}

                {!this.props.hideCommandBar && (
                    <CommandBar 
                        className={this._getClassName("menu-bar")}
                        items={this._getCommandMenuItems()} 
                        farItems={this._getFarCommandMenuItems()} />
                )}
            </div>
        );
    }

    private _getCommandMenuItems(): IContextualMenuItem[] {
        let menuItems: IContextualMenuItem[] = [];
        
        if (this.props.refreshWorkItems) {
            menuItems.push({
                key: "refresh", name: "Refresh", title: "Refresh workitems", iconProps: {iconName: "Refresh"},
                onClick: async (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this.setState({loading: true});
                    const workItems = await this.props.refreshWorkItems();
                    this._updateState({items: workItems, filteredItems: this._sortAndFilterWorkItems(workItems, this.state.sortColumn, this.state.sortOrder, this.state.filterText)});
                    this.setState({loading: false});
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

    private _getFarCommandMenuItems(): IContextualMenuItem[] {
        let menuItems: IContextualMenuItem[] = [
            {
                key: "resultCount", 
                name: `${this.state.filteredItems.length} results`, 
                className: this._getClassName("result-count")
            }
        ];
        
        if (this.props.farCommandMenuItems && this.props.farCommandMenuItems.length > 0) {
            menuItems = menuItems.concat(this.props.farCommandMenuItems);
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

    private _renderWorkItemGrid(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }
        else if (this.state.filteredItems.length === 0) {
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

    private _sortAndFilterWorkItems(workItems: WorkItem[], sortColumn: string, sortOrder: SortOrder, filterText: string): WorkItem[] {
        let items = (workItems || []).slice();
        if (sortColumn) {
            items = items.sort((w1: WorkItem, w2: WorkItem) => {
                if (Utils_String.equals(sortColumn, "System.Id", true)) {
                    return sortOrder === SortOrder.DESC ? ((w1.id > w2.id) ? -1 : 1) : ((w1.id > w2.id) ? 1 : -1);
                }            
                else {
                    let v1 = w1.fields[sortColumn];
                    let v2 = w2.fields[sortColumn];
                    return sortOrder === SortOrder.DESC ? -1 * Utils_String.ignoreCaseComparer(v1, v2) : Utils_String.ignoreCaseComparer(v1, v2);
                }
            });
        }

        if (filterText == null || filterText.trim() === "") {
            return items;
        }
        else {
            return items.filter((workItem: WorkItem) => {
                return `${workItem.id}` === filterText || this._doesAnyFieldValueContains(workItem, this.props.fieldColumns, filterText)
            });
        }
    }

    private _doesAnyFieldValueContains(workItem: WorkItem, fields: WorkItemField[], text: string): boolean {
        for (const field of fields) {
            if (Utils_String.caseInsensitiveContains(workItem.fields[field.referenceName] || "", text)) {
                return true;
            }
        }

        return false;
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
                isSorted: this.state.sortColumn && Utils_String.equals(this.state.sortColumn, f.referenceName, true),
                isSortedDescending: this.state.sortOrder && this.state.sortOrder === SortOrder.DESC
            }
        });        
    }

    @autobind
    private _onColumnHeaderClick(ev?: React.MouseEvent<HTMLElement>, column?: IColumn) {
        if (!this.props.disableSort) {
            const sortOrder = column.isSortedDescending ? SortOrder.ASC : SortOrder.DESC;
            this._updateState({sortColumn: column.fieldName, sortOrder: sortOrder, filteredItems: this._sortAndFilterWorkItems(this.state.items, column.fieldName, sortOrder, this.state.filterText)});
        }
    }

    @autobind
    private _onRenderCell(item: WorkItem, index?: number, column?: IColumn): React.ReactNode {
        let text: string;
        switch (column.fieldName.toLowerCase()) {
            case "system.id":  
                text = item.id.toString();
                break;
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
            case "system.state":
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
            case "system.assignedto":  // check isidentity flag
                return <IdentityView identityDistinctName={item.fields[column.fieldName]} />;                      
            default:
                text = item.fields[column.fieldName];  
                break;          
        }

        return <div className="overflow-ellipsis" title={text}>{text}</div>;
    }        

    private _getWiql(workItems?: WorkItem[]): string {
        const fieldStr = this.props.fieldColumns.map(f => `[${f.referenceName}]`).join(",");
        const ids = (workItems || this.state.filteredItems).map(w => w.id).join(",");
        const sortColumn = this.state.sortColumn || "System.CreatedDate";
        const sortOrder = this.state.sortOrder === SortOrder.DESC ? "DESC" : "";

        return `SELECT ${fieldStr}
                 FROM WorkItems 
                 WHERE [System.TeamProject] = @project 
                 AND [System.ID] IN (${ids}) 
                 ORDER BY [${sortColumn}] ${sortOrder}`;
    }

    @autobind
    private _updateFilterText(filterText: string): void {
        if (this._searchTimeout) {
            clearTimeout(this._searchTimeout);
            this._searchTimeout = null;
        }

        this._searchTimeout = setTimeout(() => {
            this._searchTimeout = null;
            this._updateState({filterText: filterText, filteredItems: this._sortAndFilterWorkItems(this.state.items, this.state.sortColumn, this.state.sortOrder, filterText)});
        }, 200)
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

    @autobind
    private _onStoreChanged() {
         if (this._context.stores.workItemColorStore.isLoaded()) {
             this._updateState({workItemTypeAndStateColors: this._context.stores.workItemColorStore.getAll()});
         }
    }

    private _getClassName(className?: string): string {
        if (className) {
            return `work-items-grid-${className}`;
        }
        else {
            return "work-items-grid";
        }        
    }
}