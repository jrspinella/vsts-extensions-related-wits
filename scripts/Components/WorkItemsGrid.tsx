import "../../css/WorkItemsGrid.scss";

import * as React from "react";

import {IWorkItemsGridProps, IWorkItemsGridState, SortOrder, ColumnPosition, ColumnType, IColumnsProps, IColumnProps} from "./Interfaces/WorkItemsGrid.Props";

import { DetailsList } from "OfficeFabric/DetailsList";
import { DetailsListLayoutMode, IColumn, CheckboxVisibility, ConstrainMode } from "OfficeFabric/components/DetailsList/DetailsList.Props";
import { SelectionMode } from "OfficeFabric/utilities/selection/interfaces";
import { Selection } from "OfficeFabric/utilities/selection/Selection";
import { autobind } from "OfficeFabric/Utilities";
import { IContextualMenuItem, ContextualMenuItemType } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";
import { ContextualMenu } from "OfficeFabric/ContextualMenu";
import { CommandBar } from "OfficeFabric/CommandBar";
import { SearchBox } from "OfficeFabric/SearchBox";

import { Loading } from "VSTS_Extension/components/Loading";
import { IdentityView } from "VSTS_Extension/components/IdentityView";
import { MessagePanel, MessageType } from "VSTS_Extension/components/MessagePanel";
import { FluxContext } from "./Interfaces/FluxContext";
import * as WorkItemsHelpers from "../Helpers/WorkItemsHelpers";

import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import Utils_Core = require("VSS/Utils/Core");
import { WorkItem, WorkItemField, FieldType, WorkItemType, WorkItemStateColor } from "TFS/WorkItemTracking/Contracts";



export class WorkItemsGrid extends React.Component<IWorkItemsGridProps, IWorkItemsGridState> {
    static defaultProps = {
        refreshWorkItems: null,
        selectionMode: SelectionMode.multiple,
        onItemInvoked: (item: WorkItem, index: number) => {
            WorkItemsHelpers.openWorkItemDialog(null, item);
        },
        columnsProps: {
            disableSort: false,
            disableColumnResize: false,
            extraColumns: []
        },
        commandBarProps: {
            hideSearchBox: false,
            hideCommandBar: false,
            extraCommandMenuItems: [],
            farCommandMenuItems: []
        },
        contextMenuProps: {
            disableContextMenu: false,
            extraContextMenuItems: null
        }
    } as IWorkItemsGridProps;

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
            filteredItems: this.props.items || [],
            items: this.props.items || [],
            sortColumn: null,
            sortOrder: SortOrder.ASC,
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

    public render(): JSX.Element {
        return (
            <div className={this._getClassName()}>
                {this._renderCommandBar()}
                {this._renderWorkItemGrid()}
                {this.state.isContextMenuVisible && !this.props.contextMenuProps.disableContextMenu && (
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
                {!this.props.commandBarProps.hideSearchBox && (
                    <SearchBox 
                        className={this._getClassName("searchbox")}
                        value={this.state.filterText || ""}
                        onSearch={this._updateFilterText}
                        onChange={this._updateFilterText} />
                )}

                {!this.props.commandBarProps.hideCommandBar && (
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
        
        if (this.props.commandBarProps.extraCommandMenuItems && this.props.commandBarProps.extraCommandMenuItems.length > 0) {
            menuItems.push({ key: "divider", name: "Divider", itemType: ContextualMenuItemType.Divider });
            menuItems = menuItems.concat(this.props.commandBarProps.extraCommandMenuItems);
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
        
        if (this.props.commandBarProps.farCommandMenuItems && this.props.commandBarProps.farCommandMenuItems.length > 0) {
            menuItems = menuItems.concat(this.props.commandBarProps.farCommandMenuItems);
        }

        return menuItems;
    }

    private _getContextMenuItems(): IContextualMenuItem[] {
        const selectedWorkItems = this._selection.getSelection() as WorkItem[];

        let contextMenuItems: IContextualMenuItem[] = [
            {
                key: "OpenQuery", name: "Open as query", title: "Open selected workitems as a query", iconProps: {iconName: "OpenInNewWindow"}, 
                disabled: this._selection.getSelectedCount() == 0,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {                    
                    const url = `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_workitems?_a=query&wiql=${encodeURIComponent(this._getWiql(selectedWorkItems))}`;
                    window.open(url, "_blank");
                }
            }
        ];

        if (this.props.contextMenuProps.extraContextMenuItems) {
            const extraContextMenus = this.props.contextMenuProps.extraContextMenuItems(selectedWorkItems);
            if (extraContextMenus && extraContextMenus.length > 0) {
                contextMenuItems.push({ key: "divider", name: "Divider", itemType: ContextualMenuItemType.Divider });
                contextMenuItems = contextMenuItems.concat(extraContextMenus);
            }
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
                            this.props.onItemInvoked(item, index);
                        }}
                        selection={ this._selection }
                        onItemContextMenu={this._showContextMenu}
                        onColumnHeaderClick={this._onColumnHeaderClick}
                    />;
        }
    }

    private _sortAndFilterWorkItems(workItems: WorkItem[], sortColumn: IColumn, sortOrder: SortOrder, filterText: string): WorkItem[] {
        let items = (workItems || []).slice();
        if (sortColumn && sortColumn.data.comparer) {
            items = items.sort((w1: WorkItem, w2: WorkItem) => sortColumn.data.comparer(w1, w2, sortColumn.key, sortOrder));
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
        let columns: IColumn[] = [];
        let leftColumns: IColumn[] = [];
        let rightColumns: IColumn[] = [];

        let extraColumnMapper = (ec: IColumnProps) => {
            return {
                fieldName: ec.key,
                key: ec.key,
                name: ec.name,
                minWidth: ec.minWidth || 50,
                maxWidth: ec.maxWidth || 100,
                isResizable: !this.props.columnsProps.disableColumnResize,
                data: { type: ColumnType.Custom, renderer: ec.renderer, comparer: ec.comparer },
                isSorted: this.state.sortColumn && Utils_String.equals(this.state.sortColumn.key, ec.key, true),
                isSortedDescending: this.state.sortOrder && this.state.sortOrder === SortOrder.DESC
            }
        }

        if (this.props.columnsProps.extraColumns && this.props.columnsProps.extraColumns.length > 0) {
            leftColumns = this.props.columnsProps.extraColumns.map(extraColumnMapper);
            rightColumns = this.props.columnsProps.extraColumns.filter(ec => ec.position !== ColumnPosition.FarLeft).map(extraColumnMapper);            
        }

        if (leftColumns.length > 0) {
            columns = columns.concat(leftColumns);
        }

        columns = columns.concat(this.props.fieldColumns.map(f => {
            const columnSize = WorkItemsHelpers.getColumnSize(f);
            return {
                fieldName: f.referenceName,
                key: f.referenceName,
                name: f.name,
                data: { type: ColumnType.Field, renderer: WorkItemsHelpers.workItemFieldCellRenderer, comparer: WorkItemsHelpers.workItemFieldValueComparer },
                minWidth: columnSize.minWidth,
                maxWidth: columnSize.maxWidth,
                isResizable: !this.props.columnsProps.disableColumnResize,
                isSorted: this.state.sortColumn && Utils_String.equals(this.state.sortColumn.key, f.referenceName, true),
                isSortedDescending: this.state.sortOrder && this.state.sortOrder === SortOrder.DESC
            }
        }));

        if (rightColumns.length > 0) {
            columns = columns.concat(rightColumns);
        }

        return columns;
    }

    @autobind
    private _onColumnHeaderClick(ev: React.MouseEvent<HTMLElement>, column: IColumn) {
        if (!this.props.columnsProps.disableSort && column.data.comparer) {
            const sortOrder = column.isSortedDescending ? SortOrder.ASC : SortOrder.DESC;
            this._updateState({sortColumn: column, sortOrder: sortOrder, filteredItems: this._sortAndFilterWorkItems(this.state.items, column, sortOrder, this.state.filterText)});
        }
    }

    @autobind
    private _onRenderCell(item: WorkItem, index: number, column: IColumn): React.ReactNode {
        if (column.data.type === ColumnType.Custom) {
            return column.data.renderer ? column.data.renderer(item, index) : null;
        }
        else {
            return column.data.renderer(item, index, column, {workItemTypeAndStateColors: this.state.workItemTypeAndStateColors});
        }        
    }        

    private _getWiql(workItems?: WorkItem[]): string {
        const fieldStr = this.props.fieldColumns.map(f => `[${f.referenceName}]`).join(",");
        const ids = (workItems || this.state.filteredItems).map(w => w.id).join(",");
        const sortColumn = this.state.sortColumn && this.state.sortColumn.data.type === ColumnType.Field ? this.state.sortColumn.key : "System.CreatedDate";
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
        if (!this.props.contextMenuProps.disableContextMenu) {
            if (!this._selection.isIndexSelected(index)) {
                // if not already selected, unselect every other row and select this one
                this._selection.setAllSelected(false);
                this._selection.setIndexSelected(index, true, true);
            }        
            this._updateState({contextMenuTarget: e, isContextMenuVisible: true});
        }
    }

    @autobind
    private _hideContextMenu(e?: any) {
        this._updateState({contextMenuTarget: null, isContextMenuVisible: false});
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