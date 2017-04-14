import "../css/app.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { IWorkItemNotificationListener, IWorkItemChangedArgs, IWorkItemFieldChangedArgs, IWorkItemLoadedArgs } from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService } from "TFS/WorkItemTracking/Services";
import { WorkItem, WorkItemQueryResult, WorkItemRelation } from "TFS/WorkItemTracking/Contracts";
import * as WitClient from "TFS/WorkItemTracking/RestClient";
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");

import { autobind } from "OfficeFabric/Utilities";
import { Fabric } from "OfficeFabric/Fabric";
import { CommandBar } from "OfficeFabric/CommandBar";
import { IContextualMenuItem } from "OfficeFabric/components/ContextualMenu/ContextualMenu.Props";

import { Loading } from "./Loading";
import { MessagePanel, MessageType } from "./MessagePanel";
import { InputError } from "./InputError";
import { UserPreferenceModel } from "./Models";
import { UserPreferences } from "./UserPreferences";

interface IRelatedWitsState {
    loading: boolean;
    listItems: IListItem[];
    isWorkItemLoaded?: boolean;
    isNew?: boolean;
    settings?: UserPreferenceModel;
    settingsPanelOpen?: boolean;
}

interface IListItem {
    workItem: WorkItem;
    isLinked: boolean;
}

export class RelatedWits extends React.Component<void, IRelatedWitsState> {

    constructor(props: void, context?: any) {
        super(props, context);

        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: IWorkItemLoadedArgs) => {
                if (args.isNew) {
                    this.setState({...this.state, isWorkItemLoaded: true, isNew: true});
                }
                else {
                    this._refreshList();
                }
            },
            onUnloaded: (args: IWorkItemChangedArgs) => {
                this.setState({...this.state, isWorkItemLoaded: false, workItems: []});
            },
            onSaved: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onRefreshed: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
            onReset: (args: IWorkItemChangedArgs) => {
                this._refreshList();
            },
        } as IWorkItemNotificationListener);

        this.state = {
            loading: true,
            listItems: [],            
        } as IRelatedWitsState;
    }

    public render(): JSX.Element {
        if (!this.state.isWorkItemLoaded) {
            return null;
        }
        else if (this.state.isNew) {
            return <MessagePanel message="Please save the workitem to get the list of related work items." messageType={MessageType.Info} />;
        }
        else {
            return (
                <Fabric className="fabric-container">
                    <CommandBar className="all-view-menu-toolbar" items={this._getMenuItems()} />
                    {this._renderList()}
                </Fabric>
            );
        }        
    }

    private _getMenuItems(): IContextualMenuItem[] {
         return [                      
            {
                key: "refresh", name: "Refresh", title: "Refresh list", iconProps: {iconName: "Refresh"},
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this._refreshList();
                }
            },
            {
                key: "settings", name: "Settings", title: "Toggle settings panel", iconProps: {iconName: "Settings"}, 
                disabled: this.state.settings == null, checked: this.state.settingsPanelOpen,
                onClick: (event?: React.MouseEvent<HTMLElement>, menuItem?: IContextualMenuItem) => {
                    this.setState({...this.state, settingsPanelOpen: !(this.state.settingsPanelOpen)});
                }
            }
         ] as IContextualMenuItem[];
    }

    private _renderList(): JSX.Element {
        if (this.state.loading) {
            return <Loading />;
        }        
        else if (!this.state.listItems || this.state.listItems.length === 0) {
            return <MessagePanel message="No related work items found." messageType={MessageType.Info} />;
        }
        else {
            return null;
        }
    }

    private async _refreshList(): Promise<void> {
        this.setState({...this.state, isWorkItemLoaded: true, isNew: false, loading: true, listItems: []});

        if (!this.state.settings) {
            const workItemFormService = await WorkItemFormService.getService();
            const workItemType = await workItemFormService.getFieldValue("System.WorkItemType") as string;
            let settings = await UserPreferences.readUserSetting(workItemType);
            this.setState({...this.state, settings: settings});
        }


        this.setState({...this.state, loading: false, listItems: []});
    } 

    private async _getWorkItems(fieldsToSeek: string[], sortByField: string): Promise<IListItem[]> {
        let data: string[] = await this._createWiql(fieldsToSeek, sortByField);
        let queryResults = await WitClient.getClient().queryByWiql({ query: data[1] }, data[0], null, false, 20);
        return this._readWorkItemsFromQueryResults(queryResults);
    }

    private async _createWiql(fieldsToSeek: string[], sortByField: string): Promise<string[]> {
        let fieldValuesToRead = fieldsToSeek.concat(["System.ID"]);
        let workItemFormService = await this._ensureWorkItemFormService()
        let fieldValues = await workItemFormService.getFieldValues(fieldValuesToRead);
        let witId = fieldValues["System.ID"]; 
        // Generate fields to retrieve part
        let fieldsToRetrieveString: string = "";
        $.each(Constants.DEFAULT_FIELDS_TO_RETRIEVE, (i: number, fieldRefName: string) => {
            fieldsToRetrieveString = `${fieldsToRetrieveString}[${fieldRefName}],`
        });
        // remove last comma
        fieldsToRetrieveString = fieldsToRetrieveString.substring(0, fieldsToRetrieveString.length - 1);

        // Generate fields to seek part
        let fieldsToSeekString: string = "";
        $.each(fieldsToSeek, (i: number, fieldRefName: string) => {
            let fieldValue = fieldValues[fieldRefName];
            if (Utils_String.equals(fieldRefName, "System.Tags", true) && fieldValue) {
                fieldsToSeekString = fieldsToSeekString + " (";
                let fieldValueStr = fieldValue.toString();
                $.each(fieldValueStr.split(";"), (i: number, v: string) => {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] CONTAINS '${v}' OR`
                });
                if (fieldsToSeekString) {
                    // remove last OR
                    fieldsToSeekString = fieldsToSeekString.substring(0, fieldsToSeekString.length - 3) + ") AND";
                }
            }
            else if (!Utils_String.equals(fieldRefName, "System.TeamProject", true)) {
                if (Utils_String.equals(typeof(fieldValue), "string", true) && fieldValue) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = '${fieldValue}' AND`
                }
                else if (Utils_String.equals(typeof(fieldValue), "number", true) && fieldValue != null) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = ${fieldValue} AND`
                }
                else if (Utils_String.equals(typeof(fieldValue), "boolean", true) && fieldValue != null) {
                    fieldsToSeekString = `${fieldsToSeekString} [${fieldRefName}] = ${fieldValue} AND`
                }
            }
        });
        if (fieldsToSeekString) {
            // remove last OR
            fieldsToSeekString = fieldsToSeekString.substring(0, fieldsToSeekString.length - 4);
        }
        let fieldsToSeekPredicate = fieldsToSeekString ? `AND ${fieldsToSeekString}` : "";
        let wiql = `SELECT ${fieldsToRetrieveString} FROM workitems where [System.TeamProject] = @project AND [System.ID] <> ${witId} ${fieldsToSeekPredicate} order by [${sortByField}] desc`;

        return [fieldValues["System.TeamProject"] as string, wiql];
    }


    private async _readWorkItemsFromQueryResults(queryResults: WorkItemQueryResult): Promise<IListItem[]> {
        if (queryResults.workItems && queryResults.workItems.length > 0) {
            var workItemIdMap: IDictionaryNumberTo<WorkItemReference> = {};
            var workItemIds = $.map(queryResults.workItems, (elem: WorkItemReference) => {
                return elem.id;
            });
            var fields = $.map(queryResults.columns, (elem: WorkItemFieldReference) => {
                return elem.referenceName;
            });

            $.each(queryResults.workItems, (i:number, w: WorkItemReference) => {
                workItemIdMap[w.id] = w;
            })

            let workItems = await WitClient.getClient().getWorkItems(workItemIds, fields);
            // sort the workitems in the same order as they got retrieved from query
            var sortedWorkItems = workItems.sort((w1: WorkItem, w2: WorkItem) => {
                                        if (workItemIds.indexOf(w1.id) < workItemIds.indexOf(w2.id)) { return -1 }
                                        if (workItemIds.indexOf(w1.id) > workItemIds.indexOf(w2.id)) { return 1 }
                                        return 0;
                                    });

            let workItemFormService = await WorkItemFormService.getService();
            let relations = await workItemFormService.getWorkItemRelations();
            return $.map(sortedWorkItems, (w: WorkItem) => $.extend(w, { 
                url: workItemIdMap[w.id].url,
                isLinked: Utils_Array.arrayContains(w.url, relations, (url: string, relation: WorkItemRelation) => {
                    return Utils_String.equals(relation.url, url, true);
                })
            }));
        }
        else {
            return [];
        }
    }
}

export function init() {
    ReactDOM.render(<RelatedWits />, $("#ext-container")[0]);
}