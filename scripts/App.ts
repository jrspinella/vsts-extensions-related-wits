import Q = require("q");
import Service = require("VSS/Service");
import Utils_Core = require("VSS/Utils/Core");
import Utils_String = require("VSS/Utils/String");
import Utils_Array = require("VSS/Utils/Array");
import {BaseControl} from "VSS/Controls";
import * as WitExtensionContracts  from "TFS/WorkItemTracking/ExtensionContracts";
import * as WitContracts from "TFS/WorkItemTracking/Contracts";
import {IWorkItemFormService, WorkItemFormService, WorkItemFormNavigationService, IWorkItemFormNavigationService} from "TFS/WorkItemTracking/Services";
import * as WitClient from "TFS/WorkItemTracking/RestClient"
import {StatusIndicator} from "VSS/Controls/StatusIndicator";
import {RelatedWitsControl, RelatedFieldsControl} from "scripts/Controls";
import {RelatedWitsControlOptions, RelatedFieldsControlOptions, Constants, Strings, UserPreferenceModel} from "scripts/Models";
import {UserPreferences} from "scripts/UserPreferences";

export class App {   
    // private _wiql: string = 
    //     "SELECT [System.Id],[System.WorkItemType],[System.Title],[System.AssignedTo],[System.State],[System.Tags] " + 
    //     "FROM WorkItems WHERE [System.TeamProject] = @project AND [System.WorkItemType] = '{0}' AND [System.Tags] CONTAINS '1' AND [System.State] = 'Active' AND [System.AreaPath] = 'Personal'"
    // ";
    private _statusIndicator: StatusIndicator;
    private _relatedFieldsControl: RelatedFieldsControl;
    private _relatedWitsControl: RelatedWitsControl;    
    private _workItemFormService: IWorkItemFormService;

    public initialize(): void {
        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: WitExtensionContracts.IWorkItemLoadedArgs) => {
                if (!args.isNew) {
                    this._render();
                }
                else {
                    this._clear();
                    this._toggleNewWorkItemMessage(true);
                }
            },
            onUnloaded: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                this._clear();
            },
            onSaved: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                this._render();
            },
            onRefreshed: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                this._render();
            },
        } as WitExtensionContracts.IWorkItemNotificationListener);     
    }

    private _toggleNewWorkItemMessage(show: boolean): void {
        if (show) {
            $(".new-workitem-message").show();
        }
        else {
            $(".new-workitem-message").hide();
        }
    }

    private _render(): void {
        this._clear();        
        this._showLoading(); 

        this._getWorkItemType()
            .then(
                (workItemType: string) =>  {
                    // read fields to seek from user preferences
                    UserPreferences.readUserSetting(workItemType)
                        .then((model: UserPreferenceModel) => {
                            let fieldsToSeek = model.fields || Constants.DEFAULT_FIELDS_TO_SEEK;
                            let sortByField = model.sortByField || "System.ChangedDate";

                            this._getWorkItemFields().then(
                                (fields: WitContracts.WorkItemField[]) => {
                                    // show fields control
                                    this._relatedFieldsControl = <RelatedFieldsControl>BaseControl.createIn(RelatedFieldsControl, $(".related-fields-container"), {
                                        selectedFields: fieldsToSeek.slice(),  // create a clone
                                        sortByField: sortByField,
                                        allFields: fields,
                                        savePreferences: (model: UserPreferenceModel) => {
                                            UserPreferences.writeUserSetting(workItemType, model);
                                        },
                                        refresh: (fields: string[], sortByField) => {
                                            this._showLoading();
                                            this._renderRelatedWitsControl(fields, sortByField);
                                        }
                                    });

                                    this._renderRelatedWitsControl(fieldsToSeek, sortByField);
                                }
                            )
                        }
                    )
            });
    }

    private _renderRelatedWitsControl(fieldsToSeek: string[], sortByField: string): void {
        // get related work items based on fields to seek
        this._getWorkItems(fieldsToSeek, sortByField)
            .then(
                (workItems: WitContracts.WorkItem[]) => {
                    if (!this._relatedWitsControl) {                                 
                        this._relatedWitsControl = <RelatedWitsControl>BaseControl.createIn(RelatedWitsControl, $(".related-wits-container"), {
                            workItems: workItems,
                            openWorkItem: (workItemId: number, newTab: boolean) => {
                                WorkItemFormNavigationService.getService().then((workItemNavSvc: IWorkItemFormNavigationService) => {
                                    // if control is pressed, open the work item in a new tab
                                    workItemNavSvc.openWorkItem(workItemId, newTab);
                                });
                            }
                        });
                    }
                    else {
                        this._relatedWitsControl.refresh(workItems);
                    }
                    this._hideLoading();
                }
            );
    }

    private _showLoading(): void {
        if (!this._statusIndicator) {
            this._statusIndicator = <StatusIndicator>BaseControl.createIn(StatusIndicator, $(".ext-container"), {
                center: true, 
                throttleMinTime: 0, 
                imageClass: Strings.LOADING_ICON, 
                message: Strings.Loading_Text 
            });
        }
        this._statusIndicator.start();
    }

    private _hideLoading(): void {
        if (this._statusIndicator) {
            this._statusIndicator.complete();
        }
    }

    private _ensureWorkItemFormService(): IPromise<IWorkItemFormService> {
        var defer = Q.defer();

        if (!this._workItemFormService) {
            WorkItemFormService.getService().then(
                (workItemFormService: IWorkItemFormService) => {
                    this._workItemFormService = workItemFormService;
                    defer.resolve(this._workItemFormService);
                }
            );
        }
        else {
            Utils_Core.delay(this, 0, () => {
                defer.resolve(this._workItemFormService);
            });
        }

        return defer.promise;
    }

    private _getWorkItemFields(): IPromise<WitContracts.WorkItemField[]> {
        var defer = Q.defer();

        this._ensureWorkItemFormService()
            .then((workItemFormService: IWorkItemFormService) => workItemFormService.getFields())
            .then((fields: WitContracts.WorkItemField[]) => defer.resolve(fields));

        return defer.promise;
    }

    private _getWorkItemType(): IPromise<string> {
        var defer = Q.defer();

        this._ensureWorkItemFormService()
            .then((workItemFormService: IWorkItemFormService) => workItemFormService.getFieldValue("System.WorkItemType"))
            .then((workItemType: string) => defer.resolve(workItemType));

        return defer.promise;
    }

    private _clear(): void {
        this._toggleNewWorkItemMessage(false);
        if (this._relatedFieldsControl) {
            this._relatedFieldsControl.dispose();
            this._relatedFieldsControl = null;
        }
        if (this._relatedWitsControl) {
            this._relatedWitsControl.dispose();
            this._relatedWitsControl = null;
        }
    }

    private _createWiql(fieldsToSeek: string[], sortByField: string): IPromise<string[]> {
        var defer = Q.defer();
        let fieldValuesToRead = fieldsToSeek.concat(["System.ID"]);
        this._ensureWorkItemFormService()
            .then((workItemFormService: IWorkItemFormService) => workItemFormService.getFieldValues(fieldValuesToRead))
            .then((fieldValues: IDictionaryStringTo<Object>) => {    
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

                defer.resolve([fieldValues["System.TeamProject"], wiql]);
            });

        return defer.promise;
    }

    private _getWorkItems(fieldsToSeek: string[], sortByField: string): IPromise<WitContracts.WorkItem[]> {
        var defer = Q.defer();

        this._createWiql(fieldsToSeek, sortByField)
            .then((data: string[]) => WitClient.getClient().queryByWiql({ query: data[1] }, data[0], null, false, 20))
            .then((queryResults: WitContracts.WorkItemQueryResult) => this._loadWorkItems(queryResults))
            .then((workItems: WitContracts.WorkItem[]) => defer.resolve(workItems));
            
        return defer.promise;
    }

    private _loadWorkItems(queryResults: WitContracts.WorkItemQueryResult): IPromise<WitContracts.WorkItem[]> {
        var defer = Q.defer();

        if(queryResults.workItems && queryResults.workItems.length > 0) {
            var workItemIds = $.map(queryResults.workItems, (elem: WitContracts.WorkItemReference) => {
                return elem.id;
            });
            var fields = $.map(queryResults.columns, (elem: WitContracts.WorkItemFieldReference) => {
                return elem.referenceName;
            });

            WitClient.getClient().getWorkItems(workItemIds, fields)
                .then((workItems: WitContracts.WorkItem[]) => {
                    var sortedWorkItems = workItems.sort((w1: WitContracts.WorkItem, w2: WitContracts.WorkItem) => {
                                                if (workItemIds.indexOf(w1.id) < workItemIds.indexOf(w2.id)) { return -1 }
                                                if (workItemIds.indexOf(w1.id) > workItemIds.indexOf(w2.id)) { return 1 }
                                                return 0;
                                            });
                    defer.resolve(sortedWorkItems);
                });
        }
        else {
            Utils_Core.delay(this, 0, () => {
                defer.resolve([]);
            });
        }

        return defer.promise;
    }
}