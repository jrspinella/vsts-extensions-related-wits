import * as WitContracts from "TFS/WorkItemTracking/Contracts";

export interface RelatedWitsControlOptions {
    workItems:  RelatedWitReference[];
    openWorkItem: (workItemId: number, newTab: boolean) => void;
    linkWorkItem: (workItem: RelatedWitReference, relationType: string, comment: string) => void;
}

export interface RelatedWitReference extends WitContracts.WorkItem {
    url: string;
    isLinked: boolean;
}

export interface UserPreferenceModel {
    fields: string[];
    sortByField: string;
    top?: number;
}

export interface RelatedFieldsControlOptions {
    selectedFields: string[];
    allFields: WitContracts.WorkItemField[];
    sortByField: string;
    savePreferences: (model: UserPreferenceModel) => void;
    refresh: (fields: string[], sortByField: string) => void;
}

export interface AddLinkDialogResult {
    relationType: WitContracts.WorkItemRelationType;
    comment: string;
}

export class Constants {
    public static StorageKey = "RelatedWorkItemsFields";
    public static UserScope = { scopeType: "User" };

    public static DEFAULT_SORT_BY_FIELD = "System.ChangedDate";
    public static DEFAULT_RESULT_SIZE = 20;

    public static DEFAULT_FIELDS_TO_RETRIEVE = [
        "System.ID",
        "System.WorkItemType",
        "System.Title",
        "System.AssignedTo",
        "System.AreaPath",
        "System.State",
        "System.Tags"
    ];

    public static DEFAULT_FIELDS_TO_SEEK = [
        "System.WorkItemType",
        "System.Tags",
        "System.State",
        "System.AreaPath"
    ];

    public static QueryableFieldTypes = [
        WitContracts.FieldType.Boolean,
        WitContracts.FieldType.Double,
        WitContracts.FieldType.Integer,
        WitContracts.FieldType.String,
        WitContracts.FieldType.TreePath        
    ];

    public static SortableFieldTypes = [
        WitContracts.FieldType.DateTime,
        WitContracts.FieldType.Double,
        WitContracts.FieldType.Integer,
        WitContracts.FieldType.String,
        WitContracts.FieldType.TreePath
    ];

    public static ExcludedFields = [
        "System.AttachedFiles",
        "System.AttachedFileCount",
        "System.ExternalLinkCount",
        "System.HyperLinkCount",
        "System.BISLinks",
        "System.LinkedFiles",
        "System.PersonId",
        "System.RelatedLinks",
        "System.RelatedLinkCount",
        "System.TeamProject",
        "System.Rev",
        "System.Watermark",
        "Microsoft.VSTS.Build.IntegrationBuild"
    ];
}