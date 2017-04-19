import * as WitContracts from "TFS/WorkItemTracking/Contracts";

export interface UserPreferenceModel {
    fields: string[];
    sortByField: string;
    top?: number;
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