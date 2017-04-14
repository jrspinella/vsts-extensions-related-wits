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

    public static DEFAULT_FIELDS_TO_RETRIEVE = [
        "System.ID",
        "System.WorkItemType",
        "System.Title",
        "System.AssignedTo",
        "System.State",
        "System.Tags"
    ];
    public static DEFAULT_FIELDS_TO_SEEK = [
        "System.TeamProject",
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

    public static ExcludedSortableFields = [
        "System.AttachedFiles",
        "System.BISLinks",
        "System.LinkedFiles",
        "System.PersonId",
        "System.RelatedLinks"
    ];
}

export class Strings {
    public static NoWorkItemsFound = "No related work items found";
    public static AddFieldPlaceholder = "Add field";
    public static SavePreferenceTitle = "Save look up field preferences for this work item type";
    public static RefreshList = "Refresh the work item list";
    public static RemoveItemTitle = "Remove field";
    public static NeedAtleastOneField = "At least one look up field should be specified";
    public static SortBy = "Sort By :";
    public static AddLink = "Link to this workitem";
    public static AddLinkDialogTitle = "Add link to this workitem";
    public static AddLinkDialogOkText = "Add Link";
    public static AddLinkTitle = "Link added by Related workitems extension";
    public static LinkTypeLabel = "Link type";
    public static LinkCommentLabel = "Comment";
}

export interface IdentityReference {
    id?: string;
    displayName: string;
    uniqueName?: string;
    isIdentity?: boolean;
}