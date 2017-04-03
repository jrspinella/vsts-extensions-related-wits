import * as Utils_String from "VSS/Utils/String";
import {IdentityReference} from "./Models";

export function fieldNameComparer(a: string, b: string): number {
    let aUpper = a.toUpperCase();
    let bUpper = b.toUpperCase();

    if (aUpper < bUpper) { return -1 }
    if (aUpper > bUpper) { return 1 }
    return 0;
}

export class WorkItemTypeColorHelper {
    private static _workItemTypeColors: IDictionaryStringTo<string> = {
        "Bug": "#CC293D",
        "Task": "#F2CB1D",
        "Requirement": "#009CCC",
        "Feature": "#773B93",
        "Epic": "#FF7B00",
        "User Story": "#009CCC",
        "Product Backlog Item": "#009CCC"
    };

    public static parseColor(type: string): string {
        if (WorkItemTypeColorHelper._workItemTypeColors[type]) {
            return WorkItemTypeColorHelper._workItemTypeColors[type];
        }
        else {
            return "#FF9D00";
        }
    }
}

export class StateColorHelper {
    private static _stateColors: IDictionaryStringTo<string> = {
        "New": "#B2B2B2",
        "Active": "#007acc",
        "Resolved": "#ff9d00",
        "Closed": "#339933",
        "Requested": "#b2b2b2",
        "Accepted": "#007acc",
        "Removed": "#ffffff",
        "Ready": "#007acc",
        "Design": "#b2b2b2",
        "Inactive": "#339933",
        "In Planning": "#007acc",
        "In Progress": "#007acc",
        "Completed": "#339933"
    };

    public static parseColor(state: string): string {
        if (StateColorHelper._stateColors[state]) {
            return StateColorHelper._stateColors[state];
        }
        else {
            return "#007acc";
        }
    }
}

export class IdentityHelper {
    private static IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START = "<";
    private static IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END = ">";
    private static AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START = "<<";
    private static AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END = ">>";
    private static TFS_GROUP_PREFIX = "id:";
    private static AAD_IDENTITY_USER_PREFIX = "user:";
    private static AAD_IDENTITY_GROUP_PREFIX = "group:";
    private static IDENTITY_UNIQUENAME_SEPARATOR = "\\";

    public static parseIdentity(identityValue: string): IdentityReference {
        if (!identityValue) { 
            return  { displayName: "<Unassigned>" }; 
        }

        var i = identityValue.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START);
        var j = identityValue.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);
        var name = identityValue;
        var displayName = name;
        var alias = "";
        var rightPart = "";
        var id = "";
        if (i >= 0 && j > i) {
            displayName = $.trim(name.substr(0, i));
            rightPart = $.trim(name.substr(i + 1, j - i - 1)); //gets string in the <>
            var vsIdFromAlias: string = IdentityHelper.getVsIdFromGroupUniqueName(rightPart); // if it has vsid in unique name (for TFS groups)

            if (rightPart.indexOf("@") !== -1 || rightPart.indexOf("\\") !== -1 || vsIdFromAlias) {
                // if its a valid alias
                alias = rightPart;

                // If the alias component is just a guid then this is not a uniqueName but
                // vsId which is used only for TFS groups
                if (vsIdFromAlias != "") {
                    id = vsIdFromAlias;
                    alias = "";
                }
            }
            else {
                // if its not a valid alias, treat it as a non-identity string
                displayName = identityValue;
            }
        }

        return {
            id: id,
            displayName: displayName,
            uniqueName: alias,
            isIdentity: displayName != identityValue
        };
    }

    // Given a group uniquename that looks like id:2465ce16-6260-47a2-bdff-5fe4bc912c04\Build Administrators or id:2465ce16-6260-47a2-bdff-5fe4bc912c04, it will get the tfid from unique name
    private static getVsIdFromGroupUniqueName(str: string): string {
        var leftPart: string;
        if (!str) { return ""; }

        var vsid = "";
        var i = str.lastIndexOf(IdentityHelper.IDENTITY_UNIQUENAME_SEPARATOR);
        if (i === -1) {
            leftPart = str;
        }
        else {
            leftPart = str.substr(0, i);
        }

        if (Utils_String.startsWith(leftPart, IdentityHelper.TFS_GROUP_PREFIX)) {
            var rightPart = $.trim(leftPart.substr(3));
            vsid = Utils_String.isGuid(rightPart) ? rightPart : "";
        }

        return vsid;
    }
}