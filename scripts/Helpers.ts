import * as Utils_String from "VSS/Utils/String";
import {IdentityReference} from "./Models";

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

/**
 * Parse a distinct display name string into an entity reference object
 * 
 * @param name A distinct display name for an identity
 */
export function parseUniquefiedIdentityName(name: string): {displayName: string, uniqueName: string, imageUrl: string} {
    if (!name) { 
        return {
            displayName: "",
            uniqueName: "",
            imageUrl: ""
        };
    }
    
    let i = name.lastIndexOf("<");
    let j = name.lastIndexOf(">");
    let displayName = name;
    let alias = "";
    let rightPart = "";
    let id = "";
    if (i >= 0 && j > i && j === name.length - 1) {
        displayName = $.trim(name.substr(0, i));
        rightPart = $.trim(name.substr(i + 1, j - i - 1)); //gets string in the <>
        let vsIdFromAlias: string = getVsIdFromGroupUniqueName(rightPart); // if it has vsid in unique name (for TFS groups)

        if (rightPart.indexOf("@") !== -1 || rightPart.indexOf("\\") !== -1 || vsIdFromAlias || Utils_String.isGuid(rightPart)) {
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
            displayName = name;
        }
    }

    let imageUrl = "";
    if (id) {
        imageUrl = `${VSS.getWebContext().host.uri}/_api/_common/IdentityImage?id=${id}`;
    }
    else if(alias) {
        imageUrl = `${VSS.getWebContext().host.uri}/_api/_common/IdentityImage?identifier=${alias}&identifierType=0`;
    }

    return {
        displayName: displayName,
        uniqueName: alias,
        imageUrl: imageUrl
    };
}

export function getVsIdFromGroupUniqueName(str: string): string {
    let leftPart: string;
    if (!str) { return ""; }

    let vsid = "";
    let i = str.lastIndexOf("\\");
    if (i === -1) {
        leftPart = str;
    }
    else {
        leftPart = str.substr(0, i);
    }

    if (Utils_String.startsWith(leftPart, "id:")) {
        let rightPart = $.trim(leftPart.substr(3));
        vsid = Utils_String.isGuid(rightPart) ? rightPart : "";
    }

    return vsid;
}
