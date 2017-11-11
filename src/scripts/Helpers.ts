import { WorkItem } from "TFS/WorkItemTracking/Contracts";

export async function confirmAction(condition: boolean, msg: string): Promise<boolean> {
    if (condition) {
        let dialogService: IHostDialogService = await VSS.getService(VSS.ServiceIds.Dialog) as IHostDialogService;
        try {
            await dialogService.openMessageDialog(msg, { useBowtieStyle: true });
            return true;
        }
        catch (e) {
            // user selected "No"" in dialog
            return false;
        }
    }

    return true;
}

export function getQueryUrl(workItems: WorkItem[], fields: string[]) {
    const fieldStr = fields.join(",");
    const ids = (workItems).map(w => w.id).join(",");

    const wiql = `SELECT ${fieldStr}
             FROM WorkItems 
             WHERE [System.ID] IN (${ids})`;

    return `${VSS.getWebContext().host.uri}/${VSS.getWebContext().project.id}/_queries/query/?wiql=${encodeURIComponent(wiql)}`;
}