import Service = require("VSS/Service");
import {Constants, UserPreferenceModel} from "./Models";

export class UserPreferences {
    /**
    * Read user extension data for a particular work item type
    * @param workItemType
    */
    public static async readUserSetting(workItemType: string): Promise<UserPreferenceModel> {
        const dataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
        try {
            const model = await dataService.getValue<UserPreferenceModel>(`${Constants.StorageKey}_${workItemType}`, Constants.UserScope);
            return model || {} as UserPreferenceModel;
        }
        catch (e) {
            return {} as UserPreferenceModel;
        }
    }

    /**
    * Write user extension data for a particular work item type
    */
    public static async writeUserSetting(workItemType: string, model: UserPreferenceModel): Promise<void> {
        const dataService = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
        try {
            await dataService.setValue<UserPreferenceModel>(`${Constants.StorageKey}_${workItemType}`, model, Constants.UserScope);
        }
        catch (e) {
            return null;
        }
    }
}