import Q = require("q");
import Service = require("VSS/Service");
import {Constants, UserPreferenceModel} from "scripts/Models";

export class UserPreferences {
    /**
    * Read user extension data for a particular work item type
    * @param workItemType
    */
    public static readUserSetting(workItemType: string): IPromise<UserPreferenceModel> {
        var defer = Q.defer();

        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData)
            .then((dataService: IExtensionDataService) => dataService.getValue<UserPreferenceModel>(`${Constants.StorageKey}_${workItemType}`, Constants.UserScope))
            .then((model: UserPreferenceModel) => defer.resolve(model || {}));

        return defer.promise;
    }

    /**
    * Write user extension data for a particular work item type
    */
    public static writeUserSetting(workItemType: string, model: UserPreferenceModel): void {
        VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData)
            .then((dataService: IExtensionDataService) => dataService.setValue<UserPreferenceModel>(`${Constants.StorageKey}_${workItemType}`, model, Constants.UserScope));
    }
}