import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDataService } from "./IDataService";
import MockDataService from "./MockDataService";
import SPDataService from "./SPDataService";

export default class DataServiceProvider {
    public static get(context: BaseComponentContext, useMock: boolean = true): IDataService {
        console.log("DataServiceProvider - get - EnvironmentType (0-Test; 1-Local; 2-SharePoint; 3-ClassicSharePoint ): " + Environment.type);
        if (useMock === false) {
            return new SPDataService(context);
        } else {
            return (Environment.type !== EnvironmentType.Local) ? new SPDataService(context) : new MockDataService();
        }
    }
}