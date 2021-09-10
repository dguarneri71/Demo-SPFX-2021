import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDataService } from "./IDataService";
import { ICamlQuery, IItemAddResult, IList, ISiteGroupInfo, ISiteUser, IWeb, sp, Web } from "@pnp/sp/presets/all";

export default class SPDataService implements IDataService {
    // DG - 09/09/2021 - Using PnP/PnPjs
    // Uso BaseComponentContext invece di WebPartContext perchè così il modulo SPDataService funziona anche con le estensioni.
    private context: BaseComponentContext;

    // Configurazione servizio "sp"
    constructor(context: BaseComponentContext) {
        this.context = context;
        sp.setup({
            spfxContext: this.context
        });
    }
    //////////// DG - 10/09/2021
}