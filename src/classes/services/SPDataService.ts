import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IDataService } from "./IDataService";
import { ICamlQuery, IItemAddResult, IList, ISite, ISiteGroupInfo, ISiteUser, IWeb, Site, sp, Web } from "@pnp/sp/presets/all";

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

    public loadDocuments(siteUrl: string, libraryId: string): Promise<any[]> {
        return new Promise<any[]>((res, reject) => {
            // if a site URL has been specified, use that site, otherwise assume,
            // that the selected list is in the current site
            const web: IWeb = siteUrl ? Web(siteUrl) : sp.web;
            web.lists
                .getById(libraryId)
                // FileLeafRef contains the name of the file, FileRef contains the
                // server-relative URL of the file to be used in the document link
                .items.select('FileLeafRef', 'FileRef')
                .orderBy('Modified', false)
                .get()
                // show retrieved documents, if any
                .then(docs => {
                    res(docs);
                })
                // show error
                .catch(err => {
                    reject(err);
                });
        });
    }

    public getSiteCollectionId(siteUrl?: string): Promise<string | void> {
        return new Promise<string>((resolve: (siteId: string) => void, reject: (error: any) => void): void => {
            const site: ISite = Site(siteUrl);
            site.select('Id').get()
                .then(({ Id }): void => {
                    resolve(Id);
                })
                .catch(err => reject(err));
        });
    }

    public getSiteId(siteUrl?: string): Promise<string | void> {
        return new Promise<string>((resolve: (siteId: string) => void, reject: (error: any) => void): void => {
            const web: IWeb = Web(siteUrl);
            web.select('Id').get()
                .then(({ Id }): void => {
                    resolve(Id);
                })
                .catch(err => reject(err));
        });
    }
}