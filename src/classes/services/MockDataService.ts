import { TestItem } from "../commons/TestItem";
import { IDataService } from "./IDataService";

export default class MockDataService implements IDataService {     
    public loadDocuments(siteUrl: string, libraryId: string): Promise<any[]> {
        return new Promise<any[]>((res, reject) => {
            let docs = [];
            docs.push({'FileLeafRef': 'demo1.pdf', 'FileRef': '/demo1.pdf'});
            docs.push({'FileLeafRef': 'demo2.pdf', 'FileRef': '/demo2.pdf'});
            docs.push({'FileLeafRef': 'demo3.pdf', 'FileRef': '/demo3.pdf'});
            res(docs);
        });
    }

    public getSiteCollectionId(siteUrl?: string): Promise<string | void> {
        return Promise.resolve();
    }

    public getSiteId(siteUrl?: string): Promise<string | void> {
        return Promise.resolve();
    }

    public loadItems(siteUrl: string, libraryId: string): Promise<TestItem[]> {
        throw new Error("Method not implemented.");
    }   
}