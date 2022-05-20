import { TestItem } from "../commons/TestItem";

export interface IDataService {
    loadDocuments(siteUrl: string, libraryId: string): Promise<any[]>;
    getSiteCollectionId(siteUrl?: string): Promise<void | string>;
    getSiteId(siteUrl?: string): Promise<void | string>;

    loadItems(siteUrl: string, libraryId: string): Promise<TestItem[]>;
}