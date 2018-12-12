/// <reference types="sharepoint" />
/**
 * Serializes a TermSet to a conforming JSON that can be used to initialize SP.Taxonomy.TermSet
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the termset to serialize
 */
export declare function termSetToJson(tset: SP.Taxonomy.TermSet): {
    Id: string;
    Name: string;
    LastModifiedDate: Date;
    CreatedDate: Date;
    Contact: string;
    Description: string;
    IsOpenForTermCreation: boolean;
    Stakeholders: string[];
    CustomSortOrder: string;
    CustomProperties: {
        [key: string]: string;
    };
    IsAvailableForTagging: boolean;
    Owner: string;
};
/**
 * Serializes a Term to a conforming JSON that can be used to initialize SP.Taxonomy.Term
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the term to serialize
 */
export declare function termToJson(term: SP.Taxonomy.Term): {
    Id: string;
    Name: string;
    Description: string;
    IsDeprecated: boolean;
    IsKeyword: boolean;
    IsPinned: boolean;
    IsReused: boolean;
    IsRoot: boolean;
    PathOfTerm: string;
    TermsCount: number;
    IsSourceTerm: boolean;
    IsAvailableForTagging: boolean;
    LocalCustomProperties: {
        [key: string]: string;
    };
    CustomProperties: {
        [key: string]: string;
    };
    CustomSortOrder: string;
};
export interface IValue {
    id: string | number;
    label: string;
}
export interface ITermStore {
    createTerm(parentTerm: SP.Taxonomy.Term, name: string, locale: number, guid: SP.Guid, properties?: IValue[]): Promise<SP.Taxonomy.Term>;
    getParentTermById(termId: string): Promise<SP.Taxonomy.Term>;
    getTerms(termIds: string[]): Promise<SP.Taxonomy.Term[]>;
    getTermsByTermId(termId: string): Promise<SP.Taxonomy.Term[]>;
    getTermsByTermSetId(termSetId: string): Promise<SP.Taxonomy.Term[]>;
    getAllTermsByTermSetId(termSetId: string): Promise<SP.Taxonomy.Term[]>;
    termSetIdFromTaxonomyField(fieldInternalName: string): Promise<string>;
    getTermsByIds(ids: string[]): Promise<SP.Taxonomy.TermCollection>;
    getParentTermByTerm(term: SP.Taxonomy.Term): Promise<SP.Taxonomy.Term>;
    getTopLevelParentOfTerm(id: string): Promise<SP.Taxonomy.Term>;
    getTermParents(termId: string): Promise<SP.Taxonomy.Term[]>;
    getParentThatSatisfies(id: string, fn: (v: SP.Taxonomy.Term) => boolean): Promise<SP.Taxonomy.Term>;
    getTermLabelsById(termId: string): Promise<SP.Taxonomy.Label[]>;
    getLabelsForTerms(terms: SP.Taxonomy.Term[]): Promise<SP.Taxonomy.LabelCollection[]>;
    getTermsSubTreeFlat(termId: string, list?: SP.Taxonomy.Term[]): Promise<SP.Taxonomy.Term[]>;
    getAllTermSetsInSiteCollectionGroup(createIfMissing: boolean): Promise<SP.Taxonomy.TermSet[]>;
    getTermLabels(term: SP.Taxonomy.Term): Promise<SP.Taxonomy.Label[]>;
    getTermById(id: string): Promise<SP.Taxonomy.Term>;
    getSiteCollectionTermGroup(createIfMissing: any): Promise<SP.Taxonomy.TermGroup>;
}
export declare type TermStoreContextCallback<T> = (ctx: SP.ClientContext, session: SP.Taxonomy.TaxonomySession, tstore: SP.Taxonomy.TermStore, execute: (result: T) => void) => void;
export declare function createTermStore(): ITermStore;
