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
    /**
     * Create a child term given the parent {SP.Taxonomy.Term}
     * @param parentTerm the parent term
     * @param name the name of the term
     * @param locale the locale number default 1033 (en-US)
     * @param guid the guid of the term
     * @param properties any custom properties to add to the term
     */
    createTerm(parentTerm: SP.Taxonomy.Term, name: string, locale: number, guid: SP.Guid, properties?: IValue[]): Promise<SP.Taxonomy.Term>;
    /**
     * Get the parent term by term id
     * @param termId the term id to find its parent
     */
    getParentTermById(termId: string): Promise<SP.Taxonomy.Term>;
    /**
     * Get terms given their guids
     * @param termIds array of term guids (string)
     */
    getTerms(termIds: string[]): Promise<SP.Taxonomy.Term[]>;
    /**
     * get child terms of term given the parent term id
     * @param termId parent term id
     */
    getTermsByTermId(termId: string): Promise<SP.Taxonomy.Term[]>;
    /**
     * Get immediate child terms given term set id
     * @param termSetId the term set id
     */
    getTermsByTermSetId(termSetId: string): Promise<SP.Taxonomy.Term[]>;
    /**
     * All child terms of given termset id
     * @param termSetId the term set id
     */
    getAllTermsByTermSetId(termSetId: string): Promise<SP.Taxonomy.Term[]>;
    /**
     * Get term set id from taxonomy field name
     */
    termSetIdFromTaxonomyField(fieldInternalName: string): Promise<string>;
    /**
     * get terms collection given term ids
     * @param ids term ids
     * @see {ITermStore#getTerms}
     */
    getTermsByIds(ids: string[]): Promise<SP.Taxonomy.TermCollection>;
    /**
     * Get parent term by term
     * @param term {SP.Taxonomy.Term} the term to get its parent
     */
    getParentTermByTerm(term: SP.Taxonomy.Term): Promise<SP.Taxonomy.Term>;
    /**
     * Get top level parent of a term.
     * This method executes recursively until it finds the top level parent of a
     * term.
     * @param id the term id
     */
    getTopLevelParentOfTerm(id: string): Promise<SP.Taxonomy.Term>;
    /**
     * Get term parents (all term parents) given a term id
     * @param termId the term id to find its parents
     */
    getTermParents(termId: string): Promise<SP.Taxonomy.Term[]>;
    /**
     * Get term parents that satisfy a specific condition (using a callback)
     * @param id the id of the term
     * @param fn the callback
     */
    getParentThatSatisfies(id: string, fn: (v: SP.Taxonomy.Term) => boolean): Promise<SP.Taxonomy.Term>;
    /**
     * Get the term label by id
     * @param termId the term id
     */
    getTermLabelsById(termId: string): Promise<SP.Taxonomy.Label[]>;
    /**
     * Get labels of given terms
     * @param terms the terms to find their labels
     */
    getLabelsForTerms(terms: SP.Taxonomy.Term[]): Promise<SP.Taxonomy.LabelCollection[]>;
    /**
     * Get all child terms under a specific term as a flat array
     * @param termId the term id
     * @param list initial array defaults to []
     */
    getTermsSubTreeFlat(termId: string, list?: SP.Taxonomy.Term[]): Promise<SP.Taxonomy.Term[]>;
    /**
     * get all term sets in a group, create the group if missing
     * @param createIfMissing whether to create the group if it cannot be found
     */
    getAllTermSetsInSiteCollectionGroup(createIfMissing: boolean): Promise<SP.Taxonomy.TermSet[]>;
    /**
     * get all term labels
     * @param term the term to get its labels (different locales if any)
     */
    getTermLabels(term: SP.Taxonomy.Term): Promise<SP.Taxonomy.Label[]>;
    /**
     * Get term by id
     * @param id the term id
     */
    getTermById(id: string): Promise<SP.Taxonomy.Term>;
    /**
     * Get site collection term group
     * @param createIfMissing whether to create site collection term group or not
     */
    getSiteCollectionTermGroup(createIfMissing: any): Promise<SP.Taxonomy.TermGroup>;
}
export declare type TermStoreContextCallback<T> = (ctx: SP.ClientContext, session: SP.Taxonomy.TaxonomySession, tstore: SP.Taxonomy.TermStore, execute: (result: T) => void) => void;
export declare function createTermStore(): ITermStore;
