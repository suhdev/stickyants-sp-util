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
