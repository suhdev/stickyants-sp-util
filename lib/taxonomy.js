"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Serializes a TermSet to a conforming JSON that can be used to initialize SP.Taxonomy.TermSet
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the termset to serialize
 */
function termSetToJson(tset) {
    return {
        Id: tset.get_id().toString(),
        Name: tset.get_name(),
        LastModifiedDate: tset.get_lastModifiedDate(),
        CreatedDate: tset.get_createdDate(),
        Contact: tset.get_contact(),
        Description: tset.get_description(),
        IsOpenForTermCreation: tset.get_isOpenForTermCreation(),
        Stakeholders: tset.get_stakeholders(),
        CustomSortOrder: tset.get_customSortOrder(),
        CustomProperties: tset.get_customProperties(),
        IsAvailableForTagging: tset.get_isAvailableForTagging(),
        Owner: tset.get_owner(),
    };
}
exports.termSetToJson = termSetToJson;
/**
 * Serializes a Term to a conforming JSON that can be used to initialize SP.Taxonomy.Term
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the term to serialize
 */
function termToJson(term) {
    return {
        Id: term.get_id().toString(),
        Name: term.get_name(),
        Description: term.get_description(),
        IsDeprecated: term.get_isDeprecated(),
        IsKeyword: term.get_isKeyword(),
        IsPinned: term.get_isPinned(),
        IsReused: term.get_isReused(),
        IsRoot: term.get_isRoot(),
        PathOfTerm: term.get_pathOfTerm(),
        TermsCount: term.get_termsCount(),
        IsSourceTerm: term.get_isSourceTerm(),
        IsAvailableForTagging: term.get_isAvailableForTagging(),
        LocalCustomProperties: term.get_localCustomProperties(),
        CustomProperties: term.get_customProperties(),
        CustomSortOrder: term.get_customSortOrder(),
    };
}
exports.termToJson = termToJson;
