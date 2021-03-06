import { executeOnContext } from "./execute";
import * as find from 'lodash.find';
import * as startsWith from 'lodash.startswith';
import * as some from 'lodash.some'; 

/**
 * Serializes a TermSet to a conforming JSON that can be used to initialize SP.Taxonomy.TermSet 
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the termset to serialize
 */
export function termSetToJson(tset:SP.Taxonomy.TermSet){
    return {
        Id:tset.get_id().toString(), 
        Name:tset.get_name(), 
        LastModifiedDate:tset.get_lastModifiedDate(), 
        CreatedDate:tset.get_createdDate(),
        Contact:tset.get_contact(), 
        Description:tset.get_description(),
        IsOpenForTermCreation:tset.get_isOpenForTermCreation(), 
        Stakeholders:tset.get_stakeholders(),
        CustomSortOrder:tset.get_customSortOrder(), 
        CustomProperties:tset.get_customProperties(), 
        IsAvailableForTagging:tset.get_isAvailableForTagging(), 
        Owner:tset.get_owner(), 
    };
}

/**
 * Serializes a Term to a conforming JSON that can be used to initialize SP.Taxonomy.Term 
 * This is useful to store the term in a cahce, e.g. IndexedDB
 * @param tset the term to serialize
 */
export function termToJson(term:SP.Taxonomy.Term){
    return {
        Id:term.get_id().toString(), 
        Name:term.get_name(), 
        Description:term.get_description(),
        IsDeprecated:term.get_isDeprecated(), 
        IsKeyword:term.get_isKeyword(), 
        IsPinned:term.get_isPinned(), 
        IsReused:term.get_isReused(), 
        IsRoot:term.get_isRoot(), 
        PathOfTerm:term.get_pathOfTerm(), 
        TermsCount:term.get_termsCount(), 
        IsSourceTerm:term.get_isSourceTerm(), 
        IsAvailableForTagging:term.get_isAvailableForTagging(), 
        LocalCustomProperties:term.get_localCustomProperties(), 
        CustomProperties:term.get_customProperties(),
        CustomSortOrder:term.get_customSortOrder(), 
    }; 
}

export interface IValue {
    id:string|number;
    label:string; 
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
    createTerm(parentTerm:SP.Taxonomy.Term, name:string, locale:number, guid:SP.Guid, properties?:IValue[]):Promise<SP.Taxonomy.Term>; 
    /**
     * Get the parent term by term id 
     * @param termId the term id to find its parent
     */
    getParentTermById(termId:string):Promise<SP.Taxonomy.Term>; 
    /**
     * Get terms given their guids
     * @param termIds array of term guids (string)
     */
    getTerms(termIds:string[]):Promise<SP.Taxonomy.Term[]>;
    /**
     * get child terms of term given the parent term id
     * @param termId parent term id 
     */
    getTermsByTermId(termId:string):Promise<SP.Taxonomy.Term[]>;
    /**
     * Get immediate child terms given term set id
     * @param termSetId the term set id 
     */
    getTermsByTermSetId(termSetId:string):Promise<SP.Taxonomy.Term[]>;  
    /**
     * All child terms of given termset id 
     * @param termSetId the term set id 
     */
    getAllTermsByTermSetId(termSetId: string):Promise<SP.Taxonomy.Term[]>;
    /**
     * Get term set id from taxonomy field name 
     */
    termSetIdFromTaxonomyField(fieldInternalName:string):Promise<string>;
    /**
     * get terms collection given term ids 
     * @param ids term ids
     * @see {ITermStore#getTerms}
     */
    getTermsByIds(ids:string[]):Promise<SP.Taxonomy.TermCollection>; 
    /**
     * Get parent term by term
     * @param term {SP.Taxonomy.Term} the term to get its parent
     */
    getParentTermByTerm(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Term>; 
    /**
     * Get top level parent of a term. 
     * This method executes recursively until it finds the top level parent of a
     * term. 
     * @param id the term id
     */
    getTopLevelParentOfTerm(id:string):Promise<SP.Taxonomy.Term>;
    /**
     * Get term parents (all term parents) given a term id
     * @param termId the term id to find its parents
     */
    getTermParents(termId:string):Promise<SP.Taxonomy.Term[]>;
    /**
     * Get term parents that satisfy a specific condition (using a callback)
     * @param id the id of the term 
     * @param fn the callback 
     */
    getParentThatSatisfies(id:string,fn:(v:SP.Taxonomy.Term)=>boolean):Promise<SP.Taxonomy.Term>;
    /**
     * Get the term label by id 
     * @param termId the term id 
     */
    getTermLabelsById(termId:string):Promise<SP.Taxonomy.Label[]>;
    /**
     * Get labels of given terms 
     * @param terms the terms to find their labels 
     */
    getLabelsForTerms(terms:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.LabelCollection[]>; 
    /**
     * Get all child terms under a specific term as a flat array
     * @param termId the term id 
     * @param list initial array defaults to [] 
     */
    getTermsSubTreeFlat(termId:string,list?:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.Term[]>;
    /**
     * get all term sets in a group, create the group if missing 
     * @param createIfMissing whether to create the group if it cannot be found 
     */
    getAllTermSetsInSiteCollectionGroup(createIfMissing:boolean):Promise<SP.Taxonomy.TermSet[]>;
    /**
     * get all term labels 
     * @param term the term to get its labels (different locales if any)
     */
    getTermLabels(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Label[]>;
    /**
     * Get term by id
     * @param id the term id 
     */
    getTermById(id:string):Promise<SP.Taxonomy.Term>;
    /**
     * Get site collection term group 
     * @param createIfMissing whether to create site collection term group or not
     */
    getSiteCollectionTermGroup(createIfMissing):Promise<SP.Taxonomy.TermGroup>;
}

export type TermStoreContextCallback<T> = (ctx: SP.ClientContext, session: SP.Taxonomy.TaxonomySession, tstore: SP.Taxonomy.TermStore, execute: (result: T) => void) => void; 

export function createTermStore():ITermStore{
    var cache = {
        termSetByTermId:{}, 
        termsByTermSetId:{},
        parentsByTermId:{}
    };

    function createExecutionContext<T>(fn: TermStoreContextCallback<T>,siteUrl?:string){
        let ctx = new SP.ClientContext(siteUrl||_spPageContextInfo.siteAbsoluteUrl); 
        let session = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx); 
        let tstore = session.getDefaultSiteCollectionTermStore(); 
        return new Promise<T>((resolve,reject)=>{
            fn(ctx, session, tstore, function (result:any) {
                ctx.executeQueryAsync(()=>{
                    resolve(result); 
                }, reject);
            }); 
        });
    }

    function getTermsByTermSetId(termSetId:string):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx,session,store,execute)=>{
            let tset = store.getTermSet(new SP.Guid(termSetId));
            let terms = tset.get_terms();
            ctx.load(terms); 
            execute(terms);
        })
        .then((terms:SP.Taxonomy.TermCollection)=>{
            return terms.get_data(); 
        }); 
    }

    function getTermsByTermId(termId:string):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let tset = store.getTerm(new SP.Guid(termId));
            let terms = tset.get_terms();
            ctx.load(terms);
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return terms.get_data();
        }); 
    }

    function getTermParentPaths(path:string){
        var paths = path.split(';');
        paths.pop(); 
        var prev = ''; 
        var pp = []; 
        for(var p of paths){
            prev = prev?prev +';'+p:p; 
            pp.push(prev); 
        }
        return pp; 
    }

    async function getTermParents(termId:string):Promise<SP.Taxonomy.Term[]>{
        var termPaths:string[] = []; 
        if (cache.parentsByTermId[termId]){
            return cache.parentsByTermId[termId]; 
        }
        return createExecutionContext<SP.Taxonomy.Term[]>(async (ctx, session, store, execute) => {
            let term = store.getTerm(new SP.Guid(termId));
            let tset = term.get_termSet(); 
            let terms = tset.getAllTerms(); 
            ctx.load(terms); 
            ctx.load(term); 
            await executeOnContext(ctx);
            termPaths = getTermParentPaths(term.get_pathOfTerm()); 
            var parents = terms.get_data().filter(e=>{
                return some(termPaths,v=>v === e.get_pathOfTerm()); 
            }); 
            execute(cache.parentsByTermId[termId] = parents);
        });
    }

    function getTermsByTermIds(...termIds:string[]):Promise<SP.Taxonomy.Term[][]>{
        var taxTerms:SP.Taxonomy.TermCollection[] = []; 
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let terms:SP.Taxonomy.TermCollection; 
            for(var termId of termIds){
                let tset = store.getTerm(new SP.Guid(termId));
                terms = tset.get_terms();
                taxTerms.push(terms); 
                ctx.load(terms);
            }
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return taxTerms.map(e=>e.get_data());
        }); 
    }

    function getTermSetByTermId(termId:string){
        return createExecutionContext<SP.Taxonomy.TermSet>((ctx, session, store, execute) => {
            var term = store.getTerm(new SP.Guid(termId)); 
            var tset = term.get_termSet(); 
            ctx.load(tset); 
            execute(tset);
        }); 
    }

    async function getTermsSubTreeFlat(termId:string,list:SP.Taxonomy.Term[] = []):Promise<SP.Taxonomy.Term[]>{
        var isTermSetCached = cache.termSetByTermId[termId]?true:false; 
        let termset:SP.Taxonomy.TermSet = cache.termSetByTermId[termId] || await getTermSetByTermId(termId); 
        cache.termSetByTermId[termId] = termset; 
        var isTermsCached = cache.termsByTermSetId[termset.get_id().toString()]?true:false; 
        let terms:SP.Taxonomy.Term[] = cache.termsByTermSetId[termset.get_id().toString()] || await new Promise((res,rej)=>{
            var ctx = termset.get_context();
            var terms = termset.getAllTerms();
            ctx.load(terms);
            ctx.executeQueryAsync(()=>{
                res(terms.get_data()); 
            },(c,err)=>{
                rej(err); 
            })
        }); 
        cache.termsByTermSetId[termset.get_id().toString()] = terms; 
        if (!isTermsCached){
            for(var t of terms){
                cache.termSetByTermId[t.get_id().toString()] = termset; 
            }
        }
        var parentTerm = find(terms,e=>e.get_id().toString() === termId); 
        if (parentTerm){
            return terms.filter(e=> startsWith(e.get_pathOfTerm(),parentTerm.get_pathOfTerm())); 
        }
        return [];
    }

    function getAllTermsByTermSetId(termSetId: string) {
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx, session, store, execute) => {
            let tset = store.getTermSet(new SP.Guid(termSetId));
            let terms = tset.getAllTerms();
            ctx.load(terms);
            execute(terms);
        })
        .then((terms: SP.Taxonomy.TermCollection) => {
            return terms.get_data();
        }); 
    }

    function getAllTermSetsInSiteCollectionGroup(createIfMissing:boolean = false):Promise<SP.Taxonomy.TermSet[]>{
        return createExecutionContext<SP.Taxonomy.TermSetCollection>((ctx, session, store, execute) => {
            let groups = store.getSiteCollectionGroup(ctx.get_site(),createIfMissing); 
            let terms = groups.get_termSets();
            ctx.load(terms);
            execute(terms); 
        },_spPageContextInfo.siteAbsoluteUrl)
        .then((terms: SP.Taxonomy.TermSetCollection) => {
            return terms.get_data();
        });
    }

    function getTermLabels(term:SP.Taxonomy.Term):Promise<SP.Taxonomy.Label[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection>((ctx,session,store,execute)=>{
            let labels = term.get_labels(); 
            ctx.load(labels); 
            execute(labels); 
        })
        .then((labels:SP.Taxonomy.LabelCollection)=>{
            return labels.get_data();
        }); 
    }

    function getTermLabelsById(termId:string):Promise<SP.Taxonomy.Label[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(termId)); 
            let labels = term.get_labels(); 
            ctx.load(labels); 
            execute(labels); 
        })
        .then((labels:SP.Taxonomy.LabelCollection)=>{
            return labels.get_data();
        });
    }

    function getSiteCollectionTermGroup(createIfMissing:boolean = false):Promise<SP.Taxonomy.TermGroup>{
        return createExecutionContext<SP.Taxonomy.TermGroup>((ctx,session,store,execute)=>{
            let group:SP.Taxonomy.TermGroup = store.getSiteCollectionGroup(ctx.get_site(),createIfMissing);
            ctx.load(group);
            execute(group); 
        },_spPageContextInfo.siteAbsoluteUrl);
    }

    function getLabelsForTerms(terms:SP.Taxonomy.Term[]):Promise<SP.Taxonomy.LabelCollection[]>{
        return createExecutionContext<SP.Taxonomy.LabelCollection[]>((ctx,session,store,execute)=>{
            let labels:SP.Taxonomy.LabelCollection[] = [];
            terms.forEach((term)=>{
                let l = term.get_labels(); 
                ctx.load(l); 
                labels.push(l); 
            });
            execute(labels); 
        });
    }

    function getTerms(termIds:string[]):Promise<SP.Taxonomy.Term[]>{
        return createExecutionContext<SP.Taxonomy.Term[]>((ctx,session,store,execute)=>{
            let terms:SP.Taxonomy.Term[] = [];
            termIds.forEach((id)=>{
                let t = store.getTerm(new SP.Guid(id)); 
                ctx.load(t); 
                terms.push(t); 
            });
            execute(terms); 
        });
    }

    async function getTopLevelParentOfTerm(id:string):Promise<SP.Taxonomy.Term>{
        try {
            var term = await getTermById(id);
            var path = term.get_pathOfTerm().split(';'); 
            if (path.length === 1){
                return term; 
            }else {
                path.pop(); 
                var parent:SP.Taxonomy.Term = term; 
                for(var i= 0; i<path.length;i++){
                    parent = parent.get_parent(); 
                }
                var ctx = parent.get_context(); 
                ctx.load(parent); 
                return new Promise<SP.Taxonomy.Term>((res,rej)=>{
                    ctx.executeQueryAsync(()=>{
                        res(parent); 
                    },(c,err)=>{
                        rej(err); 
                    });
                });
            }
        }catch(err){
            throw err; 
        }
    }

    async function getParentThatSatisfies(id:string,fn:(v:SP.Taxonomy.Term)=>boolean):Promise<SP.Taxonomy.Term>{
        try {
            var term = await getTermById(id);
            var path = term.get_pathOfTerm().split(';'); 
            if (path.length === 1){
                return term; 
            }else {
                path.pop(); 
                var parent:SP.Taxonomy.Term = term; 
                for(var i= 0; i<path.length;i++){
                    var parent = await getParentTermByTerm(parent); 
                    if (parent && fn(parent)){
                        return parent; 
                    }
                }
                return null; 
            }
        }catch(err){
            throw err; 
        }
    }

    
    function termSetIdFromTaxonomyField(fieldInternalName:string):Promise<string>{
        return new Promise<string>((res,rej)=>{
            let ctx = new SP.ClientContext(); 
            let web = ctx.get_web(); 
            let fields = web.get_availableFields(); 
            let field:SP.Taxonomy.TaxonomyField = fields.getByInternalNameOrTitle(fieldInternalName) as any; 
            ctx.load(field); 
            ctx.executeQueryAsync(()=>{
                res(field.get_termSetId().toString()); 
            },rej); 
        }); 
    }

    function getParentTermById(id:string){
        return createExecutionContext<SP.Taxonomy.Term>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(id)); 
            let parent = term.get_parent(); 
            ctx.load(parent);  
            execute(parent);
        });
    }

    function getParentTermByTerm(term:SP.Taxonomy.Term){
        return new Promise<SP.Taxonomy.Term>((res,rej)=>{
            var ctx = term.get_context(); 
            var parent = term.get_parent(); 
            ctx.load(parent); 
            ctx.executeQueryAsync(()=>{
                res(parent); 
            },(c,err)=>{
                rej(err); 
            })
        }); 
    }

    function getTermById(id:string){
        return createExecutionContext<SP.Taxonomy.Term>((ctx,session,store,execute)=>{
            let term = store.getTerm(new SP.Guid(id)); 
            ctx.load(term); ; 
            execute(term);
        });
    }

    function createTerm(parentTerm:SP.Taxonomy.Term, name:string, locale:number, guid:SP.Guid, properties:IValue[] = []){
        return new Promise<SP.Taxonomy.Term>((res,rej)=>{
            var ctx = parentTerm.get_context(); 
            var store = parentTerm.get_termStore(); 
            let term = parentTerm.createTerm(name,locale,guid);
            properties.forEach((e)=>{
                term.setLocalCustomProperty(e.id as string,e.label);
            });
            
            store.commitAll();
            ctx.executeQueryAsync(()=>{
                res(term); 
            },(c,err)=>rej(err)); 
        });
    }

    function getTermsByIds(ids:string[]){
        return createExecutionContext<SP.Taxonomy.TermCollection>((ctx,session,store,execute)=>{
            let terms = store.getTermsById(ids.map(e=>new SP.Guid(e)));
            ctx.load(terms); ; 
            execute(terms);
        })
    }

    return {
        createTerm, 
        getTerms,
        getTermsByTermId,
        getTermsByTermSetId, 
        getAllTermsByTermSetId, 
        getTermsByIds,
        getTermById,
        getTermLabelsById,
        getTermParents,
        getParentTermByTerm,
        getTermsSubTreeFlat,
        getTopLevelParentOfTerm,
        termSetIdFromTaxonomyField,
        getParentThatSatisfies,
        getAllTermSetsInSiteCollectionGroup,
        getLabelsForTerms,
        getTermLabels,
        getSiteCollectionTermGroup,
        getParentTermById
    }; 
}

