"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var execute_1 = require("./execute");
var find = require("lodash.find");
var startsWith = require("lodash.startswith");
var some = require("lodash.some");
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
function createTermStore() {
    var cache = {
        termSetByTermId: {},
        termsByTermSetId: {},
        parentsByTermId: {}
    };
    function createExecutionContext(fn, siteUrl) {
        var ctx = new SP.ClientContext(siteUrl || _spPageContextInfo.siteAbsoluteUrl);
        var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);
        var tstore = session.getDefaultSiteCollectionTermStore();
        return new Promise(function (resolve, reject) {
            fn(ctx, session, tstore, function (result) {
                ctx.executeQueryAsync(function () {
                    resolve(result);
                }, reject);
            });
        });
    }
    function getTermsByTermSetId(termSetId) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var tset = store.getTermSet(new SP.Guid(termSetId));
            var terms = tset.get_terms();
            ctx.load(terms);
            execute(terms);
        })
            .then(function (terms) {
            return terms.get_data();
        });
    }
    function getTermsByTermId(termId) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var tset = store.getTerm(new SP.Guid(termId));
            var terms = tset.get_terms();
            ctx.load(terms);
            execute(terms);
        })
            .then(function (terms) {
            return terms.get_data();
        });
    }
    function getTermParentPaths(path) {
        var paths = path.split(';');
        paths.pop();
        var prev = '';
        var pp = [];
        for (var _i = 0, paths_1 = paths; _i < paths_1.length; _i++) {
            var p = paths_1[_i];
            prev = prev ? prev + ';' + p : p;
            pp.push(prev);
        }
        return pp;
    }
    function getTermParents(termId) {
        return __awaiter(this, void 0, void 0, function () {
            var termPaths;
            var _this = this;
            return __generator(this, function (_a) {
                termPaths = [];
                if (cache.parentsByTermId[termId]) {
                    return [2 /*return*/, cache.parentsByTermId[termId]];
                }
                return [2 /*return*/, createExecutionContext(function (ctx, session, store, execute) { return __awaiter(_this, void 0, void 0, function () {
                        var term, tset, terms, parents;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    term = store.getTerm(new SP.Guid(termId));
                                    tset = term.get_termSet();
                                    terms = tset.getAllTerms();
                                    ctx.load(terms);
                                    ctx.load(term);
                                    return [4 /*yield*/, execute_1.executeOnContext(ctx)];
                                case 1:
                                    _a.sent();
                                    termPaths = getTermParentPaths(term.get_pathOfTerm());
                                    parents = terms.get_data().filter(function (e) {
                                        return some(termPaths, function (v) { return v === e.get_pathOfTerm(); });
                                    });
                                    execute(cache.parentsByTermId[termId] = parents);
                                    return [2 /*return*/];
                            }
                        });
                    }); })];
            });
        });
    }
    function getTermsByTermIds() {
        var termIds = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            termIds[_i] = arguments[_i];
        }
        var taxTerms = [];
        return createExecutionContext(function (ctx, session, store, execute) {
            var terms;
            for (var _i = 0, termIds_1 = termIds; _i < termIds_1.length; _i++) {
                var termId = termIds_1[_i];
                var tset = store.getTerm(new SP.Guid(termId));
                terms = tset.get_terms();
                taxTerms.push(terms);
                ctx.load(terms);
            }
            execute(terms);
        })
            .then(function (terms) {
            return taxTerms.map(function (e) { return e.get_data(); });
        });
    }
    function getTermSetByTermId(termId) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var term = store.getTerm(new SP.Guid(termId));
            var tset = term.get_termSet();
            ctx.load(tset);
            execute(tset);
        });
    }
    function getTermsSubTreeFlat(termId, list) {
        if (list === void 0) { list = []; }
        return __awaiter(this, void 0, void 0, function () {
            var isTermSetCached, termset, _a, isTermsCached, terms, _b, _i, terms_1, t, parentTerm;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        isTermSetCached = cache.termSetByTermId[termId] ? true : false;
                        _a = cache.termSetByTermId[termId];
                        if (_a) return [3 /*break*/, 2];
                        return [4 /*yield*/, getTermSetByTermId(termId)];
                    case 1:
                        _a = (_c.sent());
                        _c.label = 2;
                    case 2:
                        termset = _a;
                        cache.termSetByTermId[termId] = termset;
                        isTermsCached = cache.termsByTermSetId[termset.get_id().toString()] ? true : false;
                        _b = cache.termsByTermSetId[termset.get_id().toString()];
                        if (_b) return [3 /*break*/, 4];
                        return [4 /*yield*/, new Promise(function (res, rej) {
                                var ctx = termset.get_context();
                                var terms = termset.getAllTerms();
                                ctx.load(terms);
                                ctx.executeQueryAsync(function () {
                                    res(terms.get_data());
                                }, function (c, err) {
                                    rej(err);
                                });
                            })];
                    case 3:
                        _b = (_c.sent());
                        _c.label = 4;
                    case 4:
                        terms = _b;
                        cache.termsByTermSetId[termset.get_id().toString()] = terms;
                        if (!isTermsCached) {
                            for (_i = 0, terms_1 = terms; _i < terms_1.length; _i++) {
                                t = terms_1[_i];
                                cache.termSetByTermId[t.get_id().toString()] = termset;
                            }
                        }
                        parentTerm = find(terms, function (e) { return e.get_id().toString() === termId; });
                        if (parentTerm) {
                            return [2 /*return*/, terms.filter(function (e) { return startsWith(e.get_pathOfTerm(), parentTerm.get_pathOfTerm()); })];
                        }
                        return [2 /*return*/, []];
                }
            });
        });
    }
    function getAllTermsByTermSetId(termSetId) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var tset = store.getTermSet(new SP.Guid(termSetId));
            var terms = tset.getAllTerms();
            ctx.load(terms);
            execute(terms);
        })
            .then(function (terms) {
            return terms.get_data();
        });
    }
    function getAllTermSetsInSiteCollectionGroup(createIfMissing) {
        if (createIfMissing === void 0) { createIfMissing = false; }
        return createExecutionContext(function (ctx, session, store, execute) {
            var groups = store.getSiteCollectionGroup(ctx.get_site(), createIfMissing);
            var terms = groups.get_termSets();
            ctx.load(terms);
            execute(terms);
        }, _spPageContextInfo.siteAbsoluteUrl)
            .then(function (terms) {
            return terms.get_data();
        });
    }
    function getTermLabels(term) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var labels = term.get_labels();
            ctx.load(labels);
            execute(labels);
        })
            .then(function (labels) {
            return labels.get_data();
        });
    }
    function getTermLabelsById(termId) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var term = store.getTerm(new SP.Guid(termId));
            var labels = term.get_labels();
            ctx.load(labels);
            execute(labels);
        })
            .then(function (labels) {
            return labels.get_data();
        });
    }
    function getSiteCollectionTermGroup(createIfMissing) {
        if (createIfMissing === void 0) { createIfMissing = false; }
        return createExecutionContext(function (ctx, session, store, execute) {
            var group = store.getSiteCollectionGroup(ctx.get_site(), createIfMissing);
            ctx.load(group);
            execute(group);
        }, _spPageContextInfo.siteAbsoluteUrl);
    }
    function getLabelsForTerms(terms) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var labels = [];
            terms.forEach(function (term) {
                var l = term.get_labels();
                ctx.load(l);
                labels.push(l);
            });
            execute(labels);
        });
    }
    function getTerms(termIds) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var terms = [];
            termIds.forEach(function (id) {
                var t = store.getTerm(new SP.Guid(id));
                ctx.load(t);
                terms.push(t);
            });
            execute(terms);
        });
    }
    function getTopLevelParentOfTerm(id) {
        return __awaiter(this, void 0, void 0, function () {
            var term, path, parent, i, ctx, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, getTermById(id)];
                    case 1:
                        term = _a.sent();
                        path = term.get_pathOfTerm().split(';');
                        if (path.length === 1) {
                            return [2 /*return*/, term];
                        }
                        else {
                            path.pop();
                            parent = term;
                            for (i = 0; i < path.length; i++) {
                                parent = parent.get_parent();
                            }
                            ctx = parent.get_context();
                            ctx.load(parent);
                            return [2 /*return*/, new Promise(function (res, rej) {
                                    ctx.executeQueryAsync(function () {
                                        res(parent);
                                    }, function (c, err) {
                                        rej(err);
                                    });
                                })];
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        err_1 = _a.sent();
                        throw err_1;
                    case 3: return [2 /*return*/];
                }
            });
        });
    }
    function getParentThatSatisfies(id, fn) {
        return __awaiter(this, void 0, void 0, function () {
            var term, path, parent, i, parent, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        return [4 /*yield*/, getTermById(id)];
                    case 1:
                        term = _a.sent();
                        path = term.get_pathOfTerm().split(';');
                        if (!(path.length === 1)) return [3 /*break*/, 2];
                        return [2 /*return*/, term];
                    case 2:
                        path.pop();
                        parent = term;
                        i = 0;
                        _a.label = 3;
                    case 3:
                        if (!(i < path.length)) return [3 /*break*/, 6];
                        return [4 /*yield*/, getParentTermByTerm(parent)];
                    case 4:
                        parent = _a.sent();
                        if (parent && fn(parent)) {
                            return [2 /*return*/, parent];
                        }
                        _a.label = 5;
                    case 5:
                        i++;
                        return [3 /*break*/, 3];
                    case 6: return [2 /*return*/, null];
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        err_2 = _a.sent();
                        throw err_2;
                    case 9: return [2 /*return*/];
                }
            });
        });
    }
    function termSetIdFromTaxonomyField(fieldInternalName) {
        return new Promise(function (res, rej) {
            var ctx = new SP.ClientContext();
            var web = ctx.get_web();
            var fields = web.get_availableFields();
            var field = fields.getByInternalNameOrTitle(fieldInternalName);
            ctx.load(field);
            ctx.executeQueryAsync(function () {
                res(field.get_termSetId().toString());
            }, rej);
        });
    }
    function getParentTermById(id) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var term = store.getTerm(new SP.Guid(id));
            var parent = term.get_parent();
            ctx.load(parent);
            execute(parent);
        });
    }
    function getParentTermByTerm(term) {
        return new Promise(function (res, rej) {
            var ctx = term.get_context();
            var parent = term.get_parent();
            ctx.load(parent);
            ctx.executeQueryAsync(function () {
                res(parent);
            }, function (c, err) {
                rej(err);
            });
        });
    }
    function getTermById(id) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var term = store.getTerm(new SP.Guid(id));
            ctx.load(term);
            ;
            execute(term);
        });
    }
    function createTerm(parentTerm, name, locale, guid, properties) {
        if (properties === void 0) { properties = []; }
        return new Promise(function (res, rej) {
            var ctx = parentTerm.get_context();
            var store = parentTerm.get_termStore();
            var term = parentTerm.createTerm(name, locale, guid);
            properties.forEach(function (e) {
                term.setLocalCustomProperty(e.id, e.label);
            });
            store.commitAll();
            ctx.executeQueryAsync(function () {
                res(term);
            }, function (c, err) { return rej(err); });
        });
    }
    function getTermsByIds(ids) {
        return createExecutionContext(function (ctx, session, store, execute) {
            var terms = store.getTermsById(ids.map(function (e) { return new SP.Guid(e); }));
            ctx.load(terms);
            ;
            execute(terms);
        });
    }
    return {
        createTerm: createTerm,
        getTerms: getTerms,
        getTermsByTermId: getTermsByTermId,
        getTermsByTermSetId: getTermsByTermSetId,
        getAllTermsByTermSetId: getAllTermsByTermSetId,
        getTermsByIds: getTermsByIds,
        getTermById: getTermById,
        getTermLabelsById: getTermLabelsById,
        getTermParents: getTermParents,
        getParentTermByTerm: getParentTermByTerm,
        getTermsSubTreeFlat: getTermsSubTreeFlat,
        getTopLevelParentOfTerm: getTopLevelParentOfTerm,
        termSetIdFromTaxonomyField: termSetIdFromTaxonomyField,
        getParentThatSatisfies: getParentThatSatisfies,
        getAllTermSetsInSiteCollectionGroup: getAllTermSetsInSiteCollectionGroup,
        getLabelsForTerms: getLabelsForTerms,
        getTermLabels: getTermLabels,
        getSiteCollectionTermGroup: getSiteCollectionTermGroup,
        getParentTermById: getParentTermById
    };
}
exports.createTermStore = createTermStore;
