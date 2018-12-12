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
/**
 * Extracts the image url from `PublishingPageImage` field value
 * @param str the `PublishingPageImage` field value
 */
function extractUrlFromPublishingPageImage(str) {
    var k = null;
    str && str.replace(/src="([\s\S]+?)"/, function (e, v) {
        return k = v;
    });
    return k;
}
exports.extractUrlFromPublishingPageImage = extractUrlFromPublishingPageImage;
/**
 * Get page layout by name/title
 * @param ctx the SP.ClientContext to execute this request on
 * @param name the name/Title of the page layout
 */
function getPageLayoutItemByName(ctx, name) {
    return __awaiter(this, void 0, void 0, function () {
        var rootWeb, list, query, layouts, data;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    rootWeb = ctx.get_site().get_rootWeb();
                    list = rootWeb.get_lists().getByTitle('Master Page Gallery');
                    query = new SP.CamlQuery();
                    query.set_viewXml("<View><Query><Where><Eq><FieldRef Name=\"Title\" /><Value Type=\"Text\">" + name + "</Value></Eq></Where></Query></View>");
                    layouts = list.getItems(query);
                    ctx.load(layouts);
                    return [4 /*yield*/, execute_1.executeOnContext(ctx)];
                case 1:
                    _a.sent();
                    data = layouts.get_data();
                    if (data.length) {
                        return [2 /*return*/, data[0]];
                    }
                    return [2 /*return*/, null];
            }
        });
    });
}
exports.getPageLayoutItemByName = getPageLayoutItemByName;
/**
 * Get page layout item by the associated content type name
 * @param ctx the context to execute on
 * @param name the associated content type name
 */
function getPageLayoutItemByAssociatedContentType(ctx, name) {
    return __awaiter(this, void 0, void 0, function () {
        var rootWeb, list, query, layouts;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    rootWeb = ctx.get_site().get_rootWeb();
                    list = rootWeb.get_lists().getByTitle('Master Page Gallery');
                    query = new SP.CamlQuery();
                    query.set_viewXml("<View><Query><Where><Contains><FieldRef Name=\"PublishingAssociatedContentType\" /><Value Type=\"Text\">" + name + "</Value></Contains></Where></Query></View>");
                    layouts = list.getItems(query);
                    ctx.load(layouts);
                    return [4 /*yield*/, execute_1.executeOnContext(ctx)];
                case 1:
                    _a.sent();
                    if (layouts.get_data().length) {
                        return [2 /*return*/, layouts.get_data()[0]];
                    }
                    else {
                        return [2 /*return*/, null];
                    }
                    return [2 /*return*/];
            }
        });
    });
}
exports.getPageLayoutItemByAssociatedContentType = getPageLayoutItemByAssociatedContentType;
/**
 * Creates a publishing page given the title, filename and page layout item
 * @param ctx {SP.ClientContext} the client context to execute on
 * @param title the title of the page
 * @param fileName the file name of the page
 * @param layout {SP.ListItem} the page layout item
 *
 * @see {getPageLayoutItemByAssociatedContentType}
 */
function createPublishingPageWithLayout(ctx, fileName, layout) {
    if (typeof fileName !== "string" || !fileName || !fileName.trim()) {
        throw new Error("Invalid file name provided, expected truthy string but got " + typeof fileName + " - '" + fileName + "'");
    }
    var web = ctx.get_web();
    var newPublishingPage = SP.Publishing.PublishingWeb.getPublishingWeb(ctx, web);
    var pageInfo = new SP.Publishing.PublishingPageInformation();
    fileName = fileName.replace(/\.aspx$/i, '') + '.aspx';
    pageInfo.set_name(fileName);
    pageInfo.set_pageLayoutListItem(layout);
    var newPage = newPublishingPage.addPublishingPage(pageInfo);
    ctx.load(newPage);
    return execute_1.executeOnContext(ctx, newPage);
}
exports.createPublishingPageWithLayout = createPublishingPageWithLayout;
/**
 * Creates a publishing page given the context, page file name, and the content type name
 * @param ctx the client context to execute on
 * @param fileName the page file name
 * @param contentType the content type name
 */
function createPublishingPage(ctx, fileName, contentType) {
    return __awaiter(this, void 0, void 0, function () {
        var layout;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getPageLayoutItemByAssociatedContentType(ctx, contentType)];
                case 1:
                    layout = _a.sent();
                    return [2 /*return*/, createPublishingPageWithLayout(ctx, fileName, layout)];
            }
        });
    });
}
exports.createPublishingPage = createPublishingPage;
