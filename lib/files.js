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
 * Get root folder of a list
 * @param ctx the client context to execute on
 * @param list the list to get its root folder
 * @returns {Promise<SP.Folder}
 */
function getRootFolder(ctx, list) {
    var folder = list.get_rootFolder();
    ctx.load(folder);
    return execute_1.executeOnContext(ctx, folder);
}
exports.getRootFolder = getRootFolder;
/**
 * Get list parent web server relative url
 * @param ctx the client context to execute on
 * @param list the list to get its parent web url
 * @returns {Promise<string>}
 */
function getListParentWebUrl(ctx, list) {
    return __awaiter(this, void 0, void 0, function () {
        var web;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    web = list.get_parentWeb();
                    ctx.load(web);
                    return [4 /*yield*/, execute_1.executeOnContext(ctx)];
                case 1:
                    _a.sent();
                    return [2 /*return*/, web.get_serverRelativeUrl()];
            }
        });
    });
}
exports.getListParentWebUrl = getListParentWebUrl;
/**
 * Checks if a folder exists in a list or not
 * @param ctx the client context to execute on
 * @param list the list to check the folder path on
 * @param folderName the folder name
 * @returns {Promise<boolean>}
 */
function folderExistsInList(ctx, list, folderName) {
    return __awaiter(this, void 0, void 0, function () {
        var rootFolder, url, rf, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRootFolder(ctx, list)];
                case 1:
                    rootFolder = _a.sent();
                    url = rootFolder.get_serverRelativeUrl();
                    rf = ctx.get_web().getFolderByServerRelativeUrl(url + "/" + folderName);
                    ctx.load(rf);
                    _a.label = 2;
                case 2:
                    _a.trys.push([2, 4, , 5]);
                    return [4 /*yield*/, execute_1.executeOnContext(ctx)];
                case 3:
                    _a.sent();
                    return [2 /*return*/, true];
                case 4:
                    err_1 = _a.sent();
                    return [2 /*return*/, false];
                case 5: return [2 /*return*/];
            }
        });
    });
}
exports.folderExistsInList = folderExistsInList;
/**
 * Get folder in a list by folder name
 * @param ctx the client context to execute on
 * @param list the list to get the folder in
 * @param folderName the folder name
 * @returns {Promise<SP.Folder>}
 */
function getFolderInListByName(ctx, list, folderName) {
    return __awaiter(this, void 0, void 0, function () {
        var rootFolder, f;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getRootFolder(ctx, list)];
                case 1:
                    rootFolder = _a.sent();
                    f = ctx.get_web().getFolderByServerRelativeUrl(rootFolder.get_serverRelativeUrl() + "/" + folderName);
                    ctx.load(f);
                    return [2 /*return*/, execute_1.executeOnContext(ctx, f)];
            }
        });
    });
}
exports.getFolderInListByName = getFolderInListByName;
/**
 * Creates a folder in a list
 * @param ctx the client context to execute on
 * @param list the list to create the folder in
 * @param folderName the folder name
 * @returns {Promise<SP.Folder>}
 */
function createFolderInList(ctx, list, folderName) {
    var leafNames = folderName.split('/');
    var leafName = leafNames[leafNames.length - 1];
    var item = new SP.ListItemCreationInformation();
    item.set_leafName(folderName);
    item.set_underlyingObjectType(SP.FileSystemObjectType.folder);
    var et = list.addItem(item);
    et.set_item('Title', leafName);
    et.update();
    ctx.load(et);
    return execute_1.executeOnContext(ctx);
}
exports.createFolderInList = createFolderInList;
/**
 * Creates a folder in list if it doesn't exist, this can create folder to any level i.e. folder1/folder2/folder3
 * @param ctx the client context to execute on
 * @param list the list to create the folders on
 * @param folderName the folder name
 */
function createFolderInListIfNotExist(ctx, list, folderName) {
    var folders = ("" + folderName).split('/');
    var ff = [];
    folders.forEach(function (e, i) {
        ff.push(folders.slice(0, i + 1).join('/'));
    });
    return ff.reduce(function (prev, curr) {
        return prev.then(function () {
            return folderExistsInList(ctx, list, curr).then(function (yes) {
                if (!yes) {
                    return createFolderInList(ctx, list, curr);
                }
            });
        });
    }, new Promise(function (r) {
        r();
    }))
        .then(function () {
        return getFolderInListByName(ctx, list, "" + folderName);
    });
}
exports.createFolderInListIfNotExist = createFolderInListIfNotExist;