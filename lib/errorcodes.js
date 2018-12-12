"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * JSOM error codes, the possible values of error codes that JSOM
 * can return
 */
var SPErrorCode;
(function (SPErrorCode) {
    /**
     * The accessed item has been deleted
     */
    SPErrorCode[SPErrorCode["ItemDoesNotExist"] = -2147024809] = "ItemDoesNotExist";
    /**
     * Some unknown error
     */
    SPErrorCode[SPErrorCode["GenericError"] = -1] = "GenericError";
    /**
     * User does not have permission to access/perform operation
     */
    SPErrorCode[SPErrorCode["AccessDenied"] = -2147024891] = "AccessDenied";
    /**
     * A document with the same name already exist and the user is attempting to add new file/folder
     */
    SPErrorCode[SPErrorCode["DocAlreadyExists"] = -2130575257] = "DocAlreadyExists";
    /**
     * A version that is more recent has been saved
     */
    SPErrorCode[SPErrorCode["VersionConflict"] = -2130575339] = "VersionConflict";
    /**
     * List item is in recycle bin
     */
    SPErrorCode[SPErrorCode["ListItemDeleted"] = -2130575338] = "ListItemDeleted";
    /**
     * A field value that has been provided is invalid
     */
    SPErrorCode[SPErrorCode["InvalidFieldValue"] = -2130575155] = "InvalidFieldValue";
    /**
     * Operation not supported
     */
    SPErrorCode[SPErrorCode["NotSupported"] = -2147024846] = "NotSupported";
    /**
     *
     */
    SPErrorCode[SPErrorCode["Redirect"] = -2130575152] = "Redirect";
    /**
     *
     */
    SPErrorCode[SPErrorCode["NotSupportedRequestVersion"] = -2130575151] = "NotSupportedRequestVersion";
    /**
     * A specific field validation has failed
     */
    SPErrorCode[SPErrorCode["FieldValueFailedValidation"] = -2130575163] = "FieldValueFailedValidation";
    /**
     * Item update failed validation rules
     */
    SPErrorCode[SPErrorCode["ItemValueFailedValidation"] = -2130575162] = "ItemValueFailedValidation";
})(SPErrorCode = exports.SPErrorCode || (exports.SPErrorCode = {}));
