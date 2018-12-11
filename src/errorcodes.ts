/**
 * JSOM error codes, the possible values of error codes that JSOM 
 * can return 
 */
export enum SPErrorCode {
    /**
     * The accessed item has been deleted
     */
    ItemDoesNotExist = -2147024809,
    /**
     * Some unknown error
     */
    GenericError = -1,
    /**
     * User does not have permission to access/perform operation 
     */
    AccessDenied = -2147024891,
    /**
     * A document with the same name already exist and the user is attempting to add new file/folder
     */
    DocAlreadyExists = -2130575257,
    /**
     * A version that is more recent has been saved 
     */
    VersionConflict = -2130575339,
    /**
     * List item is in recycle bin 
     */
    ListItemDeleted = -2130575338,
    /**
     * A field value that has been provided is invalid
     */
    InvalidFieldValue = -2130575155,
    /**
     * Operation not supported
     */
    NotSupported = -2147024846,
    /**
     * 
     */
    Redirect = -2130575152,
    /**
     * 
     */
    NotSupportedRequestVersion = -2130575151,
    /**
     * A specific field validation has failed 
     */
    FieldValueFailedValidation = -2130575163,
    /**
     * Item update failed validation rules
     */
    ItemValueFailedValidation = -2130575162,
}
