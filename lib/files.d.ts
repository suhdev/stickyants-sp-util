/// <reference types="sharepoint" />
/**
 * Get root folder of a list
 * @param ctx the client context to execute on
 * @param list the list to get its root folder
 * @returns {Promise<SP.Folder}
 */
export declare function getRootFolder(ctx: SP.ClientContext, list: SP.List): Promise<SP.Folder>;
/**
 * Get list parent web server relative url
 * @param ctx the client context to execute on
 * @param list the list to get its parent web url
 * @returns {Promise<string>}
 */
export declare function getListParentWebUrl(ctx: SP.ClientContext, list: SP.List): Promise<string>;
/**
 * Checks if a folder exists in a list or not
 * @param ctx the client context to execute on
 * @param list the list to check the folder path on
 * @param folderName the folder name
 * @returns {Promise<boolean>}
 */
export declare function folderExistsInList(ctx: SP.ClientContext, list: SP.List, folderName: string): Promise<boolean>;
/**
 * Get folder in a list by folder name
 * @param ctx the client context to execute on
 * @param list the list to get the folder in
 * @param folderName the folder name
 * @returns {Promise<SP.Folder>}
 */
export declare function getFolderInListByName(ctx: SP.ClientContext, list: SP.List, folderName: string): Promise<any>;
/**
 * Creates a folder in a list
 * @param ctx the client context to execute on
 * @param list the list to create the folder in
 * @param folderName the folder name
 * @returns {Promise<SP.Folder>}
 */
export declare function createFolderInList(ctx: SP.ClientContext, list: SP.List, folderName: string): Promise<SP.Folder>;
/**
 * Creates a folder in list if it doesn't exist, this can create folder to any level i.e. folder1/folder2/folder3
 * @param ctx the client context to execute on
 * @param list the list to create the folders on
 * @param folderName the folder name
 */
export declare function createFolderInListIfNotExist(ctx: SP.ClientContext, list: SP.List, folderName: string): Promise<SP.Folder>;
