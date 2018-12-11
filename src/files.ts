import { executeOnContext } from "./execute";

/**
 * Get root folder of a list 
 * @param ctx the client context to execute on 
 * @param list the list to get its root folder 
 * @returns {Promise<SP.Folder}
 */
export function getRootFolder(ctx:SP.ClientContext,list:SP.List):Promise<SP.Folder>{
    var folder = list.get_rootFolder();
    ctx.load(folder); 
    return executeOnContext(ctx,folder); 
}

/**
 * Get list parent web server relative url  
 * @param ctx the client context to execute on
 * @param list the list to get its parent web url 
 * @returns {Promise<string>}
 */
export async function getListParentWebUrl(ctx:SP.ClientContext, list:SP.List){
    var web = list.get_parentWeb();
    ctx.load(web);
    await executeOnContext(ctx); 
    return web.get_serverRelativeUrl();
}

/**
 * Checks if a folder exists in a list or not 
 * @param ctx the client context to execute on 
 * @param list the list to check the folder path on 
 * @param folderName the folder name 
 * @returns {Promise<boolean>}
 */
export async function folderExistsInList(ctx:SP.ClientContext,list:SP.List,folderName:string):Promise<boolean>{
    const rootFolder = await getRootFolder(ctx,list); 
    var url = rootFolder.get_serverRelativeUrl(); 
    let rf = ctx.get_web().getFolderByServerRelativeUrl(`${url}/${folderName}`); 
    ctx.load(rf); 
    try {
        await executeOnContext(ctx); 
        return true;  
    }catch(err){
        return false; 
    }
}

/**
 * Get folder in a list by folder name 
 * @param ctx the client context to execute on 
 * @param list the list to get the folder in 
 * @param folderName the folder name 
 * @returns {Promise<SP.Folder>}
 */
export async function getFolderInListByName(ctx:SP.ClientContext,list:SP.List,folderName:string):Promise<any>{
    const rootFolder = await getRootFolder(ctx,list); 
    let f = ctx.get_web().getFolderByServerRelativeUrl(`${rootFolder.get_serverRelativeUrl()}/${folderName}`); //get_folders().getByUrl(folderName); 
    ctx.load(f); 
    return executeOnContext(ctx,f);
}

/**
 * Creates a folder in a list 
 * @param ctx the client context to execute on 
 * @param list the list to create the folder in 
 * @param folderName the folder name 
 * @returns {Promise<SP.Folder>}
 */
export function createFolderInList(ctx:SP.ClientContext,list:SP.List,folderName:string):Promise<SP.Folder>{
    var leafNames = folderName.split('/'); 
    var leafName = leafNames[leafNames.length-1]; 
    var item = new SP.ListItemCreationInformation(); 
    item.set_leafName(folderName); 
    item.set_underlyingObjectType(SP.FileSystemObjectType.folder); 
    var et = list.addItem(item);
    et.set_item('Title',leafName); 
    et.update(); 
    ctx.load(et); 
    return executeOnContext(ctx); 
}

/**
 * Creates a folder in list if it doesn't exist, this can create folder to any level i.e. folder1/folder2/folder3
 * @param ctx the client context to execute on 
 * @param list the list to create the folders on 
 * @param folderName the folder name 
 */
export function createFolderInListIfNotExist(ctx:SP.ClientContext,list:SP.List,folderName:string):Promise<SP.Folder>{
    let folders = `${folderName}`.split('/');
    var ff:string[] = []; 
    folders.forEach((e,i)=>{
        ff.push(folders.slice(0,i+1).join('/'));
    });
    return ff.reduce((prev,curr)=>{
        return prev.then(()=>{
            return folderExistsInList(ctx,list,curr).then((yes)=>{
                if (!yes){
                    return createFolderInList(ctx,list,curr);
                }
            });
        });
    },new Promise<any>((r)=>{
        r();
    }))
    .then(()=>{
        return getFolderInListByName(ctx,list,`${folderName}`); 
    });
}