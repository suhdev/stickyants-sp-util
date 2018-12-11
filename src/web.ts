import { executeOnContext } from "./execute";

/**
 * Get web properties using JSOM
 * @param webUrl the web url (absolute/relative) 
 */
export async function getWebProperties<T>(webUrl:string = _spPageContextInfo.webAbsoluteUrl){
    let ctx = new SP.ClientContext(webUrl); 
    let web = ctx.get_web(); 
    let props = web.get_allProperties();
    ctx.load(props); 
    await executeOnContext(ctx);
    return props.get_fieldValues(); 
}

/**
 * Set web properties 
 * @param webUrl the web url (absolute/relative) 
 * @param dict a dictionary of key/value pairs to use as new values for the web properties. 
 */
export function setWebProperties(webUrl:string = _spPageContextInfo.webAbsoluteUrl, dict:Dictionary<any>){
    let ctx = new SP.ClientContext(webUrl); 
    let web = ctx.get_web(); 
    let props = web.get_allProperties();
    for(var key in dict){
        if (Object.hasOwnProperty.call(dict,key)){
            props.set_item(key,dict[key]); 
        }
    }
    web.update();
    return executeOnContext(ctx); 
}