import { executeOnContext } from "./execute";

/**
 * Extracts the image url from `PublishingPageImage` field value 
 * @param str the `PublishingPageImage` field value
 */
export function extractUrlFromPublishingPageImage(str:string){
    var k = null;
    str && str.replace(/src="([\s\S]+?)"/,(e,v)=>{
        return k = v; 
    });
    return k; 
}

/**
 * Get page layout by name/title 
 * @param ctx the SP.ClientContext to execute this request on 
 * @param name the name/Title of the page layout
 */
export async function getPageLayoutItemByName(ctx:SP.ClientContext, name:string):Promise<SP.ListItem>{
    var rootWeb = ctx.get_site().get_rootWeb(); 
    var list = rootWeb.get_lists().getByTitle('Master Page Gallery'); 
    var query = new SP.CamlQuery(); 
    query.set_viewXml(`<View><Query><Where><Eq><FieldRef Name="Title" /><Value Type="Text">${name}</Value></Eq></Where></Query></View>`);
    var layouts = list.getItems(query); 
    ctx.load(layouts); 
    await executeOnContext(ctx);
    const data = layouts.get_data(); 
    if(data.length){
        return data[0]; 
    }
    return null; 
}

/**
 * Get page layout item by the associated content type name 
 * @param ctx the context to execute on 
 * @param name the associated content type name 
 */
export async function getPageLayoutItemByAssociatedContentType(ctx:SP.ClientContext,name:string):Promise<SP.ListItem>{
    var rootWeb = ctx.get_site().get_rootWeb(); 
    var list = rootWeb.get_lists().getByTitle('Master Page Gallery'); 
    var query = new SP.CamlQuery(); 
    query.set_viewXml(`<View><Query><Where><Contains><FieldRef Name="PublishingAssociatedContentType" /><Value Type="Text">${name}</Value></Contains></Where></Query></View>`);
    var layouts = list.getItems(query); 
    ctx.load(layouts);
    await executeOnContext(ctx);  
    if(layouts.get_data().length){
        return layouts.get_data()[0] 
    }else {
        return null; 
    }
}

/**
 * Creates a publishing page given the title, filename and page layout item 
 * @param ctx {SP.ClientContext} the client context to execute on 
 * @param title the title of the page 
 * @param fileName the file name of the page
 * @param layout {SP.ListItem} the page layout item 
 * 
 * @see {getPageLayoutItemByAssociatedContentType}
 */
export function createPublishingPageWithLayout(ctx:SP.ClientContext, fileName:string,  layout:SP.ListItem){
    if (typeof fileName !== "string" || !fileName || !fileName.trim()){
        throw new Error(`Invalid file name provided, expected truthy string but got ${typeof fileName} - '${fileName}'`); 
    }
    var web = ctx.get_web();
    var newPublishingPage = SP.Publishing.PublishingWeb.getPublishingWeb(ctx, web); 
    var pageInfo = new SP.Publishing.PublishingPageInformation(); 
    fileName = fileName.replace(/\.aspx$/i,'') +'.aspx'; 
    pageInfo.set_name(fileName);  
    pageInfo.set_pageLayoutListItem(layout);  
    var newPage = newPublishingPage.addPublishingPage(pageInfo);
    ctx.load(newPage); 
    return executeOnContext(ctx,newPage); 
}

/**
 * Creates a publishing page given the context, page file name, and the content type name
 * @param ctx the client context to execute on
 * @param fileName the page file name 
 * @param contentType the content type name 
 */
export async function createPublishingPage(ctx:SP.ClientContext,fileName:string,contentType:string){
    var layout = await getPageLayoutItemByAssociatedContentType(ctx,contentType); 
    return createPublishingPageWithLayout(ctx,fileName,layout); 
}