/// <reference types="sharepoint" />
/**
 * Extracts the image url from `PublishingPageImage` field value
 * @param str the `PublishingPageImage` field value
 */
export declare function extractUrlFromPublishingPageImage(str: string): any;
/**
 * Get page layout by name/title
 * @param ctx the SP.ClientContext to execute this request on
 * @param name the name/Title of the page layout
 */
export declare function getPageLayoutItemByName(ctx: SP.ClientContext, name: string): Promise<SP.ListItem>;
/**
 * Get page layout item by the associated content type name
 * @param ctx the context to execute on
 * @param name the associated content type name
 */
export declare function getPageLayoutItemByAssociatedContentType(ctx: SP.ClientContext, name: string): Promise<SP.ListItem>;
/**
 * Creates a publishing page given the title, filename and page layout item
 * @param ctx {SP.ClientContext} the client context to execute on
 * @param title the title of the page
 * @param fileName the file name of the page
 * @param layout {SP.ListItem} the page layout item
 *
 * @see {getPageLayoutItemByAssociatedContentType}
 */
export declare function createPublishingPageWithLayout(ctx: SP.ClientContext, fileName: string, layout: SP.ListItem): Promise<SP.Publishing.PublishingPage>;
/**
 * Creates a publishing page given the context, page file name, and the content type name
 * @param ctx the client context to execute on
 * @param fileName the page file name
 * @param contentType the content type name
 */
export declare function createPublishingPage(ctx: SP.ClientContext, fileName: string, contentType: string): Promise<SP.Publishing.PublishingPage>;
