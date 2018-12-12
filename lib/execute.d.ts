/// <reference types="sharepoint" />
/**
 * Calls `executeQueryAsync` and returns a promise so it can be chained into
 * a promise sequence or used in an async function i.e. awaited
 * @param ctx {SP.ClientContext} the client context to call `executeQueryAsync` on
 * @param value {any} the value to be returned by the generated promise (if it succeeds).
 */
export declare function executeOnContext<T>(ctx: SP.ClientContext, value?: T): Promise<T>;
