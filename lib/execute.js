"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Calls `executeQueryAsync` and returns a promise so it can be chained into
 * a promise sequence or used in an async function i.e. awaited
 * @param ctx {SP.ClientContext} the client context to call `executeQueryAsync` on
 * @param value {any} the value to be returned by the generated promise (if it succeeds).
 */
function executeOnContext(ctx, value) {
    return new Promise(function (res, rej) {
        ctx.executeQueryAsync(function () {
            res(value);
        }, function (c, err) {
            rej(err);
        });
    });
}
exports.executeOnContext = executeOnContext;
