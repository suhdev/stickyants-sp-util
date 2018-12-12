"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Get the previous major version of an item/page/document
 * @param version the UI version
 */
function getPreviousMajorVersion(version, draftsNo) {
    if (draftsNo === void 0) { draftsNo = 512; }
    return Math.floor(version / draftsNo) * draftsNo;
}
exports.getPreviousMajorVersion = getPreviousMajorVersion;
/**
 * Get the next major version of an item/page/document
 * @param version the UI version
 */
function getNextMajorVersion(version, draftsNo) {
    if (draftsNo === void 0) { draftsNo = 512; }
    return (Math.floor(version / draftsNo) + 1) * draftsNo;
}
exports.getNextMajorVersion = getNextMajorVersion;
