/**
 * Get the previous major version of an item/page/document
 * @param version the UI version 
 */
export function getPreviousMajorVersion(version:number, draftsNo = 512){
    return Math.floor(version/draftsNo)*draftsNo; 
}

/**
 * Get the next major version of an item/page/document
 * @param version the UI version 
 */
export function getNextMajorVersion(version:number, draftsNo = 512){
    return (Math.floor(version/draftsNo)+1)*draftsNo; 
}