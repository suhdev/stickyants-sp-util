/**
 * Get user picture using their email address
 * @param email the email address of the user
 */
export function getUserPictureLink(email:string){
    return `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${email}&UA=0&size=HR64x64`;
}

/**
 * Get user delve profile link using their email address
 * @param email the email address of the user
 */
export function getUserDelveLink(email:string){
    
    return `${location.protocol}//${location.host.replace(/(.*)?\.sharepoint\.*com$/gi,'$1')}-my.sharepoint.com/_layouts/15/me.aspx/?p=${email}&v=work`;
}


/**
 * Creates a link to a specific version of the current page given a ui version
 * @param uiVersion the UIVersion to generate a link to 
 */
export function createVersionLink(uiVersion:number){
    return `${location.protocol}//${location.host}${location.pathname}?PageVersion=${uiVersion}`; 
}