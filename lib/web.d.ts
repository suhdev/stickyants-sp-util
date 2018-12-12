/**
 * Get web properties using JSOM
 * @param webUrl the web url (absolute/relative)
 */
export declare function getWebProperties<T>(webUrl?: string): Promise<any>;
/**
 * Set web properties
 * @param webUrl the web url (absolute/relative)
 * @param dict a dictionary of key/value pairs to use as new values for the web properties.
 */
export declare function setWebProperties(webUrl: string, dict: Dictionary<any>): Promise<{}>;
