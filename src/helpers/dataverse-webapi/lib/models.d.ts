export declare class WebApiConfig {
    version: string;
    accessToken?: string;
    url?: string;
    /**
     * Constructor
     * @param config WebApiConfig
     */
    constructor(version: string, accessToken?: string, url?: string);
}
export interface WebApiRequestResult {
    error: boolean;
    response: any;
    headers?: unknown;
}
export interface WebApiRequestConfig {
    method: string;
    contentType: string;
    body?: string;
    queryString: string;
    apiConfig: WebApiConfig;
    queryOptions?: QueryOptions;
}
export interface QueryOptions {
    maxPageSize?: number;
    impersonateUserId?: string;
    representation?: boolean;
    customHeaders?: Record<string, string>;
}
export interface Entity {
    [propName: string]: string | number | boolean | undefined | null | Entity | Entity[];
}
export interface RetrieveMultipleResponse {
    value: Entity[];
    '@odata.nextlink': string;
}
export interface ChangeSet {
    queryString: string;
    entity: Entity;
    method: string;
}
export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}
