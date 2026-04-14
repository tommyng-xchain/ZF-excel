import { ChangeSet, Entity, FunctionInput, QueryOptions, RetrieveMultipleResponse, WebApiConfig, WebApiRequestConfig, WebApiRequestResult } from './models';
type RequestCallback = (config: WebApiRequestConfig, callback: (result: WebApiRequestResult) => void) => void;
export declare function getHeaders(config: WebApiRequestConfig): Record<string, string>;
/**
 * Retrieve a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
export declare function retrieve(apiConfig: WebApiConfig, entitySet: string, id: string, submitRequest: RequestCallback, queryString?: string, queryOptions?: QueryOptions): Promise<Entity>;
/**
 * Retrieve a single property of a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param property Property to retrieve
 */
export declare function retrieveProperty(apiConfig: WebApiConfig, entitySet: string, id: string, submitRequest: RequestCallback, property: string): Promise<Entity>;
/**
 * Retrieve columns for a related navigation property of a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param property Navigation property to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
export declare function retrieveNavigationProperties(apiConfig: WebApiConfig, entitySet: string, id: string, submitRequest: RequestCallback, property: string, queryString?: string, queryOptions?: QueryOptions): Promise<Entity>;
/**
 * Retrieve multiple records from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
export declare function retrieveMultiple(apiConfig: WebApiConfig, entitySet: string, submitRequest: RequestCallback, queryString?: string, queryOptions?: QueryOptions): Promise<RetrieveMultipleResponse>;
/**
 * Retrieve next page from a retrieveMultiple request
 * @param apiConfig WebApiConfig object
 * @param url Query from the @odata.nextlink property of a retrieveMultiple
 * @param queryOptions Various query options for the query
 */
export declare function retrieveMultipleNextPage(apiConfig: WebApiConfig, url: string, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<RetrieveMultipleResponse>;
/**
 * Create a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param entity Entity to create
 * @param queryOptions Various query options for the query
 */
export declare function create(apiConfig: WebApiConfig, entitySet: string, entity: Entity, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<void>;
/**
 * Create a record in Dataverse and return data
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param entity Entity to create
 * @param select Select odata query parameter
 * @param queryOptions Various query options for the query
 */
export declare function createWithReturnData(apiConfig: WebApiConfig, entitySet: string, entity: Entity, select: string, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<Entity>;
/**
 * Update a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param entity Entity fields to update
 * @param queryOptions Various query options for the query
 */
export declare function update(apiConfig: WebApiConfig, entitySet: string, id: string, entity: Entity, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<void>;
/**
 * Create a record in Dataverse and return data
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param id Id of record to update
 * @param entity Entity fields to update
 * @param select Select odata query parameter
 * @param queryOptions Various query options for the query
 */
export declare function updateWithReturnData(apiConfig: WebApiConfig, entitySet: string, id: string, entity: Entity, select: string, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<Entity>;
/**
 * Update a single property of a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param attribute Attribute to update
 * @param queryOptions Various query options for the query
 */
export declare function updateProperty(apiConfig: WebApiConfig, entitySet: string, id: string, attribute: string, value: string | number | boolean, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<void>;
/**
 * Delete a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to delete
 * @param id Id of record to delete
 */
export declare function deleteRecord(apiConfig: WebApiConfig, entitySet: string, id: string, submitRequest: RequestCallback): Promise<void>;
/**
 * Delete a property from a record in Dataverse. Non navigation properties only
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param attribute Attribute to delete
 */
export declare function deleteProperty(apiConfig: WebApiConfig, entitySet: string, id: string, attribute: string, submitRequest: RequestCallback): Promise<void>;
/**
 * Associate two records
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity for primary record
 * @param id Id of primary record
 * @param relationship Schema name of relationship
 * @param relatedEntitySet Type of entity for secondary record
 * @param relatedEntityId Id of secondary record
 * @param queryOptions Various query options for the query
 */
export declare function associate(apiConfig: WebApiConfig, entitySet: string, id: string, relationship: string, relatedEntitySet: string, relatedEntityId: string, submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<void>;
/**
 * Disassociate two records
 * @param apiConfig WebApiConfig obje
 * @param entitySet Type of entity for primary record
 * @param id  Id of primary record
 * @param property Schema name of property or relationship
 * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
 */
export declare function disassociate(apiConfig: WebApiConfig, entitySet: string, id: string, property: string, submitRequest: RequestCallback, relatedEntityId?: string): Promise<void>;
/**
 * Execute a default or custom bound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to run the action against
 * @param id Id of record to run the action against
 * @param actionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
export declare function boundAction(apiConfig: WebApiConfig, entitySet: string, id: string, actionName: string, submitRequest: RequestCallback, inputs?: Record<string, unknown>, queryOptions?: QueryOptions): Promise<unknown>;
/**
 * Execute a default or custom unbound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param actionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
export declare function unboundAction(apiConfig: WebApiConfig, actionName: string, submitRequest: RequestCallback, inputs?: Record<string, unknown>, queryOptions?: QueryOptions): Promise<unknown>;
/**
 * Execute a default or custom bound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to run the action against
 * @param id Id of record to run the action against
 * @param functionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
export declare function boundFunction(apiConfig: WebApiConfig, entitySet: string, id: string, functionName: string, submitRequest: RequestCallback, inputs?: FunctionInput[], queryOptions?: QueryOptions): Promise<unknown>;
/**
 * Execute an unbound function in Dataverse
 * @param apiConfig WebApiConfig object
 * @param functionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
export declare function unboundFunction(apiConfig: WebApiConfig, functionName: string, submitRequest: RequestCallback, inputs?: FunctionInput[], queryOptions?: QueryOptions): Promise<unknown>;
/**
 * Execute a batch operation in Dataverse
 * @param apiConfig WebApiConfig object
 * @param batchId Unique batch id for the operation
 * @param changeSetId Unique change set id for any changesets in the operation
 * @param changeSets Array of change sets (create or update) for the operation
 * @param batchGets Array of get requests for the operation
 * @param queryOptions Various query options for the query
 */
export declare function batchOperation(apiConfig: WebApiConfig, batchId: string, changeSetId: string, changeSets: ChangeSet[], batchGets: string[], submitRequest: RequestCallback, queryOptions?: QueryOptions): Promise<unknown>;
export {};
