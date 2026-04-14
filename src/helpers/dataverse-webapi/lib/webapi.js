"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.batchOperation = exports.unboundFunction = exports.boundFunction = exports.unboundAction = exports.boundAction = exports.disassociate = exports.associate = exports.deleteProperty = exports.deleteRecord = exports.updateProperty = exports.updateWithReturnData = exports.update = exports.createWithReturnData = exports.create = exports.retrieveMultipleNextPage = exports.retrieveMultiple = exports.retrieveNavigationProperties = exports.retrieveProperty = exports.retrieve = exports.getHeaders = void 0;
function parseGuid(id) {
    if (id === null || id === 'undefined' || id === '') {
        return '';
    }
    id = id.replace(/[{}]/g, '');
    if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(id)) {
        return id.toUpperCase();
    }
    else {
        throw Error(`Id ${id} is not a valid GUID`);
    }
}
function getHeaders(config) {
    let headers = {};
    headers.Accept = 'application/json';
    headers['OData-MaxVersion'] = '4.0';
    headers['OData-Version'] = '4.0';
    //Prefer
    headers['Prefer'] = 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"';
    headers['Content-Type'] = config.contentType;
    headers['If-None-Match'] = 'null';
    if (config.apiConfig.accessToken != null) {
        headers.Authorization = `Bearer ${config.apiConfig.accessToken}`;
    }
    headers.Prefer = getPreferHeader(config.queryOptions);
    if (config.queryOptions != null && typeof config.queryOptions !== 'undefined') {
        if (config.queryOptions.impersonateUserId != null) {
            headers.CallerObjectId = config.queryOptions.impersonateUserId;
        }
        if (config.queryOptions.customHeaders != null) {
            headers = { ...headers, ...config.queryOptions.customHeaders };
        }
    }
    return headers;
}
exports.getHeaders = getHeaders;
function getPreferHeader(queryOptions) {
    const prefer = ['odata.include-annotations="*"'];
    // add max page size to prefer request header
    if (queryOptions?.maxPageSize) {
        prefer.push(`odata.maxpagesize=${queryOptions.maxPageSize}`);
    }
    // add formatted values to prefer request header
    if (queryOptions?.representation) {
        prefer.push('return=representation');
    }
    // add choice FormattedValue values to  prefer request header
        // prefer.push('odata.include-annotations=OData.Community.Display.V1.FormattedValue');
    return prefer.join(',');
}
function getFunctionInputs(queryString, inputs) {
    if (inputs == null) {
        return queryString + ')';
    }
    const aliases = [];
    for (const input of inputs) {
        queryString += input.name;
        if (input.alias) {
            queryString += `=@${input.alias},`;
            aliases.push(`@${input.alias}=${input.value}`);
        }
        else {
            queryString += `=${input.value},`;
        }
    }
    queryString = queryString.substr(0, queryString.length - 1) + ')';
    if (aliases.length > 0) {
        queryString += `?${aliases.join('&')}`;
    }
    return queryString;
}
function handleError(result) {
    try {
        return JSON.parse(result).error;
    }
    catch (e) {
        console.log('handleError; Unexpected Error; ' + JSON.stringify(result) + ";" + e.message);
        return new Error('Unexpected Error; ');
    }
}
/**
 * Retrieve a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
function retrieve(apiConfig, entitySet, id, submitRequest, queryString, queryOptions) {
    if (queryString != null && !/^[?]/.test(queryString)) {
        queryString = `?${queryString}`;
    }
    id = parseGuid(id);
    const query = queryString != null ? `${entitySet}(${id})${queryString}` : `${entitySet}(${id})`;
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: query,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.retrieve = retrieve;
/**
 * Retrieve a single property of a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param property Property to retrieve
 */
function retrieveProperty(apiConfig, entitySet, id, submitRequest, property) {
    id = parseGuid(id);
    const query = `${entitySet}(${id})/${property}`;
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: query,
        apiConfig: apiConfig,
        queryOptions: {}
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.retrieveProperty = retrieveProperty;
/**
 * Retrieve columns for a related navigation property of a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param id Id of record to retrieve
 * @param property Navigation property to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
function retrieveNavigationProperties(apiConfig, entitySet, id, submitRequest, property, queryString, queryOptions) {
    id = parseGuid(id);
    if (queryString != null && !/^[?]/.test(queryString)) {
        queryString = `?${queryString}`;
    }
    const query = queryString != null ? `${entitySet}(${id})/${property}${queryString}` : `${entitySet}(${id})/${property}`;
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: query,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.retrieveNavigationProperties = retrieveNavigationProperties;
/**
 * Retrieve multiple records from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to retrieve
 * @param queryString OData query string parameters
 * @param queryOptions Various query options for the query
 */
function retrieveMultiple(apiConfig, entitySet, submitRequest, queryString, queryOptions) {
    if (queryString != null && !/^[?]/.test(queryString)) {
        queryString = `?${queryString}`;
    }
    const query = queryString != null ? entitySet + queryString : entitySet;
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: query,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.retrieveMultiple = retrieveMultiple;
/**
 * Retrieve next page from a retrieveMultiple request
 * @param apiConfig WebApiConfig object
 * @param url Query from the @odata.nextlink property of a retrieveMultiple
 * @param queryOptions Various query options for the query
 */
function retrieveMultipleNextPage(apiConfig, url, submitRequest, queryOptions) {
    apiConfig.url = url;
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: '',
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.retrieveMultipleNextPage = retrieveMultipleNextPage;
/**
 * Create a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param entity Entity to create
 * @param queryOptions Various query options for the query
 */
function create(apiConfig, entitySet, entity, submitRequest, queryOptions) {
    const config = {
        method: 'POST',
        contentType: 'application/json; charset=utf-8',
        queryString: entitySet,
        body: JSON.stringify(entity),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.create = create;
/**
 * Create a record in Dataverse and return data
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param entity Entity to create
 * @param select Select odata query parameter
 * @param queryOptions Various query options for the query
 */
function createWithReturnData(apiConfig, entitySet, entity, select, submitRequest, queryOptions) {
    if (select != null && !/^[?]/.test(select)) {
        select = `?${select}`;
    }
    // set representation
    if (queryOptions == null) {
        queryOptions = {};
    }
    queryOptions.representation = true;
    const config = {
        method: 'POST',
        contentType: 'application/json; charset=utf-8',
        queryString: entitySet + select,
        body: JSON.stringify(entity),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.createWithReturnData = createWithReturnData;
/**
 * Update a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param entity Entity fields to update
 * @param queryOptions Various query options for the query
 */
function update(apiConfig, entitySet, id, entity, submitRequest, queryOptions) {
    id = parseGuid(id);
    const config = {
        method: 'PATCH',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})`,
        body: JSON.stringify(entity),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.update = update;
/**
 * Create a record in Dataverse and return data
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to create
 * @param id Id of record to update
 * @param entity Entity fields to update
 * @param select Select odata query parameter
 * @param queryOptions Various query options for the query
 */
function updateWithReturnData(apiConfig, entitySet, id, entity, select, submitRequest, queryOptions) {
    id = parseGuid(id);
    if (select != null && !/^[?]/.test(select)) {
        select = `?${select}`;
    }
    // set representation
    if (queryOptions == null) {
        queryOptions = {};
    }
    queryOptions.representation = true;
    const config = {
        method: 'PATCH',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})${select}`,
        body: JSON.stringify(entity),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve(JSON.parse(result.response));
            }
        });
    });
}
exports.updateWithReturnData = updateWithReturnData;
/**
 * Update a single property of a record in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param attribute Attribute to update
 * @param queryOptions Various query options for the query
 */
function updateProperty(apiConfig, entitySet, id, attribute, value, submitRequest, queryOptions) {
    id = parseGuid(id);
    const config = {
        method: 'PUT',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})/${attribute}`,
        body: JSON.stringify({ value: value }),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.updateProperty = updateProperty;
/**
 * Delete a record from Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to delete
 * @param id Id of record to delete
 */
function deleteRecord(apiConfig, entitySet, id, submitRequest) {
    id = parseGuid(id);
    const config = {
        method: 'DELETE',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})`,
        apiConfig: apiConfig
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.deleteRecord = deleteRecord;
/**
 * Delete a property from a record in Dataverse. Non navigation properties only
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to update
 * @param id Id of record to update
 * @param attribute Attribute to delete
 */
function deleteProperty(apiConfig, entitySet, id, attribute, submitRequest) {
    id = parseGuid(id);
    const queryString = `/${attribute}`;
    const config = {
        method: 'DELETE',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})${queryString}`,
        apiConfig: apiConfig
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.deleteProperty = deleteProperty;
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
function associate(apiConfig, entitySet, id, relationship, relatedEntitySet, relatedEntityId, submitRequest, queryOptions) {
    id = parseGuid(id);
    const related = {
        '@odata.id': `${apiConfig.url}/${relatedEntitySet}(${relatedEntityId})`
    };
    const config = {
        method: 'POST',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})/${relationship}/$ref`,
        body: JSON.stringify(related),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.associate = associate;
/**
 * Disassociate two records
 * @param apiConfig WebApiConfig obje
 * @param entitySet Type of entity for primary record
 * @param id  Id of primary record
 * @param property Schema name of property or relationship
 * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
 */
function disassociate(apiConfig, entitySet, id, property, submitRequest, relatedEntityId) {
    id = parseGuid(id);
    let queryString = property;
    if (relatedEntityId != null) {
        queryString += `(${relatedEntityId})`;
    }
    queryString += '/$ref';
    const config = {
        method: 'DELETE',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})/${queryString}`,
        apiConfig: apiConfig
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                resolve();
            }
        });
    });
}
exports.disassociate = disassociate;
/**
 * Execute a default or custom bound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to run the action against
 * @param id Id of record to run the action against
 * @param actionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
function boundAction(apiConfig, entitySet, id, actionName, submitRequest, inputs, queryOptions) {
    id = parseGuid(id);
    const config = {
        method: 'POST',
        contentType: 'application/json; charset=utf-8',
        queryString: `${entitySet}(${id})/Microsoft.Dynamics.CRM.${actionName}`,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    if (inputs != null) {
        config.body = JSON.stringify(inputs);
    }
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                if (result.response) {
                    resolve(JSON.parse(result.response));
                }
                else {
                    resolve(null);
                }
            }
        });
    });
}
exports.boundAction = boundAction;
/**
 * Execute a default or custom unbound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param actionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
function unboundAction(apiConfig, actionName, submitRequest, inputs, queryOptions) {
    const config = {
        method: 'POST',
        contentType: 'application/json; charset=utf-8',
        queryString: actionName,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    if (inputs != null) {
        config.body = JSON.stringify(inputs);
    }
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                if (result.response) {
                    resolve(JSON.parse(result.response));
                }
                else {
                    resolve(null);
                }
            }
        });
    });
}
exports.unboundAction = unboundAction;
/**
 * Execute a default or custom bound action in Dataverse
 * @param apiConfig WebApiConfig object
 * @param entitySet Type of entity to run the action against
 * @param id Id of record to run the action against
 * @param functionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
function boundFunction(apiConfig, entitySet, id, functionName, submitRequest, inputs, queryOptions) {
    id = parseGuid(id);
    let queryString = `${entitySet}(${id})/Microsoft.Dynamics.CRM.${functionName}(`;
    queryString = getFunctionInputs(queryString, inputs);
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: queryString,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                if (result.response) {
                    resolve(JSON.parse(result.response));
                }
                else {
                    resolve(null);
                }
            }
        });
    });
}
exports.boundFunction = boundFunction;
/**
 * Execute an unbound function in Dataverse
 * @param apiConfig WebApiConfig object
 * @param functionName Name of the action to run
 * @param inputs Any inputs required by the action
 * @param queryOptions Various query options for the query
 */
function unboundFunction(apiConfig, functionName, submitRequest, inputs, queryOptions) {
    let queryString = `${functionName}(`;
    queryString = getFunctionInputs(queryString, inputs);
    const config = {
        method: 'GET',
        contentType: 'application/json; charset=utf-8',
        queryString: queryString,
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                if (result.response) {
                    resolve(JSON.parse(result.response));
                }
                else {
                    resolve(null);
                }
            }
        });
    });
}
exports.unboundFunction = unboundFunction;
/**
 * Execute a batch operation in Dataverse
 * @param apiConfig WebApiConfig object
 * @param batchId Unique batch id for the operation
 * @param changeSetId Unique change set id for any changesets in the operation
 * @param changeSets Array of change sets (create or update) for the operation
 * @param batchGets Array of get requests for the operation
 * @param queryOptions Various query options for the query
 */
function batchOperation(apiConfig, batchId, changeSetId, changeSets, batchGets, submitRequest, queryOptions) {
    // build post body
    const body = [];
    if (changeSets.length > 0) {
        body.push(`--batch_${batchId}`);
        body.push(`Content-Type: multipart/mixed;boundary=changeset_${changeSetId}`);
        body.push('');
    }
    // push change sets to body
    for (let i = 0; i < changeSets.length; i++) {
        body.push(`--changeset_${changeSetId}`);
        body.push('Content-Type: application/http');
        body.push('Content-Transfer-Encoding:binary');
        body.push(`Content-ID: ${i + 1}`);
        body.push('');
        body.push(`${changeSets[i].method} ${apiConfig.url}/${changeSets[i].queryString} HTTP/1.1`);
        body.push('Content-Type: application/json;type=entry');
        body.push('');
        body.push(JSON.stringify(changeSets[i].entity));
    }
    if (changeSets.length > 0) {
        body.push(`--changeset_${changeSetId}--`);
        body.push('');
    }
    // push get requests to body
    for (const get of batchGets) {
        body.push(`--batch_${batchId}`);
        body.push('Content-Type: application/http');
        body.push('Content-Transfer-Encoding:binary');
        body.push('');
        body.push(`GET ${apiConfig.url}/${get} HTTP/1.1`);
        body.push('Accept: application/json');
        body.push('');
    }
    if (batchGets.length > 0) {
        body.push('');
    }
    body.push(`--batch_${batchId}--`);
    const config = {
        method: 'POST',
        contentType: `multipart/mixed;boundary=batch_${batchId}`,
        queryString: '$batch',
        body: body.join('\r\n'),
        apiConfig: apiConfig,
        queryOptions: queryOptions
    };
    return new Promise((resolve, reject) => {
        submitRequest(config, (result) => {
            if (result.error) {
                reject(handleError(result.response));
            }
            else {
                if (result.response) {
                    resolve(result.response);
                }
                else {
                    resolve(null);
                }
            }
        });
    });
}
exports.batchOperation = batchOperation;
