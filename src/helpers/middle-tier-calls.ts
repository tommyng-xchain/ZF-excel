// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/

import * as $ from "jquery";
import { WebApiConfig } from "./dataverse-webapi/lib/node";
import * as Wwebapi from "./dataverse-webapi/lib";
import { showMessage } from "./message-helper";
import * as Benz from "./benz/type";

/* global global, console, require */

export var token: string = null;

export async function callGetUserData(middletierToken: string, args?): Promise<any> {

  try {
    const response = await $.ajax({
      type: "GET",
      url: `https://${global.environment_name}.api.crm5.dynamics.com/api/data/v9.2/`,
      headers: { Authorization: "Bearer " + middletierToken },
      cache: false,
    });
    return response;
  } catch (err) {
    console.log(`Error from middle tier. \n${err.responseText || err.message}`);
    throw err;
  }
}

// export async function callGetData(middletierToken: string): Promise<any> {
//   try {
//     const url = "https://org8390b622.api.crm5.dynamics.com/api/data/v9.2/benz_testexcels";
//     console.log(`on callGetData.... `+ url);

//     const response = await $.ajax({
//       type: "GET",
//       url: url,
//       headers: { Authorization: "Bearer " + middletierToken },
//       cache: false,
//     });
//     return response;
//   } catch (err) {
//     console.log(`Error from middle tier. \n${err.responseText || err.message}`);
//     throw err;
//   }
// }
export async function callRetrieveData(token: string, args?): Promise<any> {
  console.log("callRetrieveData...");

  let response = null;
  let callapiconfig = global.Callapiaction.action;

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);

    return await Wwebapi.retrieve(
      config,
      callapiconfig.entitySet,
      callapiconfig.entityid,
      callapiconfig.queryString,
      callapiconfig.queryOptions
    );
  } catch (err) {
    console.log(`callRetrieveData Error ${JSON.stringify(response)}`);
    console.log(`callRetrieveData Error from middle tier. \n${err.stack}`);
    throw err;
  }
}

export async function callRetrieveMultipleData(token: string, args?): Promise<any> {
  console.log("callRetrieveMultipleData...");

  let response = null;
  let callapiconfig = args ?? global.Callapiaction.action;

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);

    return await Wwebapi.retrieveMultiple(
      config,
      callapiconfig.entitySet,
      callapiconfig.queryString,
      callapiconfig.queryOptions
    );
  } catch (err) {
    console.log(`callRetrieveMultipleData Error ${JSON.stringify(response)}`);
    console.log(`callRetrieveMultipleData Error from middle tier. \n${err.stack}`);
    throw err;
  }
}
export async function RetrieveMultipleDataAsync(token: string, args): Promise<any> {
  console.log("RetrieveMultipleData...");

  let response = null;

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);

    return await Wwebapi.retrieveMultiple(config, args.entitySet, args.queryString, args.queryOptions);
  } catch (err) {
    console.log(`RetrieveMultipleData Error ${JSON.stringify(response)}`);
    console.log(`RetrieveMultipleData Error from middle tier. \n${err.stack}`);
    throw err;
  }
}
export async function callCreate(token: string, args?): Promise<any> {
  console.log("callCreate...");
  if (!token) {
    console.log("callCreate...token error");
    return;
  }
  let response = null;
  let callapiconfig = global.Callapiaction.action;
  let environment_name = global.environment_name;

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);
    console.log("entity...");
    console.log(callapiconfig);
    console.log(typeof callapiconfig.queryString);
    console.log(callapiconfig.queryString);
    if (!callapiconfig.queryString) {
      throw new Error("queryString is empty");
    }
    return await Wwebapi.createWithReturnData(
      config,
      callapiconfig.entitySet,
      callapiconfig.queryString,
      callapiconfig.queryOptions
    );
  } catch (err) {
    console.log(`callCreate Error ${JSON.stringify(response)}`);
    console.log(err);
    console.log(err.stack);
    console.log(err.message);
    showMessage({ style: "error", message: "Error: " + err.message });
    // throw err;
  }
}

export async function callUpdate(token: string, args?): Promise<any> {
  console.log("callUpdate...");
  if (!token) {
    console.log("callUpdate...token error");
    return;
  }
  let response = null;
  let callapiconfig = global.Callapiaction.action;
  let environment_name = global.environment_name;
  // delete callapiconfig.queryString["benz_prototypesalesmeasureid"];

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);
    console.log("entity...");
    console.log(callapiconfig);
    console.log(typeof callapiconfig.queryString);
    console.log(callapiconfig.queryString);
    if (!callapiconfig.queryString) {
      throw new Error("queryString is empty");
    }
    console.log("callapiconfig.id");
    console.log(callapiconfig.id);
    return await Wwebapi.update(
      config,
      callapiconfig.entitySet,
      callapiconfig.id,
      callapiconfig.queryString,
      callapiconfig.queryOptions
    );
  } catch (err) {
    console.log(`callUpdate Error ${JSON.stringify(response)}`);
    console.error(err.stack);
    console.error(err.message);
    throw err;
  }
}

export async function callDisassociate(token: string, args?): Promise<any> {
  console.log("callDisassociate...");
  if (!token) {
    console.log("callDisassociate...token error");
    return;
  }
  let response = null;
  let callapiconfig = global.Callapiaction.action;
  let environment_name = global.environment_name;
  // delete callapiconfig.queryString["benz_prototypesalesmeasureid"];

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);
    console.log("entity...");
    console.log(callapiconfig);
    console.log(typeof callapiconfig.relatedEntityId);
    console.log(callapiconfig.relatedEntityId);
    if (!callapiconfig.relatedEntityId) {
      throw new Error("relatedEntityId is empty");
    }
    console.log("callapiconfig.id");
    console.log(callapiconfig.id);
    return await Wwebapi.disassociate(
      config,
      callapiconfig.entitySet,
      callapiconfig.id,
      callapiconfig.property,
      callapiconfig.relatedEntityId
    );
  } catch (err) {
    console.log(`callUpdate Error ${JSON.stringify(response)}`);
    console.error(err.stack);
    console.error(err.message);
    throw err;
  }
}

export async function MapAssociate(token: string, callapiconfig): Promise<any> {
  console.log("MapAssociate...");

  let response = null;

  try {
    const domain = `https:///${global.environment_name}.api.crm5.dynamics.com`;
    // const url = `${domain}/api/data/v9.2/benz_testexcels?$select=benz_name`;
    // console.log(`on callGetData.... `+url);
    const config = new WebApiConfig("9.2", token, domain);
    console.log("entity...");
    console.log(callapiconfig);
    callapiconfig = callapiconfig.Callapiaction.action;
    console.log("entity...");
    console.log(callapiconfig);
    return await Wwebapi.associate(
      config,
      callapiconfig.entitySet,
      callapiconfig.id.toString(),
      callapiconfig.relationship,
      callapiconfig.relatedEntitySet,
      callapiconfig.relatedEntityId,
      callapiconfig.queryOptions
    );
  } catch (err) {
    console.log(`MapAssociate Error ${JSON.stringify(response)}`);
    console.log(response);
    console.log(err.stack);
    console.log(err.message);
    throw err;
  }
}
