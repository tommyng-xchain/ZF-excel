/* global console, global */

import * as authdialog from "./fallbackauthdialog";
//callGetUserData
import {
  MapAssociate,
  RetrieveMultipleDataAsync,
  callCreate,
  callRetrieveData,
  callRetrieveMultipleData,
  callUpdate,
  callDisassociate,
} from "./middle-tier-calls";
import { handleClientSideErrors } from "./error-handler";
import { getGlobalVariable } from "../commands/commands";
import * as Account from "./benz/account";

var action = "";

export async function setAccount(t) {
  Account.setAccount(t);
}

export async function getAccessToken(): Promise<string> {
  return global.ApiAccessToken;
}

export async function get_Data__benz_carModel(callback): Promise<void> {
  let response: any;
  // $select=benz_prototypemodeldesignationid,benz_name,_benz_prototypemodeldesignationtypeclass_value
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: "benz_prototypemodeldesignations",
      queryString: "",
      queryOptions: "",
    },
  };
  try {
    action = "get_Data__benz_carModel";

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
    });
    if (global.ApiAccessToken) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.error("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (!response) {
      console.error("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.error("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.error("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__benz_carModel);
      }
    } else {
      throw exception;
    }
    console.error("EXCEPTION: " + JSON.stringify(exception));
  }
}

export async function get_Data__benz_carModelGroup(callback): Promise<void> {
  let response: any;

  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: "benz_prototypemodelgroups",
      queryString: "$select=benz_prototypemodelgroupid,benz_name",
      queryOptions: "",
    },
  };
  try {
    action = "get_Data__benz_carModelGroup";
    console.log("get_Data__benz_carModelGroup... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    // if (!response) {
    //   console.log("no response...");
    //   authdialog.dialogFallback(callback, callRetrieveMultipleData);

    //   // throw new Error("Middle tier didn't respond");
    // }
    // else if (response.claims) {
    //   console.log("getUserData 2... ");
    //   // Microsoft Graph requires an additional form of authentication. Have the Office host
    //   // get a new token using the Claims string, which tells AAD to prompt the user for all
    //   // required forms of authentication.
    //   let mfaMiddletierToken: string = await OfficeRuntime.auth.getAccessToken({
    //     authChallenge: response.claims,
    //   });
    //   response = callRetrieveMultipleData(mfaMiddletierToken, "benz_supporttypes");
    // }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("get_Data__benz_carModelGroup error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__benz_carModelGroup);
      }
    } else {
      throw exception;
    }
    console.log("EXCEPTION: " + JSON.stringify(exception));
  }
}

export async function get_Data__benz_generalofferspecificoffer(callback): Promise<void> {
  let response: any;

  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet:
        "EntityDefinitions(LogicalName='benz_prototypesalesmeasure')/Attributes(LogicalName='benz_generalofferspecificoffer')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      queryString: "?$select=LogicalName&$expand=OptionSet($select=Options)",
      queryOptions: "",
    },
  };
  try {
    action = "get_Data__benz_generalofferspecificoffer";
    console.log("get_Data__benz_generalofferspecificoffer... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    // if (!response) {
    //   console.log("no response...");
    //   authdialog.dialogFallback(callback, callRetrieveMultipleData);

    //   // throw new Error("Middle tier didn't respond");
    // }
    // else if (response.claims) {
    //   console.log("getUserData 2... ");
    //   // Microsoft Graph requires an additional form of authentication. Have the Office host
    //   // get a new token using the Claims string, which tells AAD to prompt the user for all
    //   // required forms of authentication.
    //   let mfaMiddletierToken: string = await OfficeRuntime.auth.getAccessToken({
    //     authChallenge: response.claims,
    //   });
    //   response = callRetrieveMultipleData(mfaMiddletierToken, "benz_supporttypes");
    // }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("get_Data__benz_carModelGroup error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__benz_generalofferspecificoffer);
      }
    } else {
      throw exception;
    }
    console.log("EXCEPTION: " + JSON.stringify(exception));
  }
}
export async function get_Data__benz_fsProduct(callback): Promise<void> {
  let response: any;
  // $select=benz_prototypemodeldesignationid,benz_name,_benz_prototypemodeldesignationtypeclass_value
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet: "benz_fsproducts",
      queryString: "$select=benz_fsproductid,benz_name,benz_nameoffinanceproduct,_benz_financeproduct_value",
      queryOptions: "",
    },
  };
  try {
    action = "get_Data__benz_fsProduct";

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
    });
    if (global.ApiAccessToken) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.error("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    // AAD errors are returned to the client with HTTP code 200, so they do not trigger
    // the catch block below.
    if (!response) {
      console.error("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.error("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.error("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__benz_fsProduct);
      }
    } else {
      throw exception;
    }
    console.error("EXCEPTION: " + JSON.stringify(exception));
  }
}

export async function handleAADErrors(response: any, callback: any, callGetData: any, args?) {
  console.log(`retryGetMiddletierToken >0 ... ${action}`);
  authdialog.dialogFallback(callback, callGetData, args);
}

export async function retrieve_Data(callback): Promise<any> {
  let response: any;
  try {
    action = "retrieve_Data";
    console.log("retrieve_Data... ");
    console.log(global.Callapiaction);

    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveData(global.ApiAccessToken);
      } catch (e) {
        console.log("callCreate error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveData);
    } else {
      return await callback(response);
    }
  } catch (exception) {
    console.log("retrieve_Data error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, retrieve_Data);
      }
    } else {
      throw exception;
    }
    console.log("retrieve_Data EXCEPTION: " + JSON.stringify(exception));
  }
}
export async function post_Data(callback): Promise<any> {
  let response: any;
  try {
    action = "post_Data";
    console.log("post_Data... ");
    console.log(global.Callapiaction);
    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(global.ApiAccessToken);
      } catch (e) {
        console.log("callCreate error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(middletierToken);
      } catch (e) {
        console.log("callCreate error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callCreate error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callCreate);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callCreate);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callCreate);
    } else {
      return await callback(response);
    }
  } catch (exception) {
    console.log("get_Data__benz_carModelGroup error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, post_Data);
      }
    } else {
      throw exception;
    }
    console.log("post_Data EXCEPTION: " + JSON.stringify(exception));
  }
}
export async function post_Data_ReturnData(): Promise<any> {
  let response: any;
  try {
    action = "post_Data_ReturnData";
    console.log("post_Data_ReturnData... ");
    console.log(global.Callapiaction);
    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(global.ApiAccessToken);
      } catch (e) {
        console.error("callCreate error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(middletierToken);
      } catch (e) {
        console.error("callCreate error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callCreate(getGlobalVariable("accessToken"));
      } catch (e) {
        console.error("getGlobalVariable accessToken callCreate error");
      }
    }
  } catch (exception) {
    console.error("post_Data_ReturnData error");
    console.error("post_Data_ReturnData EXCEPTION: " + JSON.stringify(exception));
  }
  return response;
}

export async function update_Data_ReturnData(): Promise<any> {
  let response: any;
  try {
    action = "update_Data_ReturnData";
    console.log("update_Data_ReturnData... ");
    console.log(global.Callapiaction);
    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callUpdate(global.ApiAccessToken);
      } catch (e) {
        console.error("callUpdate error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callUpdate(middletierToken);
      } catch (e) {
        console.error("callUpdate error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callUpdate(getGlobalVariable("accessToken"));
      } catch (e) {
        console.error("getGlobalVariable accessToken callUpdate error");
      }
    }
  } catch (exception) {
    console.error("post_Data_ReturnData error");
    console.error("post_Data_ReturnData EXCEPTION: " + JSON.stringify(exception));
  }
  return response;
}
export async function get_Data__sm_memotype_mbhk(callback): Promise<void> {
  let response: any;
  global.Callapiaction = {
    name: "callapiaction",
    action: {
      entitySet:
        "EntityDefinitions(LogicalName='benz_prototypesalesmeasuremainform')/Attributes(LogicalName='benz_memotypeformbhk')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      queryString: "?$select=LogicalName&$expand=OptionSet($select=Options)",
      queryOptions: "",
    },
  };
  try {
    action = "get_Data__sm_memotype_mbhk";
    console.log("get_Data__sm_memotype_mbhk... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("get_Data__sm_memotype_mbhk error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__sm_memotype_mbhk);
      }
    } else {
      throw exception;
    }
    console.log("EXCEPTION: " + JSON.stringify(exception));
  }
}

export async function get_Count__byModth(callback): Promise<void> {
  let response: any;
  try {
    action = "get_Count__sm_byModth";
    console.log("get_Count__sm_byModth... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("get_Count__sm_byModth error");
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Count__byModth);
      }
    } else {
      throw exception;
    }
    console.log("EXCEPTION: " + JSON.stringify(exception));
  }
}

export async function getCountByModth(): Promise<any> {
  let response: any = null;
  try {
    action = "get_Count__sm_byModth";
    console.log("get_Count__sm_byModth... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.error("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.error("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.error("handleAADErrors...");
      // handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.error("handleAADErrors...");
      // handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.error("response is empty");
      // handleAADErrors(response, callback, callRetrieveMultipleData);
    }
  } catch (exception) {
    console.error("get_Count__sm_byModth error");
    console.error("EXCEPTION: " + JSON.stringify(exception));
  }
  return response;
}
export async function post_MapAssociate(conf) {
  let response: any;
  try {
    action = "get_Count__sm_byModth";
    console.log("get_Count__sm_byModth... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await MapAssociate(global.ApiAccessToken, conf);
      } catch (e) {
        console.log("MapAssociate error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await MapAssociate(middletierToken, conf);
      } catch (e) {
        console.log("MapAssociate error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await MapAssociate(getGlobalVariable("accessToken"), conf);
      } catch (e) {
        console.log("getGlobalVariable accessToken MapAssociate error");
      }
    }
  } catch (e) {
    console.error("get_Count__sm_byModth error");
    console.error("EXCEPTION: " + JSON.stringify(e));
    console.error(response);
  }
  return response;
}
export async function post_Disassociate(conf) {
  let response: any;
  try {
    action = "get_Count__sm_byModth";
    console.log("get_Count__sm_byModth... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callDisassociate(global.ApiAccessToken, conf);
      } catch (e) {
        console.log("callDisassociate error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callDisassociate(middletierToken, conf);
      } catch (e) {
        console.log("callDisassociate error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callDisassociate(getGlobalVariable("accessToken"), conf);
      } catch (e) {
        console.log("getGlobalVariable accessToken callDisassociate error");
      }
    }
  } catch (e) {
    console.error("get_Count__sm_byModth error");
    console.error("EXCEPTION: " + JSON.stringify(e));
    console.error(response);
  }
  return response;
}
export async function get_Data__main(callback): Promise<void> {
  let response: any;
  try {
    action = "get_Data__sm_main";
    console.log("get_Data__sm_main... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }

    console.log(`response len:${response}`);
    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("get_Data__sm_main error");
    console.log("EXCEPTION: " + JSON.stringify(exception));

    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, get_Data__main);
      }
    } else {
      throw exception;
    }
  }
}
export async function get_Data__main_return(): Promise<void> {
  let response: any = null;
  try {
    action = "get_Data__sm_main";
    console.log("get_Data__sm_main... ");

    let middletierToken: string;
    await getAccessToken().then((value) => {
      middletierToken = value;
      // console.log("after OfficeRuntime.auth.getAccessToken ");
      // console.log(value);
    });
    // console.log("get global.ApiAccessToken...", global.ApiAccessToken);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (middletierToken) {
      console.log("has middletierToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(middletierToken);
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    } else if (getGlobalVariable("accessToken")) {
      console.log("has GlobalVariable accessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("getGlobalVariable accessToken callRetrieveMultipleData error");
      }
    }
  } catch (exception) {
    console.log("get_Data__sm_main error");
    console.log("EXCEPTION: " + JSON.stringify(exception));
  }
  return response;
}
export async function RetrieveMultipleData(callback): Promise<void> {
  let response: any;
  try {
    console.log("RetrieveMultipleData");

    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        response = await callRetrieveMultipleData(global.ApiAccessToken ?? getGlobalVariable("accessToken"));
      } catch (e) {
        console.log("callRetrieveMultipleData error");
      }
    }

    if (!response) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.error) {
      console.log("handleAADErrors...");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else if (response.length === 0) {
      console.log("response is empty");
      handleAADErrors(response, callback, callRetrieveMultipleData);
    } else {
      callback(response);
    }
  } catch (exception) {
    console.log("RetrieveMultipleData error");
    console.log("EXCEPTION: " + JSON.stringify(exception));

    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (handleClientSideErrors(exception)) {
        console.log(`handleClientSideErrors... ${action}`);
        authdialog.dialogFallback(callback, RetrieveMultipleData);
      }
    } else {
      throw exception;
    }
  }
}

export async function RetrieveAndReturnMultipleData(callback, args?: any): Promise<any> {
  let response: any = null;
  try {
    console.log("RetrieveMultipleData", args);

    if (!global.ApiAccessToken) {
      authdialog.dialogFallback(callback, RetrieveAndReturnMultipleData, args);
    }

    console.log("has global.ApiAccessToken... ");
    try {
      //benz_supporttypes
      return await callRetrieveMultipleData(global.ApiAccessToken ?? getGlobalVariable("accessToken"), args);
    } catch (e) {
      console.log("callRetrieveMultipleData error");
      
    }
  } catch (exception) {
    console.error("RetrieveMultipleData error");
    console.error("EXCEPTION: " + JSON.stringify(exception));
    console.error(exception.stack);
  }
  return response;
}

export async function retrieve_DataReturn(obj?: any): Promise<any> {
  let response: any;
  try {
    console.log(global.Callapiaction);
    if (global.ApiAccessToken) {
      console.log("has global.ApiAccessToken... ");
      try {
        //benz_supporttypes
        return await callRetrieveData(global.ApiAccessToken, obj);
      } catch (e) {
        console.log("callCreate error");
      }
    }
  } catch (exception) {
    console.error("retrieve_Data error");
    console.error("retrieve_Data EXCEPTION: " + JSON.stringify(exception));
  }
  return response;
}
