/* global global, require, console */
import { jwtDecode } from "jwt-decode";
import { retrieve_DataReturn, RetrieveAndReturnMultipleData, RetrieveMultipleData } from "../dataverse-data-helper";
import { urlencoded } from "express";

var config = require("./config/init.json");

const user_conf = config.user.info;
const teams_conf = config.user.teams;

export async function setAccount(t) {
  let userToken = jwtDecode(t); // Using the https://www.npmjs.com/package/jwt-decode library.
  console.log("**userToken setAccount", userToken);
  global.PowerAccount = {
    email: userToken["unique_name"].toString(),
    fullname: userToken["name"].toString(),
    azureactivedirectoryobjectid: userToken["oid"].toString(),
    token: t.toString(),
  };
  await setAccountRoles();
}

export async function setAccountRoles() {
  const res = await RetrieveAndReturnMultipleData(setAccountRoles, {
    entitySet: user_conf.entitySet,
    queryString: user_conf.queryString.replace("{value}", global.PowerAccount.azureactivedirectoryobjectid.toString()),
    queryOptions: "",
  });
  setAccountRolesRes(res);
}

export async function setAccountRolesRes(res) {
  console.log("setUserRoles");
  global.PowerAccount.roles = res.value[0]["systemuserroles_association"].map((e) => e.name);

  global.PowerAccount["teams"] = res.value[0]["teammembership_association"].map(
    (e) => `'${e.name.replace("&", "%26")}'`
  );
  await setteamsRoles();
}

export async function setteamsRoles() {
  const teams = global.PowerAccount["teams"];
  console.log("setteamsRoles", teams);
  const res = await RetrieveAndReturnMultipleData(setteamsRoles, {
    entitySet: teams_conf.entitySet,
    queryString: teams_conf.queryString.replace("{value}", teams.join(",")),
    queryOptions: "",
  });
  setteamrolesToUser(res);
  console.log(global.PowerAccount);
  // await RetrieveMultipleData(setteamroles);
}

export async function setteamrolesToUser(res) {
  console.log("setteamrolesToUser");
  for (const v of res.value) {
    global.PowerAccount.roles = global.PowerAccount.roles.concat(v["teamroles_association"].map((e) => e.name));
  }
  console.log(global.PowerAccount);
}

// export async function getRolus(){
//   /* system user       "azureactivedirectoryobjectid"
//   https://org8390b622.api.crm5.dynamics.com/api/data/v9.2/systemusers?$filter=azureactivedirectoryobjectid%20eq%20%27f7661210-228c-4c4f-8a8b-584a1cea0089%27&$expand=systemuserroles_association($select=name),teammembership_association($select=name)

// https://org8390b622.api.crm5.dynamics.com/api/data/v9.2/teams(247415c1-671c-ef11-840a-002248eecc9d)?$expand=teamroles_association($select=name)
//   */
// }
