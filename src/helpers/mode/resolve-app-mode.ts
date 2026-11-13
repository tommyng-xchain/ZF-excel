import * as Benz from "../benz/type";

/* global __APP_MODE__ */
declare const __APP_MODE__: string;

export function getBuildAppMode(): Benz.AppMode {
  return __APP_MODE__ === Benz.AppMode.UAT ? Benz.AppMode.UAT : Benz.AppMode.PRODUCTION;
}

export function getBuildAppConfig(): {
  clientId: string;
  environment_name: string;
  authority: string;
} {
  if (getBuildAppMode() === Benz.AppMode.UAT) {
    return require("./uat/app.json");
  }
  return require("./production/app.json");
}
