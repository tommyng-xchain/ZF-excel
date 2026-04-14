"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.WebApiConfig = void 0;
class WebApiConfig {
    version;
    accessToken;
    url;
    /**
     * Constructor
     * @param config WebApiConfig
     */
    constructor(version, accessToken, url) {
        // If URL not provided, get it from Xrm.Context
        if (url == null) {
            const context = typeof GetGlobalContext !== 'undefined' ? GetGlobalContext() : Xrm.Utility.getGlobalContext();
            url = `${context.getClientUrl()}/api/data/v${version}`;
            this.url = url;
        }
        else {
            this.url = `${url}/api/data/v${version}`;
            this.url = this.url.replace('//', '/');
        }
        this.version = version;
        this.accessToken = accessToken;
    }
}
exports.WebApiConfig = WebApiConfig;
