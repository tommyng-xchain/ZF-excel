/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

import * as Benz from "./type";
/* global Excel */
import { PublicClientApplication, Configuration } from "@azure/msal-browser";

declare global {
  interface ObjectConstructor {
    typedKeys<T>(obj: T): Array<keyof T>;
  }
}
Object.typedKeys = Object.keys as any;

declare global {
  var mode: Benz.AppMode;

  var modules: any;

  var clientId: string;

  var environment_name: string;

  var authority: string;

  var PublicClientApp: PublicClientApplication;

  var msalConfig: Configuration;

  var ApiAccessToken: string;

  var action_back: any;

  var CurrentModule: string;

  var CurrentConfig: Benz.FormConfig;

  var CarModels: Benz.CarModel[];

  var SupportTypes: Benz.SupportType[];

  var TypeClasses: Benz.TypeClasses[];

  var CarModelGroups: Benz.CarModelGroup[];

  var Generalofferspecificoffers: Benz.Generalofferspecificoffer[];

  var mbhk: Benz.Generalofferspecificoffer[];

  var Generalofferspecificoffers: Benz.Generalofferspecificoffer[];

  var Memotypes_mbhk: Benz.Memotype_mbhk[];

  var Memotypes_mbfs: Benz.Memotype_mbfs[];

  var Memotypes_zf: Benz.Memotype_zf[];

  var Departments: Benz.Departments[];

  var BusinessUnits: Benz.BusinessUnits[];

  var Month: Benz.DropdownChoices[];

  var Callapiaction: any;

  var readyCallapiaction: any[];

  var MainFormObject: any;

  // var dataValidation: Object;
  var dataChoices: Object;

  var Account: any;

  var PowerAccount: Benz.Account;

  var AccountID: string;

  var form: Benz.Form;

  var MemoNo: string;

  var perpostedItems: object[];

  var perpostrelationship_items: any;

  var TableDataChangeAction: Excel.TableChangedEventArgs[];

  var homeAccountId: string;

  var currentSelectRangeAddress: string;

  var lastnumber: number;

  var choicesSets: Benz.choicesSets;

  var tableConfigs: Benz.tableConfig;

  var pagetable: {};

  var WorksheetTableConfig: {};

  var Table__model: any;

  var Table__modelgroup: any;

  var onChangeDataValidation: Benz.onChangeDataValidation_getdata;

  var modulesOChange: Benz.modulesOChange[];

  var loadingModal: any;

  var RowMeasureID: any[];

  var RowMeasureName: any[];

  var updatingRecord: string;

  var isZFRecord: boolean;

  var CurrentLineNumber: number;

  var getSheet: boolean;

  var updating: boolean;

  var RowSpecialCase: any[];

  var EnvironmentVariable: {};

  var ErrorMsg: string;
}
