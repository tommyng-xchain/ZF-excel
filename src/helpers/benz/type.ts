import { BENZ } from "./factory";

export enum AppMode {
  /**
   * @remarks
   *
   */
  UAT = "uat",
  /**
   * @remarks
   *
   */
  MIGRATION = "migration",
  /**
   * @remarks
   *
   */
  PRODUCTION = "production",
}

/* global Excel */
export type modules = {
  config: FormConfig;
  object: BENZ;
};

export enum action {
  create = "create",
  update = "update",
}
export type Account = {
  azureactivedirectoryobjectid: string;
  token: string;
  id?: string;
  fullname: string;
  email: string;
  roles?: string[];
  Teams?: string[];
  BusinessUnit?: string;
};

export type FormConfig = {
  action: action;
  module: string;
  moduleTypeFullName: string;
  specialcase: specialcase;
  memoid_set: string[];
  permissions: permissions;
  relationship: tablerelationship[];
  Table_Main_LogicalName: string;
  Table_Item_LogicalName: string;
  Table_item_main_relationship_LogicalName: string;
  Table_linenumber_LogicalName: string;
  StartButtonLabel: string;
  layout: FormLayout[];
  table: Table;
};

export type permissions = {
  start_button: string[];
  create_button: string[];
  Update_button: string[];
  GetData_button: string[];
  ImportData_button: string[];
};

export type tablerelationship = {
  id: string;
  to_LogicalName: string;
  from_LogicalName: string;
  from_FieldLogicalName: string;
  relationship_LogicalName: string;
  querys: query[];
  action: {};
};

export type specialcase = {
  available: boolean;
  target: string;
  targetTitle: string;
};

export type FormLayout = {
  mandatory: boolean;
  specialcase: boolean;
  address: string;
  rowIndex: number;
  index: number;
  EntityLogicalName: string;
  LogicalName: string;
  SchemaName?: string;
  ValueLogicalName: string;
  AttributeType: string;
  LookupInoutType: string;
  LookupEntityLogicalName: string;
  LookupEntityFieldLogicalName: string;
  LookupActionType: "create" | "Append to";
  LookupAllowUsingExist: boolean;
  Label: string;
  title: string;
  formulas: string;
  values: string;
  type: string;
  id: string;
  relationship: relationship;
  enable: boolean;
  ConditionalFormat: TableColumnsDatavalidation;
  cell: Cell;
  outOfTarget: outOfTarget;
  module: string;
  action: string[];
  needFixedValue?: boolean;
  fixedValue?: string;
  isFormula?: boolean;
  canPass?: boolean;
  setToString?: boolean;
  MemoNameGen?: boolean;
  useEnvironmentVaule?: string;
};

export type Cell = {
  columnHidden?: boolean;
  inputtype: string;
  query: query[];
  valuesMap: {};
  numberFormat: string;
  format: format;
  address: string;
  values: string;
  dataValidation: DataValidation;
  formulas: string[][];
  conditionalFormats: any;
  WorksheetSingleClicked: onchange_dataValidation[];
  TableChanged: onchange_dataValidation[];
  WorksheetChanged: onchange_dataValidation[];
  getvalueregex: string;
  querys: query[];
  pad: number;
  fouceCheckChange?: boolean;
  defalutValues?: string;
  checkSubmitTo?: boolean;
};

export type onchange_dataValidation = {
  type: string;
  tablefield: string[];
  columns: string[];
  separator: string;
  EntityLogicalName: string;
  EntityLogicalName1: string;
  FieldLogicalNames: string[];
  queryString: string;
  queryString1: string;
  queryStringCheck: string;
  queryOptions: string;
  regexp: string;
  querys: query[];
  continue: Continue;
  setFieldData: setFieldData[];
  targetvalue: string;
  errorAlert: errorAlert;
  findKey: string;
  ErrorMsg: string;
};
enum DataValidationAlertStyle {
  success = "success",
  error = "error",
  warning = "warning",
}

export type errorAlert = {
  message: string;
  showAlert?: boolean;
  style: DataValidationAlertStyle | "success" | "error" | "warning";
};

export type setFieldData = {
  ToLogicalName: string[];
  FromLogicalNames: string[];
  settype: "value" | "list" | "date";
  operator:
    | Excel.DataValidationOperator
    | "Between"
    | "NotBetween"
    | "EqualTo"
    | "NotEqualTo"
    | "GreaterThan"
    | "LessThan"
    | "GreaterThanOrEqualTo"
    | "LessThanOrEqualTo";
  ExceedCheck?: number;
  canPass?: boolean;
};
export type query = {
  id: string;
  type: string;
  add: boolean;
  EntityLogicalName: string;
  LogicalName: string;
  entitySet: string;
  queryString: string;
  queryOptions: string;
  expand: query;
  followFieldLogicalName: string;
  getvalueregex: string;
  getApiValue: string;
  valueset: string[];
  settoLogicalName: string;
  to_LogicalName: string;
  from_LogicalName: string;
  from_FieldLogicalName: string;
  relationship_LogicalName: string;
};
export type DataValidation = {
  target: string;
  rule: Excel.DataValidationRule;
  errorAlert: Excel.DataValidationErrorAlert;
  prompt: any;
  ignoreBlanks: any;
};
export type format = {
  fill: Excel.RangeFill;
  font: Excel.RangeFont;
  type: string;
  columnWidth: number;
  protection: protection;
  verticalAlignment: string;
  horizontalAlignment: string;
};

export type protection = {
  locked: boolean;
};
export type Table = {
  name: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  hasHeaders: boolean;
  columns: TableColumns[];
};

export type TableColumns = {
  mandatory: boolean;
  specialcase: boolean;
  address: string;
  rowIndex: number;
  index: number;
  EntityLogicalName: string;
  LogicalName: string;
  SchemaName?: string;
  ValueLogicalName: string;
  AttributeType: string;
  LookupInoutType: string;
  LookupEntityLogicalName: string;
  LookupEntityFieldLogicalName: string;
  LookupActionType: "create" | "Append to";
  LookupAllowUsingExist: boolean;
  Label: string;
  title: string;
  formulas: string;
  values: string;
  type: string;
  id: string;
  relationship: relationship;
  enable: boolean;
  ConditionalFormat: TableColumnsDatavalidation;
  cell: Cell;
  outOfTarget: outOfTarget;
  needFixedValue?: boolean;
  fixedValue?: string;
  isFormula?: boolean;
  canPass?: boolean;
  defalutValues?: string;
  setToString?: boolean;
  MemoNameGen?: boolean;
  useEnvironmentVaule?: string;
};

export type outOfTarget = {
  value: string;
  targetLogicalName: string;
};
export type TableColumnsDatavalidation = {
  separate: string;
  regexp: string;
  TrueCell: Cell;
  FalseCell: Cell;
};
export type relationship = {
  type: string;
  LogicalName: string;
  RelatedLogicalName: string;
};

export type benz_salesmeasure = {
  benz_name: string;
};
export type choices = {
  id: string;
  name: string;
};
export type SM_Form = {
  main_form: object;
  items: SM_Form_item[];
};
export type Form = {
  main: any;
  items: any[];
  relationship_items: any;
  object: {};
};

export type tablerelationshipItem = {
  id: string;
};

export type SM_Form_item = {
  in_model_group: Array<string>[];
  ex_model_group: Array<string>[];
  in_model: Array<string>[];
  ex_model: Array<string>[];
  data: object;
};

export type CarModel = choices & {};

export type CarModelGroup = choices & {};

export type SupportType = choices & {};

export type TypeClasses = choices & {};

export type Generalofferspecificoffer = choices & {};

export type Memotype = choices & {
  value: string;
};
export type DropdownChoices = choices & {
  value: string;
};

export type Memotype_mbhk = Memotype & {};

export type Memotype_mbfs = Memotype & {};

export type Memotype_zf = Memotype & {};

export type Departments = Memotype & {};

export type BusinessUnits = Memotype & {};

export type initConfig = {
  applyString: object;
  forms: formInitConfig[];
  user: queryUser;
  init_object: initObject[];
};
export type formInitConfig = {
  id: string;
  module: string;
  company: string;
};

export type User = {
  roles: userRoles;
};

export type userRoles = {};

export type queryUser = {
  roles: query;
};

export type initObject = {
  index: number;
  key: string;
  entitySet: string;
  queryString: string;
  queryOptions: string;
  Type: string;
  oType: string;
  fields: string[];
  columns: string[];
  searchPanes: boolean;
};

export type choicesSets = {
  [key: string]: DropdownChoices[];
};

export type tableConfig = {};

export type onChangeDataValidation_getdata = {
  args: changeArgs;
  config: FormConfig;
  columnIndex: number;
  type: string;
};

export type changeArgs =
  | Excel.TableChangedEventArgs
  | Excel.WorksheetChangedEventArgs
  | Excel.WorksheetSingleClickedEventArgs;

export type modulesOChange = {
  address: string;
  row: number;
  Inclusion: DropdownChoices[];
  InclusionGroup: DropdownChoices[];
  Exclusion: DropdownChoices[];
  ExclusionGroup: DropdownChoices[];
  model: DropdownChoices[];
};

export type mainformobject = {
  Field: FormLayout;
  Range: Excel.Range;
};
export type mainformobjectKey = {
  [key: string]: mainformobject;
};

export type Continue = {
  parameters1?: string;
  parameters2?: string;
  operator?: "==" | "!=" | "And" | "Or";
  targetLogicalName?: string;
  value?: string;
  continue?: Continue[];
};

export type queryUpdateObject = {
  config: FormConfig;
  object: object;
};

export type dataValidationCheck = {
  config: FormConfig;
  onchangeconfig: onchange_dataValidation;
  value: any;
  fconfig: FormLayout | TableColumns;
  row?: number;
};

export type passConfig = {
  passVaildation: passVaildation;
};

export type passVaildation = {
  mode: string;
  moduleName: string;
  actionKey: any[];
  LogicalName: any[];
};
