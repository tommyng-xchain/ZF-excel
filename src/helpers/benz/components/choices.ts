/* global global, console*/
import * as BenzType from "../type";
export function set_choicesSets(key: string, obj: BenzType.DropdownChoices[]) {
  try {
    console.log(`set choices list ${key.toString().toLowerCase()}`);
    console.log(obj);
    global.choicesSets[key.toString().toLowerCase()] = obj;
  } catch (e) {
    console.error(`Error: Can't set choices list - ${key.toString().toLowerCase()}`);
    console.error(obj);
    console.error(e.stack);
  }
}

export function get_choicesSets(key): BenzType.DropdownChoices[] {
  try {
    console.log(`get choicesSets ${key.toString().toLowerCase()}`);
    console.log(global.choicesSets[key.toString().toLowerCase()]);
    return global.choicesSets[key.toString().toLowerCase()];
  } catch (e) {
    console.error(`Error: Can't get choices list - ${key.toString().toLowerCase()}`);
    console.error(e.stack);
  }
}

export function set_dataChoices(key, obj) {
  try {
    console.log(`set choices ${key.toString().toLowerCase()} `);
    console.log(obj);

    global.dataChoices[key.toString().toLowerCase()] = obj;
    console.log(global.dataChoices);
  } catch (e) {
    console.error(`Error: Can't set choices list - ${key.toString().toLowerCase()}`);
    console.error(obj);
    console.error(e.stack);
  }
}

export function get_choices_list(key) {
  try {
    console.log(`get choices list - ${key.toString().toLowerCase()}`);
    console.log(global.dataChoices[key.toString().toLowerCase()]);
    return global.dataChoices[key.toString().toLowerCase()];
  } catch (e) {
    console.error(`Error: Can't get choices list - ${key}`);
    console.error(e.stack);
  }
}

export function get_choices(key, value): string {
  try {
    console.log(`get choices ${key.toString().toLowerCase()} ${value}`);
    console.log(global.dataChoices[key.toString().toLowerCase()]);
    console.log(global.dataChoices);
    if (global.dataChoices[key.toString().toLowerCase()]) {
      const choi = global.dataChoices[key.toString().toLowerCase()];
      console.log(choi);
      const val = choi.find((ele) => ele.name == value.toString());
      console.log(val);
      return val.value;
    } else {
      return null;
    }
  } catch (e) {
    console.error(`Error: Can't get choices list - ${key.toString().toLowerCase()}`);
    console.error(e.stack);
    return null;
  }
}
export function get_choicesByValue(choiceskey, key, value, output): string {
  try {
    console.log(`get choices ${choiceskey.toString().toLowerCase()} ${value}`);
    console.log(global.dataChoices[choiceskey.toString().toLowerCase()]);
    console.log(global.dataChoices);
    if (global.dataChoices[choiceskey.toString().toLowerCase()]) {
      const choi = global.dataChoices[choiceskey.toString().toLowerCase()];
      console.log(choi);
      const val = choi.find((ele) => ele[key] == value.toString());
      console.log(val);
      return val[output];
    } else {
      return null;
    }
  } catch (e) {
    console.error(`Error: Can't get choices list - ${key.toString().toLowerCase()}`);
    console.error(e.stack);
    return null;
  }
}
export function get_dataValidation_source(key): string {
  try {
    console.log(`get choices source - ${key.toString().toLowerCase()}`);
    console.log(global.dataChoices[key.toString().toLowerCase()]);
    console.log(global.dataChoices);
    return global.dataChoices[key.toString().toLowerCase()].map((option) => option.name ?? option).join(",") ?? "";
  } catch (e) {
    console.error(`Error: Can't get choices source - ${key}`);
    console.error(e.stack);
    return "";
  }
}
