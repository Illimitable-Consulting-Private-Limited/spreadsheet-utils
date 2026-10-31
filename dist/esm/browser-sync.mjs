// src/schema/InvalidError.js
var InvalidError = class extends Error {
  constructor(reason) {
    super("invalid");
    this.reason = reason;
  }
};

// src/schema/types/Number.js
function NumberType(value) {
  if (typeof value === "string") {
    const stringifiedValue = value;
    value = Number(value);
    if (String(value) !== stringifiedValue) {
      throw new InvalidError("not_a_number");
    }
  }
  if (typeof value !== "number") {
    throw new InvalidError("not_a_number");
  }
  if (isNaN(value)) {
    throw new InvalidError("invalid_number");
  }
  if (!isFinite(value)) {
    throw new InvalidError("out_of_bounds");
  }
  return value;
}

// src/schema/types/String.js
function StringType(value) {
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    if (isNaN(value)) {
      throw new InvalidError("invalid_number");
    }
    if (!isFinite(value)) {
      throw new InvalidError("out_of_bounds");
    }
    return String(value);
  }
  throw new InvalidError("not_a_string");
}

// src/schema/types/Boolean.js
function BooleanType(value) {
  if (typeof value === "boolean") {
    return value;
  }
  throw new InvalidError("not_a_boolean");
}

// src/schema/types/Date.js
function DateType(value) {
  if (value instanceof Date) {
    if (isNaN(value.valueOf())) {
      throw new InvalidError("out_of_bounds");
    }
    return value;
  }
  throw new InvalidError("not_a_date");
}

// src/utils/isObject.js
var objectConstructor = {}.constructor;
function isObject(object) {
  return object !== void 0 && object !== null && object.constructor === objectConstructor;
}

// src/utils/checkpoint.js
var latestCheckpointTimestamp;
function checkpoint(name) {
  const now = Date.now();
  const shouldOutputLog = typeof global !== "undefined" ? Boolean(global.READ_EXCEL_FILE_CHECKPOINTS) : typeof window !== "undefined" ? Boolean(window.READ_EXCEL_FILE_CHECKPOINTS) : false;
  if (shouldOutputLog) {
    if (latestCheckpointTimestamp) {
      console.log("  -", now - latestCheckpointTimestamp, "ms");
    }
    console.log("*", name);
  }
  latestCheckpointTimestamp = now;
}

// src/schema/parseSheetData.js
var EMPTY_CELL_VALUE = null;
function parseSheetData(data, schema, optionsCustom) {
  checkpoint("parse sheet data using schema");
  const objects = [];
  let errors = [];
  const parsedRows = parseSheetDataWithPerRowErrors(data, schema, optionsCustom);
  let parsedRowIndex = 0;
  for (const { object, errors: rowErrors } of parsedRows) {
    if (rowErrors) {
      errors = errors.concat(
        rowErrors.map(
          // Add row number property to each row error.
          (rowError) => ({ ...rowError, row: parsedRowIndex + 1 })
        )
      );
    } else {
      objects.push(object);
    }
    parsedRowIndex++;
  }
  checkpoint("end");
  if (errors.length > 0) {
    return { errors };
  }
  return { objects, rows: objects };
}
function parseSheetDataWithPerRowErrors(data, schema, optionsCustom) {
  validateSchema(schema);
  const options = applyDefaultOptions(optionsCustom);
  const [columns, ...dataRows] = data;
  return dataRows.map((row) => parseDataRow(row, schema, columns, options));
}
function parseDataRow(dataRow, schema, columns, options) {
  const schemaEntry = {
    schema
  };
  const {
    value,
    isEmptyValue: isEmptyValue2,
    errors,
    children
  } = parseProperty(dataRow, schemaEntry, void 0, columns, options);
  const dummyParentObject = {
    // The "dummy" parent object has a "dummy" value.
    // This value is irrelevant because it won't be read anywhere.
    value: PARSED_OBJECT_TREE_START,
    // The "dummy" parent object is empty if the parsed row is empty.
    isEmptyValue: isEmptyValue2,
    // The "dummy" object has the same errors as the parsed row.
    errors,
    // The parsed object by default is not required to have any data
    // so the "dummy" object is not required.
    isRequired: void 0
  };
  const requiredErrors = runPendingRequiredValidations(
    schemaEntry,
    value,
    isEmptyValue2,
    errors,
    children,
    // Simulate a "dummy" parent object for the top-level object.
    dummyParentObject.isRequired,
    dummyParentObject.value,
    dummyParentObject.isEmptyValue,
    dummyParentObject.errors,
    columns
  );
  if (errors || requiredErrors) {
    return {
      errors: (errors || []).concat(requiredErrors || [])
    };
  }
  return {
    object: transformValue(value, isEmptyValue2, void 0, options)
  };
}
function parseObject(row, schema, path, columns, options) {
  const object = {};
  let isEmptyObject = true;
  let errors = [];
  const children = [];
  for (const key of Object.keys(schema)) {
    const child = parseProperty(row, schema[key], getPropertyPath(key, path), columns, options);
    if (child.errors) {
      errors = errors.concat(child.errors);
    } else {
      object[key] = transformValue(child.value, child.isEmptyValue, getPropertyPath(key, path), options);
      if (isEmptyObject && !child.isEmptyValue) {
        isEmptyObject = false;
      }
    }
    children.push({
      ...child,
      // `schemaEntry` will be used when running `required` validation of this property (later),
      schemaEntry: schema[key]
    });
  }
  if (errors.length > 0) {
    return {
      // Return the errors.
      errors,
      // Return the `children` because `required` validations still have to be run (later).
      children
    };
  }
  return {
    value: object,
    isEmptyValue: isEmptyObject,
    // Return the `children` because `required` validations still have to be run (later).
    children
  };
}
function parseProperty(row, schemaEntry, path, columns, options) {
  const columnIndex = schemaEntry.column ? columns.indexOf(schemaEntry.column) : void 0;
  const isMissingColumn = schemaEntry.column ? columnIndex < 0 : void 0;
  const {
    value,
    isEmptyValue: isEmptyValue2,
    errors,
    children
  } = schemaEntry.column ? isMissingColumn ? { value: options.propertyValueWhenColumnIsMissing, isEmptyValue: true } : parseCellValueWithPossibleErrors(row[columnIndex], schemaEntry, columnIndex, options) : parseObject(
    row,
    schemaEntry.schema,
    path,
    columns,
    options
  );
  if (errors) {
    return {
      // Return the errors.
      errors,
      // Return the `children` because `required` validations still have to be run (later).
      children
    };
  }
  return {
    value,
    isEmptyValue: isEmptyValue2,
    // Return the `children` because `required` validations still have to be run (later).
    children
  };
}
function parseCellValueWithPossibleErrors(cellValue, schemaEntry, columnIndex, options) {
  const {
    value,
    isEmptyValue: isEmptyValue2,
    error: errorMessage,
    reason: errorReason
  } = parseCellValue(cellValue, schemaEntry, options);
  if (errorMessage) {
    const error = createError({
      error: errorMessage,
      reason: errorReason,
      column: schemaEntry.column,
      columnIndex,
      valueType: schemaEntry.type,
      value: cellValue
    });
    return {
      errors: [error]
    };
  }
  return {
    value,
    isEmptyValue: isEmptyValue2
  };
}
function parseCellValue(cellValue, schemaEntry, options) {
  if (cellValue === void 0) {
    return {
      value: options.propertyValueWhenColumnIsMissing,
      isEmptyValue: true
    };
  }
  if (cellValue === EMPTY_CELL_VALUE) {
    return {
      value: options.propertyValueWhenCellIsEmpty,
      isEmptyValue: true
    };
  }
  if (Array.isArray(schemaEntry.type)) {
    return parseArrayValue(cellValue, schemaEntry, options);
  }
  return parseValue(cellValue, schemaEntry, options);
}
function parseArrayValue(value, schemaEntry, options) {
  if (typeof value !== "string") {
    return {
      error: "not_a_string"
    };
  }
  let isEmptyArray = true;
  const errors = [];
  const reasons = [];
  const values = parseSeparatedSubstrings(value, options.separatorCharacter).map((substring) => {
    if (errors.length > 0) {
      return;
    }
    if (!substring) {
      errors.push("invalid");
      reasons.push("syntax");
      return;
    }
    const {
      value: value2,
      isEmptyValue: isEmptyValue2,
      error,
      reason
    } = parseValue(substring, schemaEntry, options);
    if (error) {
      errors.push(error);
      reasons.push(reason);
      return;
    }
    if (isEmptyArray && !isEmptyValue2) {
      isEmptyArray = false;
    }
    return value2;
  });
  if (errors.length > 0) {
    return {
      error: errors[0],
      reason: reasons[0]
    };
  }
  return {
    value: values,
    isEmptyValue: isEmptyArray
  };
}
function parseValue(value, schemaEntry, options) {
  if (value === EMPTY_CELL_VALUE) {
    return {
      value,
      isEmptyValue: true
    };
  }
  let result;
  if (schemaEntry.type) {
    result = parseValueOfType(
      value,
      // Get the type of the value.
      //
      // Handle the case if it's a comma-separated value.
      // Example `type`: String[]
      // Example Input Value: 'Barack Obama, "String, with, colons", Donald Trump'
      // Example Parsed Value: ['Barack Obama', 'String, with, colons', 'Donald Trump']
      //
      Array.isArray(schemaEntry.type) ? schemaEntry.type[0] : schemaEntry.type,
      options
    );
  } else {
    result = { value };
  }
  if (result.error) {
    return result;
  }
  if (value === EMPTY_CELL_VALUE) {
    return {
      value,
      isEmptyValue: true
    };
  }
  if (schemaEntry.oneOf) {
    const errorAndReason = validateOneOf(result.value, schemaEntry.oneOf);
    if (errorAndReason) {
      return errorAndReason;
    }
  }
  if (schemaEntry.validate) {
    try {
      schemaEntry.validate(result.value);
    } catch (error) {
      return { error: error.message };
    }
  }
  return {
    value: result.value,
    isEmptyValue: isEmptyValue(result.value)
  };
}
function validateOneOf(value, oneOf) {
  if (oneOf.indexOf(value) < 0) {
    return { error: "invalid", reason: "unknown" };
  }
}
function parseValueOfType(value, type) {
  switch (type) {
    case String:
      return parseValueUsingTypeParser(value, StringType);
    case Number:
      return parseValueUsingTypeParser(value, NumberType);
    case Date:
      return parseValueUsingTypeParser(value, DateType);
    case Boolean:
      return parseValueUsingTypeParser(value, BooleanType);
    default:
      if (typeof type !== "function") {
        throw new Error(`Unsupported schema \`type\`: ${type && type.name || type}`);
      }
      return parseValueUsingTypeParser(value, type);
  }
}
function parseValueUsingTypeParser(value, type) {
  try {
    const parsedValue = type(value);
    if (parsedValue === void 0) {
      return { value: EMPTY_CELL_VALUE };
    }
    return { value: parsedValue };
  } catch (error) {
    const result = { error: error.message };
    if (error.reason) {
      result.reason = error.reason;
    }
    return result;
  }
}
function getNextSubstring(string, separatorCharacter, startIndex) {
  let i = 0;
  let substring = "";
  while (startIndex + i < string.length) {
    const character = string[startIndex + i];
    if (character === separatorCharacter) {
      return [substring, i];
    } else {
      substring += character;
      i++;
    }
  }
  return [substring, i];
}
function parseSeparatedSubstrings(string, separatorCharacter) {
  const elements = [];
  let index = 0;
  while (index < string.length) {
    const [substring, length] = getNextSubstring(string, separatorCharacter, index);
    index += length + separatorCharacter.length;
    elements.push(substring.trim());
  }
  return elements;
}
function transformValue(value, isEmptyValue2, path, options) {
  if (isEmptyValue2) {
    if (isObject(value)) {
      return options.transformEmptyObject(value, { path });
    } else if (Array.isArray(value)) {
      return options.transformEmptyArray(value, { path });
    }
  }
  return value;
}
function getPropertyPath(propertyName, parentObjectPath) {
  return `${parentObjectPath ? parentObjectPath + "." : ""}${propertyName}`;
}
function runPendingRequiredValidations(schemaEntry, value, isEmptyValue2, errors, children, parentObjectIsRequired, parentObjectValue, parentObjectValueIsEmpty, parentObjectErrors, columns) {
  let requiredErrors = [];
  const isRequired = isPropertyRequired(
    schemaEntry,
    parentObjectIsRequired,
    parentObjectValue,
    parentObjectValueIsEmpty,
    parentObjectErrors
  );
  if (isRequired && isEmptyValue2) {
    requiredErrors.push(createError({
      error: "required",
      column: schemaEntry.column,
      columnIndex: columns.indexOf(schemaEntry.column),
      valueType: schemaEntry.type,
      value
    }));
  }
  if (children) {
    for (const child of children) {
      const requiredErrorsOfChild = runPendingRequiredValidations(
        child.schemaEntry,
        child.value,
        child.isEmptyValue,
        child.errors,
        child.children,
        // The following properties describe the parent object of the `child`,
        // i.e. the current (iterated) object.
        isRequired,
        value,
        isEmptyValue2,
        errors,
        columns
      );
      if (requiredErrorsOfChild) {
        requiredErrors = requiredErrors.concat(requiredErrorsOfChild);
      }
    }
  }
  if (requiredErrors.length > 0) {
    return requiredErrors;
  }
}
function isPropertyRequired(schemaEntry, parentObjectIsRequired, parentObjectValue, parentObjectValueIsEmpty, parentObjectErrors) {
  if (parentObjectIsRequired === false && (parentObjectValueIsEmpty || parentObjectErrors)) {
    return false;
  }
  return schemaEntry.required && (typeof schemaEntry.required === "boolean" ? schemaEntry.required : (
    // If there were any non-`required` errors when parsing the parent object,
    // the `parentObject` will be `undefined`. In that case, "complex" `required()`
    // validations — the ones where `required` is a function — can't really be run
    // because those validations assume a fully and correctly parsed parent object
    // be passed as an argument, and the thing is that the `parentObject` is unknown.
    // As a result, only "basic" `required` validations could be run,
    // i.e. the ones where `required` is just a boolean, and "complex" `required`
    // validations, i.e. the ones where `required` is a functions, should be skipped,
    // because it's better to skip some `required` errors than to trigger falsy ones.
    parentObjectErrors ? false : schemaEntry.required(parentObjectValue)
  ));
}
function createError({
  column,
  columnIndex,
  valueType,
  value,
  error: errorMessage,
  reason
}) {
  const error = {
    error: errorMessage,
    column,
    columnIndex,
    value
  };
  if (reason) {
    error.reason = reason;
  }
  if (valueType) {
    error.type = valueType;
  }
  return error;
}
function validateSchema(schema) {
  for (const key of Object.keys(schema)) {
    const schemaEntry = schema[key];
    if (typeof schemaEntry.type === "object" && !Array.isArray(schemaEntry.type)) {
      throw new Error("When defining a nested schema, use a `schema` property instead of a `type` property");
    }
    if (!schemaEntry.schema) {
      if (!schemaEntry.column) {
        throw new Error(`"column" not defined for schema entry "${key}".`);
      }
    }
  }
  validateObjectSchemaRequiredProperty(schema, void 0);
}
function validateObjectSchemaRequiredProperty(schema, required) {
  if (required !== void 0 && required !== false) {
    throw new Error(`In a schema, a nested object can have a \`required\` property but the only allowed value is \`undefined\` or \`false\`. Otherwise, a "required" error for a nested object would have to include a specific \`column\` title and a nested object doesn't have one. You've specified the following \`required\`: ${required}`);
  }
  for (const key of Object.keys(schema)) {
    if (isObject(schema[key].schema)) {
      if (schema[key].column) {
        throw new Error(`In a schema, \`column\` property is only allowed when describing a property value rather than a nested object. Key: ${key}. Schema:
${JSON.stringify(schema[key], null, 2)}`);
      }
      validateObjectSchemaRequiredProperty(schema[key].schema, schema[key].required);
    }
  }
}
function isEmptyValue(value) {
  return value === void 0 || value === null;
}
var DEFAULT_OPTIONS = {
  propertyValueWhenColumnIsMissing: void 0,
  propertyValueWhenCellIsEmpty: null,
  // shouldSkipRequiredValidationWhenColumnIsMissing: () => false,
  // `transformEmptyObject(object, { path })` applies to both the top-level object
  // and any of its nested objects.
  transformEmptyObject: () => null,
  transformEmptyArray: () => null,
  separatorCharacter: ","
};
function applyDefaultOptions(options) {
  if (options) {
    return {
      ...DEFAULT_OPTIONS,
      ...options
    };
  } else {
    return DEFAULT_OPTIONS;
  }
}
var PARSED_OBJECT_TREE_START = {};

// src/core/shapeReadResult.js
function convertRowsToObjects(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const seen = /* @__PURE__ */ new Map();
  const headers = (rows[0] || []).map((h, i) => {
    let base = h != null ? String(h).trim() : "";
    if (!base) base = `column${i + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count > 0 ? `${base}_${count + 1}` : base;
  });
  const result = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      obj[headers[c]] = row[c] !== void 0 ? row[c] : null;
    }
    result.push(obj);
  }
  return result;
}
function shapeReadResult(tabularResult, options = {}) {
  const worksheets = tabularResult.worksheets || [];
  if (options.format === "tabular" || options.worksheets === true) {
    return tabularResult;
  }
  const asObjects = options.asObjects === true || options.header === true;
  if (options.allSheets === true) {
    return worksheets.map((ws, i) => {
      const sheetData = ws.data || [];
      return {
        name: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
        data: asObjects ? convertRowsToObjects(sheetData) : sheetData
      };
    });
  }
  let selectedWorksheet = worksheets[0];
  if (options.sheet !== void 0) {
    if (typeof options.sheet === "number") {
      const index = options.sheet - 1;
      if (index < 0 || index >= worksheets.length) {
        throw new Error(`Sheet not found: ${options.sheet}`);
      }
      selectedWorksheet = worksheets[index];
    } else if (typeof options.sheet === "string") {
      const found = worksheets.find(
        (ws) => (ws.worksheetName || ws.sheet) === options.sheet
      );
      if (!found) {
        throw new Error(`Sheet not found: "${options.sheet}"`);
      }
      selectedWorksheet = found;
    }
  }
  let data = selectedWorksheet && selectedWorksheet.data ? selectedWorksheet.data : [];
  if (options.maxColumns && data.length > 0) {
    data = data.map((r) => r.length > options.maxColumns ? r.slice(0, options.maxColumns) : r);
  }
  if (options.maxRows && data.length > options.maxRows) {
    data = data.slice(0, options.maxRows);
  }
  if (typeof options.onRow === "function") {
    data.forEach((row, i) => options.onRow(row, i));
  }
  if (options.schema) {
    return parseSheetData(data, options.schema, options);
  }
  if (asObjects) {
    return convertRowsToObjects(data);
  }
  return data;
}

// src/sync/unzipSync.js
import { unzipSync } from "fflate";

// src/zip/UnzipError.js
var UnzipError = class extends Error {
};
function createUnzipError(error) {
  const unzipError = new UnzipError(error.message);
  if (error.stack) {
    unzipError.stack = error.stack;
  }
  if (Error.captureStackTrace) {
    Error.captureStackTrace(unzipError, createUnzipError);
  }
  unzipError.cause = error;
  return unzipError;
}

// src/sync/unzipSync.js
var DEFAULT_MAX_BYTES = 256 * 1024 * 1024;
function isFlateError(error) {
  return Boolean(error) && typeof error.code === "number";
}
function unzipBytesSync(bytes, options = {}) {
  const filter = options.filter;
  const maxBytes = options.maxBytes || DEFAULT_MAX_BYTES;
  let total = 0;
  try {
    return unzipSync(bytes, {
      filter: (file) => {
        if (filter && !filter(file)) {
          return false;
        }
        if (file.originalSize) {
          total += file.originalSize;
          if (total > maxBytes) {
            throw new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`);
          }
        }
        return true;
      }
    });
  } catch (error) {
    throw isFlateError(error) ? createUnzipError(error) : error;
  }
}
function unzipArchiveSyncAsPromise(bytes, options = {}) {
  const nameFilter = options.filter;
  try {
    return Promise.resolve(unzipBytesSync(bytes, {
      maxBytes: options.maxBytes,
      filter: nameFilter ? (file) => nameFilter(file.name) : void 0
    }));
  } catch (error) {
    return Promise.reject(error);
  }
}
function zipEntryNamesSync(bytes) {
  const names = [];
  try {
    unzipSync(bytes, {
      filter: (file) => {
        names.push(file.name);
        return false;
      }
    });
  } catch {
  }
  return names;
}

// src/utils/helpers.js
var columnNames = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
var getColumnNameCache = {};
var getColumnName = function(num) {
  if (typeof num === "string") {
    num = parseInt(num);
  }
  let result = getColumnNameCache[num];
  if (typeof result !== "undefined") {
    return result;
  } else {
    let letter = "";
    let quotient = num + 1;
    let remainder;
    while (quotient > 0) {
      remainder = (quotient - 1) % 26;
      letter = columnNames[remainder] + letter;
      quotient = (quotient - remainder) / 26 | 0;
    }
    return getColumnNameCache[num] = letter;
  }
};
var getCellNameFromCoords = function(x, y) {
  if (typeof x === "string") {
    x = parseInt(x);
  }
  if (typeof y === "string") {
    y = parseInt(y);
  }
  return getColumnName(x) + (y + 1);
};
var getCoordsFromCellName = function(name) {
  if (!name) {
    return [];
  }
  let letter = null;
  let number = [];
  let size = name.length;
  let c;
  for (let i = 0; i < size; i++) {
    c = name.charCodeAt(i);
    if (c >= 65 && c <= 90) {
      letter = letter * 26 + (c - 64);
    } else if (c >= 48 && c <= 57) {
      number.push(c - 48);
    }
  }
  let total = null;
  size = number.length;
  if (size) {
    for (let i = 0; i < size; i++) {
      total = total * 10 + number[i];
    }
    if (total) {
      total = total - 1;
    }
  }
  if (letter === null) {
    return [letter, total];
  }
  return [letter - 1, total];
};
var getTextContent = function(node) {
  if (!node) return "";
  if (typeof node === "string") return node;
  if (Array.isArray(node)) {
    return node.map((n) => getTextContent(n)).join("");
  }
  if (node.type === "#text" && node.props) {
    const textProp = node.props.find((p) => p.name === "textContent");
    if (textProp) return textProp.value;
  }
  if (node.children) {
    return getTextContent(node.children);
  }
  return "";
};
var parseAttributes = function(element) {
  const attrs = {};
  if (element && element.props) {
    element.props.forEach((attr) => {
      attrs[attr.name] = attr.value;
    });
  }
  if (element && element.attributes) {
    element.attributes.forEach((attr) => {
      attrs[attr.name] = attr.value;
    });
  }
  return attrs;
};
var readUInt16LE = function(buffer, offset) {
  return buffer[offset] | buffer[offset + 1] << 8;
};
var readUInt32LE = function(buffer, offset) {
  return (buffer[offset] | buffer[offset + 1] << 8 | buffer[offset + 2] << 16 | buffer[offset + 3] << 24) >>> 0;
};
var readFloat64LE = function(buffer, offset) {
  const view = new DataView(buffer.buffer, buffer.byteOffset + offset, 8);
  return view.getFloat64(0, true);
};
var convertWidthToPixels = function(width, unit = "pt") {
  if (!width || isNaN(width)) return 100;
  switch (unit) {
    case "pt":
      return Math.round(parseFloat(width) * 7.15);
    case "char":
      return Math.round(parseFloat(width) * 8.43);
    case "px":
      return Math.round(parseFloat(width));
    default:
      return Math.round(parseFloat(width) * 7.15);
  }
};

// src/xlsx/parseSpreadsheetInfo.js
function parseSpreadsheetInfo(content, parseXml2) {
  const state = createInitialState();
  return parseXml2(
    content,
    state,
    onOpenTag,
    null,
    null
  ).then(() => getResultFromState(state));
  function createInitialState() {
    return {
      workbookPr: void 0,
      sheets: []
    };
  }
  function getResultFromState(state2) {
    return {
      epoch1904: state2.workbookPr ? state2.workbookPr.epoch1904 : false,
      sheets: state2.sheets
    };
  }
  function onOpenTag(tagName, attributes, state2) {
    if (tagName === "workbookPr") {
      if (!state2.workbookPr) {
        state2.workbookPr = {
          epoch1904: attributes.date1904 === "1"
        };
      }
    } else if (tagName === "sheet") {
      if (attributes.name) {
        state2.sheets.push({
          // `sheetId` attribute value is an arbitrary, `1`-based unique positive integer
          // assigned to a worksheet, typically starting at `1` for the first sheet.
          //  Deleting and adding new sheets might cause the sheetId values to become non-sequential.
          // For example, `sheetId`s could be `1`, `2`, `4`, if sheet `3` was deleted.
          id: Number(attributes.sheetId),
          name: attributes.name,
          relationId: attributes.id
        });
      }
    }
  }
}

// src/xlsx/parseFilePaths.js
function parseFilePaths(content, parseXml2) {
  const RELATIONSHIPS_BASE_URL_TRANSITIONAL_STANDARD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const RELATIONSHIPS_BASE_URL_STRICT_STANDARD = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
  const state = createInitialState();
  return parseXml2(
    content,
    state,
    onOpenTag,
    null,
    null
  ).then(() => getResultFromState(state));
  function createInitialState() {
    return {
      sheets: {},
      sharedStrings: void 0,
      styles: void 0
    };
  }
  function getResultFromState(state2) {
    return state2;
  }
  function onOpenTag(tagName, attributes, state2) {
    if (tagName === "Relationship") {
      addFilePathForRelation(state2, attributes.Id, attributes.Type, attributes.Target);
    }
  }
  function addFilePathForRelation(state2, id, type, target) {
    switch (type) {
      case RELATIONSHIPS_BASE_URL_TRANSITIONAL_STANDARD + "styles":
      case RELATIONSHIPS_BASE_URL_STRICT_STANDARD + "styles":
        state2.styles = getFilePathFromRelationTarget(target);
        break;
      case RELATIONSHIPS_BASE_URL_TRANSITIONAL_STANDARD + "sharedStrings":
      case RELATIONSHIPS_BASE_URL_STRICT_STANDARD + "sharedStrings":
        state2.sharedStrings = getFilePathFromRelationTarget(target);
        break;
      case RELATIONSHIPS_BASE_URL_TRANSITIONAL_STANDARD + "worksheet":
      case RELATIONSHIPS_BASE_URL_STRICT_STANDARD + "worksheet":
        state2.sheets[id] = getFilePathFromRelationTarget(target);
        break;
    }
  }
  function getFilePathFromRelationTarget(path) {
    if (path[0] === "/") {
      return path.slice("/".length);
    }
    return "xl/" + path;
  }
}

// src/xlsx/parseStyles.js
function parseStyles(content, parseXml2) {
  const state = createInitialState();
  return parseXml2(
    content,
    state,
    onOpenTag,
    onCloseTag,
    null
  ).then(() => getResultFromState(state));
  function createInitialState() {
    return {
      numberFormats: {},
      baseStyles: [],
      // The first `164` elements of the `styles` array are going to be `undefined`
      // because those represent the built-in "default" styles in `.xlsx` specification.
      // These "default" styles have IDs from `0` to `163`, i.e. according to their index.
      styles: [],
      cellStyleXfs: false,
      cellXfs: false
    };
  }
  function getResultFromState(state2) {
    return state2.styles.map((style) => {
      if (style.xfId) {
        const { xfId, ...styleProperties } = style;
        return {
          ...state2.baseStyles[xfId],
          ...styleProperties
        };
      } else {
        return style;
      }
    });
  }
  function onOpenTag(tagName, attributes, state2) {
    if (tagName === "numFmt") {
      state2.numberFormats[attributes.numFmtId] = {
        id: Number(attributes.numFmtId),
        template: attributes.formatCode
      };
    } else if (tagName === "cellStyleXfs") {
      state2.cellStyleXfs = true;
    } else if (tagName === "cellXfs") {
      state2.cellXfs = true;
    } else if (tagName === "xf") {
      if (state2.cellStyleXfs) {
        state2.baseStyles.push(parseCellStyle(attributes));
      } else if (state2.cellXfs) {
        const style = parseCellStyle(attributes, state2.numberFormats);
        if (attributes.xfId) {
          style.xfId = Number(attributes.xfId);
        }
        state2.styles.push(style);
      }
    }
  }
  function onCloseTag(tagName, state2) {
    if (tagName === "cellStyleXfs") {
      state2.cellStyleXfs = false;
    } else if (tagName === "cellXfs") {
      state2.cellXfs = false;
    }
  }
  function parseCellStyle(attributes, numberFormats) {
    const style = {};
    const { numFmtId } = attributes;
    if (numFmtId) {
      if (numberFormats && numberFormats[numFmtId]) {
        style.numberFormat = numberFormats[numFmtId];
      } else {
        style.numberFormat = { id: Number(numFmtId) };
      }
    }
    return style;
  }
}

// src/xlsx/parseSharedStrings.js
function parseSharedStrings(content, parseXml2) {
  const state = createInitialState();
  return parseXml2(
    content,
    state,
    onOpenTag,
    onCloseTag,
    onText
  ).then(() => getResultFromState(state));
  function createInitialState() {
    return {
      si: void 0,
      strings: []
    };
  }
  function getResultFromState(state2) {
    return state2.strings;
  }
  function onOpenTag(tagName, attributes, state2) {
    if (tagName === "si") {
      state2.si = createInitialStateInSharedString();
    } else if (state2.si) {
      onOpenTagInSharedString(tagName, attributes, state2.si);
    }
  }
  function onCloseTag(tagName, state2) {
    if (tagName === "si") {
      state2.strings.push(state2.si.string);
      state2.si = void 0;
    } else if (state2.si) {
      onCloseTagInSharedString(tagName, state2.si);
    }
  }
  function onText(text, state2) {
    if (state2.si) {
      onTextInSharedString(text, state2.si);
    }
  }
  function createInitialStateInSharedString() {
    return {
      t: false,
      r: false,
      rPh: false,
      string: ""
    };
  }
  function onOpenTagInSharedString(tagName, attributes, state2) {
    if (tagName === "t") {
      state2.t = true;
    } else if (tagName === "r") {
      state2.r = true;
    } else if (tagName === "rPh") {
      state2.rPh = true;
    }
  }
  function onCloseTagInSharedString(tagName, state2) {
    if (tagName === "t") {
      state2.t = false;
    } else if (tagName === "r") {
      state2.r = false;
    } else if (tagName === "rPh") {
      state2.rPh = false;
    }
  }
  function onTextInSharedString(text, state2) {
    if (state2.rPh) {
    } else if (state2.t) {
      if (state2.r) {
        state2.string += text;
      } else {
        state2.string = text;
      }
    }
  }
}

// src/xlsx/parseExcelTimestamp.js
function parseExcelTimestamp(excelSerialDate, epoch1904) {
  const NUMBER_OF_LEAP_YEARS_BETWEEN_1900_AND_1970 = 17;
  const JANUARY_0TH_1900_DAY = 1;
  const ERRONEOUS_FEBRUARY_29_1990_DAY = 1;
  const DAY = 24 * 60 * 60 * 1e3;
  const DAYS_IN_YEAR = 365;
  if (epoch1904) {
    excelSerialDate += (1904 - 1900) * DAYS_IN_YEAR + JANUARY_0TH_1900_DAY + ERRONEOUS_FEBRUARY_29_1990_DAY;
  }
  const daysBeforeUnixEpoch2 = JANUARY_0TH_1900_DAY + ERRONEOUS_FEBRUARY_29_1990_DAY + (1970 - 1900) * DAYS_IN_YEAR + NUMBER_OF_LEAP_YEARS_BETWEEN_1900_AND_1970;
  return Math.floor((excelSerialDate - daysBeforeUnixEpoch2) * DAY);
}

// src/xlsx/isDateFormat.js
var DATE_FORMAT_SPECIFIC_LOCALE_PREFIX = /^\[\$-[^\]]+\]/;
var DATE_FORMAT_ALLOW_ANY_OTHER_TEXT_SUFFIX = /;@$/;
var IS_DATE_FORMAT_CACHE = {};
function isDateFormatCached(template) {
  if (template in IS_DATE_FORMAT_CACHE) {
    return IS_DATE_FORMAT_CACHE[template];
  }
  template = template.toLowerCase();
  template = template.replace(DATE_FORMAT_SPECIFIC_LOCALE_PREFIX, "");
  template = template.replace(DATE_FORMAT_ALLOW_ANY_OTHER_TEXT_SUFFIX, "");
  const tokens = template.split(/\W+/);
  const result = tokens.length === 0 ? false : tokens.every((token) => DATE_TEMPLATE_TOKENS.indexOf(token) >= 0);
  IS_DATE_FORMAT_CACHE[template] = result;
  return result;
}
var DATE_TEMPLATE_TOKENS = [
  // Seconds (min two digits). Example: "05".
  "ss",
  // Minutes (min two digits). Example: "05". Could also be "Months". Weird.
  "mm",
  // Hours. Example: "1".
  "h",
  // Hours (min two digits). Example: "01".
  "hh",
  // "AM" part of "AM/PM". Lowercased just in case.
  "am",
  // "PM" part of "AM/PM". Lowercased just in case.
  "pm",
  // Day. Example: "1"
  "d",
  // Day (min two digits). Example: "01"
  "dd",
  // Month (numeric). Example: "1".
  "m",
  // Month (numeric, min two digits). Example: "01". Could also be "Minutes". Weird.
  "mm",
  // Month (shortened month name). Example: "Jan".
  "mmm",
  // Month (full month name). Example: "January".
  "mmmm",
  // Two-digit year. Example: "20".
  "yy",
  // Full year. Example: "2020".
  "yyyy",
  // I don't have any idea what "e" means.
  // It's used in "built-in" XLSX formats:
  // * 27 '[$-404]e/m/d';
  // * 36 '[$-404]e/m/d';
  // * 50 '[$-404]e/m/d';
  // * 57 '[$-404]e/m/d';
  "e"
];

// src/xlsx/isDateFormatStyle.js
function isDateFormatStyle(style, defaultDateFormat, shouldGuessDateFormatFromNumberFormatTemplate) {
  if (!style.numberFormat) {
    return false;
  }
  if (
    // Whether it's a "number format" that's conventionally used for storing date timestamps.
    BUILT_IN_DATE_FORMAT_IDS.indexOf(style.numberFormat.id) >= 0 || // Whether it's a "number format" that uses a "formatting template"
    // that the developer is certain is a date formatting template.
    defaultDateFormat && style.numberFormat.template === defaultDateFormat || // Whether the "smart formatting template" feature is not disabled
    // and it has detected that it's a date formatting template by looking at it.
    shouldGuessDateFormatFromNumberFormatTemplate && style.numberFormat.template && isDateFormatCached(style.numberFormat.template)
  ) {
    return true;
  }
  return false;
}
var LOCALE_INDEPENDENT_BUILT_IN_DATE_FORMAT_IDS = [
  14,
  // mm-dd-yy
  15,
  // d-mmm-yy
  16,
  // d-mmm
  17,
  // mmm-yy
  18,
  // h:mm AM/PM
  19,
  // h:mm:ss AM/PM
  20,
  // h:mm
  21,
  // h:mm:ss
  22,
  // m/d/yy h:mm
  45,
  // mm:ss
  46,
  // [h]:mm:ss
  47
  // mmss.0
];
var MAINLAND_CHINESE_OR_TAIWANESE_LOCALE_BUILT_IN_DATE_FORMAT_IDS = [
  27,
  // [$-404]e/m/d OR yyyy"年"m"月"
  28,
  // [$-404]e"年"m"月"d"日" OR m"月"d"日"
  29,
  // [$-404]e"年"m"月"d"日" OR m"月"d"日"
  30,
  // m/d/yy OR m-d-yy
  31,
  // yyyy"年"m"月"d"日" OR yyyy"年"m"月"d"日"
  32,
  // hh"時"mm"分" OR h"时"mm"分"
  33,
  // hh"時"mm"分"ss"秒" OR h"时"mm"分"ss"秒"
  34,
  // 上午/下午hh"時"mm"分" OR 上午/下午h"时"mm"分"
  35,
  // 上午/下午hh"時"mm"分"ss"秒" OR 上午/下午h"时"mm"分"ss"秒"
  36,
  // [$-404]e/m/d OR yyyy"年"m"月"
  50,
  // [$-404]e/m/d OR yyyy"年"m"月"
  51,
  // [$-404]e"年"m"月"d"日" OR m"月"d"日"
  52,
  // 上午/下午hh"時"mm"分" OR yyyy"年"m"月"
  53,
  // 上午/下午hh"時"mm"分"ss"秒" OR m"月"d"日"
  54,
  // [$-404]e"年"m"月"d"日" OR m"月"d"日"
  55,
  // 上午/下午hh"時"mm"分" OR 上午/下午h"时"mm"分"
  56,
  // 上午/下午hh"時"mm"分"ss"秒" OR 上午/下午h"时"mm"分"ss"秒"
  57,
  // [$-404]e/m/d OR yyyy"年"m"月"
  58
  // [$-404]e"年"m"月"d"日" OR m"月"d"日"
];
var JAPANESE_OR_KOREAN_LOCALE_BUILT_IN_DATE_FORMAT_IDS = [
  27,
  // [$-411]ge.m.d OR yyyy"年" mm"月" dd"日"
  28,
  // [$-411]ggge"年"m"月"d"日" OR mm-dd
  29,
  // [$-411]ggge"年"m"月"d"日" OR mm-dd
  30,
  // m/d/yy OR mm-dd-yy
  31,
  // yyyy"年"m"月"d"日" OR yyyy"년" mm"월" dd"일"
  32,
  // h"時"mm"分" OR h"시" mm"분"
  33,
  // h"時"mm"分"ss"秒" OR h"시" mm"분" ss"초"
  34,
  // yyyy"年"m"月" OR yyyy-mm-dd
  35,
  // m"月"d"日" OR yyyy-mm-dd
  36,
  // [$-411]ge.m.d OR yyyy"年" mm"月" dd"日"
  50,
  // [$-411]ge.m.d OR yyyy"年" mm"月" dd"日"
  51,
  // [$-411]ggge"年"m"月"d"日" OR mm-dd
  52,
  // yyyy"年"m"月" OR yyyy-mm-dd
  53,
  // m"月"d"日" OR yyyy-mm-dd
  54,
  // [$-411]ggge"年"m"月"d"日" OR mm-dd
  55,
  // yyyy"年"m"月" OR yyyy-mm-dd
  56,
  // m"月"d"日" OR yyyy-mm-dd
  57,
  // [$-411]ge.m.d OR yyyy"年" mm"月" dd"日"
  58
  // [$-411]ggge"年"m"月"d"日" OR mm-dd
];
var THAI_LOCALE_BUILT_IN_DATE_FORMAT_IDS = [
  71,
  // ว/ด/ปปปป
  72,
  // ว-ดดด-ปป
  73,
  // ว-ดดด
  74,
  // ดดด-ปป
  75,
  // ช:นน
  76,
  // ช:นน:ทท
  77,
  // ว/ด/ปปปป ช:นน
  78,
  // นน:ทท
  79,
  // [ช]:นน:ทท
  80,
  // นน:ทท.0
  81
  // d/m/bb
];
var BUILT_IN_DATE_FORMAT_IDS = LOCALE_INDEPENDENT_BUILT_IN_DATE_FORMAT_IDS.concat(
  // Add Mainland Chinese or Taiwanese date format IDs that haven't already been added.
  MAINLAND_CHINESE_OR_TAIWANESE_LOCALE_BUILT_IN_DATE_FORMAT_IDS
).concat(
  // Add Japanese or Korean date format IDs that haven't already been added.
  JAPANESE_OR_KOREAN_LOCALE_BUILT_IN_DATE_FORMAT_IDS.filter(
    (numberFormatId) => MAINLAND_CHINESE_OR_TAIWANESE_LOCALE_BUILT_IN_DATE_FORMAT_IDS.indexOf(numberFormatId) < 0
  )
).concat(
  // Add Thai date format IDs that haven't already been added.
  THAI_LOCALE_BUILT_IN_DATE_FORMAT_IDS.filter(
    (numberFormatId) => MAINLAND_CHINESE_OR_TAIWANESE_LOCALE_BUILT_IN_DATE_FORMAT_IDS.indexOf(numberFormatId) < 0
  ).filter(
    (numberFormatId) => JAPANESE_OR_KOREAN_LOCALE_BUILT_IN_DATE_FORMAT_IDS.indexOf(numberFormatId) < 0
  )
);

// src/xlsx/parseCell.js
var EMPTY_CELL_VALUE2 = null;
var EMPTY_CELL = [null, EMPTY_CELL_VALUE2];
function parseCell(t, s, v, inlineString, [
  sharedStrings,
  styles,
  epoch1904,
  defaultDateFormat,
  dateTemplateParser,
  parseNumberCustom
]) {
  switch (t || "n") {
    // `t="str"` means that the cell value is calculated using a formula.
    // The formula is defined as the text of a child `<f/>` element.
    //
    // It could optionally include a `<v/>` element whose text is the cached result
    // of the calculation from the last time the file was saved in a spreadsheet editor application.
    //
    // An optional `<v/>` element holds a pre-computed result of the formula defined by `<f/>`.
    //
    // Example:
    //
    // <c r="B3" t="str">
    // 	<f>CONCATENATE(C1,D1)</f>
    // 	<v>C1ValueD1Value</v>
    // </c>
    //
    // Here's a guide on formulas in XLSX files:
    // https://github.com/MiniMax-AI/skills/blob/main/skills/minimax-xlsx/references/validate.md
    //
    case "str":
      if (v === void 0) {
        return "VALUE_MISSING";
      }
      if (!v) {
        return EMPTY_CELL;
      }
      return ["s", v];
    // `t="inlineStr"` means that `<is/>` holds the string value.
    //
    // Inside a `<c t="inlineStr"/>`, the specification requires there to exist an `<is/>` element,
    // and within that `<is/>` element it requires to exist a `<t/>` element.
    //
    // Example:
    //
    // <c r="A1" s="1" t="inlineStr">
    //   <is>
    //     <t>
    //       Test 123
    //     </t>
    //   </is>
    // </c>
    //
    case "inlineStr":
      if (inlineString === void 0) {
        return "VALUE_MISSING";
      }
      return ["s", inlineString];
    // `type="s"` means that the string value is stored in the Shared Strings Table.
    // This way it attempts to compress the `.xlsx` file by reusing all string values
    // in case they repeat throughout the spreadsheet.
    //
    // This optimization can't be used when writing an `.xlsx` file in a "streaming"
    // fashion, i.e. when the entire spreadsheet data is not known in adavance
    // at the start of writing the file.
    // But it can be used in all other situations. And hence, it is used.
    // So this is the most common cell type, actually.
    //
    // Example:
    //
    // <c r="A3" t="s">
    //   <v>3</v>
    // </c>
    //
    case "s":
      if (!v) {
        return "VALUE_MISSING";
      }
      const sharedStringIndex = Number(v);
      if (isNaN(sharedStringIndex) || sharedStrings[sharedStringIndex] === void 0) {
        return "VALUE_INVALID";
      }
      return ["s", sharedStrings[sharedStringIndex]];
    // Boolean (TRUE/FALSE) values are stored as either "1" or "0" in cells of type "b".
    //
    // Example:
    //
    // <c r="A1" t="b">
    //   <v>1</v>
    // </c>
    //
    case "b":
      if (!v) {
        return "VALUE_MISSING";
      }
      if (v === "1") {
        return ["b", true];
      }
      if (v === "0") {
        return ["b", false];
      }
      return "VALUE_INVALID";
    // If cell type is "e", the `<v/>` element's text is an error code string (required).
    //
    // Example:
    //
    // <c r="A1" t="e">
    //   <f>1/0</f>
    //   <v>#DIV/0!</v>
    // </c>
    //
    case "e":
      if (!v) {
        return "VALUE_MISSING";
      }
      return ["e", v];
    // XLSX supports date cells of type "d", though it seems like it (almost?) never
    // uses type "d" for storing dates, preferring type "n" and numeric timestamp instead.
    // The value of a "d" cell is supposedly a string in "ISO 8601" format.
    // I haven't seen an `.xlsx` file having such cells.
    //
    // Example:
    //
    // <c r="A1" s="1" t="d">
    //   <v>
    //     2021-06-10T00:47:45.700Z
    //   </v>
    // </c>
    //
    case "d":
      if (!v) {
        return EMPTY_CELL;
      }
      const parsedDate = new Date(v);
      if (isNaN(parsedDate.valueOf())) {
        return "VALUE_INVALID";
      }
      return ["d", parsedDate.getTime()];
    // type "n" is used for numeric cells.
    //
    // An optional `s` attribute defines how this number should be formatted — 
    // it should be a zero-based index of the style (XF record) in `styles.xml`.
    //
    // Example:
    //
    // <c r="A1" s="1" t="n">
    //   <v>123.45</v>
    // </c>
    //
    case "n":
      if (!v) {
        return EMPTY_CELL;
      }
      if (s) {
        const styleId = Number(s);
        if (isNaN(styleId) || styles[styleId] === void 0) {
          return "FORMAT_INVALID";
        }
        if (isDateFormatStyle(styles[styleId], defaultDateFormat, dateTemplateParser)) {
          const timestamp = Number(v);
          if (isNaN(timestamp)) {
            return "VALUE_INVALID";
          }
          return ["d", parseExcelTimestamp(timestamp, epoch1904)];
        }
      }
      if (parseNumberCustom) {
        return ["n", v];
      }
      const number = Number(v);
      if (isNaN(number)) {
        return "VALUE_INVALID";
      }
      return ["n", number];
    default:
      return "TYPE_INVALID";
  }
}

// src/xlsx/parseCellAddress.js
function parseCellAddress(cellAddress) {
  let columnNumber = 0;
  let i = 0;
  while (i < cellAddress.length) {
    const charCode = cellAddress.charCodeAt(i);
    if (charCode >= 48 && charCode <= 57) {
      const rowNumber = Number(cellAddress.slice(i));
      if (isNaN(rowNumber)) {
        invalidCellAddress(cellAddress);
      }
      return [
        // Row number (starting at `1`).
        rowNumber,
        // Column number (starting at `1`).
        columnNumber
      ];
    }
    columnNumber *= 26;
    columnNumber += cellAddress.charCodeAt(i) - 64;
    i++;
  }
  invalidCellAddress(cellAddress);
}
function invalidCellAddress(cellAddress) {
  throw new Error(`<c r="${cellAddress}">`);
}

// src/xlsx/InvalidSpreadsheetError.js
var InvalidSpreadsheetError = class extends Error {
  constructor(message) {
    super(message);
    this.name = "InvalidSpreadsheetError";
  }
};

// src/xlsx/parseSheet.js
var EMPTY_CELL_VALUE3 = null;
function parseSheet(content, parseXml2, { sharedStrings, styles, epoch1904, options }) {
  const parseCellParameters = [
    sharedStrings,
    styles,
    epoch1904,
    options.dateFormat,
    // defaultDateFormat
    options.smartDateParser !== false,
    // dateTemplateParser
    options.parseNumber
    // parseNumberCustom
  ];
  let rows = [];
  let errors = [];
  const state = createInitialState();
  return parseXml2(
    content,
    state,
    onOpenTag,
    onCloseTag,
    onText,
    onProgress
  ).then(() => {
    const { rowCount, columnCount, dataRowCount, dataColumnCount } = state.sheetData;
    if (dataRowCount < rowCount) {
      rows = rows.slice(0, dataRowCount);
    }
    if (dataColumnCount < columnCount) {
      let i = 0;
      while (i < rows.length) {
        if (rows[i].length > dataColumnCount) {
          rows[i] = rows[i].slice(0, dataColumnCount);
        }
        i++;
      }
    }
    const startedAt = Date.now();
    for (const row of rows) {
      while (row.length < dataColumnCount) {
        row.push(EMPTY_CELL_VALUE3);
      }
    }
    return rows;
  });
  function createInitialState() {
    return {
      dimension: void 0,
      sheetData: void 0
    };
  }
  function getRowsFromState(state2) {
    return state2.sheetData.rows;
  }
  function setRowsInState(state2, rows2) {
    state2.sheetData.rowIndexShift += state2.sheetData.rows.length - rows2.length;
    state2.sheetData.rows = rows2;
  }
  function getErrorsFromState(state2) {
    return state2.sheetData.errors;
  }
  function setErrorsInState(state2, errors2) {
    state2.sheetData.errors = errors2;
  }
  const THROW_ON_FIRST_CELL_ERROR = true;
  function throwInvalidCellError({ row, column, error }) {
    throw new InvalidSpreadsheetError(`<c/> at row ${row}, col ${column}: ${error}`);
  }
  function onProgress(end) {
    const rowsRead = getRowsFromState(state);
    const errorsEncountered = getErrorsFromState(state);
    if (end) {
      rows = rows.concat(rowsRead);
      errors = errors.concat(errorsEncountered);
      if (errors.length > 0) {
        throwInvalidCellError(errors[0]);
      }
    } else {
      if (rowsRead.length > 1) {
        const finalizedRows = rowsRead.slice(0, -1);
        rows = rows.concat(finalizedRows);
        errors = errors.concat(errorsEncountered);
        setRowsInState(state, rowsRead.slice(-1));
        setErrorsInState(state, []);
      }
    }
  }
  function onOpenTag(tagName, attributes, state2) {
    if (tagName === "dimension") {
      state2.dimension = parseSheetDimensionRef(attributes.ref);
    } else if (tagName === "sheetData") {
      state2.sheetData = createInitialStateInSheetData();
    } else if (state2.sheetData) {
      onOpenTagInSheetData(tagName, attributes, state2.sheetData);
    }
  }
  function onCloseTag(tagName, state2) {
    if (state2.sheetData) {
      onCloseTagInSheetData(tagName, state2.sheetData);
    }
  }
  function onText(text, state2) {
    if (state2.sheetData) {
      onTextInSheetData(text, state2.sheetData);
    }
  }
  function parseSheetDimensionRef(ref) {
    let dimensions = ref.split(":").map(parseCellAddress);
    if (dimensions.length === 1) {
      dimensions = [dimensions[0], dimensions[0]];
    }
    return dimensions;
  }
  function createInitialStateInSheetData() {
    return {
      c: void 0,
      rows: [],
      row: void 0,
      rowNumber: void 0,
      // How many rows have been removed from the start of `state.rows`
      // as part of `onProgress()` handler calls.
      rowIndexShift: 0,
      // Current position in the sheet.
      cursor: [0, 0],
      // Total row count.
      rowCount: 0,
      // Total column count.
      columnCount: 0,
      // Non-empty row count.
      dataRowCount: 0,
      // Non-empty column count.
      dataColumnCount: 0,
      // Cell with errors.
      errors: []
    };
  }
  function onOpenTagInSheetData(tagName, attributes, state2) {
    if (tagName === "row") {
      if (attributes.r) {
        state2.rowNumber = Number(attributes.r);
      }
      state2.row = [];
    } else if (tagName === "c") {
      state2.c = createInitialStateInCell();
      state2.c.attributes = attributes;
    } else if (state2.c) {
      onOpenTagInCell(tagName, attributes, state2.c);
    }
  }
  function onCloseTagInSheetData(tagName, state2) {
    if (tagName === "row") {
      if (state2.rowNumber) {
        const previousRowNumber = state2.rowIndexShift + state2.rows.length;
        if (state2.rowNumber <= previousRowNumber) {
          throw new InvalidSpreadsheetError(`Out-of-place <row/> number ${state2.rowNumber} follows <row/> number ${previousRowNumber}`);
        }
        while (state2.rowNumber > state2.rowIndexShift + state2.rows.length + 1) {
          state2.rows.push([]);
        }
      }
      if (options && options.maxColumns && state2.row.length > options.maxColumns) {
        state2.row = state2.row.slice(0, options.maxColumns);
      }
      state2.rows.push(state2.row);
      if (options && options.maxRows && state2.rows.length >= options.maxRows && options.throwOnLimit) {
        throw new Error(`Row limit of ${options.maxRows} exceeded`);
      }
      if (state2.row.length > 0) {
        state2.dataRowCount = state2.rowNumber;
      }
      if (state2.rowNumber > state2.rowCount) {
        state2.rowCount = state2.rowNumber;
      }
      state2.row = void 0;
      state2.rowNumber = void 0;
    } else if (tagName === "c") {
      const cell = parseCellFromXmlData(state2.c);
      if (cell.row < state2.cursor[0] || cell.row === state2.cursor[0] && cell.column <= state2.cursor[1]) {
        throw new InvalidSpreadsheetError(`Out-of-place <c/> at row ${cell.row} col ${cell.column} follows <c/> at row ${state2.cursor[0]} col ${state2.cursor[1]}`);
      }
      state2.cursor[0] = cell.row;
      state2.cursor[1] = cell.column;
      if (!state2.rowNumber) {
        state2.rowNumber = cell.row;
      }
      if (cell.error) {
        if (THROW_ON_FIRST_CELL_ERROR) {
          throwInvalidCellError(cell);
        }
        state2.errors.push(cell);
      } else if (cell.value !== EMPTY_CELL_VALUE3) {
        while (cell.column > state2.row.length + 1) {
          state2.row.push(EMPTY_CELL_VALUE3);
        }
        state2.row.push(cell.value);
        if (cell.column > state2.dataColumnCount) {
          state2.dataColumnCount = cell.column;
        }
      }
      if (cell.column > state2.columnCount) {
        state2.columnCount = cell.column;
      }
      state2.c = void 0;
    } else if (state2.c) {
      onCloseTagInCell(tagName, state2.c);
    }
  }
  function onTextInSheetData(text, state2) {
    if (state2.c) {
      onTextInCell(text, state2.c);
    }
  }
  function parseCellFromXmlData({
    attributes,
    inlineString,
    vText
  }) {
    const [row, column] = parseCellAddress(attributes.r);
    const errorOrTypeAndValue = parseCellAndTrimValue(
      attributes.t,
      attributes.s,
      vText,
      inlineString,
      parseCellParameters,
      options.trim !== false
    );
    if (typeof errorOrTypeAndValue === "string") {
      return {
        row,
        column,
        error: errorOrTypeAndValue
        // // Report the "raw" unparsed value of the cell for potential debugging.
        // // Also report the cell type and the format in case of a numeric value.
        // //
        // // For "inline string" cells, the value should actually be the `inlineString` argument
        // // rather than `vText` argument, but the only case when it could throw an error
        // // when parsing an "inline string" cell is `VALUE_MISSING` which means that
        // // `inlineString` argument is `undefined`, same as `vText` argument in this case,
        // // so the resulting `value` property is correct anyway.
        // //
        // value: vText,
        // type: attributes.t,
        // formatId: attributes.s
      };
    }
    return {
      row,
      column,
      value: parseCellValue2(errorOrTypeAndValue[1], errorOrTypeAndValue[0])
    };
  }
  function parseCellAndTrimValue(t, s, v, inlineString, parameters, trimStrings) {
    const errorOrTypeAndValue = parseCellWithRepairAbility(
      t,
      s,
      v,
      inlineString,
      parameters
    );
    if (Array.isArray(errorOrTypeAndValue) && errorOrTypeAndValue[0] === "s") {
      if (trimStrings) {
        errorOrTypeAndValue[1] = errorOrTypeAndValue[1].trim();
      }
      if (errorOrTypeAndValue[1] === "") {
        return EMPTY_CELL;
      }
    }
    return errorOrTypeAndValue;
  }
  function parseCellWithRepairAbility(t, s, v, inlineString, parameters) {
    const errorOrTypeAndValue = parseCell(
      t,
      s,
      v,
      inlineString,
      parameters
    );
    if (errorOrTypeAndValue === "VALUE_MISSING") {
      switch (t || "n") {
        // * If the cell is defined by a formula.
        // * Or contains an inline string.
        // * Or contains a shared string.
        // * Or contains a boolean value.
        case "str":
        case "inlineStr":
        case "s":
        case "b":
          return EMPTY_CELL;
      }
    }
    if (t === "e") {
      return EMPTY_CELL;
    }
    return errorOrTypeAndValue;
  }
  function parseCellValue2(value, type) {
    if (type === "n") {
      if (options.parseNumber) {
        return options.parseNumber(value);
      }
      return value;
    } else if (type === "d") {
      return new Date(value);
    } else {
      return value;
    }
  }
  function createInitialStateInCell() {
    return {
      v: false,
      is: false,
      t: false,
      r: false,
      rPh: false,
      vText: void 0,
      inlineString: void 0,
      attributes: void 0
    };
  }
  function onOpenTagInCell(tagName, attributes, state2) {
    if (tagName === "v") {
      state2.v = true;
    } else if (tagName === "is") {
      state2.is = true;
      state2.inlineString = "";
    } else if (tagName === "t") {
      state2.t = true;
    } else if (tagName === "r") {
      state2.r = true;
    } else if (tagName === "rPh") {
      state2.rPh = true;
    }
  }
  function onCloseTagInCell(tagName, state2) {
    if (tagName === "v") {
      state2.v = false;
      state2.vText || (state2.vText = "");
    } else if (tagName === "is") {
      state2.is = false;
    } else if (tagName === "t") {
      state2.t = false;
    } else if (tagName === "r") {
      state2.r = false;
    } else if (tagName === "rPh") {
      state2.rPh = false;
    }
  }
  function onTextInCell(text, state2) {
    if (state2.v) {
      state2.vText = text;
    } else if (state2.is) {
      if (state2.rPh) {
      } else if (state2.t) {
        if (state2.r) {
          state2.inlineString += text;
        } else {
          state2.inlineString = text;
        }
      }
    }
  }
  function getSheetDimensions(cells) {
    let minRow = cells.length === 0 ? 0 : 1;
    let minCol = cells.length === 0 ? 0 : 1;
    let maxRow = 0;
    let maxCol = 0;
    for (const cell of cells) {
      if (maxRow < cell.row) {
        maxRow = cell.row;
      }
      if (maxCol < cell.column) {
        maxCol = cell.column;
      }
    }
    return [
      [minRow, minCol],
      [maxRow, maxCol]
    ];
  }
}

// src/utils/convertValuesFromUint8ArraysToStrings.js
import { strFromU8 as strFromU8_ } from "fflate";
function convertValuesFromUint8ArraysToStrings(entries) {
  checkpoint("convert files to strings");
  const convertedEntries = {};
  for (const key of Object.keys(entries)) {
    convertedEntries[key] = strFromU8(entries[key]);
  }
  return convertedEntries;
}
function strFromU8(data) {
  if (typeof TextDecoder !== "undefined") {
    return new TextDecoder().decode(data);
  } else {
    return strFromU8_(data);
  }
}

// src/utils/isPromise.js
function isPromise(anything) {
  return typeof anything === "object" && typeof anything.then === "function";
}

// src/xlsx/SheetNotFoundError.js
var SheetNotFoundError = class extends Error {
  constructor(sheet, sheets) {
    super(`Sheet not found: ${typeof sheet === "number" ? sheet + ". Sheet count: " + sheets.length : sheet + ". Available sheets: " + sheets.join(", ")}`);
    this.name = "SheetNotFoundError";
    this.sheet = sheet;
    this.sheets = sheets;
  }
};

// src/xlsx/parseSpreadsheetContents.js
var CAN_USE_WORKER = false;
function parseSpreadsheetContents(parseXml2, contents_, options = {}) {
  const contents = convertValuesFromUint8ArraysToStrings(contents_);
  checkpoint("parse spreadsheet info and file paths");
  return readFiles(
    getXmlFilesAtFixedPaths(),
    contents,
    parseXml2
  ).then(({ spreadsheetInfo, filePaths }) => {
    checkpoint('parse "shared strings" and "styles"');
    return readFiles(
      getXmlFilesAtNonFixedPaths(filePaths),
      contents,
      parseXml2
    ).then(({ sharedStrings, styles }) => {
      const sheetRelationIdsToRead = options.sheets ? options.sheets.map((sheet) => getSheetRelationId(sheet, spreadsheetInfo.sheets)) : spreadsheetInfo.sheets.map((_) => _.relationId);
      checkpoint(`parse sheet${sheetRelationIdsToRead.length === 1 ? "" : "s"} data`);
      return readFiles(
        getSheetDataXmlFiles(filePaths, sheetRelationIdsToRead, {
          sharedStrings,
          styles,
          epoch1904: spreadsheetInfo.epoch1904,
          options
        }),
        contents,
        parseXml2
      ).then((sheetsData) => {
        checkpoint("end");
        return sheetRelationIdsToRead.map((sheetRelationId) => ({
          sheet: getSheetNameByRelationId(sheetRelationId, spreadsheetInfo.sheets),
          data: sheetsData[sheetRelationId]
        }));
      });
    });
  });
}
function parseSpreadsheetContentsInWorker(createWorkerFunction, parseXml2, contents, options) {
  if (!(options && options.parseNumber)) {
    options = {
      ...options,
      parseNumber: null
    };
  }
  if (!createWorkerFunction || !CAN_USE_WORKER) {
    return parseSpreadsheetContents(parseXml2, contents, options);
  }
}
function getSheetRelationId(sheet, sheets) {
  if (typeof sheet === "string") {
    for (const _sheet of sheets) {
      if (_sheet.name === sheet) {
        return _sheet.relationId;
      }
    }
  } else {
    if (sheet <= sheets.length) {
      return sheets[sheet - 1].relationId;
    }
  }
  throw new SheetNotFoundError(sheet, sheets.map((_) => _.name));
}
function getSheetNameByRelationId(sheetRelationId, sheets) {
  for (const sheet of sheets) {
    if (sheet.relationId === sheetRelationId) {
      return sheet.name;
    }
  }
  throw new Error(`Sheet relation ID not found: ${sheetRelationId}`);
}
function getXmlFilesAtFixedPaths() {
  return {
    // Read the paths to certain files inside the `.xlsx` file, which is itself just a `.zip` archive.
    // These paths aren't standardized between different spreadsheet editors.
    // https://github.com/tidyverse/readxl/issues/104
    "xl/_rels/workbook.xml.rels": {
      name: "filePaths",
      parse: parseFilePaths
    },
    // General info on the spreadsheet.
    "xl/workbook.xml": {
      name: "spreadsheetInfo",
      parse: parseSpreadsheetInfo
    }
  };
}
function getXmlFilesAtNonFixedPaths(filePaths) {
  return {
    // The usual file path for "shared strings" is "xl/sharedStrings.xml".
    [filePaths.sharedStrings || "xl/sharedStrings.xml"]: {
      name: "sharedStrings",
      // `parseSharedStrings()` returns a `Promise`.
      parse: parseSharedStrings,
      // It seems that "sharedStrings.xml" is not required to exist.
      // For example, that could be the case when a spreadsheet doesn't contain any strings.
      // https://github.com/catamphetamine/read-excel-file/issues/85
      fallback: Promise.resolve([])
    },
    // The usual file path for "styles" is "xl/styles.xml".
    [filePaths.styles || "xl/styles.xml"]: {
      name: "styles",
      parse: parseStyles,
      fallback: {}
    }
  };
}
function getSheetDataXmlFiles(filePaths, sheetRelationIdsToRead, sheetDataParserParameters) {
  return Object.keys(filePaths.sheets).filter((sheetRelationId) => sheetRelationIdsToRead.includes(sheetRelationId)).reduce((filesInfo, sheetRelationId) => ({
    ...filesInfo,
    [filePaths.sheets[sheetRelationId]]: {
      name: sheetRelationId,
      // `parseSheet()` returns a `Promise`.
      parse: (content, parseXml2) => parseSheet(content, parseXml2, sheetDataParserParameters)
    }
  }), {});
}
function readFiles(filesInfo, contents, parseXml2) {
  const results = {};
  for (const filePath of Object.keys(filesInfo)) {
    const fileInfo = filesInfo[filePath];
    results[fileInfo.name] = contents[filePath] === void 0 ? fileInfo.fallback === void 0 ? (() => {
      throw new InvalidSpreadsheetError(`"${filePath}" file not found inside the \`.xlsx\` file`);
    })() : fileInfo.fallback : fileInfo.parse(contents[filePath], parseXml2);
  }
  const promises = [];
  for (const name of Object.keys(results)) {
    if (isPromise(results[name])) {
      promises.push(results[name].then((result) => {
        results[name] = result;
      }));
    }
  }
  if (promises.length > 0) {
    return Promise.all(promises).then(() => results);
  }
  return results;
}

// src/saxen/parser.js
function Parser_(options) {
  var fromCharCode = String.fromCharCode;
  var hasOwnProperty = Object.prototype.hasOwnProperty;
  var ENTITY_PATTERN = /&#(\d+);|&#x([0-9a-f]+);|&(\w+);/ig;
  var ENTITY_MAPPING = {
    "amp": "&",
    "apos": "'",
    "gt": ">",
    "lt": "<",
    "quot": '"'
  };
  Object.keys(ENTITY_MAPPING).forEach(function(k) {
    ENTITY_MAPPING[k.toUpperCase()] = ENTITY_MAPPING[k];
  });
  function replaceEntities(_, d, x, z) {
    if (z) {
      if (hasOwnProperty.call(ENTITY_MAPPING, z)) {
        return ENTITY_MAPPING[z];
      } else {
        return "&" + z + ";";
      }
    }
    if (d) {
      return fromCharCode(d);
    }
    return fromCharCode(parseInt(x, 16));
  }
  function decodeEntities(s) {
    if (s.length > 3 && s.indexOf("&") !== -1) {
      return s.replace(ENTITY_PATTERN, replaceEntities);
    }
    return s;
  }
  var NON_WHITESPACE_OUTSIDE_ROOT_NODE = "non-whitespace outside of root node";
  function error(msg) {
    return new Error(msg);
  }
  function missingNamespaceForPrefix(prefix) {
    return "missing namespace for prefix <" + prefix + ">";
  }
  function getter(getFn) {
    return {
      "get": getFn,
      "enumerable": true
    };
  }
  function cloneNsMatrix(nsMatrix) {
    var clone = {}, key;
    for (key in nsMatrix) {
      clone[key] = nsMatrix[key];
    }
    return clone;
  }
  var NAME_CACHE = /* @__PURE__ */ Symbol("nameCache");
  function uriPrefix(prefix) {
    return prefix + "$uri";
  }
  function buildNsMatrix(nsUriToPrefix) {
    var nsMatrix = {}, uri, prefix;
    for (uri in nsUriToPrefix) {
      prefix = nsUriToPrefix[uri];
      nsMatrix[prefix] = prefix;
      nsMatrix[uriPrefix(prefix)] = uri;
    }
    return nsMatrix;
  }
  function noopGetContext() {
    return { line: 0, column: 0 };
  }
  function throwFunc(err) {
    throw err;
  }
  function Parser(options2) {
    if (!this) {
      return new Parser(options2);
    }
    var proxy = options2 && options2["proxy"];
    var onText, onOpenTag, onCloseTag, onCDATA, onError = throwFunc, onWarning, onComment, onQuestion, onAttention;
    var getContext = noopGetContext;
    var streaming = false;
    var rootTagFound = false;
    var leftoverXml = "";
    var maybeNS = false;
    var isNamespace = false;
    var returnError = null;
    var parseStop = false;
    var nsMatrixStack, nsMatrix, nodeStack;
    var nsUriToPrefix;
    function handleError(err) {
      if (!(err instanceof Error)) {
        err = error(err);
      }
      returnError = err;
      onError(err, getContext);
    }
    function handleWarning(err) {
      if (!onWarning) {
        return;
      }
      if (!(err instanceof Error)) {
        err = error(err);
      }
      onWarning(err, getContext);
    }
    this["on"] = function(name, cb) {
      if (typeof cb !== "function") {
        throw error("required args <name, cb>");
      }
      switch (name) {
        case "openTag":
          onOpenTag = cb;
          break;
        case "text":
          onText = cb;
          break;
        case "closeTag":
          onCloseTag = cb;
          break;
        case "error":
          onError = cb;
          break;
        case "warn":
          onWarning = cb;
          break;
        case "cdata":
          onCDATA = cb;
          break;
        case "attention":
          onAttention = cb;
          break;
        // <!XXXXX zzzz="eeee">
        case "question":
          onQuestion = cb;
          break;
        // <? ....  ?>
        case "comment":
          onComment = cb;
          break;
        default:
          throw error("unsupported event: " + name);
      }
      return this;
    };
    this["ns"] = function(nsMap) {
      if (typeof nsMap === "undefined") {
        nsMap = {};
      }
      if (typeof nsMap !== "object") {
        throw error("required args <nsMap={}>");
      }
      var _nsUriToPrefix = {}, k;
      for (k in nsMap) {
        _nsUriToPrefix[k] = nsMap[k];
      }
      isNamespace = true;
      nsUriToPrefix = _nsUriToPrefix;
      return this;
    };
    function resetState() {
      nsMatrixStack = isNamespace ? [] : null;
      nsMatrix = isNamespace ? buildNsMatrix(nsUriToPrefix) : null;
      nodeStack = [];
      getContext = noopGetContext;
      parseStop = false;
      returnError = null;
      rootTagFound = false;
      leftoverXml = "";
    }
    this["parse"] = function(xml) {
      if (typeof xml !== "string") {
        throw error("required args <xml=string>");
      }
      if (streaming) {
        throw error("parse during stream; call end() first");
      }
      resetState();
      parse2(xml);
      getContext = noopGetContext;
      parseStop = false;
      return returnError;
    };
    this["write"] = function(xml) {
      if (typeof xml !== "string") {
        throw error("required args <xml=string>");
      }
      if (!streaming) {
        resetState();
        streaming = true;
      }
      if (!returnError) {
        leftoverXml = parse2(leftoverXml + xml, true) || "";
      }
      return this;
    };
    this["end"] = function() {
      if (!streaming) {
        resetState();
      }
      streaming = false;
      if (!returnError) {
        parse2(leftoverXml);
      }
      leftoverXml = "";
      getContext = noopGetContext;
      parseStop = false;
      return returnError;
    };
    this["stop"] = function() {
      parseStop = true;
    };
    function parse2(xml, streaming2 = false) {
      var elNameCache = null, elNameCacheMatrix = null;
      var _nsMatrix, anonymousNsCount = 0, tagStart = false, tagEnd = false, i = 0, j = 0, x, y, q, w, v, xmlns, elementName, _elementName, elementProxy;
      var attrsString = "", attrsStart = 0, cachedAttrs;
      function normalizeAttrName(name, defaultAlias) {
        var w2 = name.indexOf(":");
        if (w2 === -1) {
          return name;
        }
        var nsName = nsMatrix[name.substring(0, w2)];
        if (!nsName) {
          handleWarning(missingNamespaceForPrefix(name.substring(0, w2)));
          return null;
        }
        return defaultAlias === nsName ? name.substr(w2 + 1) : nsName + name.substr(w2);
      }
      function getAttrs() {
        if (cachedAttrs !== null) {
          return cachedAttrs;
        }
        var nsUri, nsUriPrefix, defaultAlias = isNamespace && nsMatrix["xmlns"], attrList = isNamespace && maybeNS ? [] : null, i2 = attrsStart, s = attrsString, l = s.length, hasNewMatrix, newalias, value, alias, name, attrs = {}, seenAttrs = /* @__PURE__ */ new Set(), skipAttr, w2, j2;
        parseAttr:
          for (; i2 < l; i2++) {
            skipAttr = false;
            w2 = s.charCodeAt(i2);
            if (w2 === 32 || w2 < 14 && w2 > 8) {
              continue;
            }
            if (w2 < 65 || w2 > 122 || w2 > 90 && w2 < 97) {
              if (w2 !== 95 && w2 !== 58) {
                handleWarning("illegal first char attribute name");
                skipAttr = true;
              }
            }
            for (j2 = i2 + 1; j2 < l; j2++) {
              w2 = s.charCodeAt(j2);
              if (w2 > 96 && w2 < 123 || w2 > 64 && w2 < 91 || w2 > 47 && w2 < 59 || w2 === 46 || // '.'
              w2 === 45 || // '-'
              w2 === 95) {
                continue;
              }
              if (w2 === 32 || w2 < 14 && w2 > 8) {
                handleWarning("missing attribute value");
                i2 = j2;
                continue parseAttr;
              }
              if (w2 === 61) {
                break;
              }
              handleWarning("illegal attribute name char");
              skipAttr = true;
            }
            name = s.substring(i2, j2);
            if (name === "xmlns:xmlns") {
              handleWarning("illegal declaration of xmlns");
              skipAttr = true;
            }
            w2 = s.charCodeAt(j2 + 1);
            if (w2 === 34) {
              j2 = s.indexOf('"', i2 = j2 + 2);
              if (j2 === -1) {
                j2 = s.indexOf("'", i2);
                if (j2 !== -1) {
                  handleWarning("attribute value quote missmatch");
                  skipAttr = true;
                }
              }
            } else if (w2 === 39) {
              j2 = s.indexOf("'", i2 = j2 + 2);
              if (j2 === -1) {
                j2 = s.indexOf('"', i2);
                if (j2 !== -1) {
                  handleWarning("attribute value quote missmatch");
                  skipAttr = true;
                }
              }
            } else {
              handleWarning("missing attribute value quotes");
              skipAttr = true;
              for (j2 = j2 + 1; j2 < l; j2++) {
                w2 = s.charCodeAt(j2 + 1);
                if (w2 === 32 || w2 < 14 && w2 > 8) {
                  break;
                }
              }
            }
            if (j2 === -1) {
              handleWarning("missing closing quotes");
              j2 = l;
              skipAttr = true;
            }
            if (!skipAttr) {
              value = s.substring(i2, j2);
            }
            i2 = j2;
            for (; j2 + 1 < l; j2++) {
              w2 = s.charCodeAt(j2 + 1);
              if (w2 === 32 || w2 < 14 && w2 > 8) {
                break;
              }
              if (i2 === j2) {
                handleWarning("illegal character after attribute end");
                skipAttr = true;
              }
            }
            i2 = j2 + 1;
            if (skipAttr) {
              continue parseAttr;
            }
            if (seenAttrs.has(name)) {
              handleWarning("attribute <" + name + "> already defined");
              continue;
            }
            seenAttrs.add(name);
            if (!isNamespace) {
              attrs[name] = value;
              continue;
            }
            if (maybeNS) {
              newalias = name === "xmlns" ? "xmlns" : name.charCodeAt(0) === 120 && name.substr(0, 6) === "xmlns:" ? name.substr(6) : null;
              if (newalias !== null) {
                nsUri = decodeEntities(value);
                nsUriPrefix = uriPrefix(newalias);
                alias = nsUriToPrefix[nsUri];
                if (!alias) {
                  if (newalias === "xmlns" || nsUriPrefix in nsMatrix && nsMatrix[nsUriPrefix] !== nsUri) {
                    do {
                      alias = "ns" + anonymousNsCount++;
                    } while (typeof nsMatrix[alias] !== "undefined");
                  } else {
                    alias = newalias;
                  }
                  nsUriToPrefix[nsUri] = alias;
                }
                if (nsMatrix[newalias] !== alias) {
                  if (!hasNewMatrix) {
                    nsMatrix = cloneNsMatrix(nsMatrix);
                    hasNewMatrix = true;
                  }
                  nsMatrix[newalias] = alias;
                  if (newalias === "xmlns") {
                    nsMatrix[uriPrefix(alias)] = nsUri;
                    defaultAlias = alias;
                  }
                  nsMatrix[nsUriPrefix] = nsUri;
                }
                attrs[name] = value;
                continue;
              }
              attrList.push(name, value);
              continue;
            }
            name = normalizeAttrName(name, defaultAlias);
            if (name === null) {
              continue;
            }
            attrs[name] = value;
          }
        if (maybeNS) {
          for (i2 = 0, l = attrList.length; i2 < l; i2++) {
            name = normalizeAttrName(attrList[i2++], defaultAlias);
            value = attrList[i2];
            if (name === null) {
              continue;
            }
            attrs[name] = value;
          }
        }
        return cachedAttrs = attrs;
      }
      function getParseContext() {
        var splitsRe = /(\r\n|\r|\n)/g;
        var line = 0;
        var column = 0;
        var startOfLine = 0;
        var endOfLine = j;
        var match;
        var data;
        while (i >= startOfLine) {
          match = splitsRe.exec(xml);
          if (!match) {
            break;
          }
          endOfLine = match[0].length + match.index;
          if (endOfLine > i) {
            break;
          }
          line += 1;
          startOfLine = endOfLine;
        }
        if (i == -1) {
          column = endOfLine;
          data = xml.substring(j);
        } else if (j === 0) {
          data = xml.substring(j, i);
        } else {
          column = i - startOfLine;
          data = j == -1 ? xml.substring(i) : xml.substring(i, j + 1);
        }
        return {
          "data": data,
          "line": line,
          "column": column
        };
      }
      getContext = getParseContext;
      if (proxy) {
        elementProxy = Object.create({}, {
          "name": getter(function() {
            return elementName;
          }),
          "originalName": getter(function() {
            return _elementName;
          }),
          "attrs": getter(getAttrs),
          "ns": getter(function() {
            return nsMatrix;
          })
        });
      }
      while (j !== -1) {
        if (xml.charCodeAt(j) === 60) {
          i = j;
        } else {
          i = xml.indexOf("<", j);
        }
        if (i === -1) {
          if (streaming2) {
            return xml.substring(j);
          }
          if (nodeStack.length) {
            return handleError("unexpected end of file");
          }
          if (!rootTagFound) {
            return handleError("missing start tag");
          }
          if (j < xml.length) {
            if (xml.substring(j).trim()) {
              handleWarning(NON_WHITESPACE_OUTSIDE_ROOT_NODE);
            }
          }
          return;
        }
        if (!rootTagFound) {
          rootTagFound = true;
        }
        if (j !== i) {
          if (nodeStack.length) {
            if (onText) {
              onText(xml.substring(j, i), decodeEntities, getContext);
              if (parseStop) {
                return;
              }
            }
          } else {
            if (xml.substring(j, i).trim()) {
              handleWarning(NON_WHITESPACE_OUTSIDE_ROOT_NODE);
              if (parseStop) {
                return;
              }
            }
          }
        }
        w = xml.charCodeAt(i + 1);
        if (w === 33) {
          q = xml.charCodeAt(i + 2);
          if (q === 91 && xml.substr(i + 3, 6) === "CDATA[") {
            j = xml.indexOf("]]>", i);
            if (j === -1) {
              if (streaming2) {
                return xml.substring(i);
              }
              return handleError("unclosed cdata");
            }
            if (onCDATA) {
              onCDATA(xml.substring(i + 9, j), getContext);
              if (parseStop) {
                return;
              }
            }
            j += 3;
            continue;
          }
          if (q === 45 && xml.charCodeAt(i + 3) === 45) {
            j = xml.indexOf("-->", i);
            if (j === -1) {
              if (streaming2) {
                return xml.substring(i);
              }
              return handleError("unclosed comment");
            }
            if (onComment) {
              onComment(xml.substring(i + 4, j), decodeEntities, getContext);
              if (parseStop) {
                return;
              }
            }
            j += 3;
            continue;
          }
        }
        if (w === 63) {
          j = xml.indexOf("?>", i);
          if (j === -1) {
            if (streaming2) {
              return xml.substring(i);
            }
            return handleError("unclosed question");
          }
          if (onQuestion) {
            onQuestion(xml.substring(i, j + 2), getContext);
            if (parseStop) {
              return;
            }
          }
          j += 2;
          continue;
        }
        for (x = i + 1; ; x++) {
          v = xml.charCodeAt(x);
          if (isNaN(v)) {
            if (streaming2) {
              return xml.substring(i);
            }
            j = -1;
            return handleError("unclosed tag");
          }
          if (v === 34) {
            q = xml.indexOf('"', x + 1);
            x = q !== -1 ? q : x;
          } else if (v === 39) {
            q = xml.indexOf("'", x + 1);
            x = q !== -1 ? q : x;
          } else if (v === 62) {
            j = x;
            break;
          }
        }
        if (w === 33) {
          if (onAttention) {
            onAttention(xml.substring(i, j + 1), decodeEntities, getContext);
            if (parseStop) {
              return;
            }
          }
          j += 1;
          continue;
        }
        cachedAttrs = {};
        if (w === 47) {
          tagStart = false;
          tagEnd = true;
          if (!nodeStack.length) {
            return handleError("missing open tag");
          }
          x = elementName = nodeStack.pop();
          q = i + 2 + x.length;
          if (xml.substring(i + 2, q) !== x) {
            return handleError("closing tag mismatch");
          }
          for (; q < j; q++) {
            w = xml.charCodeAt(q);
            if (w === 32 || w > 8 && w < 14) {
              continue;
            }
            return handleError("close tag");
          }
        } else {
          if (xml.charCodeAt(j - 1) === 47) {
            x = elementName = xml.substring(i + 1, j - 1);
            tagStart = true;
            tagEnd = true;
          } else {
            x = elementName = xml.substring(i + 1, j);
            tagStart = true;
            tagEnd = false;
          }
          if (!(w > 96 && w < 123 || w > 64 && w < 91 || w === 95 || w === 58)) {
            return handleError("illegal first char nodeName");
          }
          for (q = 1, y = x.length; q < y; q++) {
            w = x.charCodeAt(q);
            if (w > 96 && w < 123 || w > 64 && w < 91 || w > 47 && w < 59 || w === 45 || w === 95 || w == 46) {
              continue;
            }
            if (w === 32 || w < 14 && w > 8) {
              elementName = x.substring(0, q);
              cachedAttrs = null;
              break;
            }
            return handleError("invalid nodeName");
          }
          if (!tagEnd) {
            nodeStack.push(elementName);
          }
        }
        if (isNamespace) {
          _nsMatrix = nsMatrix;
          if (tagStart) {
            if (!tagEnd) {
              nsMatrixStack.push(_nsMatrix);
            }
            if (cachedAttrs === null) {
              if (maybeNS = x.indexOf("xmlns", q) !== -1) {
                attrsStart = q;
                attrsString = x;
                getAttrs();
                maybeNS = false;
              }
            }
          }
          _elementName = elementName;
          if (elNameCacheMatrix !== nsMatrix) {
            elNameCache = nsMatrix[NAME_CACHE];
            if (elNameCache === void 0) {
              elNameCache = nsMatrix[NAME_CACHE] = {};
            }
            elNameCacheMatrix = nsMatrix;
          }
          var _cachedName = elNameCache[elementName];
          if (_cachedName !== void 0) {
            elementName = _cachedName;
          } else {
            w = elementName.indexOf(":");
            if (w !== -1) {
              xmlns = nsMatrix[elementName.substring(0, w)];
              if (!xmlns) {
                return handleError("missing namespace on <" + _elementName + ">");
              }
              elementName = elementName.substr(w + 1);
            } else {
              xmlns = nsMatrix["xmlns"];
            }
            if (xmlns) {
              elementName = xmlns + ":" + elementName;
            }
            elNameCache[_elementName] = elementName;
          }
        }
        if (tagStart) {
          attrsStart = q;
          attrsString = x;
          if (onOpenTag) {
            if (proxy) {
              onOpenTag(elementProxy, decodeEntities, tagEnd, getContext);
            } else {
              onOpenTag(elementName, getAttrs, decodeEntities, tagEnd, getContext);
            }
            if (parseStop) {
              return;
            }
          }
        }
        if (tagEnd) {
          if (onCloseTag) {
            onCloseTag(proxy ? elementProxy : elementName, decodeEntities, tagStart, getContext);
            if (parseStop) {
              return;
            }
          }
          if (isNamespace) {
            if (!tagStart) {
              nsMatrix = nsMatrixStack.pop();
            } else {
              nsMatrix = _nsMatrix;
            }
          }
        }
        j += 1;
      }
    }
  }
  return new Parser(options);
}

// src/xml/parseXmlStream.saxen.js
function parseXmlStream(state, onOpenTag, onCloseTag, onText) {
  let errored = false;
  const mustNotHaveErrored = () => {
    if (errored) {
      throw new Error("Errored");
    }
  };
  let resolvePromise;
  const parser2 = new Parser_();
  const write = (xml) => {
    mustNotHaveErrored();
    parser2.write(xml);
  };
  const end = () => {
    mustNotHaveErrored();
    parser2.end();
    resolvePromise();
  };
  const promise = new Promise((resolve, reject) => {
    resolvePromise = resolve;
    const onerror = (error) => {
      errored = true;
      throw error;
    };
    const ontext = (text, decodeEntities) => {
      if (onText) {
        onText(decodeEntities(text), state);
      }
    };
    const onopentag = (elementName, getAttributes, decodeEntities, selfClosing, getContext) => {
      if (onOpenTag) {
        const attributes = getAttributes();
        for (const name in attributes) {
          attributes[trimXmlnsPrefix(name, true)] = decodeEntities(attributes[name]);
        }
        onOpenTag(
          trimXmlnsPrefix(elementName),
          attributes,
          state
        );
      }
    };
    const onclosetag = (elementName) => {
      if (onCloseTag) {
        onCloseTag(trimXmlnsPrefix(elementName), state);
      }
    };
    parser2.on("error", onerror);
    parser2.on("text", ontext);
    parser2.on("openTag", onopentag);
    parser2.on("closeTag", onclosetag);
  });
  return { promise, write, end };
}
function trimXmlnsPrefix(string, isAttributeName) {
  let i = 0;
  while (i < string.length) {
    if (string[i] === ":") {
      if (isAttributeName && i === 5 && string.slice(0, 5) === "xmlns") {
      } else {
        return string.slice(i + 1);
      }
    }
    i++;
  }
  return string;
}

// src/xml/parseXml.js
function parseXml(xml, state, onOpenTag, onCloseTag, onText, onProgress) {
  const parser2 = parseXmlStream(state, onOpenTag, onCloseTag, onText);
  if (onProgress) {
    parseXmlInChunks(parser2, xml, onProgress);
  } else {
    parser2.write(xml);
    parser2.end();
  }
  return parser2.promise.then(
    (result) => result,
    (error) => {
      const spreadsheetError = new InvalidSpreadsheetError(error.message);
      spreadsheetError.stack = error.stack;
      spreadsheetError.cause = error;
      throw spreadsheetError;
    }
  );
  function parseXmlInChunks(parser3, xml2, onProgress2, nonBlocking) {
    const MAX_CHUNK_PROCESSING_TIME = 7;
    const INITIAL_CHUNK_SIZE = 64 * 1024;
    let chunksCount = 0;
    let chunkSize = INITIAL_CHUNK_SIZE;
    const parseNextChunk = () => {
      chunksCount++;
      const startedAt = Date.now();
      if (xml2.length > chunkSize) {
        parser3.write(xml2.slice(0, chunkSize));
        if (onProgress2) {
          onProgress2(false);
        }
        xml2 = xml2.slice(chunkSize);
        const chunkProcessingTime = Date.now() - startedAt;
        if (chunkProcessingTime < MAX_CHUNK_PROCESSING_TIME * 0.5) {
          chunkSize *= 2;
        } else if (chunkProcessingTime > MAX_CHUNK_PROCESSING_TIME) {
          chunkSize /= 2;
        }
        return true;
      } else {
        parser3.write(xml2);
        parser3.end();
        if (onProgress2) {
          onProgress2(true);
        }
        return false;
      }
    };
    const loop = () => {
      if (parseNextChunk()) {
        if (nonBlocking) {
          if (typeof setImmediate !== "undefined") {
            setImmediate(loop);
          } else {
            setTimeout(loop, 0);
          }
        } else {
          loop();
        }
      } else {
      }
    };
    loop();
  }
}

// src/xlsx/file/InvalidInputError.js
var MESSAGES = {
  XLS_FILE_NOT_SUPPORTED: "You passed a legacy `.xls` file. Only `.xlsx` files are supported",
  FILE_NOT_SUPPORTED: "Doesn't look like an `.xlsx` file",
  INVALID_ZIP: "Couldn't unzip `.xlsx` file contents",
  NO_DATA: "No data"
};
var InvalidInputError = class extends Error {
  /**
   * Creates an `InvalidInputError` instance.
   * @param {string} code
   * @param {any} [cause]
   */
  constructor(code, cause) {
    super(MESSAGES[code] || code);
    this.code = code;
    this.name = "InvalidInputError";
    this.cause = cause;
  }
};

// src/xlsx/file/createFileTypeDetector.js
var ZIP_FILE_SIGNATURE = [80, 75];
var XLS_FILE_SIGNATURE = [208, 207, 17, 224];
var FILE_TYPE_SIGNATURES = [
  ZIP_FILE_SIGNATURE,
  XLS_FILE_SIGNATURE
];
var XLSX_FILE_TYPE = FILE_TYPE_SIGNATURES.indexOf(ZIP_FILE_SIGNATURE);
var XLS_FILE_TYPE = FILE_TYPE_SIGNATURES.indexOf(XLS_FILE_SIGNATURE);
function createFileTypeDetector() {
  let type;
  let possibleTypes = indexesOf(FILE_TYPE_SIGNATURES);
  let i = 0;
  return (byte) => {
    if (isNaN(type)) {
      let t;
      possibleTypes = possibleTypes.filter((typeIndex) => {
        if (byte === FILE_TYPE_SIGNATURES[typeIndex][i]) {
          if (FILE_TYPE_SIGNATURES[typeIndex].length === i + 1) {
            t = typeIndex;
          }
          return true;
        }
      });
      if (possibleTypes.length === 1) {
        type = t;
      } else if (possibleTypes.length === 0) {
        type = -1;
      }
    }
    i++;
    return type;
  };
}
function indexesOf(array) {
  const indexes = [];
  let i = 0;
  while (i < array.length) {
    indexes.push(i);
    i++;
  }
  return indexes;
}

// src/xlsx/file/validateLeadingBytes.js
function validateLeadingBytes(bytes) {
  const fileTypeDetector = createFileTypeDetector();
  for (const byte of bytes) {
    if (validateByte(byte, fileTypeDetector)) {
      return;
    }
  }
  noFileTypeCouldBeDetermined(bytes.length);
}
function validateByte(byte, fileTypeDetector) {
  const fileType = fileTypeDetector(byte);
  if (fileType !== void 0) {
    if (fileType === XLS_FILE_TYPE) {
      throw new InvalidInputError("XLS_FILE_NOT_SUPPORTED");
    }
    if (fileType < 0) {
      throw new InvalidInputError("FILE_NOT_SUPPORTED");
    }
    return true;
  }
}
function noFileTypeCouldBeDetermined(byteCount) {
  throw new InvalidInputError(byteCount === 0 ? "NO_DATA" : "FILE_NOT_SUPPORTED");
}

// src/export/filterZipArchiveEntry.js
function filterZipArchiveEntry({ path }) {
  return path.endsWith(".xml") || path.endsWith(".xml.rels");
}

// src/sync/parseXLSXSync.js
async function parseXLSXSync(bytes, options = {}) {
  validateLeadingBytes(bytes);
  let contents;
  try {
    contents = unzipBytesSync(bytes, {
      filter: (file) => filterZipArchiveEntry({ path: file.name }),
      maxBytes: options.maxUncompressedSize || options.maxBytes
    });
  } catch (error) {
    if (error instanceof UnzipError) {
      throw new InvalidInputError("INVALID_ZIP", error.cause);
    }
    throw error;
  }
  const sheets = await parseSpreadsheetContentsInWorker(void 0, parseXml, contents, options);
  const worksheets = (sheets || []).map((s) => {
    const data = s.data || [];
    let maxCols = 0;
    for (const row of data) {
      if (row && row.length > maxCols) {
        maxCols = row.length;
      }
    }
    const columns = [];
    for (let c = 0; c < maxCols; c++) {
      columns.push({ title: getColumnName(c), width: 100 });
    }
    return { worksheetName: s.sheet || s.name || "Sheet1", data, columns };
  });
  const result = { worksheets };
  if (options.onload) {
    options.onload(result);
  }
  return result;
}

// src/sync/writeXlsxFileSync.js
import { zipSync, strToU8 } from "fflate";

// src/write/xlsx/helpers/features/getAdditionalContent.js
function _createForOfIteratorHelperLoose(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0;
  }
}
function _arrayLikeToArray(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function getAdditionalContent(fileName, features, sheetOptions, properties) {
  var content = "";
  for (var _iterator = _createForOfIteratorHelperLoose(features), _step; !(_step = _iterator()).done; ) {
    var feature = _step.value;
    var transform5 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform5 && transform5.insert) {
      var insertedContent = transform5.insert(sheetOptions, properties);
      if (insertedContent) {
        content += insertedContent;
      }
    }
  }
  return content;
}

// src/write/xlsx/helpers/features/transformContent.js
function _createForOfIteratorHelperLoose2(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray2(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray2(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray2(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray2(r, a) : void 0;
  }
}
function _arrayLikeToArray2(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function transformContent(content, fileName, features, sheetOptions, properties) {
  for (var _iterator = _createForOfIteratorHelperLoose2(features), _step; !(_step = _iterator()).done; ) {
    var feature = _step.value;
    var transform5 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform5 && transform5.transform) {
      content = transform5.transform(content, sheetOptions, properties);
    }
  }
  return content;
}

// src/write/xml/escapeXmlSpecialCharacters.js
function escapeXmlSpecialCharacters(string, _ref) {
  var isAttributeValue = _ref.isAttributeValue;
  string = replaceAll(string, "&", "&amp;");
  string = replaceAll(string, ">", "&gt;");
  string = replaceAll(string, "<", "&lt;");
  if (isAttributeValue) {
    string = replaceAll(string, "'", "&apos;");
    string = replaceAll(string, '"', "&quot;");
  }
  return string;
}
function replaceAll(string, replacedSubstring, replacementSubstring) {
  if (string.replaceAll) {
    return string.replaceAll(replacedSubstring, replacementSubstring);
  }
  return string.replace(new RegExp(replacedSubstring, "g"), replacementSubstring);
}

// src/write/xml/removeInvalidXmlCharacters.js
var INVALID_CHARACTERS = /((?:[\0-\x08\x0B\f\x0E-\x1F\uFFFD\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]))/g;
var DISCOURAGED_CHARACTERS = new RegExp("([\\x7F-\\x84]|[\\x86-\\x9F]|[\\uFDD0-\\uFDEF]|(?:\\uD83F[\\uDFFE\\uDFFF])|(?:\\uD87F[\\uDFFE\\uDFFF])|(?:\\uD8BF[\\uDFFE\\uDFFF])|(?:\\uD8FF[\\uDFFE\\uDFFF])|(?:\\uD93F[\\uDFFE\\uDFFF])|(?:\\uD97F[\\uDFFE\\uDFFF])|(?:\\uD9BF[\\uDFFE\\uDFFF])|(?:\\uD9FF[\\uDFFE\\uDFFF])|(?:\\uDA3F[\\uDFFE\\uDFFF])|(?:\\uDA7F[\\uDFFE\\uDFFF])|(?:\\uDABF[\\uDFFE\\uDFFF])|(?:\\uDAFF[\\uDFFE\\uDFFF])|(?:\\uDB3F[\\uDFFE\\uDFFF])|(?:\\uDB7F[\\uDFFE\\uDFFF])|(?:\\uDBBF[\\uDFFE\\uDFFF])|(?:\\uDBFF[\\uDFFE\\uDFFF])(?:[\\0-\\t\\x0B\\f\\x0E-\\u2027\\u202A-\\uD7FF\\uE000-\\uFFFF]|[\\uD800-\\uDBFF][\\uDC00-\\uDFFF]|[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|(?:[^\\uD800-\\uDBFF]|^)[\\uDC00-\\uDFFF]))", "g");
function removeInvalidXmlCharacters(string) {
  string = string.replace(INVALID_CHARACTERS, "");
  string = string.replace(DISCOURAGED_CHARACTERS, "");
  return string;
}

// src/write/xml/sanitizeAttributeValue.js
function sanitizeAttributeValue(attributeValue) {
  if (typeof attributeValue !== "string") {
    throw new TypeError("Argument must be a string");
  }
  attributeValue = removeInvalidXmlCharacters(attributeValue);
  return escapeXmlSpecialCharacters(attributeValue, {
    isAttributeValue: true
  });
}

// src/write/xml/sanitizeAttributeName.js
function sanitizeAttributeName(attributeName) {
  if (typeof attributeName !== "string") {
    throw new TypeError("Argument must be a string");
  }
  attributeName = removeInvalidXmlCharacters(attributeName);
  return attributeName.replace(/[^a-zA-Z_0-9-.:]/g, "").replace(/^[^a-zA-Z_]+/, "");
}

// src/write/xml/getAttributesString.js
function getAttributesString(attributes) {
  return Object.keys(attributes).map(function(name) {
    return "".concat(sanitizeAttributeName(name), '="').concat(sanitizeAttributeValue(String(attributes[name])), '"');
  }).reduce(function(combined, part) {
    return combined + " " + part;
  }, "");
}

// src/write/xml/getOpeningTagMarkup.js
function getOpeningTagMarkup(tagName, attributes) {
  return "<" + tagName + (attributes ? getAttributesString(attributes) : "") + ">";
}

// src/write/xml/getClosingTagMarkup.js
function getClosingTagMarkup(tagName) {
  return "</" + tagName + ">";
}

// src/write/xml/getSelfClosingTagMarkup.js
function getSelfClosingTagMarkup(tagName, attributes) {
  return getOpeningTagMarkup(tagName, attributes).slice(0, -1) + "/>";
}

// src/write/xlsx/helpers/features/getElementXml.js
function _createForOfIteratorHelperLoose3(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray3(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray3(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray3(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray3(r, a) : void 0;
  }
}
function _arrayLikeToArray3(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function getElementXml(fileName, tagName, attributes, innerXml, index, properties, sheetOptionsOrSheetsOptions, features) {
  for (var _iterator = _createForOfIteratorHelperLoose3(features), _step; !(_step = _iterator()).done; ) {
    var feature = _step.value;
    var transform5 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform5 && transform5.transformElementAttributes) {
      attributes = transform5.transformElementAttributes(tagName, attributes || NO_ATTRIBUTES, index, sheetOptionsOrSheetsOptions, properties);
    }
  }
  if (innerXml) {
    return getOpeningTagMarkup(tagName, attributes) + innerXml + getClosingTagMarkup(tagName);
  } else {
    return getSelfClosingTagMarkup(tagName, attributes);
  }
}
var NO_ATTRIBUTES = {};

// src/write/xlsx/files/workbook.xml.js
function generateWorkbookXml(_ref) {
  var sheetIdsAndNames = _ref.sheetIdsAndNames, features = _ref.features, sheetsOptions = _ref.sheetsOptions;
  var tag = function tag2(tagName, attributes, innerXml, index) {
    return getElementXml("xl/workbook.xml", tagName, attributes, innerXml, index, EMPTY_OBJECT, sheetsOptions, features);
  };
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' + // The order of elements is dictated by the XLSX spec.
  //
  // <workbookPr/>
  tag("workbookPr") + // <bookViews/>
  tag("bookViews", null, tag("workbookView", null, null, 0)) + // <sheets/>
  tag("sheets", null, sheetIdsAndNames.map(function(_ref2, i) {
    var sheetId = _ref2.sheetId, sheetName = _ref2.sheetName;
    return tag("sheet", {
      "r:id": "rId".concat(sheetId),
      sheetId,
      name: sheetName
    }, null, i);
  }).join("")) + // <definedNames/>
  tag("definedNames") + // <calcPr/>
  tag("calcPr") + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/workbook.xml", features, sheetsOptions) + "</workbook>";
  xml = transformContent(xml, "xl/workbook.xml", features, sheetsOptions);
  return xml;
}
var EMPTY_OBJECT = {};

// src/write/xlsx/files/workbook.xml.rels.js
function generateWorkbookXmlRels(_ref) {
  var sheetIds = _ref.sheetIds, features = _ref.features, sheetsOptions = _ref.sheetsOptions;
  var xml = '<?xml version="1.0" ?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheetIds.map(function(id) {
    return '<Relationship Id="rId'.concat(id, '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet').concat(id, '.xml"/>');
  }).join("") + '<Relationship Id="rId'.concat(sheetIds.length + 1, '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>') + '<Relationship Id="rId'.concat(sheetIds.length + 2, '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>') + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/_rels/workbook.xml.rels", features, sheetsOptions) + "</Relationships>";
  xml = transformContent(xml, "xl/_rels/workbook.xml.rels", features, sheetsOptions);
  return xml;
}

// src/write/xlsx/files/rels.js
function generateRelsXml(_ref) {
  var features = _ref.features, sheetsOptions = _ref.sheetsOptions;
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId-workbook-1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("_rels/.rels", features, sheetsOptions) + "</Relationships>";
  xml = transformContent(xml, "_rels/.rels", features, sheetsOptions);
  return xml;
}

// src/write/xlsx/files/[Content_Types].xml.js
function generateContentTypesXml(_ref) {
  var sheetIds = _ref.sheetIds, features = _ref.features, sheetsOptions = _ref.sheetsOptions;
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default ContentType="application/xml" Extension="xml"/><Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/><Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" PartName="/xl/sharedStrings.xml"/><Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>' + sheetIds.map(function(sheetId) {
    return '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet'.concat(sheetId, '.xml"/>');
  }).join("") + sheetIds.map(function(sheetId) {
    return getDrawingContentTypeXml(sheetId);
  }).join("") + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("[Content_Types].xml", features, sheetsOptions) + "</Types>";
  xml = transformContent(xml, "[Content_Types].xml", features, sheetsOptions);
  return xml;
}
function getDrawingContentTypeXml(sheetId) {
  return '<Override ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" PartName="/xl/drawings/drawing'.concat(sheetId, '.xml"/>');
}

// src/write/xlsx/files/drawing.xml.js
function generateDrawingXml(_ref) {
  var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId, sheetOptions = _ref.sheetOptions, features = _ref.features;
  var xml = DRAWING_XML_START + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/drawings/drawing{id}.xml", features, sheetOptions, {
    sheetIndex,
    sheetId
  }) + DRAWING_XML_END;
  xml = transformContent(xml, "xl/drawings/drawing{id}.xml", features, sheetOptions, {
    sheetIndex,
    sheetId
  });
  return xml;
}
var DRAWING_XML_START = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
var DRAWING_XML_END = "</xdr:wsDr>";

// src/write/xlsx/files/drawing.xml.rels.js
function generateDrawingXmlRels(_ref) {
  var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId, sheetOptions = _ref.sheetOptions, features = _ref.features;
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/drawings/_rels/drawing{id}.xml.rels", features, sheetOptions, {
    sheetIndex,
    sheetId
  }) + "</Relationships>";
  xml = transformContent(xml, "xl/drawings/_rels/drawing{id}.xml.rels", features, sheetOptions, {
    sheetIndex,
    sheetId
  });
  return xml;
}

// src/write/xlsx/files/sheet.xml/sheetViews.js
function generateSheetViews(tag, viewProperties, sheetIndex) {
  if (!hasView(viewProperties)) {
    return "";
  }
  var showGridLines = viewProperties.showGridLines, rightToLeft = viewProperties.rightToLeft, zoomScale = viewProperties.zoomScale;
  var sheetViewAttributes = {
    tabSelected: sheetIndex === 0 ? 1 : 0,
    // The first sheet is selected by default.
    workbookViewId: 0
  };
  if (showGridLines === false) {
    sheetViewAttributes.showGridLines = false;
  }
  if (rightToLeft) {
    sheetViewAttributes.rightToLeft = 1;
  }
  if (typeof zoomScale === "number") {
    sheetViewAttributes.zoomScale = Math.round(zoomScale * 100);
  }
  var sheetViewXml = tag("sheetView", sheetViewAttributes, null, 0);
  return tag("sheetViews", null, sheetViewXml);
}
function hasView(_ref) {
  var showGridLines = _ref.showGridLines, rightToLeft = _ref.rightToLeft, zoomScale = _ref.zoomScale;
  return showGridLines === false || rightToLeft || typeof zoomScale === "number";
}

// src/write/xlsx/files/sheet.xml/column.js
function generateColumnDescription(tag, column, index) {
  if (!column) {
    return "";
  }
  var width = column.width;
  if (!width) {
    return "";
  }
  var columnNumber = index + 1;
  var attributes = {
    min: columnNumber,
    max: columnNumber,
    width,
    customWidth: 1
  };
  return tag("col", attributes, null, index);
}

// src/write/xlsx/files/sheet.xml/columns.js
function generateColumnsDescription(tag, columns) {
  if (columns) {
    var columnsXml = columns.map(function(column, index) {
      return generateColumnDescription(tag, column, index);
    }).join("");
    if (columnsXml) {
      return tag("cols", null, columnsXml);
    }
  }
  return "";
}

// src/write/xml/sanitizeTextContent.js
function sanitizeTextContent(textContent) {
  if (typeof textContent !== "string") {
    throw new TypeError("Argument must be a string");
  }
  textContent = removeInvalidXmlCharacters(textContent);
  return escapeXmlSpecialCharacters(textContent, {
    isAttributeValue: false
  });
}

// src/write/xlsx/helpers/getCellAddress.js
function getCellAddress(rowIndex, columnIndex) {
  return "".concat(getColumnLetter(columnIndex)).concat(rowIndex + 1);
}
var LETTERS_COUNT = 26;
function getColumnLetter(columnIndex) {
  if (typeof columnIndex !== "number") {
    return "";
  }
  var prefix = Math.floor(columnIndex / LETTERS_COUNT);
  var letter = String.fromCharCode(97 + columnIndex % LETTERS_COUNT).toUpperCase();
  if (prefix === 0) {
    return letter;
  }
  return getColumnLetter(prefix - 1) + letter;
}

// src/write/xlsx/helpers/convertDateToSerialNumber.js
var daysBeforeUnixEpoch = 70 * 365 + 19;
var hour = 60 * 60 * 1e3;
var day = 24 * hour;
function convertDateToSerialNumber(date) {
  return date.getTime() / day + daysBeforeUnixEpoch;
}

// src/write/xlsx/files/sheet.xml/cell.js
function _slicedToArray(r, e) {
  return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray4(r, e) || _nonIterableRest();
}
function _nonIterableRest() {
  throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray4(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray4(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray4(r, a) : void 0;
  }
}
function _arrayLikeToArray4(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function _iterableToArrayLimit(r, l) {
  var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (null != t) {
    var e, n, i, u, a = [], f = true, o = false;
    try {
      if (i = (t = t.call(r)).next, 0 === l) {
        if (Object(t) !== t) return;
        f = false;
      } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = true) ;
    } catch (r2) {
      o = true, n = r2;
    } finally {
      try {
        if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return;
      } finally {
        if (o) throw n;
      }
    }
    return a;
  }
}
function _arrayWithHoles(r) {
  if (Array.isArray(r)) return r;
}
function generateCell(tag, _ref, index, rowIndex) {
  var value = _ref.value, type = _ref.type, cellStyleId = _ref.cellStyleId, findOrCreateSharedString = _ref.findOrCreateSharedString;
  if (value === null) {
    if (cellStyleId === void 0) {
      return "";
    }
  }
  var cellAttributes = {
    r: getCellAddress(rowIndex, index)
  };
  if (cellStyleId !== void 0) {
    cellAttributes.s = String(cellStyleId);
  }
  if (value === null) {
    return tag("c", cellAttributes, null, index);
  }
  if (type === Date && cellStyleId === void 0) {
    throw new Error("No `format` was specified for a `Date` value in a cell in row ".concat(rowIndex + 1, " column ").concat(index + 1, ". Either specify a `format` for this cell or specify a default global one by passing `dateFormat` option to `writeXlsxFile()` function"));
  }
  var valueTextContent = getValueTextContent(type, value, findOrCreateSharedString);
  var typeAttribute = getTypeAttribute(type);
  if (typeAttribute) {
    cellAttributes.t = typeAttribute;
  }
  var _getOpeningAndClosing = getOpeningAndClosingTags(type), _getOpeningAndClosing2 = _slicedToArray(_getOpeningAndClosing, 2), valueOpeningTags = _getOpeningAndClosing2[0], valueClosingTags = _getOpeningAndClosing2[1];
  var cellXml = valueOpeningTags + valueTextContent + valueClosingTags;
  return tag("c", cellAttributes, cellXml, index);
}
function getTypeAttribute(type) {
  switch (type) {
    case String:
      return "s";
    // I don't know why did I comment out the use of "inlineStr" XLSX type.
    // Perhaps there were some issues with it. Or perhaps there weren't and everyone else just uses "s".
    // // "inlineStr" type is used instead of "s" to avoid creating a "shared strings" index.
    // return 'inlineStr'
    case Number:
      return;
    case Date:
      return;
    case Boolean:
      return "b";
    case "Formula":
      return;
    default:
      throw new Error("Unknown type: ".concat(type && type.name || type));
  }
}
function getValueTextContent(type, value, findOrCreateSharedString) {
  switch (type) {
    case String:
      if (typeof value !== "string") {
        throw new Error("Invalid cell value: ".concat(value, ". Expected a string"));
      }
      return findOrCreateSharedString(value);
    case Number:
      if (typeof value !== "number") {
        throw new Error("Invalid cell value: ".concat(value, ". Expected a number"));
      }
      return String(value);
    case Date:
      if (!(value instanceof Date)) {
        throw new Error("Invalid cell value: ".concat(value, ". Expected a Date"));
      }
      return String(convertDateToSerialNumber(value));
    case Boolean:
      if (typeof value !== "boolean") {
        throw new Error("Invalid cell value: ".concat(value, ". Expected a boolean"));
      }
      return value ? "1" : "0";
    case "Formula":
      if (typeof value !== "string") {
        throw new Error("Invalid cell value: ".concat(value, ". Expected a string"));
      }
      return sanitizeTextContent(value);
    default:
      throw new Error("Unknown type: ".concat(type && type.name || type));
  }
}
var TAG_BRACKET_LEFT_REGEXP = /</g;
function getOpeningAndClosingTags(type) {
  var openingTags = getOpeningTags(type);
  var closingTags = openingTags.replace(TAG_BRACKET_LEFT_REGEXP, "</");
  return [openingTags, closingTags];
}
function getOpeningTags(type) {
  switch (type) {
    // case 'inlineStr':
    //   return '<is><t>'
    case "Formula":
      return "<f>";
    default:
      return "<v>";
  }
}

// src/write/xlsx/helpers/hasAlignment.js
function hasAlignment(_ref) {
  var align = _ref.align, alignVertical = _ref.alignVertical, textRotation = _ref.textRotation, indent = _ref.indent, wrap = _ref.wrap;
  return Boolean(align || alignVertical || typeof textRotation === "number" || indent || wrap);
}

// src/write/xlsx/helpers/hasBorder.js
function hasBorder(_ref) {
  var borderColor = _ref.borderColor, borderStyle = _ref.borderStyle, leftBorderColor = _ref.leftBorderColor, leftBorderStyle = _ref.leftBorderStyle, rightBorderColor = _ref.rightBorderColor, rightBorderStyle = _ref.rightBorderStyle, topBorderColor = _ref.topBorderColor, topBorderStyle = _ref.topBorderStyle, bottomBorderColor = _ref.bottomBorderColor, bottomBorderStyle = _ref.bottomBorderStyle;
  return Boolean(borderColor || borderStyle || leftBorderColor || leftBorderStyle || rightBorderColor || rightBorderStyle || topBorderColor || topBorderStyle || bottomBorderColor || bottomBorderStyle);
}

// src/write/xlsx/helpers/hasFill.js
function hasFill(_ref) {
  var backgroundColor = _ref.backgroundColor, fillPatternStyle = _ref.fillPatternStyle, fillPatternColor = _ref.fillPatternColor;
  return Boolean(backgroundColor || fillPatternStyle && fillPatternColor);
}

// src/write/xlsx/helpers/hasFont.js
function hasFont(_ref) {
  var fontFamily = _ref.fontFamily, fontSize = _ref.fontSize, fontWeight = _ref.fontWeight, fontStyle = _ref.fontStyle, textDecoration = _ref.textDecoration, textColor = _ref.textColor;
  return Boolean(fontFamily || typeof fontSize === "number" || fontWeight || fontStyle || textDecoration && Object.keys(textDecoration).length > 0 || textColor);
}

// src/write/xlsx/files/sheet.xml/helpers/getCellStyleProperties.js
function getCellStyleProperties(cell, features) {
  var align = cell.align, alignVertical = cell.alignVertical, textRotation = cell.textRotation, indent = cell.indent, wrap = cell.wrap, fontFamily = cell.fontFamily, fontSize = cell.fontSize, fontWeight = cell.fontWeight, fontStyle = cell.fontStyle, textDecoration = cell.textDecoration, textColor = cell.textColor, backgroundColor = cell.backgroundColor, fillPatternStyle = cell.fillPatternStyle, fillPatternColor = cell.fillPatternColor, borderColor = cell.borderColor, borderStyle = cell.borderStyle, leftBorderColor = cell.leftBorderColor, leftBorderStyle = cell.leftBorderStyle, rightBorderColor = cell.rightBorderColor, rightBorderStyle = cell.rightBorderStyle, topBorderColor = cell.topBorderColor, topBorderStyle = cell.topBorderStyle, bottomBorderColor = cell.bottomBorderColor, bottomBorderStyle = cell.bottomBorderStyle;
  if (hasAlignment({
    align,
    alignVertical,
    textRotation,
    indent,
    wrap
  }) || hasFont({
    fontFamily,
    fontSize,
    fontWeight,
    fontStyle,
    textDecoration,
    textColor
  }) || hasFill({
    backgroundColor,
    fillPatternStyle,
    fillPatternColor
  }) || hasBorder({
    borderColor,
    borderStyle,
    leftBorderColor,
    leftBorderStyle,
    rightBorderColor,
    rightBorderStyle,
    topBorderColor,
    topBorderStyle,
    bottomBorderColor,
    bottomBorderStyle
  })) {
    return omitUndefinedProperties({
      // alignment
      align,
      alignVertical,
      textRotation,
      indent,
      wrap,
      // font
      fontFamily,
      fontSize,
      fontWeight,
      fontStyle,
      textDecoration,
      textColor,
      // fill
      backgroundColor,
      fillPatternStyle,
      fillPatternColor,
      // border
      borderColor,
      borderStyle,
      leftBorderColor,
      leftBorderStyle,
      rightBorderColor,
      rightBorderStyle,
      topBorderColor,
      topBorderStyle,
      bottomBorderColor,
      bottomBorderStyle
    });
  }
}
function omitUndefinedProperties(object) {
  var filteredObject = {};
  for (var key in object) {
    if (object[key] !== void 0) {
      filteredObject[key] = object[key];
    }
  }
  return filteredObject;
}

// src/write/xlsx/helpers/isObject.js
var objectConstructor2 = {}.constructor;
function isObject2(object) {
  return object !== void 0 && object !== null && object.constructor === objectConstructor2;
}

// src/write/xlsx/helpers/isCellObject.js
function isCellObject(cell) {
  return isObject2(cell);
}

// src/write/xlsx/files/sheet.xml/row.js
function _typeof(o) {
  "@babel/helpers - typeof";
  return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o2) {
    return typeof o2;
  } : function(o2) {
    return o2 && "function" == typeof Symbol && o2.constructor === Symbol && o2 !== Symbol.prototype ? "symbol" : typeof o2;
  }, _typeof(o);
}
function ownKeys(e, r) {
  var t = Object.keys(e);
  if (Object.getOwnPropertySymbols) {
    var o = Object.getOwnPropertySymbols(e);
    r && (o = o.filter(function(r2) {
      return Object.getOwnPropertyDescriptor(e, r2).enumerable;
    })), t.push.apply(t, o);
  }
  return t;
}
function _objectSpread(e) {
  for (var r = 1; r < arguments.length; r++) {
    var t = null != arguments[r] ? arguments[r] : {};
    r % 2 ? ownKeys(Object(t), true).forEach(function(r2) {
      _defineProperty(e, r2, t[r2]);
    }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function(r2) {
      Object.defineProperty(e, r2, Object.getOwnPropertyDescriptor(t, r2));
    });
  }
  return e;
}
function _defineProperty(e, r, t) {
  return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: true, configurable: true, writable: true }) : e[r] = t, e;
}
function _toPropertyKey(t) {
  var i = _toPrimitive(t, "string");
  return "symbol" == _typeof(i) ? i : i + "";
}
function _toPrimitive(t, r) {
  if ("object" != _typeof(t) || !t) return t;
  var e = t[Symbol.toPrimitive];
  if (void 0 !== e) {
    var i = e.call(t, r || "default");
    if ("object" != _typeof(i)) return i;
    throw new TypeError("@@toPrimitive must return a primitive value.");
  }
  return ("string" === r ? String : Number)(t);
}
function _createForOfIteratorHelperLoose4(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray5(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray5(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray5(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray5(r, a) : void 0;
  }
}
function _arrayLikeToArray5(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function generateRow(tag, row, index, parameters) {
  var rowHeight;
  for (var _iterator = _createForOfIteratorHelperLoose4(row), _step; !(_step = _iterator()).done; ) {
    var cell = _step.value;
    if (isCellObject(cell)) {
      if (cell.height) {
        if (rowHeight === void 0 || rowHeight < cell.height) {
          rowHeight = cell.height;
        }
      }
    }
  }
  var rowIndex = index;
  var rowCellsXml = row.map(function(cell2, index2) {
    return getCellXml(tag, cell2, index2, rowIndex, parameters);
  }).join("");
  var rowAttributes = {
    r: index + 1
  };
  if (rowHeight) {
    rowAttributes.ht = rowHeight;
    rowAttributes.customHeight = 1;
  }
  return tag("row", rowAttributes, rowCellsXml, index);
}
function getCellXml(tag, cell, index, rowIndex, _ref) {
  var findOrCreateCellStyle = _ref.findOrCreateCellStyle, findOrCreateSharedString = _ref.findOrCreateSharedString, hasDefaultFont = _ref.hasDefaultFont, dateFormat = _ref.dateFormat, features = _ref.features;
  if (cell === void 0 || cell === null) {
    return "";
  }
  var cellObject = isCellObject(cell) ? cell : {
    value: cell
  };
  var cellStyleProperties = getCellStyleProperties(cellObject, features);
  var type = cellObject.type, value = cellObject.value, format = cellObject.format;
  if (isEmpty(value)) {
    value = null;
  } else {
    if (type === void 0) {
      type = detectValueType(value);
      if (type === void 0) {
        type = String;
        value = String(value);
      }
    }
  }
  if (format) {
    if (type !== Date && type !== Number && type !== String && type !== "Formula") {
      throw new Error('`format` "'.concat(format, '" was specified on a cell of type `').concat(type, '`. `format` could only be specified on a cell of type `Date`, `Number`, `String` or `"Formula"`.'));
    }
    if (type === String && format !== "@") {
      throw new Error('`format` "'.concat(format, '" was specified on a cell of type `String`. The only supported `format` for a cell of type `String` is "@".'));
    }
  } else {
    if (type === Date) {
      format = dateFormat;
    }
  }
  var hasFormat = Boolean(format);
  var hasCellStyle = Boolean(cellStyleProperties);
  var cellStyleId;
  if (hasDefaultFont || hasFormat || hasCellStyle) {
    cellStyleId = findOrCreateCellStyle(_objectSpread({
      format
    }, cellStyleProperties));
  }
  return generateCell(tag, {
    value,
    type,
    cellStyleId,
    findOrCreateSharedString
  }, index, rowIndex);
}
function isEmpty(value) {
  return value === void 0 || value === null || value === "";
}
function detectValueType(value) {
  switch (_typeof(value)) {
    case "string":
      return String;
    case "number":
      return Number;
    case "boolean":
      return Boolean;
    default:
      if (value instanceof Date) {
        return Date;
      }
  }
}

// src/write/xlsx/files/sheet.xml/sheetData.js
function generateSheetData(tag, sheetData, parameters) {
  var sheetDataXml = sheetData.map(function(row, index) {
    return generateRow(tag, row, index, parameters);
  }).join("");
  return tag("sheetData", null, sheetDataXml);
}

// src/write/xlsx/files/sheet.xml/processMergedCells.js
function processMergedCells(sheetData, features) {
  var mergedCells = [];
  var _cloneSheetData = function cloneSheetData() {
    sheetData = sheetData.slice();
    var i = 0;
    while (i < sheetData.length) {
      sheetData[i] = sheetData[i].slice();
      i++;
    }
    _cloneSheetData = function cloneSheetData2() {
      return sheetData;
    };
    return sheetData;
  };
  var rowIndex = 0;
  while (rowIndex < sheetData.length) {
    var row = sheetData[rowIndex];
    var columnIndex = 0;
    while (columnIndex < row.length) {
      var cell = row[columnIndex];
      if (cell) {
        var _cell$rowSpan = cell.rowSpan, rowSpan = _cell$rowSpan === void 0 ? 1 : _cell$rowSpan;
        var columnSpan = typeof cell.span === "number" ? cell.span : cell.columnSpan;
        if (typeof columnSpan !== "number") {
          columnSpan = 1;
        }
        if (columnSpan > 1 || rowSpan > 1) {
          processSpanningCells(sheetData, rowIndex, columnIndex, columnSpan, rowSpan, _cloneSheetData, features);
          mergedCells.push([[rowIndex, columnIndex], [rowIndex + (rowSpan ? rowSpan - 1 : 0), columnIndex + (columnSpan ? columnSpan - 1 : 0)]]);
        }
      }
      columnIndex++;
    }
    rowIndex++;
  }
  return {
    sheetData,
    mergedCells
  };
}
function processSpanningCells(sheetData, rowIndex, columnIndex, columnSpan, rowSpan, cloneSheetData, features) {
  var cellStyleProperties = getCellStyleProperties(sheetData[rowIndex][columnIndex], features);
  if (cellStyleProperties) {
    sheetData = cloneSheetData();
  }
  var i = rowIndex;
  while (i <= rowIndex + (rowSpan - 1)) {
    var j = columnIndex;
    while (j <= columnIndex + (columnSpan - 1)) {
      var cell = sheetData[i][j];
      if (i > rowIndex || j > columnIndex) {
        if (cell !== null && cell !== void 0) {
          throw new Error("[write-excel-file] When using `columnSpan` or `rowSpan` parameters, all hidden overlapped cells should be represented by `null`s or `undefined`s. Cell at row ".concat(rowIndex + 1, " and column ").concat(columnIndex + 1, " is configured with `columnSpan` ").concat(columnSpan, " and `rowSpan` ").concat(rowSpan, ". Cell at row ").concat(i + 1, " and column ").concat(j + 1, " is neither `null` nor `undefined`: ").concat(JSON.stringify(cell)));
        }
        if (cellStyleProperties) {
          sheetData[i][j] = cellStyleProperties;
        }
      }
      j++;
    }
    i++;
  }
}

// src/write/xlsx/files/sheet.xml/mergedCellsDescription.js
function _slicedToArray2(r, e) {
  return _arrayWithHoles2(r) || _iterableToArrayLimit2(r, e) || _unsupportedIterableToArray6(r, e) || _nonIterableRest2();
}
function _nonIterableRest2() {
  throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray6(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray6(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray6(r, a) : void 0;
  }
}
function _arrayLikeToArray6(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function _iterableToArrayLimit2(r, l) {
  var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (null != t) {
    var e, n, i, u, a = [], f = true, o = false;
    try {
      if (i = (t = t.call(r)).next, 0 === l) {
        if (Object(t) !== t) return;
        f = false;
      } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = true) ;
    } catch (r2) {
      o = true, n = r2;
    } finally {
      try {
        if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return;
      } finally {
        if (o) throw n;
      }
    }
    return a;
  }
}
function _arrayWithHoles2(r) {
  if (Array.isArray(r)) return r;
}
function generateMergedCellsDescription(tag, mergedCells) {
  if (mergedCells.length === 0) {
    return "";
  }
  var mergeCellsXml = mergedCells.map(function(_ref, index) {
    var _ref2 = _slicedToArray2(_ref, 2), from = _ref2[0], to = _ref2[1];
    var ref = getCellAddress(from[0], from[1]) + ":" + getCellAddress(to[0], to[1]);
    return tag("mergeCell", {
      ref
    }, null, index);
  }).join("");
  return tag("mergeCells", {
    count: mergedCells.length
  }, mergeCellsXml);
}

// src/write/xlsx/files/sheet.xml/pageMargins.js
function generatePageMargins(tag, _ref) {
  var orientation = _ref.orientation;
  if (orientation) {
    var marginLeft = 0.7;
    var marginRight = 0.7;
    var marginTop = 0.75;
    var marginBottom = 0.75;
    var header = 0.3;
    var footer = 0.3;
    return tag("pageMargins", {
      left: marginLeft,
      right: marginRight,
      top: marginTop,
      bottom: marginBottom,
      header,
      footer
    });
  }
  return "";
}

// src/write/xlsx/files/sheet.xml/pageSetup.js
function generatePageSetup(tag, _ref) {
  var orientation = _ref.orientation;
  if (orientation) {
    var paperSize = 9;
    return tag("pageSetup", {
      paperSize,
      orientation: sanitizeAttributeValue(orientation)
    });
  }
  return "";
}

// src/write/xlsx/files/sheet.xml/drawingReference.js
function generateDrawingReference(tag) {
  return tag("drawing", {
    "r:id": "rId-drawing-1"
  });
}

// src/write/xlsx/files/sheet.xml/sheet.xml.js
function generateSheetXml(sheetXmlParameters, features) {
  var sheetData_ = sheetXmlParameters.sheetData, sheetOptions = sheetXmlParameters.sheetOptions, sheetIndex = sheetXmlParameters.sheetIndex, sheetId = sheetXmlParameters.sheetId, hasDefaultFont = sheetXmlParameters.hasDefaultFont, findOrCreateCellStyle = sheetXmlParameters.findOrCreateCellStyle, findOrCreateSharedString = sheetXmlParameters.findOrCreateSharedString;
  var columns = sheetOptions.columns, dateFormat = sheetOptions.dateFormat, orientation = sheetOptions.orientation, showGridLines = sheetOptions.showGridLines, rightToLeft = sheetOptions.rightToLeft, zoomScale = sheetOptions.zoomScale;
  var _processMergedCells = processMergedCells(sheetData_, {
    features
  }), sheetData = _processMergedCells.sheetData, mergedCells = _processMergedCells.mergedCells;
  var tagFunctionProperties = {
    sheetIndex,
    sheetId
  };
  var tag = function tag2(tagName, attributes, innerXml, index) {
    return getElementXml("xl/worksheets/sheet{id}.xml", tagName, attributes, innerXml, index, tagFunctionProperties, sheetOptions, features);
  };
  var xml = '<?xml version="1.0" ?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' + // The order of elements is dictated by the XLSX spec.
  //
  // `<sheetViews/>`
  generateSheetViews(tag, {
    showGridLines,
    rightToLeft,
    zoomScale
  }, sheetIndex) + // `<cols/>`
  generateColumnsDescription(tag, columns) + // `<sheetData/>`
  generateSheetData(tag, sheetData, {
    findOrCreateCellStyle,
    findOrCreateSharedString,
    hasDefaultFont,
    dateFormat,
    features
  }) + // `<mergeCells/>`
  generateMergedCellsDescription(tag, mergedCells) + // `<pageMargins/>`
  generatePageMargins(tag, {
    orientation
  }) + // `<pageSetup/>`
  generatePageSetup(tag, {
    orientation
  }) + // `<drawing/>`
  generateDrawingReference(tag) + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/worksheets/sheet{id}.xml", features, sheetOptions, {
    sheetIndex,
    sheetId
  }) + "</worksheet>";
  xml = transformContent(xml, "xl/worksheets/sheet{id}.xml", features, sheetOptions, {
    sheetIndex,
    sheetId
  });
  return xml;
}

// src/write/xlsx/files/sheet.xml.rels.js
function generateSheetXmlRels(_ref) {
  var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId, sheetOptions = _ref.sheetOptions, features = _ref.features;
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + // Each sheet has at most one "drawing", so a globally-unique drawing ID is assumed to be the sheet ID.
  getDrawingRelationshipXml(sheetId) + // Apply any plugins that insert additional content to this XML.
  getAdditionalContent("xl/worksheets/_rels/sheet{id}.xml.rels", features, sheetOptions, {
    sheetIndex,
    sheetId
  }) + "</Relationships>";
  xml = transformContent(xml, "xl/worksheets/_rels/sheet{id}.xml.rels", features, sheetOptions, {
    sheetIndex,
    sheetId
  });
  return xml;
}
function getDrawingRelationshipXml(sheetId) {
  return '<Relationship Id="rId-drawing-1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing'.concat(sheetId, '.xml"/>');
}

// src/write/xlsx/files/sharedStrings.xml.js
function _createForOfIteratorHelperLoose5(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray7(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray7(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray7(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray7(r, a) : void 0;
  }
}
function _arrayLikeToArray7(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function generateSharedStringsXml(sharedStrings) {
  var xml = '<?xml version="1.0"?>';
  xml += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  for (var _iterator = _createForOfIteratorHelperLoose5(sharedStrings), _step; !(_step = _iterator()).done; ) {
    var string = _step.value;
    var attributes = string.trim().length === string.length ? "" : ' xml:space="preserve"';
    xml += "<si><t".concat(attributes, ">");
    xml += sanitizeTextContent(string);
    xml += "</t></si>";
  }
  xml += "</sst>";
  return xml;
}

// src/write/xlsx/helpers/getXlsxColorForHexColor.js
function getXlsxColorForHexColor(color) {
  if (color[0] !== "#") {
    throw new Error('Color "'.concat(color, '" must start with a "#"'));
  }
  return "FF".concat(color.slice("#".length).toUpperCase());
}

// src/write/xlsx/helpers/getFillXml.js
function getFillXml(fill) {
  var _ref = arguments.length > 1 && arguments[1] !== void 0 ? arguments[1] : {}, conditionalFormatting = _ref.conditionalFormatting;
  var backgroundColor = fill.backgroundColor, fillPatternStyle = fill.fillPatternStyle, fillPatternColor = fill.fillPatternColor;
  var isSolidFill = !fillPatternStyle || fillPatternStyle === "solid";
  if (!hasFill(fill)) {
    return '<fill><patternFill patternType="none"/></fill>';
  }
  var xml = "<fill>";
  xml += '<patternFill patternType="'.concat(isSolidFill ? "solid" : fillPatternStyle, '">');
  xml += '<fgColor rgb="'.concat(sanitizeAttributeValue(getXlsxColorForHexColor(isSolidFill ? backgroundColor : fillPatternColor)), '"/>');
  xml += "<bgColor ".concat(isSolidFill && !conditionalFormatting ? 'indexed="64"' : 'rgb="' + sanitizeAttributeValue(getXlsxColorForHexColor(backgroundColor)) + '"', "/>");
  xml += "</patternFill>";
  xml += "</fill>";
  return xml;
}

// src/write/xlsx/helpers/getBorderXml.js
function getBorderXml(_ref) {
  var borderColor = _ref.borderColor, borderStyle = _ref.borderStyle, leftBorderColor = _ref.leftBorderColor, leftBorderStyle = _ref.leftBorderStyle, rightBorderColor = _ref.rightBorderColor, rightBorderStyle = _ref.rightBorderStyle, topBorderColor = _ref.topBorderColor, topBorderStyle = _ref.topBorderStyle, bottomBorderColor = _ref.bottomBorderColor, bottomBorderStyle = _ref.bottomBorderStyle;
  var left = {
    style: leftBorderStyle || borderStyle,
    color: leftBorderColor || borderColor
  };
  var right = {
    style: rightBorderStyle || borderStyle,
    color: rightBorderColor || borderColor
  };
  var top = {
    style: topBorderStyle || borderStyle,
    color: topBorderColor || borderColor
  };
  var bottom = {
    style: bottomBorderStyle || borderStyle,
    color: bottomBorderColor || borderColor
  };
  var xml = "<border>";
  xml += getSideBorderXml("left", left);
  xml += getSideBorderXml("right", right);
  xml += getSideBorderXml("top", top);
  xml += getSideBorderXml("bottom", bottom);
  xml += "<diagonal/>";
  xml += "</border>";
  return xml;
}
function getSideBorderXml(side, _ref2) {
  var style = _ref2.style, color = _ref2.color;
  if (color && !style) {
    style = "thin";
  }
  var hasChildren = Boolean(color);
  return "<".concat(side) + (style ? ' style="'.concat(sanitizeAttributeValue(style), '"') : "") + (hasChildren ? ">" : "/>") + (color ? '<color rgb="'.concat(sanitizeAttributeValue(getXlsxColorForHexColor(color)), '"/>') : "") + (hasChildren ? "</".concat(side, ">") : "");
}

// src/write/xlsx/helpers/getFontXml.js
function getFontXml(font) {
  var _ref = arguments.length > 1 && arguments[1] !== void 0 ? arguments[1] : {}, isDefaultGenericFont = _ref.isDefaultGenericFont;
  var fontFamily = font.fontFamily, fontSize = font.fontSize, fontWeight = font.fontWeight, fontStyle = font.fontStyle, textDecoration = font.textDecoration, textColor = font.textColor;
  var xml = "<font>";
  if (fontFamily) {
    xml += '<name val="'.concat(sanitizeAttributeValue(fontFamily), '"/>');
  }
  if (typeof fontSize === "number") {
    xml += '<sz val="'.concat(fontSize, '"/>');
  }
  xml += '<family val="2"/>';
  if (!(fontFamily || typeof fontSize === "number")) {
    xml += '<scheme val="minor"/>';
  }
  if (fontWeight === "bold") {
    xml += "<b/>";
  }
  if (fontStyle === "italic") {
    xml += "<i/>";
  }
  if (textDecoration) {
    if (textDecoration.strikethrough) {
      xml += "<strike/>";
    }
    if (textDecoration.underline) {
      xml += "<u/>";
    } else if (textDecoration.doubleUnderline) {
      xml += '<u val="double"/>';
    }
  }
  if (textColor) {
    xml += '<color rgb="'.concat(sanitizeAttributeValue(getXlsxColorForHexColor(textColor)), '"/>');
  } else {
    xml += '<color theme="1"/>';
  }
  xml += "</font>";
  return xml;
}

// src/write/xlsx/helpers/getAlignmentXml.js
function getAlignmentXml(_ref) {
  var align = _ref.align, alignVertical = _ref.alignVertical, textRotation = _ref.textRotation, indent = _ref.indent, wrap = _ref.wrap;
  return "<alignment" + (align ? ' horizontal="'.concat(sanitizeAttributeValue(align), '"') : "") + (alignVertical ? ' vertical="'.concat(sanitizeAttributeValue(alignVertical), '"') : "") + (textRotation ? ' textRotation="'.concat(getTextRotation(validateTextRotation(textRotation)), '"') : "") + (indent ? ' indent="'.concat(sanitizeAttributeValue(String(indent)), '"') : "") + (wrap ? ' wrapText="1"' : "") + "/>";
}
function validateTextRotation(textRotation) {
  if (!(textRotation >= -90 && textRotation <= 90)) {
    throw new Error("Unsupported text rotation angle: ".concat(textRotation, ". Values from -90 to 90 are supported."));
  }
  return textRotation;
}
function getTextRotation(textRotation) {
  if (textRotation < 0) {
    return 90 - textRotation;
  }
  return textRotation;
}

// src/write/xlsx/initializeStyles.js
function _typeof2(o) {
  "@babel/helpers - typeof";
  return _typeof2 = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o2) {
    return typeof o2;
  } : function(o2) {
    return o2 && "function" == typeof Symbol && o2.constructor === Symbol && o2 !== Symbol.prototype ? "symbol" : typeof o2;
  }, _typeof2(o);
}
function ownKeys2(e, r) {
  var t = Object.keys(e);
  if (Object.getOwnPropertySymbols) {
    var o = Object.getOwnPropertySymbols(e);
    r && (o = o.filter(function(r2) {
      return Object.getOwnPropertyDescriptor(e, r2).enumerable;
    })), t.push.apply(t, o);
  }
  return t;
}
function _objectSpread2(e) {
  for (var r = 1; r < arguments.length; r++) {
    var t = null != arguments[r] ? arguments[r] : {};
    r % 2 ? ownKeys2(Object(t), true).forEach(function(r2) {
      _defineProperty2(e, r2, t[r2]);
    }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys2(Object(t)).forEach(function(r2) {
      Object.defineProperty(e, r2, Object.getOwnPropertyDescriptor(t, r2));
    });
  }
  return e;
}
function _defineProperty2(e, r, t) {
  return (r = _toPropertyKey2(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: true, configurable: true, writable: true }) : e[r] = t, e;
}
function _toPropertyKey2(t) {
  var i = _toPrimitive2(t, "string");
  return "symbol" == _typeof2(i) ? i : i + "";
}
function _toPrimitive2(t, r) {
  if ("object" != _typeof2(t) || !t) return t;
  var e = t[Symbol.toPrimitive];
  if (void 0 !== e) {
    var i = e.call(t, r || "default");
    if ("object" != _typeof2(i)) return i;
    throw new TypeError("@@toPrimitive must return a primitive value.");
  }
  return ("string" === r ? String : Number)(t);
}
var FORMAT_ID_STARTS_FROM = 100;
function initializeStyles(defaultFont) {
  var formats = [];
  var formatIdByFormat = {};
  var styles = [];
  var stylesIndex = {};
  var fonts = [];
  var fontIdByFontKey = {};
  var fills = [];
  var fillIdByFillKey = {};
  var borders = [];
  var borderIdByBorderKey = {};
  var defaultFontId = fonts.length;
  fontIdByFontKey[getKey(defaultFont)] = defaultFontId;
  fonts.push(defaultFont || {});
  var defaultFill = {};
  fillIdByFillKey[getKey(defaultFill)] = fills.length;
  fills.push(defaultFill);
  var defaultBorder = {};
  borderIdByBorderKey[getKey(defaultBorder)] = borders.length;
  borders.push(defaultBorder);
  var gray125Fill = {
    gray125: true
  };
  fills.push(gray125Fill);
  function findOrCreateCellStyle(cellStyle) {
    var format = cellStyle.format, align = cellStyle.align, alignVertical = cellStyle.alignVertical, textRotation = cellStyle.textRotation, indent = cellStyle.indent, wrap = cellStyle.wrap, fontFamily = cellStyle.fontFamily, fontSize = cellStyle.fontSize, fontWeight = cellStyle.fontWeight, fontStyle = cellStyle.fontStyle, textDecoration = cellStyle.textDecoration, textColor = cellStyle.textColor, backgroundColor = cellStyle.backgroundColor, fillPatternStyle = cellStyle.fillPatternStyle, fillPatternColor = cellStyle.fillPatternColor, borderColor = cellStyle.borderColor, borderStyle = cellStyle.borderStyle, leftBorderColor = cellStyle.leftBorderColor, leftBorderStyle = cellStyle.leftBorderStyle, rightBorderColor = cellStyle.rightBorderColor, rightBorderStyle = cellStyle.rightBorderStyle, topBorderColor = cellStyle.topBorderColor, topBorderStyle = cellStyle.topBorderStyle, bottomBorderColor = cellStyle.bottomBorderColor, bottomBorderStyle = cellStyle.bottomBorderStyle;
    var font = {
      fontFamily,
      fontSize,
      fontWeight,
      fontStyle,
      textDecoration,
      textColor
    };
    var fill = {
      backgroundColor,
      fillPatternStyle,
      fillPatternColor
    };
    var border = {
      borderColor,
      borderStyle,
      leftBorderColor,
      leftBorderStyle,
      rightBorderColor,
      rightBorderStyle,
      topBorderColor,
      topBorderStyle,
      bottomBorderColor,
      bottomBorderStyle
    };
    var alignment = {
      align,
      alignVertical,
      textRotation,
      indent,
      wrap
    };
    var formatKey = getKey(format);
    var fontKey = getKey(font);
    var fillKey = getKey(fill);
    var borderKey = getKey(border);
    var styleKey = getKey(cellStyle);
    var addStyle = function addStyle2() {
      var formatId;
      if (format) {
        formatId = formatIdByFormat[formatKey];
        if (formatId === void 0) {
          formatId = FORMAT_ID_STARTS_FROM + formats.length;
          formatIdByFormat[formatKey] = formatId;
          formats.push(format);
        }
      }
      var fontId;
      if (hasFont(font)) {
        fontId = fontIdByFontKey[fontKey];
        if (fontId === void 0) {
          fontId = fonts.length;
          fontIdByFontKey[fontKey] = fontId;
          fonts.push(_objectSpread2(_objectSpread2({}, font), {}, {
            fontSize: font.fontSize || defaultFont && defaultFont.fontSize,
            fontFamily: font.fontFamily || defaultFont && defaultFont.fontFamily
          }));
        }
      } else if (defaultFont) {
        fontId = defaultFontId;
      }
      var fillId;
      if (hasFill(fill)) {
        fillId = fillIdByFillKey[fillKey];
        if (fillId === void 0) {
          fillId = fills.length;
          fillIdByFillKey[fillKey] = fillId;
          fills.push({
            backgroundColor,
            fillPatternStyle,
            fillPatternColor
          });
        }
      }
      var borderId;
      if (hasBorder(border)) {
        borderId = borderIdByBorderKey[borderKey];
        if (borderId === void 0) {
          borderId = borders.length;
          borderIdByBorderKey[borderKey] = borderId;
          borders.push(border);
        }
      }
      var styleId = styles.length;
      stylesIndex[styleKey] = styleId;
      styles.push({
        formatId,
        fontId,
        fillId,
        borderId,
        alignment
      });
      return styleId;
    };
    if (stylesIndex[styleKey] !== void 0) {
      return stylesIndex[styleKey];
    }
    return addStyle();
  }
  findOrCreateCellStyle({});
  return {
    getCellStyles: function getCellStyles() {
      return {
        formats,
        styles,
        fonts,
        fills,
        borders
      };
    },
    findOrCreateCellStyle
  };
}
function getKey(object) {
  return JSON.stringify(object);
}

// src/write/xlsx/files/styles.xml.js
function _createForOfIteratorHelperLoose6(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray8(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray8(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray8(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray8(r, a) : void 0;
  }
}
function _arrayLikeToArray8(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function generateStylesXml(cellStyles, sheetsOptions, features) {
  var formats = cellStyles.formats, styles = cellStyles.styles, fonts = cellStyles.fonts, fills = cellStyles.fills, borders = cellStyles.borders;
  var xml = '<?xml version="1.0" ?>';
  xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  if (formats.length > 0) {
    xml += '<numFmts count="'.concat(formats.length, '">');
    for (var i = 0; i < formats.length; i++) {
      xml += '<numFmt numFmtId="'.concat(FORMAT_ID_STARTS_FROM + i, '" formatCode="').concat(sanitizeAttributeValue(formats[i]), '"/>');
    }
    xml += "</numFmts>";
  }
  xml += '<fonts count="'.concat(fonts.length, '">');
  for (var _iterator = _createForOfIteratorHelperLoose6(fonts), _step; !(_step = _iterator()).done; ) {
    var font = _step.value;
    xml += getFontXml(font);
  }
  xml += "</fonts>";
  xml += '<fills count="'.concat(fills.length, '">');
  for (var _iterator2 = _createForOfIteratorHelperLoose6(fills), _step2; !(_step2 = _iterator2()).done; ) {
    var fill = _step2.value;
    if (fill.gray125) {
      xml += "<fill>";
      xml += '<patternFill patternType="gray125"/>';
      xml += "</fill>";
    } else {
      xml += getFillXml(fill);
    }
  }
  xml += "</fills>";
  xml += '<borders count="'.concat(borders.length, '">');
  for (var _iterator3 = _createForOfIteratorHelperLoose6(borders), _step3; !(_step3 = _iterator3()).done; ) {
    var border = _step3.value;
    xml += getBorderXml(border);
  }
  xml += "</borders>";
  xml += '<cellXfs count="'.concat(styles.length, '">');
  for (var _iterator4 = _createForOfIteratorHelperLoose6(styles), _step4; !(_step4 = _iterator4()).done; ) {
    var cellStyle = _step4.value;
    var fontId = cellStyle.fontId, fillId = cellStyle.fillId, borderId = cellStyle.borderId, alignment = cellStyle.alignment, formatId = cellStyle.formatId;
    xml += "<xf " + [
      formatId === void 0 ? void 0 : 'numFmtId="'.concat(formatId, '" applyNumberFormat="1"'),
      fontId === void 0 ? void 0 : 'fontId="'.concat(fontId, '" applyFont="1"'),
      fillId === void 0 ? void 0 : 'fillId="'.concat(fillId, '" applyFill="1"'),
      borderId === void 0 ? void 0 : 'borderId="'.concat(borderId, '" applyBorder="1"'),
      hasAlignment(alignment) ? 'applyAlignment="1"' : void 0
      // 'xfId="0"'
    ].filter(function(_) {
      return _;
    }).join(" ") + ">" + // Possible horizontal alignment values:
    //  left, center, right, fill, justify, center_across, distributed.
    // Possible vertical alignment values:
    //  top, vcenter, bottom, vjustify, vdistributed.
    // https://xlsxwriter.readthedocs.io/format.html#set_align
    (hasAlignment(alignment) ? getAlignmentXml(alignment) : "") + "</xf>";
  }
  xml += "</cellXfs>";
  xml += getAdditionalContent("xl/styles.xml", features, sheetsOptions);
  xml += "</styleSheet>";
  xml = transformContent(xml, "xl/styles.xml", features, sheetsOptions);
  return xml;
}

// src/write/xml/findElementsNonRecursive_.js
function findElementsNonRecursive_(xml) {
  var _ref = arguments.length > 1 && arguments[1] !== void 0 ? arguments[1] : {}, tagName = _ref.tagName, stopAfterFirstMatch = _ref.stopAfterFirstMatch;
  var openingTagRegExp = new RegExp(getOpeningTagRegExpPattern(tagName), stopAfterFirstMatch ? void 0 : "g");
  var results = [];
  var openingTagMatch;
  while ((openingTagMatch = openingTagRegExp.exec(xml)) !== null) {
    var openingTag = openingTagMatch[0];
    var openingTagName = tagName || openingTagMatch[1];
    var attributes = {};
    var attributeRegExp = new RegExp('\\s+([^\\s=>]+)(?:="([^\\s=>]+)")', "g");
    var attributeMatch = void 0;
    while ((attributeMatch = attributeRegExp.exec(openingTag)) !== null) {
      attributes[attributeMatch[1]] = attributeMatch[2];
    }
    var result = {
      tagName: openingTagName,
      // openingTagMarkup: openingTag,
      openingTagStartIndex: openingTagMatch.index,
      openingTagEndIndex: openingTagMatch.index + openingTag.length - 1,
      openingTagAttributes: attributes,
      selfClosingTag: false,
      // closingTagMarkup: undefined,
      closingTagStartIndex: void 0,
      closingTagEndIndex: void 0
    };
    if (openingTag[openingTag.length - 2] === "/") {
      result.selfClosingTag = true;
    } else {
      var closingTagPosition = findClosingTagPosition(xml, result.openingTagEndIndex + 1, openingTagName);
      if (!closingTagPosition) {
        throw new Error("Invalid XML: opening tag was found but closing tag was not: </".concat(openingTagName, ">"));
      }
      result.closingTagStartIndex = closingTagPosition[0];
      result.closingTagEndIndex = closingTagPosition[1];
    }
    results.push(result);
    if (stopAfterFirstMatch) {
      break;
    }
    if (result.selfClosingTag) {
      openingTagRegExp.lastIndex = result.openingTagEndIndex + 1;
    } else {
      openingTagRegExp.lastIndex = result.closingTagEndIndex + 1;
    }
  }
  return results;
}
function findClosingTagPosition(xml, startFromIndex, tagName) {
  var openingOrClosingTagRegExp = new RegExp("<(/)?" + tagName + "(?:\\s+[^>]+|/)?>", "g");
  openingOrClosingTagRegExp.lastIndex = startFromIndex;
  var nestingLevel = 0;
  var openingOrClosingTagMatch;
  while ((openingOrClosingTagMatch = openingOrClosingTagRegExp.exec(xml)) !== null) {
    var openingOrClosingTag = openingOrClosingTagMatch[0];
    var closingTagMarker = openingOrClosingTagMatch[1];
    if (closingTagMarker) {
      if (nestingLevel > 0) {
        nestingLevel--;
      } else {
        return [openingOrClosingTagMatch.index, openingOrClosingTagMatch.index + openingOrClosingTag.length - 1];
      }
    } else {
      nestingLevel++;
    }
  }
}
function getOpeningTagRegExpPattern(tagName) {
  return "<" + (tagName || "([^\\s/>]+)") + "(?:\\s+[^>]+|/)?>";
}

// src/write/xml/findElement.js
function findElement(xml, tagName) {
  var elements = findElementsNonRecursive_(xml, {
    tagName,
    stopAfterFirstMatch: true
  });
  return elements[0];
}

// src/write/xml/setMarkupInsideElement.js
function setMarkupInsideElement(xml, element, replacementXml) {
  if (replacementXml) {
    if (element.selfClosingTag) {
      return xml.slice(0, element.openingTagEndIndex - "/".length) + ">" + replacementXml + "</" + element.tagName + ">" + xml.slice(element.openingTagEndIndex + 1);
    }
    return xml.slice(0, element.openingTagEndIndex + 1) + replacementXml + xml.slice(element.closingTagStartIndex);
  } else {
    if (element.selfClosingTag) {
      return xml;
    }
    return xml.slice(0, element.openingTagEndIndex) + "/>" + xml.slice(element.closingTagEndIndex + 1);
  }
}

// src/write/xml/appendMarkupInsideElement.js
function appendMarkupInsideElement(xml, element, addedXml) {
  if (element.selfClosingTag) {
    return setMarkupInsideElement(xml, element, addedXml);
  }
  return xml.slice(0, element.closingTagStartIndex) + addedXml + xml.slice(element.closingTagStartIndex);
}

// src/write/xml/getMarkupInsideElement.js
function getMarkupInsideElement(xml, element) {
  if (element.selfClosingTag) {
    return "";
  }
  return xml.substring(element.openingTagEndIndex + 1, element.closingTagStartIndex);
}

// src/write/xml/applyEnclosingElementOffset.js
function applyEnclosingElementOffset(element, enclosingElement) {
  var enclosingElementOffset = enclosingElement.openingTagEndIndex + 1;
  element.openingTagStartIndex += enclosingElementOffset;
  element.openingTagEndIndex += enclosingElementOffset;
  if (!element.selfClosingTag) {
    element.closingTagStartIndex += enclosingElementOffset;
    element.closingTagEndIndex += enclosingElementOffset;
  }
}

// src/write/xml/getChildElements.js
function _createForOfIteratorHelperLoose7(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray9(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray9(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray9(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray9(r, a) : void 0;
  }
}
function _arrayLikeToArray9(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function getChildElements(xml, element) {
  var children = findElementsNonRecursive_(getMarkupInsideElement(xml, element));
  for (var _iterator = _createForOfIteratorHelperLoose7(children), _step; !(_step = _iterator()).done; ) {
    var child = _step.value;
    applyEnclosingElementOffset(child, element);
  }
  return children;
}

// src/write/xml/prependMarkupInsideElement.js
function prependMarkupInsideElement(xml, element, addedXml) {
  if (element.selfClosingTag) {
    return setMarkupInsideElement(xml, element, addedXml);
  }
  return xml.slice(0, element.openingTagEndIndex + 1) + addedXml + xml.slice(element.openingTagEndIndex + 1);
}

// src/write/xml/insertElementMarkupAccordingToOrderOfSiblings.js
function _createForOfIteratorHelperLoose8(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray10(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray10(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray10(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray10(r, a) : void 0;
  }
}
function _arrayLikeToArray10(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function insertElementMarkupAccordingToOrderOfSiblings(xml, elementMarkup, orderOfSiblings) {
  for (var _len = arguments.length, parentElementTagNames = new Array(_len > 3 ? _len - 3 : 0), _key = 3; _key < _len; _key++) {
    parentElementTagNames[_key - 3] = arguments[_key];
  }
  if (parentElementTagNames.length === 0) {
    throw new Error("At least one parent element tag name is required");
  }
  parentElementTagNames = parentElementTagNames.slice().reverse();
  var parentElement;
  var _loop = function _loop3() {
    var parentElementTagName = _step.value;
    if (parentElement) {
      parentElement = getChildElements(xml, parentElement).find(function(_) {
        return _.tagName === parentElementTagName;
      });
    } else {
      parentElement = findElement(xml, parentElementTagName);
    }
    if (!parentElement) {
      throw new Error("Element not found: <".concat(parentElementTagName, ">"));
    }
  };
  for (var _iterator = _createForOfIteratorHelperLoose8(parentElementTagNames), _step; !(_step = _iterator()).done; ) {
    _loop();
  }
  var elementTagNameMatch = elementMarkup.match(TAG_NAME_REG_EXP);
  if (!elementTagNameMatch) {
    throw new Error("Couldn't extract tag name from markup: ".concat(elementMarkup));
  }
  var elementTagName = elementTagNameMatch[1];
  if (!orderOfSiblings || orderOfSiblings.length < 2) {
    return appendMarkupInsideElement(xml, parentElement, elementMarkup);
  }
  var children = getChildElements(xml, parentElement);
  if (children.length === 0) {
    return appendMarkupInsideElement(xml, parentElement, elementMarkup);
  }
  var elementTagNameOrder = orderOfSiblings.indexOf(elementTagName);
  if (elementTagNameOrder < 0) {
    return appendMarkupInsideElement(xml, parentElement, elementMarkup);
  }
  var tagNamesBeforeElement = orderOfSiblings.slice(0, elementTagNameOrder).reverse();
  var _loop2 = function _loop22() {
    var tagName = _step2.value;
    var precedingElement = children.find(function(_) {
      return _.tagName === tagName;
    });
    if (precedingElement) {
      return {
        v: xml.slice(0, precedingElement.selfClosingTag ? precedingElement.openingTagEndIndex + 1 : precedingElement.closingTagEndIndex + 1) + elementMarkup + xml.slice(precedingElement.selfClosingTag ? precedingElement.openingTagEndIndex + 1 : precedingElement.closingTagEndIndex + 1)
      };
    }
  }, _ret;
  for (var _iterator2 = _createForOfIteratorHelperLoose8(tagNamesBeforeElement), _step2; !(_step2 = _iterator2()).done; ) {
    _ret = _loop2();
    if (_ret) return _ret.v;
  }
  return prependMarkupInsideElement(xml, parentElement, elementMarkup);
}
var TAG_NAME_REG_EXP = new RegExp(getOpeningTagRegExpPattern());

// src/write/xlsx/helpers/orderOfSiblings.js
var orderOfSiblings_default = {
  // This is what Google AI returned for a search query: "xlsx workbook.xml elements order".
  // They say, this order is defined in `CT_Workbook` schema of ECMA-376 specification.
  "xl/workbook.xml": [["workbook", [
    "workbookPr",
    // (Workbook Properties)
    "bookViews",
    // (Book Views)
    "sheets",
    // (List of sheets)
    "pivotCaches",
    // (If applicable)
    "definedNames",
    // (Named ranges)
    "calcPr"
    // (Calculation properties)
  ]]],
  // This is what Google AI returned for a search query: "xlsx styles.xml elements order".
  // They say, this order is defined in `CT_Stylesheet` schema of ECMA-376 specification.
  "xl/styles.xml": [["styleSheet", [
    "numFmts",
    // Defines custom number formats.
    "fonts",
    // Contains font definitions (font name, size, color).
    "fills",
    // Defines cell background colors/patterns.
    "borders",
    // Defines cell borders.
    "cellStyleXfs",
    // Master cell styles (usually style 0 is "Normal").
    "cellXfs",
    // Actual cell formatting definitions (<xf>) referenced by cells.
    "cellStyles",
    // Connects named styles to cellStyleXfs.
    "dxfs",
    // Differential formats (conditional formatting).
    "tableStyles"
    // Table-specific styles.
  ]]],
  // This is what Google AI returned for a search query: "xlsx sheet.xml elements order".
  // They say, this order is defined in `CT_Worksheet` schema of ECMA-376 specification.
  "xl/worksheets/sheet{id}.xml": [["worksheet", [
    "sheetPr",
    // Sheet properties (e.g., tab color, filter configurations).
    "dimension",
    // Defines the range of used cells (e.g., ref="A1:C10").
    "sheetViews",
    // View settings like freeze panes, zoom level, and whether the sheet is selected.
    "sheetFormatPr",
    // Default row height and column width properties.
    "cols",
    // Defines column-specific properties (width, hidden status).
    "sheetData",
    // (Required) The main container for cell data, rows, and values.
    "sheetCalcPr",
    // Calculation properties.
    "sheetProtection",
    // Sheet-level security and locked status.
    "protectedRanges",
    // Specific ranges that are protected.
    "scenarios",
    // Worksheet scenarios.
    "autoFilter",
    // Filter range definitions.
    "sortState",
    // Current sort information.
    "dataConsolidate",
    // Data consolidation settings.
    "customSheetViews",
    // Custom view settings.
    "mergeCells",
    // Defines ranges of merged cells.
    "phoneticPr",
    // Phonetic information properties.
    "conditionalFormatting",
    // Rules for formatting based on cell values.
    "dataValidations",
    // Data entry rules and dropdowns.
    "hyperlinks",
    // Links to external sites or other sheet locations.
    "printOptions",
    // Defines specific printing preferences, such as centering, gridlines, and headings.
    "pageMargins",
    // Defines the white space between the worksheet data and the edges of the printed page.
    "pageSetup",
    // Printing and layout settings.
    "headerFooter",
    // Sheet header and footer content.
    "rowBreaks",
    // Manual page breaks.
    "colBreaks",
    // Manual page breaks.
    "drawing",
    // References to graphics or charts (points to drawing.xml).
    "legacyDrawing",
    // References to legacy objects like comments or form controls.
    "picture",
    // Background image properties.
    "oleObjects",
    // Embedded OLE objects.
    "controls",
    // Form or ActiveX controls.
    "tableParts",
    // Links to defined Excel tables within the sheet.
    "extLst"
    // Extension list for future-proofing and custom metadata.
  ]]]
};

// src/write/xlsx/helpers/getOrderOfSiblings.js
function getOrderOfSiblings(fileName) {
  var orderOfSiblings = orderOfSiblings_default[fileName];
  if (!orderOfSiblings) {
    return;
  }
  for (var _len = arguments.length, parentTagNames = new Array(_len > 1 ? _len - 1 : 0), _key = 1; _key < _len; _key++) {
    parentTagNames[_key - 1] = arguments[_key];
  }
  if (parentTagNames.length === 0) {
    throw new Error("At least one parent element tag name is required");
  }
  var _loop = function _loop2() {
    var tagName = _parentTagNames[_i];
    var orderOfSiblingsMatchingElement = orderOfSiblings.find(function(element) {
      if (Array.isArray(element)) {
        return element[0] === tagName;
      } else {
        return element === tagName;
      }
    });
    if (!orderOfSiblingsMatchingElement) {
      return {
        v: void 0
      };
    }
    if (Array.isArray(orderOfSiblingsMatchingElement)) {
      orderOfSiblings = orderOfSiblingsMatchingElement[1];
    } else {
      return {
        v: void 0
      };
    }
  }, _ret;
  for (var _i = 0, _parentTagNames = parentTagNames; _i < _parentTagNames.length; _i++) {
    _ret = _loop();
    if (_ret) return _ret.v;
  }
  if (orderOfSiblings) {
    return orderOfSiblings.map(function(element) {
      if (Array.isArray(element)) {
        return element[0];
      } else {
        return element;
      }
    });
  }
}

// src/write/xlsx/features/conditionalFormatting.js
function _createForOfIteratorHelperLoose9(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray11(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray11(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray11(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray11(r, a) : void 0;
  }
}
function _arrayLikeToArray11(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
var conditionalFormatting_default = {
  files: {
    transform: {
      "xl/worksheets/sheet{id}.xml": {
        // For some weird reason, Excel 2007 demands `<conditionalFormatting/>` element to go before `<drawing/>` element.
        // Because `<drawing/>` element is added before any of the `insert()`s of any features,
        // this feature has to use a `transform()` function instead of an `insert()` function.
        transform: function transform(xml, sheetOptions, _ref) {
          var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId;
          var conditionalFormatting = sheetOptions.conditionalFormatting;
          if (conditionalFormatting) {
            var conditionalFormattingRulesXml = getConditionalFormattingRulesXml(conditionalFormatting);
            return insertElementMarkupAccordingToOrderOfSiblings(xml, conditionalFormattingRulesXml, getOrderOfSiblings("xl/worksheets/sheet{id}.xml", "worksheet"), "worksheet");
          }
          return xml;
        }
      },
      "xl/styles.xml": {
        insert: function insert(sheetsOptions) {
          var sheetsConditionalFormatting = sheetsOptions.map(function(sheetOptions) {
            return sheetOptions.conditionalFormatting;
          });
          if (sheetsConditionalFormatting.some(Boolean)) {
            return getConditionalFormattingStylesXml(sheetsConditionalFormatting);
          }
        }
      }
    }
  }
};
function getConditionalFormattingRulesXml(conditionalFormattingRules) {
  var xml = "";
  var i = 0;
  for (var _iterator = _createForOfIteratorHelperLoose9(conditionalFormattingRules), _step; !(_step = _iterator()).done; ) {
    var conditionalFormattingRule = _step.value;
    var _conditionalFormattin = conditionalFormattingRule.cellRange, from = _conditionalFormattin.from, to = _conditionalFormattin.to, _conditionalFormattin2 = conditionalFormattingRule.condition, formula = _conditionalFormattin2.formula, operator = _conditionalFormattin2.operator, value = _conditionalFormattin2.value, value2 = _conditionalFormattin2.value2;
    var cellRange = getCellAddress(from.row - 1, from.column - 1) + ":" + getCellAddress(to.row - 1, to.column - 1);
    xml += '<conditionalFormatting sqref="'.concat(cellRange, '">');
    var priority = i + 1;
    var dxfId = conditionalFormattingRule._globalIndex;
    if (formula) {
      xml += '<cfRule type="expression" dxfId="'.concat(dxfId, '" priority="').concat(priority, '">');
      xml += "<formula>".concat(sanitizeTextContent(formula), "</formula>");
      xml += "</cfRule>";
    } else if (operator) {
      xml += '<cfRule type="cellIs" operator="'.concat(getXlsxOperatorName(operator), '" dxfId="').concat(dxfId, '" priority="').concat(priority, '">');
      xml += "<formula>".concat(sanitizeTextContent(formatValue(value)), "</formula>");
      if (getXlsxOperatorName(operator) === "between") {
        xml += "<formula>".concat(sanitizeTextContent(formatValue(value2)), "</formula>");
      }
      xml += "</cfRule>";
    } else {
      throw new Error("Invalid conditional formatting rule:\n".concat(JSON.stringify(conditionalFormattingRule, null, 2)));
    }
    xml += "</conditionalFormatting>";
    i++;
  }
  return xml;
}
function formatValue(value) {
  if (typeof value === "string") {
    return '"' + value + '"';
  }
  return String(value);
}
function getXlsxOperatorName(operator) {
  switch (operator) {
    case "<":
      return "lessThan";
    case ">":
      return "greaterThan";
    case "<=":
      return "lessThanOrEqual";
    case ">=":
      return "greaterThanOrEqual";
    case "=":
      return "equal";
    case "!=":
      return "notEqual";
    case "...":
      return "between";
    default:
      throw new Error("Unknown conditional formatting operator: ".concat(operator));
  }
}
function getConditionalFormattingStylesXml(sheetsConditionalFormatting) {
  var totalConditionalFormattingRulesCount = 0;
  for (var _iterator2 = _createForOfIteratorHelperLoose9(sheetsConditionalFormatting), _step2; !(_step2 = _iterator2()).done; ) {
    var conditionalFormattingOfSheet = _step2.value;
    if (conditionalFormattingOfSheet) {
      totalConditionalFormattingRulesCount += conditionalFormattingOfSheet.length;
    }
  }
  var xml = "";
  xml += '<dxfs count="'.concat(totalConditionalFormattingRulesCount, '">');
  for (var _iterator3 = _createForOfIteratorHelperLoose9(sheetsConditionalFormatting), _step3; !(_step3 = _iterator3()).done; ) {
    var _conditionalFormattingOfSheet = _step3.value;
    if (_conditionalFormattingOfSheet) {
      for (var _iterator4 = _createForOfIteratorHelperLoose9(_conditionalFormattingOfSheet), _step4; !(_step4 = _iterator4()).done; ) {
        var conditionalFormattingRule = _step4.value;
        var _conditionalFormattin3 = conditionalFormattingRule.style, fontFamily = _conditionalFormattin3.fontFamily, fontSize = _conditionalFormattin3.fontSize, fontWeight = _conditionalFormattin3.fontWeight, fontStyle = _conditionalFormattin3.fontStyle, textDecoration = _conditionalFormattin3.textDecoration, textColor = _conditionalFormattin3.textColor, backgroundColor = _conditionalFormattin3.backgroundColor, fillPatternStyle = _conditionalFormattin3.fillPatternStyle, fillPatternColor = _conditionalFormattin3.fillPatternColor, borderColor = _conditionalFormattin3.borderColor, borderStyle = _conditionalFormattin3.borderStyle, leftBorderColor = _conditionalFormattin3.leftBorderColor, leftBorderStyle = _conditionalFormattin3.leftBorderStyle, rightBorderColor = _conditionalFormattin3.rightBorderColor, rightBorderStyle = _conditionalFormattin3.rightBorderStyle, topBorderColor = _conditionalFormattin3.topBorderColor, topBorderStyle = _conditionalFormattin3.topBorderStyle, bottomBorderColor = _conditionalFormattin3.bottomBorderColor, bottomBorderStyle = _conditionalFormattin3.bottomBorderStyle;
        xml += "<dxf>";
        var font = {
          fontFamily,
          fontSize,
          fontWeight,
          fontStyle,
          textDecoration,
          textColor
        };
        if (hasFont(font)) {
          if (fontFamily) {
            throw new Error("Conditional formatting can't be used to override font family");
          }
          if (typeof fontSize === "number") {
            throw new Error("Conditional formatting can't be used to override font size");
          }
          xml += getFontXml(font);
        }
        var fill = {
          backgroundColor,
          fillPatternStyle,
          fillPatternColor
        };
        if (hasFill(fill)) {
          xml += getFillXml(fill, {
            conditionalFormatting: true
          });
        }
        var border = {
          borderColor,
          borderStyle,
          leftBorderColor,
          leftBorderStyle,
          rightBorderColor,
          rightBorderStyle,
          topBorderColor,
          topBorderStyle,
          bottomBorderColor,
          bottomBorderStyle
        };
        if (hasBorder(border)) {
          xml += getBorderXml(border);
        }
        xml += "</dxf>";
      }
    }
  }
  xml += "</dxfs>";
  return xml;
}

// src/write/xlsx/helpers/getFileExtensionForContentType.js
function getFileExtensionForContentType(contentType) {
  if (!contentType) {
    throw new Error("`contentType` is required");
  }
  var extension = contentType.toLowerCase().replace(/.*\//, "");
  if (!extension) {
    throw new Error("Unsupported `contentType`: " + contentType);
  }
  return extension;
}

// src/write/xlsx/features/images.js
function _typeof3(o) {
  "@babel/helpers - typeof";
  return _typeof3 = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o2) {
    return typeof o2;
  } : function(o2) {
    return o2 && "function" == typeof Symbol && o2.constructor === Symbol && o2 !== Symbol.prototype ? "symbol" : typeof o2;
  }, _typeof3(o);
}
function _createForOfIteratorHelperLoose10(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray12(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _toConsumableArray(r) {
  return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray12(r) || _nonIterableSpread();
}
function _nonIterableSpread() {
  throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray12(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray12(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray12(r, a) : void 0;
  }
}
function _iterableToArray(r) {
  if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r);
}
function _arrayWithoutHoles(r) {
  if (Array.isArray(r)) return _arrayLikeToArray12(r);
}
function _arrayLikeToArray12(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function ownKeys3(e, r) {
  var t = Object.keys(e);
  if (Object.getOwnPropertySymbols) {
    var o = Object.getOwnPropertySymbols(e);
    r && (o = o.filter(function(r2) {
      return Object.getOwnPropertyDescriptor(e, r2).enumerable;
    })), t.push.apply(t, o);
  }
  return t;
}
function _objectSpread3(e) {
  for (var r = 1; r < arguments.length; r++) {
    var t = null != arguments[r] ? arguments[r] : {};
    r % 2 ? ownKeys3(Object(t), true).forEach(function(r2) {
      _defineProperty3(e, r2, t[r2]);
    }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys3(Object(t)).forEach(function(r2) {
      Object.defineProperty(e, r2, Object.getOwnPropertyDescriptor(t, r2));
    });
  }
  return e;
}
function _defineProperty3(e, r, t) {
  return (r = _toPropertyKey3(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: true, configurable: true, writable: true }) : e[r] = t, e;
}
function _toPropertyKey3(t) {
  var i = _toPrimitive3(t, "string");
  return "symbol" == _typeof3(i) ? i : i + "";
}
function _toPrimitive3(t, r) {
  if ("object" != _typeof3(t) || !t) return t;
  var e = t[Symbol.toPrimitive];
  if (void 0 !== e) {
    var i = e.call(t, r || "default");
    if ("object" != _typeof3(i)) return i;
    throw new TypeError("@@toPrimitive must return a primitive value.");
  }
  return ("string" === r ? String : Number)(t);
}
var images_default = {
  files: {
    transform: {
      "[Content_Types].xml": {
        insert: function insert2(sheetsOptions) {
          var sheetsImages = sheetsOptions.map(function(sheetOptions) {
            return sheetOptions.images;
          });
          if (sheetsImages.some(Boolean)) {
            return getContentTypesXml(sheetsImages);
          }
        }
      },
      "xl/drawings/drawing{id}.xml": {
        insert: function insert3(sheetOptions, _ref) {
          var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId;
          var images = sheetOptions.images;
          if (images) {
            return getImagesDrawingXml({
              images
            });
          }
        }
      },
      "xl/drawings/_rels/drawing{id}.xml.rels": {
        insert: function insert4(sheetOptions, _ref2) {
          var sheetIndex = _ref2.sheetIndex, sheetId = _ref2.sheetId;
          var images = sheetOptions.images;
          if (images) {
            return getImagesDrawingXmlRels({
              images,
              sheetIndex
            });
          }
        }
      }
    },
    write: {
      files: function files(sheetsOptions, _ref3) {
        var read2 = _ref3.read;
        var sheetsImages = sheetsOptions.map(function(sheetOptions) {
          return sheetOptions.images;
        });
        if (sheetsImages.some(Boolean)) {
          return sheetsImages.map(function(images, sheetIndex) {
            if (images) {
              return images.reduce(function(imagesContent, image, imageIndex) {
                return _objectSpread3(_objectSpread3({}, imagesContent), {}, _defineProperty3({}, "xl/media/".concat(getImageFileName(image, {
                  sheetIndex,
                  imageIndex
                })), image.content));
              }, {});
            }
            return {};
          }).reduce(function(allImagesContent, sheetImagesContent) {
            return _objectSpread3(_objectSpread3({}, allImagesContent), sheetImagesContent);
          }, {});
        }
      }
    }
  }
};
function getContentTypesXml(sheetsImages) {
  var imagesFromAllSheets = sheetsImages.reduce(function(all, images) {
    return [].concat(_toConsumableArray(all), _toConsumableArray(images || []));
  }, []);
  var xml = "";
  for (var _iterator = _createForOfIteratorHelperLoose10(getFileExtensionContentTypes(imagesFromAllSheets)), _step; !(_step = _iterator()).done; ) {
    var _step$value = _step.value, fileExtension = _step$value.fileExtension, contentType = _step$value.contentType;
    xml += '<Default Extension="'.concat(fileExtension, '" ContentType="').concat(contentType, '"/>');
  }
  return xml;
}
function getFileExtensionContentTypes(images) {
  var fileExtensionContentTypes = [];
  var addFileExtensionContentType = function addFileExtensionContentType2(image2) {
    var fileExtension = getFileExtensionForContentType(image2.contentType);
    var existingFileExtensionContentType = fileExtensionContentTypes.find(function(_) {
      return _.fileExtension === fileExtension;
    });
    if (!existingFileExtensionContentType) {
      fileExtensionContentTypes.push({
        fileExtension,
        contentType: image2.contentType
      });
    }
  };
  for (var _iterator2 = _createForOfIteratorHelperLoose10(images), _step2; !(_step2 = _iterator2()).done; ) {
    var image = _step2.value;
    addFileExtensionContentType(image);
  }
  return fileExtensionContentTypes;
}
function getImageFileName(image, _ref4) {
  var sheetIndex = _ref4.sheetIndex, imageIndex = _ref4.imageIndex;
  var sheetNumber = sheetIndex + 1;
  var imageNumber = imageIndex + 1;
  return "sheet".concat(sheetNumber, "-image").concat(imageNumber, ".").concat(getFileExtensionForContentType(image.contentType));
}
function getImagesDrawingXml(_ref5) {
  var images = _ref5.images;
  var xml = "";
  var i = 0;
  var _loop = function _loop2() {
    var image = _step3.value;
    var imageId = i + 1;
    var pxToEmu = function pxToEmu2(px) {
      return pxToEmu_(px, image.dpi);
    };
    xml += "<xdr:oneCellAnchor>";
    xml += "<xdr:from>";
    xml += "<xdr:col>".concat(image.anchor.column - 1, "</xdr:col>");
    xml += "<xdr:colOff>".concat(typeof image.offsetX === "number" ? pxToEmu(image.offsetX) : 0, "</xdr:colOff>");
    xml += "<xdr:row>".concat(image.anchor.row - 1, "</xdr:row>");
    xml += "<xdr:rowOff>".concat(typeof image.offsetY === "number" ? pxToEmu(image.offsetY) : 0, "</xdr:rowOff>");
    xml += "</xdr:from>";
    xml += '<xdr:ext cx="'.concat(pxToEmu(image.width), '" cy="').concat(pxToEmu(image.height), '"/>');
    xml += "<xdr:pic>";
    xml += "<xdr:nvPicPr>";
    xml += '<xdr:cNvPr id="'.concat(imageId, '" name="').concat(image.title ? sanitizeAttributeValue(image.title) : "Picture " + imageId, '" descr="').concat(image.description ? sanitizeAttributeValue(image.description) : "", '"/>');
    xml += "<xdr:cNvPicPr>";
    xml += '<a:picLocks noChangeAspect="1"/>';
    xml += "</xdr:cNvPicPr>";
    xml += "</xdr:nvPicPr>";
    xml += "<xdr:blipFill>";
    xml += '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId-image-'.concat(imageId, '" cstate="print"/>');
    xml += "<a:stretch>";
    xml += "<a:fillRect/>";
    xml += "</a:stretch>";
    xml += "</xdr:blipFill>";
    xml += "<xdr:spPr>";
    xml += '<a:prstGeom prst="rect">';
    xml += "<a:avLst/>";
    xml += "</a:prstGeom>";
    xml += "</xdr:spPr>";
    xml += "</xdr:pic>";
    xml += "<xdr:clientData/>";
    xml += "</xdr:oneCellAnchor>";
    i++;
  };
  for (var _iterator3 = _createForOfIteratorHelperLoose10(images), _step3; !(_step3 = _iterator3()).done; ) {
    _loop();
  }
  return xml;
}
function getImagesDrawingXmlRels(_ref6) {
  var images = _ref6.images, sheetIndex = _ref6.sheetIndex;
  return images.map(function(image, i) {
    var imageId = i + 1;
    return '<Relationship Id="rId-image-'.concat(imageId, '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/').concat(getImageFileName(image, {
      sheetIndex,
      imageIndex: i
    }), '"/>');
  }).join("");
}
var DEFAULT_DISPLAY_DPI = 96;
var DEFAULT_IMAGE_DPI = 96;
function pxToEmu_(px, imageDpi) {
  var displayDpi = DEFAULT_DISPLAY_DPI;
  return Math.round(px * 9525 * (DEFAULT_DISPLAY_DPI / displayDpi) * (DEFAULT_IMAGE_DPI / imageDpi));
}

// src/write/xml/replaceElement.js
function replaceElement(xml, element, replacementXml) {
  if (element.selfClosingTag) {
    return xml.slice(0, element.openingTagStartIndex) + replacementXml + xml.slice(element.openingTagEndIndex + 1);
  }
  return xml.slice(0, element.openingTagStartIndex) + replacementXml + xml.slice(element.closingTagEndIndex + 1);
}

// src/write/xml/findElementInsideElement.js
function findElementInsideElement(xml, tagName, enclosingElement) {
  var element = findElement(getMarkupInsideElement(xml, enclosingElement), tagName);
  if (element) {
    applyEnclosingElementOffset(element, enclosingElement);
    return element;
  }
}

// src/write/xlsx/features/stickyRowsOrColumns.js
var stickyRowsOrColumns_default = {
  files: {
    transform: {
      "xl/worksheets/sheet{id}.xml": {
        transform: function transform2(xml, sheetOptions, _ref) {
          var sheetIndex = _ref.sheetIndex, sheetId = _ref.sheetId;
          if (hasStickyRowsOrColumns(sheetOptions)) {
            var stickyRowsCount = sheetOptions.stickyRowsCount, stickyColumnsCount = sheetOptions.stickyColumnsCount;
            var paneAttributes = {
              // The vertical position of the split, in 1/20th of a point (twips),
              // or, when frozen, the number of visible rows in the top pane(s).
              ySplit: stickyRowsCount || 0,
              // The horizontal position of the split, in 1/20th of a point (twips),
              // or, when frozen, the number of visible columns in the left pane(s).
              xSplit: stickyColumnsCount || 0,
              // The visible cell in the top-left corner of the bottom-right pane.
              topLeftCell: getCellAddress(stickyRowsCount || 0, stickyColumnsCount || 0),
              // `activePane` defines which pane is active in a "split" or "frozen" `state`.
              // Possible values: "bottomLeft", "bottomRight", "topLeft", "topRight".
              activePane: "bottomRight",
              // `state: "frozen"` indicates that the panes are frozen, meaning the split area
              // (usually top rows or leftmost columns) remains fixed while the rest of the sheet scrolls.
              // Other possible values:
              // "split" — Indicates that the window is split into separate panes that can be scrolled independently, but are not locked in place.
              // "frozenSplit" — A combination used in specific scenarios (often by Excel itself) to manage complex freezing.
              state: "frozen"
            };
            var sheetViewElement = findElement(xml, "sheetView");
            if (sheetViewElement) {
              var paneElement = findElementInsideElement(xml, "pane", sheetViewElement);
              var paneXml = getSelfClosingTagMarkup("pane", paneAttributes);
              if (paneElement) {
                xml = replaceElement(xml, paneElement, paneXml);
              } else {
                xml = appendMarkupInsideElement(xml, sheetViewElement, paneXml);
              }
            } else {
              var sheetViewsElement = findElement(xml, "sheetViews");
              if (sheetViewsElement) {
                throw new Error("xl/worksheets/sheet".concat(sheetId, ".xml: <sheetViews/> element exists but it doesn't contain any <sheetView/> elements"));
              }
              var sheetViewAttributes = {
                tabSelected: sheetIndex === 0 ? 1 : 0,
                // The first sheet is selected by default.
                workbookViewId: 0
              };
              var sheetViewsXml = getOpeningTagMarkup("sheetViews") + getOpeningTagMarkup("sheetView", sheetViewAttributes) + getSelfClosingTagMarkup("pane", paneAttributes) + getClosingTagMarkup("sheetView", sheetViewAttributes) + getClosingTagMarkup("sheetViews");
              var worksheetElement = findElement(xml, "worksheet");
              xml = prependMarkupInsideElement(xml, worksheetElement, sheetViewsXml);
            }
          }
          return xml;
        }
      },
      "xl/workbook.xml": {
        transform: function transform3(xml, sheetsOptions) {
          if (sheetsOptions.some(hasStickyRowsOrColumns)) {
            var workbookViewElement = findElement(xml, "workbookView");
            if (!workbookViewElement) {
              var bookViewsElement = findElement(xml, "bookViews");
              if (bookViewsElement) {
                throw new Error("xl/workbook.xml: <bookViews/> element exists but it doesn't contain any <workbookView/> elements");
              }
              return insertElementMarkupAccordingToOrderOfSiblings(xml, "<bookViews><workbookView/></bookViews>", getOrderOfSiblings("xl/workbook.xml", "workbook"), "workbook");
            }
          }
          return xml;
        }
      }
    }
  }
};
function hasStickyRowsOrColumns(sheetOptions) {
  return Boolean(sheetOptions.stickyRowsCount) || Boolean(sheetOptions.stickyColumnsCount);
}

// src/write/xlsx/features/rowBreaks.js
var LAST_COLUMN_INDEX = 16383;
var rowBreaks_default = {
  files: {
    transform: {
      "xl/worksheets/sheet{id}.xml": {
        // Placed by the schema order of `<worksheet/>` children (after `<pageSetup/>` and `<headerFooter/>`,
        // before `<drawing/>`), and `<drawing/>` is already in the XML by the time `transform()` runs.
        transform: function transform4(xml, sheetOptions) {
          var rows = normalizeRowBreaks(sheetOptions.rowBreaks);
          if (rows.length === 0) {
            return xml;
          }
          var breaksXml = rows.map(function(row) {
            return '<brk id="' + row + '" max="' + LAST_COLUMN_INDEX + '" man="1"/>';
          }).join("");
          var elementXml = '<rowBreaks count="' + rows.length + '" manualBreakCount="' + rows.length + '">' + breaksXml + "</rowBreaks>";
          return insertElementMarkupAccordingToOrderOfSiblings(xml, elementXml, getOrderOfSiblings("xl/worksheets/sheet{id}.xml", "worksheet"), "worksheet");
        }
      }
    }
  }
};
function normalizeRowBreaks(rowBreaks) {
  if (rowBreaks === void 0 || rowBreaks === null) {
    return [];
  }
  if (!Array.isArray(rowBreaks)) {
    throw new Error("`rowBreaks` must be an array of 1-based row numbers");
  }
  rowBreaks.forEach(function(row) {
    if (typeof row !== "number" || !Number.isInteger(row) || row < 1) {
      throw new Error("`rowBreaks` must hold positive integers (1-based row numbers), got: " + String(row));
    }
  });
  return rowBreaks.filter(function(row, i) {
    return rowBreaks.indexOf(row) === i;
  }).sort(function(a, b) {
    return a - b;
  });
}

// src/write/xlsx/getFeatures.js
function getFeatures(customFeatures) {
  var features = [];
  features.push(conditionalFormatting_default);
  features.push(images_default);
  features.push(stickyRowsOrColumns_default);
  features.push(rowBreaks_default);
  if (customFeatures) {
    customFeatures.forEach(validateFeature);
    features = features.concat(customFeatures);
  }
  return features;
}
var TRANSFORMABLE_FILES = ["[Content_Types].xml", "_rels/.rels", "xl/styles.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet{id}.xml", "xl/worksheets/_rels/sheet{id}.xml.rels", "xl/drawings/drawing{id}.xml", "xl/drawings/_rels/drawing{id}.xml.rels"];
function validateFeature(feature) {
  if (feature.files && feature.files.transform) {
    for (var _i = 0, _Object$keys = Object.keys(feature.files.transform); _i < _Object$keys.length; _i++) {
      var key = _Object$keys[_i];
      if (TRANSFORMABLE_FILES.indexOf(key) < 0) {
        throw new Error("Unknown file to transform: ".concat(key));
      }
    }
  }
}

// src/write/xlsx/validateSheetData.js
function _createForOfIteratorHelperLoose11(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray13(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray13(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray13(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray13(r, a) : void 0;
  }
}
function _arrayLikeToArray13(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function validateSheetData(data) {
  if (!Array.isArray(data)) {
    throw new TypeError("Expected sheet data to be an array of rows");
  }
  for (var _iterator = _createForOfIteratorHelperLoose11(data), _step; !(_step = _iterator()).done; ) {
    var row = _step.value;
    if (Array.isArray(row)) {
      return;
    } else {
      throw new Error("Expected each sheet data row to be an array");
    }
  }
}

// src/write/xlsx/initializeSharedStrings.js
function initializeSharedStrings() {
  var sharedStrings = [];
  var sharedStringIdByString = {};
  return {
    getSharedStrings: function getSharedStrings() {
      return sharedStrings;
    },
    findOrCreateSharedString: function findOrCreateSharedString(string) {
      var id = sharedStringIdByString[string];
      if (id === void 0) {
        id = sharedStrings.length;
        sharedStringIdByString[string] = id;
        sharedStrings.push(string);
      }
      return id;
    }
  };
}

// src/write/xlsx/validateSheetName.js
var ILLEGAL_CHARACTERS_IN_SHEET_NAME = /[\[\]\/\\:*?]+/;
function validateSheetName(sheetName) {
  if (!sheetName) {
    throw new Error("Sheet name can't be empty");
  }
  if (sheetName.length > 31) {
    throw new Error('Sheet name "'.concat(sheetName, `" can't be longer than 31 characters`));
  }
  if (ILLEGAL_CHARACTERS_IN_SHEET_NAME.test(sheetName)) {
    throw new Error('Sheet name "'.concat(sheetName, '" contains illegal characters: []/\\:*?'));
  }
}

// src/write/xlsx/initializeSheets.js
function _createForOfIteratorHelperLoose12(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray14(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray14(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray14(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray14(r, a) : void 0;
  }
}
function _arrayLikeToArray14(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function initializeSheets(sheetsData, sheetsOptions, globalOptions) {
  var _initializeSharedStri = initializeSharedStrings(), getSharedStrings = _initializeSharedStri.getSharedStrings, findOrCreateSharedString = _initializeSharedStri.findOrCreateSharedString;
  var defaultFont = getDefaultFont(globalOptions);
  var _initializeStyles = initializeStyles(defaultFont), getCellStyles = _initializeStyles.getCellStyles, findOrCreateCellStyle = _initializeStyles.findOrCreateCellStyle;
  var sheetNames = sheetsOptions.map(function(sheetOptions) {
    return sheetOptions.sheet;
  });
  for (var _iterator = _createForOfIteratorHelperLoose12(sheetNames), _step; !(_step = _iterator()).done; ) {
    var sheetName = _step.value;
    validateSheetName(sheetName);
  }
  var sheetXmlParameters = [];
  var sheetIndex = 0;
  while (sheetIndex < sheetNames.length) {
    sheetXmlParameters.push({
      sheetData: sheetsData[sheetIndex],
      sheetOptions: sheetsOptions[sheetIndex],
      sheetIndex,
      sheetId: getSheetId(sheetIndex),
      hasDefaultFont: Boolean(defaultFont),
      // hasDefaultFont: Boolean(sheetsDefaultFonts[sheetIndex]),
      // findOrCreateCellStyle: (style) => findOrCreateCellStyle(style, sheetIndex),
      findOrCreateCellStyle,
      findOrCreateSharedString
    });
    sheetIndex++;
  }
  return {
    sheets: sheetNames.map(function(sheetName2, _i) {
      var sheetIndex2 = _i;
      return {
        sheetId: getSheetId(sheetIndex2),
        sheetName: sheetName2,
        sheetXmlParameters: sheetXmlParameters[sheetIndex2]
      };
    }),
    getSharedStrings,
    getCellStyles
  };
}
function getSheetId(sheetIndex) {
  return String(sheetIndex + 1);
}
function getDefaultFont(globalOptions) {
  var fontFamily = globalOptions.fontFamily, fontSize = globalOptions.fontSize;
  if (fontFamily || typeof fontSize === "number") {
    return {
      fontFamily,
      fontSize
    };
  }
}

// src/write/getSheetData/getSheetData.js
function getSheetData(objects, columns) {
  var headerRow;
  if (columns.some(function(column) {
    return column.header !== void 0 && column.header !== null && column.header !== false;
  })) {
    headerRow = columns.map(function(_ref) {
      var header = _ref.header;
      return header !== void 0 && header !== null && header !== false ? header : null;
    });
  }
  return (headerRow ? [headerRow] : []).concat(objects.map(function(object, objectIndex) {
    return columns.map(function(_ref2) {
      var cell = _ref2.cell;
      return cell(object, objectIndex);
    });
  }));
}

// src/write/xlsx/helpers/features/getWrittenFiles.js
function _typeof4(o) {
  "@babel/helpers - typeof";
  return _typeof4 = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o2) {
    return typeof o2;
  } : function(o2) {
    return o2 && "function" == typeof Symbol && o2.constructor === Symbol && o2 !== Symbol.prototype ? "symbol" : typeof o2;
  }, _typeof4(o);
}
function ownKeys4(e, r) {
  var t = Object.keys(e);
  if (Object.getOwnPropertySymbols) {
    var o = Object.getOwnPropertySymbols(e);
    r && (o = o.filter(function(r2) {
      return Object.getOwnPropertyDescriptor(e, r2).enumerable;
    })), t.push.apply(t, o);
  }
  return t;
}
function _objectSpread4(e) {
  for (var r = 1; r < arguments.length; r++) {
    var t = null != arguments[r] ? arguments[r] : {};
    r % 2 ? ownKeys4(Object(t), true).forEach(function(r2) {
      _defineProperty4(e, r2, t[r2]);
    }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys4(Object(t)).forEach(function(r2) {
      Object.defineProperty(e, r2, Object.getOwnPropertyDescriptor(t, r2));
    });
  }
  return e;
}
function _defineProperty4(e, r, t) {
  return (r = _toPropertyKey4(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: true, configurable: true, writable: true }) : e[r] = t, e;
}
function _toPropertyKey4(t) {
  var i = _toPrimitive4(t, "string");
  return "symbol" == _typeof4(i) ? i : i + "";
}
function _toPrimitive4(t, r) {
  if ("object" != _typeof4(t) || !t) return t;
  var e = t[Symbol.toPrimitive];
  if (void 0 !== e) {
    var i = e.call(t, r || "default");
    if ("object" != _typeof4(i)) return i;
    throw new TypeError("@@toPrimitive must return a primitive value.");
  }
  return ("string" === r ? String : Number)(t);
}
function _createForOfIteratorHelperLoose13(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray15(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray15(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray15(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray15(r, a) : void 0;
  }
}
function _arrayLikeToArray15(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function getWrittenFiles(features, sheetsOptions, _ref) {
  var read2 = _ref.read;
  var writtenFiles = {};
  for (var _iterator = _createForOfIteratorHelperLoose13(features), _step; !(_step = _iterator()).done; ) {
    var feature = _step.value;
    if (feature.files && feature.files.write) {
      if (feature.files.write.files) {
        var files2 = feature.files.write.files(sheetsOptions, {
          read: read2
        });
        if (files2) {
          writtenFiles = _objectSpread4(_objectSpread4({}, writtenFiles), files2);
        }
      }
    }
  }
  return writtenFiles;
}

// src/write/xlsx/generateXlsxFileContents.js
function _typeof5(o) {
  "@babel/helpers - typeof";
  return _typeof5 = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function(o2) {
    return typeof o2;
  } : function(o2) {
    return o2 && "function" == typeof Symbol && o2.constructor === Symbol && o2 !== Symbol.prototype ? "symbol" : typeof o2;
  }, _typeof5(o);
}
var _excluded = ["data"];
function ownKeys5(e, r) {
  var t = Object.keys(e);
  if (Object.getOwnPropertySymbols) {
    var o = Object.getOwnPropertySymbols(e);
    r && (o = o.filter(function(r2) {
      return Object.getOwnPropertyDescriptor(e, r2).enumerable;
    })), t.push.apply(t, o);
  }
  return t;
}
function _objectSpread5(e) {
  for (var r = 1; r < arguments.length; r++) {
    var t = null != arguments[r] ? arguments[r] : {};
    r % 2 ? ownKeys5(Object(t), true).forEach(function(r2) {
      _defineProperty5(e, r2, t[r2]);
    }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys5(Object(t)).forEach(function(r2) {
      Object.defineProperty(e, r2, Object.getOwnPropertyDescriptor(t, r2));
    });
  }
  return e;
}
function _defineProperty5(e, r, t) {
  return (r = _toPropertyKey5(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: true, configurable: true, writable: true }) : e[r] = t, e;
}
function _toPropertyKey5(t) {
  var i = _toPrimitive5(t, "string");
  return "symbol" == _typeof5(i) ? i : i + "";
}
function _toPrimitive5(t, r) {
  if ("object" != _typeof5(t) || !t) return t;
  var e = t[Symbol.toPrimitive];
  if (void 0 !== e) {
    var i = e.call(t, r || "default");
    if ("object" != _typeof5(i)) return i;
    throw new TypeError("@@toPrimitive must return a primitive value.");
  }
  return ("string" === r ? String : Number)(t);
}
function _objectWithoutProperties(e, t) {
  if (null == e) return {};
  var o, r, i = _objectWithoutPropertiesLoose(e, t);
  if (Object.getOwnPropertySymbols) {
    var n = Object.getOwnPropertySymbols(e);
    for (r = 0; r < n.length; r++) o = n[r], -1 === t.indexOf(o) && {}.propertyIsEnumerable.call(e, o) && (i[o] = e[o]);
  }
  return i;
}
function _objectWithoutPropertiesLoose(r, e) {
  if (null == r) return {};
  var t = {};
  for (var n in r) if ({}.hasOwnProperty.call(r, n)) {
    if (-1 !== e.indexOf(n)) continue;
    t[n] = r[n];
  }
  return t;
}
function _createForOfIteratorHelperLoose14(r, e) {
  var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"];
  if (t) return (t = t.call(r)).next.bind(t);
  if (Array.isArray(r) || (t = _unsupportedIterableToArray16(r)) || e && r && "number" == typeof r.length) {
    t && (r = t);
    var o = 0;
    return function() {
      return o >= r.length ? { done: true } : { done: false, value: r[o++] };
    };
  }
  throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.");
}
function _unsupportedIterableToArray16(r, a) {
  if (r) {
    if ("string" == typeof r) return _arrayLikeToArray16(r, a);
    var t = {}.toString.call(r).slice(8, -1);
    return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray16(r, a) : void 0;
  }
}
function _arrayLikeToArray16(r, a) {
  (null == a || a > r.length) && (a = r.length);
  for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e];
  return n;
}
function generateXlsxFileContents(arg1, arg2, arg3) {
  var _getArguments = getArguments(arg1, arg2, arg3), sheetsDataAndOptions = _getArguments.sheets, options = _getArguments.options;
  var sheetsData = sheetsDataAndOptions.map(function(sheet) {
    return sheet.data;
  });
  var sheetsOptions = sheetsDataAndOptions.map(function(sheet) {
    return sheet.options;
  });
  sheetsOptions = setSheetNames(sheetsOptions);
  sheetsOptions = setConditionalFormattingRulesGlobalIndexes(sheetsOptions);
  var features = getFeatures(options.features);
  var files2 = {};
  var readFile = function readFile2(path) {
    return files2[path];
  };
  var writeFiles = function writeFiles2(filesToWrite) {
    for (var _i = 0, _Object$keys = Object.keys(filesToWrite); _i < _Object$keys.length; _i++) {
      var path = _Object$keys[_i];
      validateFilePath(path);
      validateFileContent(path, filesToWrite[path]);
      files2[path] = filesToWrite[path];
    }
  };
  var _initializeSheets = initializeSheets(sheetsData, sheetsOptions, options), sheets = _initializeSheets.sheets, getSharedStrings = _initializeSheets.getSharedStrings, getCellStyles = _initializeSheets.getCellStyles;
  files2["[Content_Types].xml"] = generateContentTypesXml({
    sheetIds: sheets.map(function(_) {
      return _.sheetId;
    }),
    features,
    sheetsOptions
  });
  files2["_rels/.rels"] = generateRelsXml({
    features,
    sheetsOptions
  });
  files2["xl/_rels/workbook.xml.rels"] = generateWorkbookXmlRels({
    sheetIds: sheets.map(function(_) {
      return _.sheetId;
    }),
    features,
    sheetsOptions
  });
  files2["xl/workbook.xml"] = generateWorkbookXml({
    sheetIdsAndNames: sheets.map(function(_ref) {
      var sheetId2 = _ref.sheetId, sheetName = _ref.sheetName;
      return {
        sheetId: sheetId2,
        sheetName
      };
    }),
    features,
    sheetsOptions
  });
  var sheetIndex = 0;
  for (var _iterator = _createForOfIteratorHelperLoose14(sheets), _step; !(_step = _iterator()).done; ) {
    var _step$value = _step.value, sheetId = _step$value.sheetId, sheetXmlParameters = _step$value.sheetXmlParameters;
    var sheetOptions = sheetsOptions[sheetIndex];
    files2["xl/worksheets/sheet".concat(sheetId, ".xml")] = generateSheetXml(sheetXmlParameters, features);
    files2["xl/worksheets/_rels/sheet".concat(sheetId, ".xml.rels")] = generateSheetXmlRels({
      sheetIndex,
      sheetId,
      sheetOptions,
      features
    });
    files2["xl/drawings/drawing".concat(sheetId, ".xml")] = generateDrawingXml({
      sheetIndex,
      sheetId,
      sheetOptions,
      features
    });
    files2["xl/drawings/_rels/drawing".concat(sheetId, ".xml.rels")] = generateDrawingXmlRels({
      sheetIndex,
      sheetId,
      sheetOptions,
      features
    });
    sheetIndex++;
  }
  writeFiles(getWrittenFiles(features, sheetsOptions, {
    read: readFile
  }));
  files2["xl/styles.xml"] = generateStylesXml(getCellStyles(), sheetsOptions, features);
  files2["xl/sharedStrings.xml"] = generateSharedStringsXml(getSharedStrings());
  removeUnusedDrawings(files2, sheets);
  return files2;
}
function removeUnusedDrawings(files2, sheets) {
  for (var _iterator2 = _createForOfIteratorHelperLoose14(sheets), _step2; !(_step2 = _iterator2()).done; ) {
    var sheetId = _step2.value.sheetId;
    if (files2["xl/drawings/drawing".concat(sheetId, ".xml")] === DRAWING_XML_START + DRAWING_XML_END) {
      delete files2["xl/drawings/drawing".concat(sheetId, ".xml")];
      delete files2["xl/drawings/_rels/drawing".concat(sheetId, ".xml.rels")];
      if (!removeSubstring(files2, "xl/worksheets/sheet".concat(sheetId, ".xml"), generateDrawingReference(getSelfClosingTagMarkup))) {
        throw new Error(COULD_NOT_REMOVE_UNUSED_DRAWINGS);
      }
      removeSubstring(files2, "xl/worksheets/_rels/sheet".concat(sheetId, ".xml.rels"), getDrawingRelationshipXml(sheetId));
      removeSubstring(files2, "[Content_Types].xml", getDrawingContentTypeXml(sheetId));
    }
  }
}
var COULD_NOT_REMOVE_UNUSED_DRAWINGS = "Couldn't remove unused drawings";
function removeSubstring(files2, key, substring) {
  if (!files2[key]) {
    throw new Error("File not found: ".concat(key));
  }
  if (files2[key].indexOf(substring) < 0) {
    throw new Error('Substring "'.concat(substring, '" not found in "').concat(key, '"'));
  }
  var stringBeforeRemoval = files2[key];
  var stringAfterRemoval = stringBeforeRemoval.replace(substring, "");
  files2[key] = stringAfterRemoval;
  return stringBeforeRemoval !== stringAfterRemoval;
}
function validateFilePath(path) {
  if (path[0] === "/") {
    throw new Error("File path must not start with a slash (/)");
  }
}
function validateFileContent(path, content) {
  if (!content) {
    throw new Error("File `content` not specified: ".concat(path));
  }
}
function getArguments(arg1, arg2, arg3) {
  if (Array.isArray(arg1)) {
    if (arg1.length === 0 || Array.isArray(arg1[0])) {
      validateSheetData(arg1);
      if (Array.isArray(arg1[0]) && arg1.length > 0 && Array.isArray(arg1[0][0])) {
        throw new Error("In order to write multiple sheets, pass an array of sheet objects");
      }
      return {
        sheets: [{
          data: arg1,
          options: arg2 || {}
        }],
        options: arg3 || {}
      };
    } else if (isObject2(arg1[0])) {
      if (isObject2(arg2) && Array.isArray(arg2.schema)) {
        throw new Error("`schema` parameter was removed, use `columns` parameter instead");
      }
      if (isObject2(arg2) && Array.isArray(arg2.columns)) {
        return getArguments(getSheetData(arg1, arg2.columns), arg2, arg3);
      }
      return {
        sheets: arg1.map(function(_ref2) {
          var data = _ref2.data, options = _objectWithoutProperties(_ref2, _excluded);
          if (!data) {
            throw new Error("`data` property is required for each sheet");
          }
          validateSheetData(data);
          return {
            data,
            options: options || {}
          };
        }),
        options: arg2 || {}
      };
    } else {
      throw new Error("Invalid first argument: must be either sheet data \u2014 an array of arrays \u2014 or an array of sheet objects");
    }
  } else {
    throw new Error("Invalid first argument: must be an array");
  }
}
function setSheetNames(sheetsOptions) {
  return sheetsOptions.map(function(sheetOptions, sheetIndex) {
    return _objectSpread5(_objectSpread5({}, sheetOptions), {}, {
      sheet: sheetOptions.sheet || "Sheet".concat(sheetIndex + 1)
    });
  });
}
function setConditionalFormattingRulesGlobalIndexes(sheetsOptions) {
  var globalIndex = 0;
  return sheetsOptions.map(function(sheetOptions) {
    return _objectSpread5(_objectSpread5({}, sheetOptions), {}, {
      conditionalFormatting: sheetOptions.conditionalFormatting && sheetOptions.conditionalFormatting.map(function(conditionalFormattingRule) {
        return _objectSpread5(_objectSpread5({}, conditionalFormattingRule), {}, {
          // `++` operator returns the value before the increment.
          _globalIndex: globalIndex++
        });
      })
    });
  });
}

// src/write/export/downloadBlob.js
function downloadBlob(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.style.display = "none";
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(function() {
    URL.revokeObjectURL(url);
    document.body.removeChild(a);
  }, 100);
}

// src/sync/writeXlsxFileSync.js
var XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
function toUint8Array(path, content) {
  if (typeof content === "string") {
    return strToU8(content);
  }
  if (content instanceof Uint8Array) {
    return content;
  }
  if (typeof ArrayBuffer !== "undefined" && content instanceof ArrayBuffer) {
    return new Uint8Array(content);
  }
  throw new Error(`${path}: a file inside the .xlsx given as a Blob or File can't be read synchronously; pass an ArrayBuffer or a Uint8Array`);
}
function writeXlsxFileSync(arg1, arg2, arg3) {
  const toUint8ArrayNow = () => {
    const files2 = generateXlsxFileContents(arg1, arg2, arg3);
    const entries = {};
    for (const path of Object.keys(files2)) {
      entries[path] = toUint8Array(path, files2[path]);
    }
    return zipSync(entries);
  };
  const toBlob = () => {
    try {
      return Promise.resolve(new Blob([toUint8ArrayNow()], { type: XLSX_MIME_TYPE }));
    } catch (error) {
      return Promise.reject(error);
    }
  };
  return {
    toUint8Array: toUint8ArrayNow,
    toBlob,
    toFile: (fileName) => toBlob().then((blob) => {
      downloadBlob(blob, fileName);
    })
  };
}

// src/utils/loader.js
var fs = null;
var FS_MODULE = "fs";
async function loadFs() {
  if (fs) return;
  try {
    fs = await import(FS_MODULE);
  } catch {
    fs = { existsSync: () => false, readFileSync: () => {
      throw new Error("File system access is not available");
    } };
  }
}
async function loadAsBuffer(input) {
  if (input instanceof Uint8Array) return input;
  if (typeof ArrayBuffer !== "undefined" && input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  if (typeof input === "string") {
    await loadFs();
    if (input.length < 4096 && !input.includes("\n") && !input.includes("\r")) {
      try {
        if (fs.existsSync(input)) {
          return new Uint8Array(fs.readFileSync(input));
        }
      } catch {
      }
    }
    return new TextEncoder().encode(input);
  }
  if (typeof File !== "undefined" && input instanceof File || typeof Blob !== "undefined" && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  if (input && typeof input.pipe === "function" && typeof input.on === "function") {
    return new Promise((resolve, reject) => {
      const chunks = [];
      input.on("data", (chunk) => chunks.push(chunk));
      input.on("end", () => {
        const total = typeof Buffer !== "undefined" ? Buffer.concat(chunks) : new Uint8Array();
        resolve(new Uint8Array(total.buffer, total.byteOffset, total.byteLength));
      });
      input.on("error", reject);
    });
  }
  throw new Error("Invalid input: expected file path, File, Blob, Uint8Array, ArrayBuffer, Buffer, or Stream");
}
async function parse(parserFn, input, options = {}) {
  try {
    const result = await parserFn(input, options);
    if (options.onload) options.onload(result);
    return result;
  } catch (error) {
    if (options.onerror) options.onerror(error);
    throw error;
  }
}

// src/utils/cfb.js
var HEADER_SIGNATURE = [208, 207, 17, 224, 161, 177, 26, 225];
var HEADER_CLSID_NULL = new Array(16).fill(0);
var MAXREGSECT = 4294967290;
var ENTRY_TYPES = {
  UNKNOWN: 0,
  STORAGE: 1,
  STREAM: 2,
  ROOT: 5
};
function parseHeader(buffer) {
  for (let i = 0; i < 8; i++) {
    if (buffer[i] !== HEADER_SIGNATURE[i]) {
      throw new Error("Invalid CFB signature");
    }
  }
  const minorVersion = readUInt16LE(buffer, 24);
  const majorVersion = readUInt16LE(buffer, 26);
  const byteOrder = readUInt16LE(buffer, 28);
  if (byteOrder !== 65534) {
    throw new Error("Invalid byte order marker");
  }
  const sectorShift = readUInt16LE(buffer, 30);
  const miniSectorShift = readUInt16LE(buffer, 32);
  const sectorSize = 1 << sectorShift;
  const miniSectorSize = 1 << miniSectorShift;
  const totalSectors = readUInt32LE(buffer, 40);
  const fatSectors = readUInt32LE(buffer, 44);
  const firstDirSector = readUInt32LE(buffer, 48);
  const transactionSignature = readUInt32LE(buffer, 52);
  const miniStreamCutoff = readUInt32LE(buffer, 56);
  const firstMiniFatSector = readUInt32LE(buffer, 60);
  const miniFatSectors = readUInt32LE(buffer, 64);
  const firstDifatSector = readUInt32LE(buffer, 68);
  const difatSectors = readUInt32LE(buffer, 72);
  const difat = [];
  for (let i = 0; i < 109; i++) {
    const sector = readUInt32LE(buffer, 76 + i * 4);
    if (sector < MAXREGSECT) {
      difat.push(sector);
    }
  }
  return {
    minorVersion,
    majorVersion,
    sectorShift,
    miniSectorShift,
    sectorSize,
    miniSectorSize,
    totalSectors,
    fatSectors,
    firstDirSector,
    transactionSignature,
    miniStreamCutoff,
    firstMiniFatSector,
    miniFatSectors,
    firstDifatSector,
    difatSectors,
    difat
  };
}
function readSector(buffer, sectorId, sectorSize) {
  const offset = (sectorId + 1) * sectorSize;
  if (offset + sectorSize > buffer.length) {
    throw new Error(`Sector ${sectorId} out of bounds`);
  }
  return buffer.slice(offset, offset + sectorSize);
}
function buildFAT(buffer, header) {
  const { sectorSize, difat, firstDifatSector, difatSectors } = header;
  const entriesPerSector = sectorSize / 4;
  const fat = [];
  let fatSectors = [...difat];
  let difatSector = firstDifatSector;
  for (let i = 0; i < difatSectors && difatSector < MAXREGSECT; i++) {
    const sector = readSector(buffer, difatSector, sectorSize);
    for (let j = 0; j < entriesPerSector - 1; j++) {
      const fatSectorNum = readUInt32LE(sector, j * 4);
      if (fatSectorNum < MAXREGSECT) {
        fatSectors.push(fatSectorNum);
      }
    }
    difatSector = readUInt32LE(sector, (entriesPerSector - 1) * 4);
  }
  for (const fatSectorNum of fatSectors) {
    const sector = readSector(buffer, fatSectorNum, sectorSize);
    for (let i = 0; i < entriesPerSector; i++) {
      fat.push(readUInt32LE(sector, i * 4));
    }
  }
  return fat;
}
function buildMiniFAT(buffer, header, fat) {
  const { firstMiniFatSector, miniFatSectors, sectorSize } = header;
  if (firstMiniFatSector >= MAXREGSECT || miniFatSectors === 0) {
    return [];
  }
  const entriesPerSector = sectorSize / 4;
  const miniFat = [];
  let sector = firstMiniFatSector;
  let count = 0;
  while (sector < MAXREGSECT && count < miniFatSectors) {
    const sectorData = readSector(buffer, sector, sectorSize);
    for (let i = 0; i < entriesPerSector; i++) {
      miniFat.push(readUInt32LE(sectorData, i * 4));
    }
    sector = fat[sector];
    count++;
  }
  return miniFat;
}
function readChain(buffer, startSector, fat, sectorSize, maxSize = Infinity) {
  if (startSector >= MAXREGSECT) {
    return new Uint8Array(0);
  }
  const chunks = [];
  let sector = startSector;
  let totalSize = 0;
  const maxChainSectors = Math.ceil(buffer.length / sectorSize) + 1;
  while (sector < MAXREGSECT && totalSize < maxSize) {
    const sectorData = readSector(buffer, sector, sectorSize);
    chunks.push(sectorData);
    totalSize += sectorSize;
    sector = fat[sector];
    if (chunks.length > maxChainSectors) {
      throw new Error("FAT chain too long - possible corruption");
    }
  }
  const result = new Uint8Array(Math.min(totalSize, maxSize));
  let offset = 0;
  for (const chunk of chunks) {
    const copySize = Math.min(chunk.length, maxSize - offset);
    result.set(chunk.slice(0, copySize), offset);
    offset += copySize;
    if (offset >= maxSize) break;
  }
  return result;
}
function readMiniChain(miniStream, startSector, miniFat, miniSectorSize, size) {
  if (startSector >= MAXREGSECT) {
    return new Uint8Array(0);
  }
  size = Math.min(size, miniStream.length);
  const result = new Uint8Array(size);
  let sector = startSector;
  let offset = 0;
  while (sector < MAXREGSECT && offset < size) {
    const sectorOffset = sector * miniSectorSize;
    const copySize = Math.min(miniSectorSize, size - offset, miniStream.length - sectorOffset);
    if (sectorOffset + copySize > miniStream.length) {
      break;
    }
    result.set(miniStream.slice(sectorOffset, sectorOffset + copySize), offset);
    offset += copySize;
    sector = miniFat[sector];
    if (offset > size * 2) {
      throw new Error("MiniFAT chain too long - possible corruption");
    }
  }
  return result.slice(0, offset);
}
function parseDirectoryEntry(buffer, offset) {
  let nameLength = readUInt16LE(buffer, offset + 64);
  if (nameLength > 64) nameLength = 64;
  let name = "";
  for (let i = 0; i < nameLength - 2; i += 2) {
    const charCode = readUInt16LE(buffer, offset + i);
    if (charCode === 0) break;
    name += String.fromCharCode(charCode);
  }
  const type = buffer[offset + 66];
  const color = buffer[offset + 67];
  const leftSibling = readUInt32LE(buffer, offset + 68);
  const rightSibling = readUInt32LE(buffer, offset + 72);
  const childId = readUInt32LE(buffer, offset + 76);
  const stateBits = readUInt32LE(buffer, offset + 96);
  const startSector = readUInt32LE(buffer, offset + 116);
  const sizeLow = readUInt32LE(buffer, offset + 120);
  const sizeHigh = readUInt32LE(buffer, offset + 124);
  const size = sizeHigh === 0 ? sizeLow : sizeLow + sizeHigh * 4294967296;
  return {
    name,
    type,
    color,
    leftSibling,
    rightSibling,
    childId,
    stateBits,
    startSector,
    size
  };
}
function parseDirectory(buffer, header, fat) {
  const { firstDirSector, sectorSize } = header;
  const entriesPerSector = sectorSize / 128;
  const entries = [];
  const dirData = readChain(buffer, firstDirSector, fat, sectorSize);
  for (let i = 0; i < dirData.length; i += 128) {
    if (i + 128 > dirData.length) break;
    const entry = parseDirectoryEntry(dirData, i);
    entries.push(entry);
  }
  return entries;
}
function buildPaths(entries) {
  const paths = new Array(entries.length).fill("");
  if (entries.length === 0) return paths;
  paths[0] = "/";
  function buildPath(index, parentPath) {
    if (index >= entries.length || index < 0) return;
    const entry = entries[index];
    if (!entry.name) return;
    const currentPath = parentPath === "/" ? "/" + entry.name : parentPath + "/" + entry.name;
    paths[index] = currentPath;
    if (entry.childId < entries.length && entry.childId !== index) {
      buildPath(entry.childId, currentPath);
    }
    if (entry.leftSibling < entries.length && entry.leftSibling !== index) {
      buildPath(entry.leftSibling, parentPath);
    }
    if (entry.rightSibling < entries.length && entry.rightSibling !== index) {
      buildPath(entry.rightSibling, parentPath);
    }
  }
  if (entries[0].childId < entries.length) {
    buildPath(entries[0].childId, "/");
  }
  return paths;
}
function read(buffer, options = {}) {
  if (!(buffer instanceof Uint8Array)) {
    if (Buffer.isBuffer(buffer)) {
      buffer = new Uint8Array(buffer);
    } else {
      throw new Error("Buffer must be Uint8Array or Node Buffer");
    }
  }
  if (buffer.length < 512) {
    throw new Error(`File too small: ${buffer.length} bytes`);
  }
  const header = parseHeader(buffer);
  const fat = buildFAT(buffer, header);
  const miniFat = buildMiniFAT(buffer, header, fat);
  const entries = parseDirectory(buffer, header, fat);
  const paths = buildPaths(entries);
  let miniStream = new Uint8Array(0);
  if (entries.length > 0 && entries[0].startSector < MAXREGSECT) {
    miniStream = readChain(
      buffer,
      entries[0].startSector,
      fat,
      header.sectorSize,
      entries[0].size
    );
  }
  const cfb = {
    header,
    fat,
    miniFat,
    entries,
    paths,
    miniStream,
    buffer
  };
  return cfb;
}
function find(cfb, path) {
  if (!path) return null;
  let searchPath = path.trim();
  if (!searchPath.startsWith("/")) {
    searchPath = "/" + searchPath;
  }
  searchPath = searchPath.toUpperCase();
  for (let i = 0; i < cfb.paths.length; i++) {
    const entryPath = cfb.paths[i].toUpperCase();
    if (entryPath === searchPath) {
      return createEntryObject(cfb, i);
    }
  }
  const altPath = searchPath.replace(/^\//, "");
  for (let i = 0; i < cfb.paths.length; i++) {
    const entryPath = cfb.paths[i].toUpperCase().replace(/^\//, "");
    if (entryPath === altPath) {
      return createEntryObject(cfb, i);
    }
  }
  return null;
}
function createEntryObject(cfb, index) {
  const entry = cfb.entries[index];
  const path = cfb.paths[index];
  if (entry.type !== ENTRY_TYPES.STREAM) {
    return {
      name: entry.name,
      type: entry.type,
      size: entry.size
    };
  }
  let content;
  if (entry.size < cfb.header.miniStreamCutoff && cfb.miniStream.length > 0) {
    content = readMiniChain(
      cfb.miniStream,
      entry.startSector,
      cfb.miniFat,
      cfb.header.miniSectorSize,
      entry.size
    );
  } else {
    content = readChain(
      cfb.buffer,
      entry.startSector,
      cfb.fat,
      cfb.header.sectorSize,
      entry.size
    );
  }
  return {
    name: entry.name,
    type: entry.type,
    size: entry.size,
    content
  };
}
function listEntries(cfb) {
  return cfb.entries.map((entry, index) => ({
    name: entry.name,
    path: cfb.paths[index],
    type: entry.type,
    size: entry.size
  })).filter((e) => e.name && e.type !== ENTRY_TYPES.UNKNOWN);
}
var cfb_default = { read, find, listEntries };

// src/utils/ptg-decoder.js
var PTG = {
  // Basic operand tokens (no class bits, 0x01-0x1F)
  tExp: 1,
  tTbl: 2,
  tAdd: 3,
  tSub: 4,
  tMul: 5,
  tDiv: 6,
  tPower: 7,
  tConcat: 8,
  tLT: 9,
  tLE: 10,
  tEQ: 11,
  tGE: 12,
  tGT: 13,
  tNE: 14,
  tIsect: 15,
  tUnion: 16,
  tRange: 17,
  tUplus: 18,
  tUminus: 19,
  tPercent: 20,
  tParen: 21,
  tMissArg: 22,
  tStr: 23,
  tAttr: 25,
  // Attribute token (control flow metadata)
  tErr: 28,
  tBool: 29,
  tInt: 30,
  tNum: 31,
  // Extended tokens with class bits (>= 0x20)
  // These are BASE values (after masking with 0x1F)
  tArray: 0,
  // Full values: 0x20, 0x40, 0x60
  tFunc: 1,
  // Full values: 0x21, 0x41, 0x61
  tFuncVar: 2,
  // Full values: 0x22, 0x42, 0x62
  tName: 3,
  // Full values: 0x23, 0x43, 0x63
  tRef: 4,
  // Full values: 0x24, 0x44, 0x64
  tArea: 5,
  // Full values: 0x25, 0x45, 0x65
  tMemArea: 6,
  // Full values: 0x26, 0x46, 0x66
  tMemErr: 7,
  // Full values: 0x27, 0x47, 0x67
  tRefN: 12,
  // Full values: 0x2C, 0x4C, 0x6C (relative reference in shared formulas)
  tAreaN: 13,
  // Full values: 0x2D, 0x4D, 0x6D (relative area in shared formulas)
  // Class modifiers (bits 5-6)
  VALUE: 32,
  REFERENCE: 64,
  ARRAY: 96
};
var FUNCTIONS = {
  0: "COUNT",
  1: "IF",
  2: "ISNA",
  3: "ISERROR",
  4: "SUM",
  5: "AVERAGE",
  6: "MIN",
  7: "MAX",
  8: "ROW",
  9: "COLUMN",
  10: "NA",
  11: "NPV",
  12: "STDEV",
  13: "DOLLAR",
  14: "FIXED",
  15: "SIN",
  16: "COS",
  17: "TAN",
  18: "ATAN",
  19: "PI",
  20: "SQRT",
  21: "EXP",
  22: "LN",
  23: "LOG10",
  24: "ABS",
  25: "INT",
  26: "SIGN",
  27: "ROUND",
  28: "LOOKUP",
  29: "INDEX",
  30: "REPT",
  31: "MID",
  32: "LEN",
  33: "VALUE",
  34: "TRUE",
  35: "FALSE",
  36: "AND",
  37: "OR",
  38: "NOT",
  39: "MOD",
  40: "DCOUNT",
  41: "DSUM",
  42: "DAVERAGE",
  43: "DMIN",
  44: "DMAX",
  45: "DSTDEV",
  46: "VAR",
  47: "DVAR",
  48: "TEXT",
  49: "LINEST",
  50: "TREND",
  51: "LOGEST",
  52: "GROWTH",
  53: "GOTO",
  54: "HALT",
  55: "RETURN",
  56: "PV",
  57: "FV",
  58: "NPER",
  59: "PMT",
  60: "RATE",
  61: "MIRR",
  62: "IRR",
  63: "RAND",
  64: "MATCH",
  65: "DATE",
  66: "TIME",
  67: "DAY",
  68: "MONTH",
  69: "YEAR",
  70: "WEEKDAY",
  71: "HOUR",
  72: "MINUTE",
  73: "SECOND",
  74: "NOW",
  75: "AREAS",
  76: "ROWS",
  77: "COLUMNS",
  78: "OFFSET",
  79: "ABSREF",
  80: "RELREF",
  81: "ARGUMENT",
  82: "SEARCH",
  83: "TRANSPOSE",
  84: "ERROR",
  85: "STEP",
  86: "TYPE",
  87: "ECHO",
  88: "SET.NAME",
  89: "CALLER",
  90: "DEREF",
  91: "WINDOWS",
  92: "SERIES",
  93: "DOCUMENTS",
  94: "ACTIVE.CELL",
  95: "SELECTION",
  96: "RESULT",
  97: "ATAN2",
  98: "ASIN",
  99: "ACOS",
  100: "CHOOSE",
  101: "HLOOKUP",
  102: "VLOOKUP",
  103: "LINKS",
  104: "INPUT",
  105: "ISREF",
  106: "GET.FORMULA",
  107: "GET.NAME",
  108: "SET.VALUE",
  109: "LOG",
  110: "EXEC",
  111: "CHAR",
  112: "LOWER",
  113: "UPPER",
  114: "PROPER",
  115: "LEFT",
  116: "RIGHT",
  117: "EXACT",
  118: "TRIM",
  119: "REPLACE",
  120: "SUBSTITUTE",
  121: "CODE",
  122: "NAMES",
  123: "DIRECTORY",
  124: "FIND",
  125: "CELL",
  126: "ISERR",
  127: "ISTEXT",
  128: "ISNUMBER",
  129: "ISBLANK",
  130: "T",
  131: "N",
  132: "FOPEN",
  133: "FCLOSE",
  134: "FSIZE",
  135: "FREADLN",
  136: "FREAD",
  137: "FWRITELN",
  138: "FWRITE",
  139: "FPOS",
  140: "DATEVALUE",
  141: "TIMEVALUE",
  142: "SLN",
  143: "SYD",
  144: "DDB",
  145: "GET.DEF",
  146: "REFTEXT",
  147: "TEXTREF",
  148: "INDIRECT",
  149: "REGISTER",
  150: "CALL",
  151: "ADD.BAR",
  152: "ADD.MENU",
  153: "ADD.COMMAND",
  154: "ENABLE.COMMAND",
  155: "CHECK.COMMAND",
  156: "RENAME.COMMAND",
  157: "SHOW.BAR",
  158: "DELETE.MENU",
  159: "DELETE.COMMAND",
  160: "GET.CHART.ITEM",
  161: "DIALOG.BOX",
  162: "CLEAN",
  163: "MDETERM",
  164: "MINVERSE",
  165: "MMULT",
  166: "FILES",
  167: "IPMT",
  168: "PPMT",
  169: "COUNTA",
  170: "CANCEL.KEY",
  171: "FOR",
  172: "WHILE",
  173: "BREAK",
  174: "NEXT",
  175: "INITIATE",
  176: "REQUEST",
  177: "POKE",
  178: "EXECUTE",
  179: "TERMINATE",
  180: "RESTART",
  181: "HELP",
  182: "GET.BAR",
  183: "PRODUCT",
  184: "FACT",
  185: "GET.CELL",
  186: "GET.WORKSPACE",
  187: "GET.WINDOW",
  188: "GET.DOCUMENT",
  189: "DPRODUCT",
  190: "ISNONTEXT",
  191: "GET.NOTE",
  192: "NOTE",
  193: "STDEVP",
  194: "VARP",
  195: "DSTDEVP",
  196: "DVARP",
  197: "TRUNC",
  198: "ISLOGICAL",
  199: "DCOUNTA",
  200: "DELETE.BAR",
  201: "UNREGISTER",
  204: "USDOLLAR",
  205: "FINDB",
  206: "SEARCHB",
  207: "REPLACEB",
  208: "LEFTB",
  209: "RIGHTB",
  210: "MIDB",
  211: "LENB",
  212: "ROUNDUP",
  213: "ROUNDDOWN",
  214: "ASC",
  215: "DBCS",
  216: "RANK",
  219: "ADDRESS",
  220: "DAYS360",
  221: "TODAY",
  222: "VDB",
  223: "ELSE",
  224: "ELSE.IF",
  225: "END.IF",
  226: "FOR.CELL",
  227: "MEDIAN",
  228: "SUMPRODUCT",
  229: "SINH",
  230: "COSH",
  231: "TANH",
  232: "ASINH",
  233: "ACOSH",
  234: "ATANH",
  235: "DGET",
  236: "CREATE.OBJECT",
  237: "VOLATILE",
  238: "LAST.ERROR",
  239: "CUSTOM.UNDO",
  240: "CUSTOM.REPEAT",
  241: "FORMULA.CONVERT",
  242: "GET.LINK.INFO",
  243: "TEXT.BOX",
  244: "INFO",
  245: "GROUP",
  246: "GET.OBJECT",
  247: "DB",
  248: "PAUSE",
  251: "RESUME",
  252: "FREQUENCY",
  253: "ADD.TOOLBAR",
  254: "DELETE.TOOLBAR",
  255: "User",
  256: "RESET.TOOLBAR",
  257: "EVALUATE",
  258: "GET.TOOLBAR",
  259: "GET.TOOL",
  260: "SPELLING.CHECK",
  261: "ERROR.TYPE",
  262: "APP.TITLE",
  263: "WINDOW.TITLE",
  264: "SAVE.TOOLBAR",
  265: "ENABLE.TOOL",
  266: "PRESS.TOOL",
  267: "REGISTER.ID",
  268: "GET.WORKBOOK",
  269: "AVEDEV",
  270: "BETADIST",
  271: "GAMMALN",
  272: "BETAINV",
  273: "BINOMDIST",
  274: "CHIDIST",
  275: "CHIINV",
  276: "COMBIN",
  277: "CONFIDENCE",
  278: "CRITBINOM",
  279: "EVEN",
  280: "EXPONDIST",
  281: "FDIST",
  282: "FINV",
  283: "FISHER",
  284: "FISHERINV",
  285: "FLOOR",
  286: "GAMMADIST",
  287: "GAMMAINV",
  288: "CEILING",
  289: "HYPGEOMDIST",
  290: "LOGNORMDIST",
  291: "LOGINV",
  292: "NEGBINOMDIST",
  293: "NORMDIST",
  294: "NORMSDIST",
  295: "NORMINV",
  296: "NORMSINV",
  297: "STANDARDIZE",
  298: "ODD",
  299: "PERMUT",
  300: "POISSON",
  301: "TDIST",
  302: "WEIBULL",
  303: "SUMXMY2",
  304: "SUMX2MY2",
  305: "SUMX2PY2",
  306: "CHITEST",
  307: "CORREL",
  308: "COVAR",
  309: "FORECAST",
  310: "FTEST",
  311: "INTERCEPT",
  312: "PEARSON",
  313: "RSQ",
  314: "STEYX",
  315: "SLOPE",
  316: "TTEST",
  317: "PROB",
  318: "DEVSQ",
  319: "GEOMEAN",
  320: "HARMEAN",
  321: "SUMSQ",
  322: "KURT",
  323: "SKEW",
  324: "ZTEST",
  325: "LARGE",
  326: "SMALL",
  327: "QUARTILE",
  328: "PERCENTILE",
  329: "PERCENTRANK",
  330: "MODE",
  331: "TRIMMEAN",
  332: "TINV",
  334: "MOVIE.COMMAND",
  335: "GET.MOVIE",
  336: "CONCATENATE",
  337: "POWER",
  338: "PIVOT.ADD.DATA",
  339: "GET.PIVOT.TABLE",
  340: "GET.PIVOT.FIELD",
  341: "GET.PIVOT.ITEM",
  342: "RADIANS",
  343: "DEGREES",
  344: "SUBTOTAL",
  345: "SUMIF",
  346: "COUNTIF",
  347: "COUNTBLANK",
  348: "SCENARIO.GET",
  349: "OPTIONS.LISTS.GET",
  350: "ISPMT",
  351: "DATEDIF",
  352: "DATESTRING",
  353: "NUMBERSTRING",
  354: "ROMAN",
  355: "OPEN.DIALOG",
  356: "SAVE.DIALOG",
  357: "VIEW.GET",
  358: "GETPIVOTDATA",
  359: "HYPERLINK",
  360: "PHONETIC",
  361: "AVERAGEA",
  362: "MAXA",
  363: "MINA",
  364: "STDEVPA",
  365: "VARPA",
  366: "STDEVA",
  367: "VARA",
  368: "BAHTTEXT",
  369: "THAIDAYOFWEEK",
  370: "THAIDIGIT",
  371: "THAIMONTHOFYEAR",
  372: "THAINUMSOUND",
  373: "THAINUMSTRING",
  374: "THAISTRINGLENGTH",
  375: "ISTHAIDIGIT",
  376: "ROUNDBAHTDOWN",
  377: "ROUNDBAHTUP",
  378: "THAIYEAR",
  379: "RTD",
  380: "CUBEVALUE",
  381: "CUBEMEMBER",
  382: "CUBEMEMBERPROPERTY",
  383: "CUBERANKEDMEMBER",
  384: "HEX2BIN",
  385: "HEX2DEC",
  386: "HEX2OCT",
  387: "DEC2BIN",
  388: "DEC2HEX",
  389: "DEC2OCT",
  390: "OCT2BIN",
  391: "OCT2HEX",
  392: "OCT2DEC",
  393: "BIN2DEC",
  394: "BIN2OCT",
  395: "BIN2HEX",
  396: "IMSUB",
  397: "IMDIV",
  398: "IMPOWER",
  399: "IMABS",
  400: "IMSQRT",
  401: "IMLN",
  402: "IMLOG2",
  403: "IMLOG10",
  404: "IMSIN",
  405: "IMCOS",
  406: "IMEXP",
  407: "IMARGUMENT",
  408: "IMCONJUGATE",
  409: "IMAGINARY",
  410: "IMREAL",
  411: "COMPLEX",
  412: "IMSUM",
  413: "IMPRODUCT",
  414: "SERIESSUM",
  415: "FACTDOUBLE",
  416: "SQRTPI",
  417: "QUOTIENT",
  418: "DELTA",
  419: "GESTEP",
  420: "ISEVEN",
  421: "ISODD",
  422: "MROUND",
  423: "ERF",
  424: "ERFC",
  425: "BESSELJ",
  426: "BESSELK",
  427: "BESSELY",
  428: "BESSELI",
  429: "XIRR",
  430: "XNPV",
  431: "PRICEMAT",
  432: "YIELDMAT",
  433: "INTRATE",
  434: "RECEIVED",
  435: "DISC",
  436: "PRICEDISC",
  437: "YIELDDISC",
  438: "TBILLEQ",
  439: "TBILLPRICE",
  440: "TBILLYIELD",
  441: "PRICE",
  442: "YIELD",
  443: "DOLLARDE",
  444: "DOLLARFR",
  445: "NOMINAL",
  446: "EFFECT",
  447: "CUMPRINC",
  448: "CUMIPMT",
  449: "EDATE",
  450: "EOMONTH",
  451: "YEARFRAC",
  452: "COUPDAYBS",
  453: "COUPDAYS",
  454: "COUPDAYSNC",
  455: "COUPNCD",
  456: "COUPNUM",
  457: "COUPPCD",
  458: "DURATION",
  459: "MDURATION",
  460: "ODDLPRICE",
  461: "ODDLYIELD",
  462: "ODDFPRICE",
  463: "ODDFYIELD",
  464: "RANDBETWEEN",
  465: "WEEKNUM",
  466: "AMORDEGRC",
  467: "AMORLINC",
  468: "CONVERT",
  724: "SHEETJS",
  469: "ACCRINT",
  470: "ACCRINTM",
  471: "WORKDAY",
  472: "NETWORKDAYS",
  473: "GCD",
  474: "MULTINOMIAL",
  475: "LCM",
  476: "FVSCHEDULE",
  477: "CUBEKPIMEMBER",
  478: "CUBESET",
  479: "CUBESETCOUNT",
  480: "IFERROR",
  481: "COUNTIFS",
  482: "SUMIFS",
  483: "AVERAGEIF",
  484: "AVERAGEIFS"
};
var FUNCTION_ARG_COUNTS = {
  2: 1,
  3: 1,
  10: 0,
  15: 1,
  16: 1,
  17: 1,
  18: 1,
  19: 0,
  20: 1,
  21: 1,
  22: 1,
  23: 1,
  24: 1,
  25: 1,
  26: 1,
  27: 2,
  30: 2,
  31: 3,
  32: 1,
  33: 1,
  34: 0,
  35: 0,
  38: 1,
  39: 2,
  40: 3,
  41: 3,
  42: 3,
  43: 3,
  44: 3,
  45: 3,
  47: 3,
  48: 2,
  53: 1,
  61: 3,
  63: 0,
  65: 3,
  66: 3,
  67: 1,
  68: 1,
  69: 1,
  70: 1,
  71: 1,
  72: 1,
  73: 1,
  74: 0,
  75: 1,
  76: 1,
  77: 1,
  79: 2,
  80: 2,
  83: 1,
  85: 0,
  86: 1,
  89: 0,
  90: 1,
  94: 0,
  95: 0,
  97: 2,
  98: 1,
  99: 1,
  101: 3,
  102: 3,
  105: 1,
  106: 1,
  108: 2,
  111: 1,
  112: 1,
  113: 1,
  114: 1,
  117: 2,
  118: 1,
  119: 4,
  121: 1,
  126: 1,
  127: 1,
  128: 1,
  129: 1,
  130: 1,
  131: 1,
  133: 1,
  134: 1,
  135: 1,
  136: 2,
  137: 2,
  138: 2,
  140: 1,
  141: 1,
  142: 3,
  143: 4,
  144: 4,
  161: 1,
  162: 1,
  163: 1,
  164: 1,
  165: 2,
  172: 1,
  175: 2,
  176: 2,
  177: 3,
  178: 2,
  179: 1,
  184: 1,
  186: 1,
  189: 3,
  190: 1,
  195: 3,
  196: 3,
  197: 1,
  198: 1,
  199: 3,
  201: 1,
  207: 4,
  210: 3,
  211: 1,
  212: 2,
  213: 2,
  214: 1,
  215: 1,
  225: 0,
  229: 1,
  230: 1,
  231: 1,
  232: 1,
  233: 1,
  234: 1,
  235: 3,
  244: 1,
  247: 4,
  252: 2,
  257: 1,
  261: 1,
  271: 1,
  273: 4,
  274: 2,
  275: 2,
  276: 2,
  277: 3,
  278: 3,
  279: 1,
  280: 3,
  281: 3,
  282: 3,
  283: 1,
  284: 1,
  285: 2,
  286: 4,
  287: 3,
  288: 2,
  289: 4,
  290: 3,
  291: 3,
  292: 3,
  293: 4,
  294: 1,
  295: 3,
  296: 1,
  297: 3,
  298: 1,
  299: 2,
  300: 3,
  301: 3,
  302: 4,
  303: 2,
  304: 2,
  305: 2,
  306: 2,
  307: 2,
  308: 2,
  309: 3,
  310: 2,
  311: 2,
  312: 2,
  313: 2,
  314: 2,
  315: 2,
  316: 4,
  325: 2,
  326: 2,
  327: 2,
  328: 2,
  331: 2,
  332: 2,
  337: 2,
  342: 1,
  343: 1,
  346: 2,
  347: 1,
  350: 4,
  351: 3,
  352: 1,
  353: 2,
  360: 1,
  368: 1,
  369: 1,
  370: 1,
  371: 1,
  372: 1,
  373: 1,
  374: 1,
  375: 1,
  376: 1,
  377: 1,
  378: 1,
  382: 3,
  385: 1,
  392: 1,
  393: 1,
  396: 2,
  397: 2,
  398: 2,
  399: 1,
  400: 1,
  401: 1,
  402: 1,
  403: 1,
  404: 1,
  405: 1,
  406: 1,
  407: 1,
  408: 1,
  409: 1,
  410: 1,
  414: 4,
  415: 1,
  416: 1,
  417: 2,
  420: 1,
  421: 1,
  422: 2,
  424: 1,
  425: 2,
  426: 2,
  427: 2,
  428: 2,
  430: 3,
  438: 3,
  439: 3,
  440: 3,
  443: 2,
  444: 2,
  445: 2,
  446: 2,
  447: 6,
  448: 6,
  449: 2,
  450: 2,
  464: 2,
  468: 3,
  476: 2,
  479: 1,
  480: 2,
  65535: 0
};
function formatColRef(col, isAbsolute) {
  const letter = getColumnName(col);
  return isAbsolute ? "$" + letter : letter;
}
function formatRowRef(row, isAbsolute) {
  return isAbsolute ? "$" + (row + 1) : (row + 1).toString();
}
function formatLoc(rowField, colField) {
  const rowRel = (colField & 32768) !== 0;
  const colRel = (colField & 16384) !== 0;
  return formatColRef(colField & 16383, !colRel) + formatRowRef(rowField, !rowRel);
}
function formatLocRel(rowField, colField, cellContext) {
  const rowRel = (colField & 32768) !== 0;
  const colRel = (colField & 16384) !== 0;
  let rowNum = rowField;
  let colNum = colField & 16383;
  if (rowRel) {
    if (rowNum & 32768) rowNum -= 65536;
    if (cellContext) rowNum += cellContext.row;
  }
  if (colRel) {
    colNum = colField & 255;
    if (colNum & 128) colNum -= 256;
    if (cellContext) colNum += cellContext.col;
  }
  return formatColRef(colNum, !colRel) + formatRowRef(rowNum, !rowRel);
}
function decodePTG(tokens, debug = false, cellContext = null) {
  if (!tokens || tokens.length === 0) {
    return "";
  }
  const stack = [];
  let pos = 0;
  while (pos < tokens.length) {
    const rawToken = tokens[pos];
    const baseToken = rawToken & 31;
    const tokenClass = rawToken & 96;
    if (debug) {
      console.log(`Pos ${pos}: Token 0x${rawToken.toString(16).padStart(2, "0")} (base=0x${baseToken.toString(16).padStart(2, "0")}) Stack before:`, [...stack]);
    }
    pos++;
    if (rawToken >= 32) {
      switch (baseToken) {
        // Cell reference (absolute storage, flags only affect $ display)
        case PTG.tRef:
          if (pos + 4 <= tokens.length) {
            const row = readUInt16LE(tokens, pos);
            const col = readUInt16LE(tokens, pos + 2);
            stack.push(formatLoc(row, col));
            pos += 4;
          }
          break;
        // Relative cell reference (used in shared formulas and names)
        // tRefN stores signed offsets from the host cell
        case PTG.tRefN:
          if (pos + 4 <= tokens.length) {
            const row = readUInt16LE(tokens, pos);
            const col = readUInt16LE(tokens, pos + 2);
            stack.push(formatLocRel(row, col, cellContext));
            pos += 4;
          }
          break;
        // Area reference (absolute storage, flags only affect $ display)
        case PTG.tArea:
          if (pos + 8 <= tokens.length) {
            const row1 = readUInt16LE(tokens, pos);
            const row2 = readUInt16LE(tokens, pos + 2);
            const col1 = readUInt16LE(tokens, pos + 4);
            const col2 = readUInt16LE(tokens, pos + 6);
            stack.push(`${formatLoc(row1, col1)}:${formatLoc(row2, col2)}`);
            pos += 8;
          }
          break;
        // Relative area reference (used in shared formulas and names)
        case PTG.tAreaN:
          if (pos + 8 <= tokens.length) {
            const row1 = readUInt16LE(tokens, pos);
            const row2 = readUInt16LE(tokens, pos + 2);
            const col1 = readUInt16LE(tokens, pos + 4);
            const col2 = readUInt16LE(tokens, pos + 6);
            const cell1 = formatLocRel(row1, col1, cellContext);
            const cell2 = formatLocRel(row2, col2, cellContext);
            stack.push(`${cell1}:${cell2}`);
            pos += 8;
          }
          break;
        // Function with fixed args
        case PTG.tFunc:
          if (pos + 2 <= tokens.length) {
            const funcIndex = readUInt16LE(tokens, pos);
            pos += 2;
            const funcName = FUNCTIONS[funcIndex] || `FUNC${funcIndex}`;
            const argCount = FUNCTION_ARG_COUNTS[funcIndex] || 1;
            if (argCount === -1) {
              const arg = stack.pop() || "";
              stack.push(`${funcName}(${arg})`);
            } else {
              const args = [];
              for (let i = 0; i < argCount; i++) {
                args.unshift(stack.pop() || "");
              }
              stack.push(`${funcName}(${args.join(",")})`);
            }
          }
          break;
        // Function with variable args
        case PTG.tFuncVar:
          if (pos + 3 <= tokens.length) {
            const argCount = tokens[pos];
            const funcIndex = readUInt16LE(tokens, pos + 1);
            pos += 3;
            const funcName = FUNCTIONS[funcIndex & 32767] || `FUNC${funcIndex & 32767}`;
            const args = [];
            for (let i = 0; i < (argCount & 127); i++) {
              args.unshift(stack.pop() || "");
            }
            stack.push(`${funcName}(${args.join(",")})`);
          }
          break;
        default:
          break;
      }
      continue;
    }
    switch (rawToken) {
      // Binary operators
      case PTG.tAdd:
        const b = stack.pop() || "0";
        const a = stack.pop() || "0";
        stack.push(`${a}+${b}`);
        break;
      case PTG.tSub:
        const sub2 = stack.pop() || "0";
        const sub1 = stack.pop() || "0";
        stack.push(`${sub1}-${sub2}`);
        break;
      case PTG.tMul:
        const mul2 = stack.pop() || "0";
        const mul1 = stack.pop() || "0";
        stack.push(`${mul1}*${mul2}`);
        break;
      case PTG.tDiv:
        const div2 = stack.pop() || "0";
        const div1 = stack.pop() || "0";
        stack.push(`${div1}/${div2}`);
        break;
      case PTG.tPower:
        const pow2 = stack.pop() || "0";
        const pow1 = stack.pop() || "0";
        stack.push(`${pow1}^${pow2}`);
        break;
      case PTG.tConcat:
        const str2 = stack.pop() || '""';
        const str1 = stack.pop() || '""';
        stack.push(`${str1}&${str2}`);
        break;
      // Comparison operators
      case PTG.tLT:
        const lt2 = stack.pop() || "0";
        const lt1 = stack.pop() || "0";
        stack.push(`${lt1}<${lt2}`);
        break;
      case PTG.tLE:
        const le2 = stack.pop() || "0";
        const le1 = stack.pop() || "0";
        stack.push(`${le1}<=${le2}`);
        break;
      case PTG.tEQ:
        const eq2 = stack.pop() || "0";
        const eq1 = stack.pop() || "0";
        stack.push(`${eq1}=${eq2}`);
        break;
      case PTG.tGE:
        const ge2 = stack.pop() || "0";
        const ge1 = stack.pop() || "0";
        stack.push(`${ge1}>=${ge2}`);
        break;
      case PTG.tGT:
        const gt2 = stack.pop() || "0";
        const gt1 = stack.pop() || "0";
        stack.push(`${gt1}>${gt2}`);
        break;
      case PTG.tNE:
        const ne2 = stack.pop() || "0";
        const ne1 = stack.pop() || "0";
        stack.push(`${ne1}<>${ne2}`);
        break;
      // Range operators
      case PTG.tRange:
        const range2 = stack.pop() || "A1";
        const range1 = stack.pop() || "A1";
        stack.push(`${range1}:${range2}`);
        break;
      case PTG.tUnion:
        const union2 = stack.pop() || "A1";
        const union1 = stack.pop() || "A1";
        stack.push(`${union1},${union2}`);
        break;
      // Unary operators
      case PTG.tUplus:
        const uplus = stack.pop() || "0";
        stack.push(`+${uplus}`);
        break;
      case PTG.tUminus:
        const uminus = stack.pop() || "0";
        stack.push(`-${uminus}`);
        break;
      case PTG.tPercent:
        const percent = stack.pop() || "0";
        stack.push(`${percent}%`);
        break;
      case PTG.tParen:
        const paren = stack.pop() || "0";
        stack.push(`(${paren})`);
        break;
      // Operands
      case PTG.tInt:
        if (pos + 2 <= tokens.length) {
          const intVal = readUInt16LE(tokens, pos);
          stack.push(intVal.toString());
          pos += 2;
        }
        break;
      case PTG.tNum:
        if (pos + 8 <= tokens.length) {
          const numVal = readFloat64LE(tokens, pos);
          stack.push(numVal.toString());
          pos += 8;
        }
        break;
      case PTG.tStr:
        if (pos < tokens.length) {
          const strLen = tokens[pos];
          pos++;
          if (pos < tokens.length) {
            const isUnicode = tokens[pos] & 1;
            pos++;
            let str = "";
            if (isUnicode) {
              for (let i = 0; i < strLen && pos + 1 < tokens.length; i++, pos += 2) {
                str += String.fromCharCode(readUInt16LE(tokens, pos));
              }
            } else {
              for (let i = 0; i < strLen && pos < tokens.length; i++, pos++) {
                str += String.fromCharCode(tokens[pos]);
              }
            }
            stack.push(`"${str}"`);
          }
        }
        break;
      case PTG.tBool:
        if (pos < tokens.length) {
          const boolVal = tokens[pos];
          stack.push(boolVal ? "TRUE" : "FALSE");
          pos++;
        }
        break;
      case PTG.tAttr:
        if (pos + 2 < tokens.length) {
          pos += 3;
        }
        break;
      case PTG.tMissArg:
        stack.push("");
        break;
      default:
        break;
    }
  }
  return stack.length > 0 ? stack[stack.length - 1] : "";
}

// src/parser/xls.js
var BIFF_RECORDS = {
  BOF: 2057,
  EOF: 10,
  BOUNDSHEET: 133,
  SST: 252,
  LABELSST: 253,
  RK: 638,
  MULRK: 189,
  NUMBER: 515,
  LABEL: 516,
  BOOLERR: 517,
  FORMULA: 6,
  SHRFMLA: 1212,
  ROW: 520,
  DIMENSIONS: 512,
  XF: 224,
  FORMAT: 1054,
  FONT: 49,
  STYLE: 659,
  COLINFO: 125,
  MERGECELLS: 229,
  HLINK: 440,
  NOTE: 28,
  TXO: 438,
  MULBLANK: 190,
  BLANK: 513,
  STRING: 519
};
function parseSST(data, offset, length) {
  const strings = [];
  let pos = offset + 8;
  const end = offset + length;
  while (pos < end) {
    if (pos + 3 > end) break;
    const strLen = readUInt16LE(data, pos);
    const flags = data[pos + 2];
    pos += 3;
    const isCompressed = (flags & 1) === 0;
    let str = "";
    if (isCompressed) {
      for (let i = 0; i < strLen && pos < end; i++, pos++) {
        str += String.fromCharCode(data[pos]);
      }
    } else {
      for (let i = 0; i < strLen && pos + 1 < end; i++, pos += 2) {
        str += String.fromCharCode(readUInt16LE(data, pos));
      }
    }
    strings.push(str);
  }
  return strings;
}
function parseBIFFRecords(data) {
  const records = [];
  let offset = 0;
  while (offset + 4 <= data.length) {
    const type = readUInt16LE(data, offset);
    const length = readUInt16LE(data, offset + 2);
    if (offset + 4 + length > data.length) break;
    records.push({
      type,
      typeName: Object.keys(BIFF_RECORDS).find((k) => BIFF_RECORDS[k] === type) || `UNKNOWN_${type.toString(16)}`,
      length,
      offset: offset + 4,
      data: data.slice(offset + 4, offset + 4 + length)
    });
    offset += 4 + length;
  }
  return records;
}
function decodeRK(rk) {
  const num = readUInt32LE(new Uint8Array([rk[0], rk[1], rk[2], rk[3]]), 0);
  const isInteger = (num & 2) !== 0;
  const div100 = (num & 1) !== 0;
  let value;
  if (isInteger) {
    value = num >> 2;
  } else {
    const bytes = new Uint8Array(8);
    bytes[4] = rk[0] & 252;
    bytes[5] = rk[1];
    bytes[6] = rk[2];
    bytes[7] = rk[3];
    value = new DataView(bytes.buffer).getFloat64(0, true);
  }
  return div100 ? value / 100 : value;
}
function parseFonts(records) {
  const fonts = [];
  for (const record of records) {
    if (record.type === BIFF_RECORDS.FONT) {
      const data = record.data;
      const height = readUInt16LE(data, 0);
      const options = readUInt16LE(data, 2);
      const colorIndex = readUInt16LE(data, 4);
      const weight = readUInt16LE(data, 6);
      const escapement = readUInt16LE(data, 8);
      const underline = data[10];
      const family = data[11];
      const charset = data[12];
      const nameLen = data[14];
      const flags = data[15];
      let name = "";
      const isCompressed = (flags & 1) === 0;
      let pos = 16;
      if (isCompressed) {
        for (let i = 0; i < nameLen && pos < data.length; i++, pos++) {
          const byte = data[pos];
          if (byte >= 32) {
            name += String.fromCharCode(byte);
          }
        }
      } else {
        for (let i = 0; i < nameLen && pos + 1 < data.length; i++, pos += 2) {
          const char = readUInt16LE(data, pos);
          if (char >= 32) {
            name += String.fromCharCode(char);
          }
        }
      }
      fonts.push({
        name,
        height: height / 20,
        // Convert to points
        bold: false,
        // Don't trust FONT record weight field - it's unreliable in XLS files
        italic: (options & 2) !== 0,
        strikeout: (options & 8) !== 0,
        underline,
        colorIndex,
        family,
        charset
      });
    }
  }
  return fonts;
}
function parseFormats(records) {
  const formats = {};
  for (const record of records) {
    if (record.type === BIFF_RECORDS.FORMAT) {
      const data = record.data;
      const formatIndex = readUInt16LE(data, 0);
      const strLen = readUInt16LE(data, 2);
      const options = data[4];
      const isUnicode = (options & 1) !== 0;
      let formatString;
      if (isUnicode) {
        const chars = [];
        for (let i = 0; i < strLen; i++) {
          chars.push(readUInt16LE(data, 5 + i * 2));
        }
        formatString = String.fromCharCode(...chars);
      } else {
        formatString = String.fromCharCode(...data.slice(5, 5 + strLen));
      }
      formats[formatIndex] = formatString;
    }
  }
  return formats;
}
function parseXF(records, formats, fonts) {
  const xfStyles = [];
  for (const record of records) {
    if (record.type === BIFF_RECORDS.XF) {
      const data = record.data;
      const fontIndex = readUInt16LE(data, 0);
      const formatIndex = readUInt16LE(data, 2);
      const align1 = data[6];
      const align2 = data[8];
      const brdbkg1 = readUInt32LE(data, 10);
      const brdbkg2 = readUInt32LE(data, 14);
      const brdbkg3 = readUInt16LE(data, 18);
      const horizontalAlign = align1 & 7;
      const wrapText = (align1 >> 3 & 1) !== 0;
      const verticalAlign = align1 >> 4 & 7;
      const borders = {
        left: brdbkg1 & 15,
        right: brdbkg1 >> 4 & 15,
        top: brdbkg1 >> 8 & 15,
        bottom: brdbkg1 >> 12 & 15
      };
      const patternType = brdbkg2 >> 26 & 63;
      const icvFore = brdbkg3 & 127;
      const icvBack = brdbkg3 >> 7 & 127;
      xfStyles.push({
        fontIndex,
        font: fonts[fontIndex] ? { ...fonts[fontIndex] } : null,
        // Clone font object to avoid shared references
        formatIndex,
        format: formats[formatIndex] || null,
        alignment: {
          horizontal: horizontalAlign,
          vertical: verticalAlign,
          wrapText
        },
        borders,
        fill: {
          patternType,
          fgColor: icvFore,
          bgColor: icvBack
        }
      });
    }
  }
  return xfStyles;
}
function parseColumnInfo(records) {
  const columnsBySheet = {};
  let currentSheet = -1;
  let inSheet = false;
  for (const record of records) {
    if (record.type === BIFF_RECORDS.BOF) {
      const data = record.data;
      const type = readUInt16LE(data, 2);
      if (type === 16) {
        currentSheet++;
        inSheet = true;
        if (!columnsBySheet[currentSheet]) {
          columnsBySheet[currentSheet] = [];
        }
      }
      continue;
    }
    if (record.type === BIFF_RECORDS.EOF) {
      if (inSheet) {
        inSheet = false;
      }
      continue;
    }
    if (!inSheet) continue;
    if (record.type === BIFF_RECORDS.COLINFO) {
      const data = record.data;
      const colFirst = readUInt16LE(data, 0);
      const colLast = readUInt16LE(data, 2);
      const width = readUInt16LE(data, 4);
      const xfIndex = readUInt16LE(data, 6);
      const options = readUInt16LE(data, 8);
      for (let col = colFirst; col <= colLast; col++) {
        columnsBySheet[currentSheet][col] = {
          width: width / 256,
          // Convert to character units
          hidden: (options & 1) !== 0,
          xfIndex
        };
      }
    }
  }
  return columnsBySheet;
}
function parseRowInfo(records) {
  const rowsBySheet = {};
  let currentSheet = -1;
  let inSheet = false;
  for (const record of records) {
    if (record.type === BIFF_RECORDS.BOF) {
      const data = record.data;
      const type = readUInt16LE(data, 2);
      if (type === 16) {
        currentSheet++;
        inSheet = true;
        if (!rowsBySheet[currentSheet]) {
          rowsBySheet[currentSheet] = {};
        }
      }
      continue;
    }
    if (record.type === BIFF_RECORDS.EOF) {
      if (inSheet) {
        inSheet = false;
      }
      continue;
    }
    if (!inSheet) continue;
    if (record.type === BIFF_RECORDS.ROW) {
      const data = record.data;
      const rowIndex = readUInt16LE(data, 0);
      const height = readUInt16LE(data, 6);
      const options = readUInt16LE(data, 12);
      rowsBySheet[currentSheet][rowIndex] = {
        height: height / 20,
        // Convert to points
        hidden: (options & 32) !== 0
      };
    }
  }
  return rowsBySheet;
}
function parseMergedCells(records) {
  const mergedCellsBySheet = {};
  let currentSheet = -1;
  let inSheet = false;
  for (const record of records) {
    if (record.type === BIFF_RECORDS.BOF) {
      const data = record.data;
      const type = readUInt16LE(data, 2);
      if (type === 16) {
        currentSheet++;
        inSheet = true;
        if (!mergedCellsBySheet[currentSheet]) {
          mergedCellsBySheet[currentSheet] = [];
        }
      }
      continue;
    }
    if (record.type === BIFF_RECORDS.EOF) {
      if (inSheet) {
        inSheet = false;
      }
      continue;
    }
    if (inSheet && record.type === BIFF_RECORDS.MERGECELLS) {
      const data = record.data;
      const count = readUInt16LE(data, 0);
      for (let i = 0; i < count; i++) {
        const offset = 2 + i * 8;
        const rowFirst = readUInt16LE(data, offset);
        const rowLast = readUInt16LE(data, offset + 2);
        const colFirst = readUInt16LE(data, offset + 4);
        const colLast = readUInt16LE(data, offset + 6);
        mergedCellsBySheet[currentSheet].push({
          rowFirst,
          rowLast,
          colFirst,
          colLast
        });
      }
    }
  }
  return mergedCellsBySheet;
}
function parseCellRecords(records, sst, xfStyles, workbookData) {
  const cells = [];
  const shrfmlaRecords = [];
  let inSheet = false;
  let sheetIndex = -1;
  let lastFormulaCell = null;
  for (const record of records) {
    const data = record.data;
    if (record.type === BIFF_RECORDS.BOF) {
      const type = readUInt16LE(data, 2);
      if (type === 16) {
        inSheet = true;
        sheetIndex++;
      }
      continue;
    }
    if (record.type === BIFF_RECORDS.EOF) {
      if (inSheet) {
        inSheet = false;
      }
      continue;
    }
    if (!inSheet) continue;
    if (record.type === BIFF_RECORDS.LABELSST) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      const sstIndex = readUInt32LE(data, 6);
      cells.push({
        row,
        col,
        value: sst[sstIndex] || "",
        type: "string",
        style: xfStyles[xfIndex] || null,
        sheetIndex
      });
    }
    if (record.type === BIFF_RECORDS.RK) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      const value = decodeRK(data.slice(6, 10));
      cells.push({
        row,
        col,
        value,
        type: "number",
        style: xfStyles[xfIndex] || null,
        sheetIndex
      });
    }
    if (record.type === BIFF_RECORDS.NUMBER) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      const value = readFloat64LE(data, 6);
      cells.push({
        row,
        col,
        value,
        type: "number",
        style: xfStyles[xfIndex] || null,
        sheetIndex
      });
    }
    if (record.type === BIFF_RECORDS.FORMULA) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      const resultBytes = data.slice(6, 14);
      let value = null;
      let formulaType = "number";
      if (resultBytes[0] === 255 && resultBytes[1] === 255) {
        formulaType = "string";
      } else {
        value = readFloat64LE(resultBytes, 0);
      }
      const formulaLen = readUInt16LE(data, 20);
      const formulaTokens = data.slice(22, 22 + formulaLen);
      lastFormulaCell = {
        row,
        col,
        value,
        type: "formula",
        formulaType,
        formula: Array.from(formulaTokens),
        // Store raw tokens
        style: xfStyles[xfIndex] || null,
        sheetIndex
      };
      cells.push(lastFormulaCell);
    }
    if (record.type === BIFF_RECORDS.STRING && lastFormulaCell) {
      const strLen = readUInt16LE(data, 0);
      const flags = data[2];
      const isCompressed = (flags & 1) === 0;
      let str = "";
      let pos = 3;
      if (isCompressed) {
        for (let i = 0; i < strLen && pos < data.length; i++, pos++) {
          str += String.fromCharCode(data[pos]);
        }
      } else {
        for (let i = 0; i < strLen && pos + 1 < data.length; i++, pos += 2) {
          str += String.fromCharCode(readUInt16LE(data, pos));
        }
      }
      lastFormulaCell.value = str;
      lastFormulaCell = null;
    }
    if (record.type === BIFF_RECORDS.SHRFMLA) {
      const firstRow = readUInt16LE(data, 0);
      const lastRow = readUInt16LE(data, 2);
      const firstCol = data[4];
      const lastCol = data[5];
      const formulaLen = readUInt16LE(data, 8);
      const sharedFormula = data.slice(10, 10 + formulaLen);
      shrfmlaRecords.push({
        sheetIndex,
        firstRow,
        lastRow,
        firstCol,
        lastCol,
        formula: Array.from(sharedFormula)
      });
    }
    if (record.type === BIFF_RECORDS.BLANK) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      cells.push({
        row,
        col,
        value: "",
        type: "blank",
        style: xfStyles[xfIndex] || null,
        sheetIndex
      });
    }
    if (record.type === BIFF_RECORDS.MULBLANK) {
      const row = readUInt16LE(data, 0);
      const colFirst = readUInt16LE(data, 2);
      const colLast = readUInt16LE(data, data.length - 2);
      let pos = 4;
      for (let col = colFirst; col <= colLast; col++) {
        const xfIndex = readUInt16LE(data, pos);
        cells.push({
          row,
          col,
          value: "",
          type: "blank",
          style: xfStyles[xfIndex] || null,
          sheetIndex
        });
        pos += 2;
      }
    }
    if (record.type === BIFF_RECORDS.MULRK) {
      const row = readUInt16LE(data, 0);
      const colFirst = readUInt16LE(data, 2);
      const colLast = readUInt16LE(data, data.length - 2);
      let pos = 4;
      for (let col = colFirst; col <= colLast; col++) {
        const xfIndex = readUInt16LE(data, pos);
        const rkValue = data.slice(pos + 2, pos + 6);
        const value = decodeRK(rkValue);
        cells.push({
          row,
          col,
          value,
          type: "number",
          style: xfStyles[xfIndex] || null,
          sheetIndex
        });
        pos += 6;
      }
    }
    if (record.type === BIFF_RECORDS.BOOLERR) {
      const row = readUInt16LE(data, 0);
      const col = readUInt16LE(data, 2);
      const xfIndex = readUInt16LE(data, 4);
      const boolErr = data[6];
      const isError = data[7];
      cells.push({
        row,
        col,
        value: isError ? `#ERROR${boolErr}` : boolErr !== 0,
        type: isError ? "error" : "boolean",
        style: xfStyles[xfIndex] || null,
        sheetIndex
      });
    }
  }
  for (const shr of shrfmlaRecords) {
    for (const cell of cells) {
      if (cell.sheetIndex === shr.sheetIndex && cell.row >= shr.firstRow && cell.row <= shr.lastRow && cell.col >= shr.firstCol && cell.col <= shr.lastCol && cell.type === "formula" && cell.formula && cell.formula.length > 0 && cell.formula[0] === 1) {
        cell.formula = shr.formula;
        cell.sharedFormulaBase = {
          row: shr.firstRow,
          col: shr.firstCol
        };
      }
    }
  }
  return cells;
}
function styleToCss(style) {
  if (!style) return "";
  const css = [];
  if (style.font) {
    if (style.font.name) css.push(`font-family: ${style.font.name}`);
    if (style.font.height) css.push(`font-size: ${style.font.height}pt`);
    if (style.font.bold) css.push("font-weight: bold");
    if (style.font.italic) css.push("font-style: italic");
    if (style.font.underline) css.push("text-decoration: underline");
    if (style.font.strikeout) css.push("text-decoration: line-through");
  }
  if (style.alignment) {
    const hAlign = ["left", "center", "right", "fill", "justify", "center-across-selection", "distributed"][style.alignment.horizontal];
    const vAlign = ["top", "middle", "bottom", "justify", "distributed"][style.alignment.vertical];
    if (hAlign) css.push(`text-align: ${hAlign}`);
    if (vAlign) css.push(`vertical-align: ${vAlign}`);
    if (style.alignment.wrapText) css.push("white-space: pre-wrap");
  }
  const excelDefaultPalette = [
    "#000000",
    "#FFFFFF",
    "#FF0000",
    "#00FF00",
    "#0000FF",
    "#FFFF00",
    "#FF00FF",
    "#00FFFF",
    "#000000",
    "#FFFFFF",
    "#FF0000",
    "#00FF00",
    "#0000FF",
    "#FFFF00",
    "#FF00FF",
    "#00FFFF",
    "#800000",
    "#008000",
    "#000080",
    "#808000",
    "#800080",
    "#008080",
    "#C0C0C0",
    "#808080",
    "#9999FF",
    "#993366",
    "#FFFFCC",
    "#CCFFFF",
    "#660066",
    "#FF8080",
    "#0066CC",
    "#CCCCFF",
    "#000080",
    "#FF00FF",
    "#FFFF00",
    "#00FFFF",
    "#800080",
    "#800000",
    "#008080",
    "#0000FF",
    "#00CCFF",
    "#CCFFFF",
    "#CCFFCC",
    "#FFFF99",
    "#99CCFF",
    "#FF99CC",
    "#CC99FF",
    "#FFCC99",
    "#3366FF",
    "#33CCCC",
    "#99CC00",
    "#FFCC00",
    "#FF9900",
    "#FF6600",
    "#666699",
    "#969696",
    "#003366",
    "#339966",
    "#003300",
    "#333300",
    "#993300",
    "#993366",
    "#333399",
    "#333333",
    // System/automatic colors (64+)
    // Index 64: Window Text/Foreground (typically black, but for fills often renders as light color)
    // Index 65: Window Background (typically white)
    "#000000",
    "#FFFFFF",
    "#000000",
    "#000000",
    "#000000",
    "#000000",
    "#000000",
    "#000000"
  ];
  if (style.fill) {
    const patternType = style.fill.patternType;
    const fgColorIdx = style.fill.fgColor;
    const bgColorIdx = style.fill.bgColor;
    if (patternType === 1 && fgColorIdx >= 0 && fgColorIdx < 64 && fgColorIdx < excelDefaultPalette.length) {
      css.push(`background-color: ${excelDefaultPalette[fgColorIdx]}`);
    } else if (patternType > 1 && bgColorIdx >= 0 && bgColorIdx < 64 && bgColorIdx < excelDefaultPalette.length) {
      css.push(`background-color: ${excelDefaultPalette[bgColorIdx]}`);
    }
  }
  const fontColorIdx = style.font?.color !== void 0 ? style.font.color : style.font?.colorIndex;
  if (fontColorIdx !== void 0 && fontColorIdx >= 0 && fontColorIdx < excelDefaultPalette.length) {
    css.push(`color: ${excelDefaultPalette[fontColorIdx]}`);
  }
  if (style.borders) {
    if (style.borders.left) css.push("border-left: 1px solid #000");
    if (style.borders.right) css.push("border-right: 1px solid #000");
    if (style.borders.top) css.push("border-top: 1px solid #000");
    if (style.borders.bottom) css.push("border-bottom: 1px solid #000");
  }
  return css.join("; ");
}
function cellsToJspreadsheet(cells, columns, rows, mergedCells, formats) {
  if (cells.length === 0) return {
    data: [],
    columns: [],
    rows: {},
    cells: {},
    style: {},
    mergeCells: {}
  };
  let maxRow = cells.length > 0 ? Math.max(...cells.map((c) => c.row)) : 0;
  let maxCol = cells.length > 0 ? Math.max(...cells.map((c) => c.col)) : 0;
  mergedCells.forEach((merge) => {
    if (merge.rowLast > maxRow) maxRow = merge.rowLast;
    if (merge.colLast > maxCol) maxCol = merge.colLast;
  });
  const data = [];
  for (let r = 0; r <= maxRow; r++) {
    data[r] = [];
    for (let c = 0; c <= maxCol; c++) {
      data[r][c] = "";
    }
  }
  const cellsObj = {};
  const styleObj = {};
  cells.forEach((cell) => {
    const address = getCellNameFromCoords(cell.col, cell.row);
    if (cell.type === "formula") {
      if (cell.formula && cell.formula.length > 0) {
        try {
          const cellContext = cell.sharedFormulaBase ? {
            row: cell.row,
            col: cell.col,
            baseRow: cell.sharedFormulaBase.row,
            baseCol: cell.sharedFormulaBase.col
          } : null;
          const formulaStr = decodePTG(new Uint8Array(cell.formula), false, cellContext);
          data[cell.row][cell.col] = formulaStr ? "=" + formulaStr : "=";
          if (!cellsObj[address]) {
            cellsObj[address] = {};
          }
          cellsObj[address].formula = formulaStr;
        } catch (err) {
          data[cell.row][cell.col] = `=ERROR`;
        }
      } else {
        data[cell.row][cell.col] = `=FORMULA(${JSON.stringify(cell.value)})`;
      }
    } else {
      data[cell.row][cell.col] = cell.value;
    }
    if (cell.style) {
      const cssString = styleToCss(cell.style);
      if (cssString) {
        styleObj[address] = cssString;
      }
      const cellProps = {};
      if (cell.style.format && !cell.style.format.toUpperCase().startsWith("GENERAL")) {
        cellProps.format = cell.style.format;
      }
      if (cell.style.alignment) {
        const hAlign = ["left", "center", "right", "fill", "justify"][cell.style.alignment.horizontal];
        if (hAlign) {
          cellProps.align = hAlign;
        }
        if (cell.style.alignment.wrapText) {
          cellProps.wrap = true;
        }
      }
      if (Object.keys(cellProps).length > 0) {
        if (cellsObj[address]) {
          cellsObj[address] = { ...cellsObj[address], ...cellProps };
        } else {
          cellsObj[address] = cellProps;
        }
      }
    }
  });
  const columnsArray = [];
  for (let c = 0; c <= maxCol; c++) {
    const col = columns[c] || {};
    const column = {
      width: col.width ? convertWidthToPixels(col.width, "char") : 100,
      title: col.title || getColumnName(c),
      type: "text"
    };
    if (col.hidden) {
      column.visible = false;
    }
    columnsArray.push(column);
  }
  const rowsObj = {};
  for (let r = 0; r <= maxRow; r++) {
    if (rows[r]) {
      const row = {
        height: rows[r].height || 21
      };
      if (rows[r].hidden) {
        row.visible = false;
      }
      rowsObj[r] = row;
    }
  }
  const mergeCellsObj = {};
  mergedCells.forEach((merge) => {
    const address = getCellNameFromCoords(merge.colFirst, merge.rowFirst);
    const colspan = merge.colLast - merge.colFirst + 1;
    const rowspan = merge.rowLast - merge.rowFirst + 1;
    mergeCellsObj[address] = [colspan, rowspan];
  });
  return {
    data,
    columns: columnsArray,
    rows: rowsObj,
    cells: cellsObj,
    style: styleObj,
    mergeCells: mergeCellsObj,
    minDimensions: [maxCol + 1, maxRow + 1]
  };
}
function parseXLSBuffer(buffer) {
  const cfb = cfb_default.read(buffer, { type: "buffer" });
  const workbookEntry = cfb_default.find(cfb, "Workbook") || cfb_default.find(cfb, "Book");
  if (!workbookEntry) {
    throw new Error("No Workbook stream found in XLS file");
  }
  const workbookData = new Uint8Array(workbookEntry.content);
  const records = parseBIFFRecords(workbookData);
  let sst = [];
  const sstRecord = records.find((r) => r.type === BIFF_RECORDS.SST);
  if (sstRecord) {
    sst = parseSST(workbookData, sstRecord.offset, sstRecord.length);
  }
  const fonts = parseFonts(records);
  const formats = parseFormats(records);
  const xfStyles = parseXF(records, formats, fonts);
  const columnsBySheet = parseColumnInfo(records);
  const rowsBySheet = parseRowInfo(records);
  const mergedCellsBySheet = parseMergedCells(records);
  const sheets = [];
  const boundsheets = records.filter((r) => r.type === BIFF_RECORDS.BOUNDSHEET);
  boundsheets.forEach((bs) => {
    const sheetState = bs.data[4];
    const nameLen = bs.data[6];
    const flags = bs.data[7];
    const isCompressed = (flags & 1) === 0;
    let name = "";
    let pos = 8;
    if (isCompressed) {
      for (let i = 0; i < nameLen && pos < bs.data.length; i++, pos++) {
        const byte = bs.data[pos];
        if (byte >= 32) {
          name += String.fromCharCode(byte);
        }
      }
    } else {
      for (let i = 0; i < nameLen && pos + 1 < bs.data.length; i++, pos += 2) {
        const char = readUInt16LE(bs.data, pos);
        if (char >= 32) {
          name += String.fromCharCode(char);
        }
      }
    }
    sheets.push({
      name,
      visibility: sheetState
      // 0=visible, 1=hidden, 2=very hidden
    });
  });
  const allCells = parseCellRecords(records, sst, xfStyles, workbookData);
  const cellsBySheet = {};
  allCells.forEach((cell) => {
    if (!cellsBySheet[cell.sheetIndex]) {
      cellsBySheet[cell.sheetIndex] = [];
    }
    cellsBySheet[cell.sheetIndex].push(cell);
  });
  const globalStyles = [];
  const styleMap = /* @__PURE__ */ new Map();
  const worksheets = sheets.map((sheet, index) => {
    const sheetCells = cellsBySheet[index] || [];
    const sheetMergedCells = mergedCellsBySheet[index] || [];
    const sheetColumns = columnsBySheet[index] || [];
    const sheetRows = rowsBySheet[index] || {};
    const sheetData = cellsToJspreadsheet(sheetCells, sheetColumns, sheetRows, sheetMergedCells, formats);
    const worksheetStyleIndices = {};
    if (sheetData.style) {
      Object.keys(sheetData.style).forEach((cellRef) => {
        const cssString = sheetData.style[cellRef];
        if (!styleMap.has(cssString)) {
          styleMap.set(cssString, globalStyles.length);
          globalStyles.push(cssString);
        }
        worksheetStyleIndices[cellRef] = styleMap.get(cssString);
      });
    }
    const worksheet = {
      worksheetName: sheet.name,
      ...sheetData,
      style: worksheetStyleIndices
    };
    if (sheet.visibility === 1) {
      worksheet.worksheetState = "hidden";
    } else if (sheet.visibility === 2) {
      worksheet.worksheetState = "veryHidden";
    }
    return worksheet;
  });
  return {
    worksheets,
    style: globalStyles
  };
}

// src/xml/xml-tree-parser.js
function parser(html, options) {
  if (!options || typeof options !== "object") {
    options = {};
  }
  const ignoreTags = new Set((options.ignore || []).map((tag) => tag.toLowerCase()));
  function isSelfClosing(type) {
    if (!type) {
      return false;
    }
    const selfClosingTags = [
      "area",
      "base",
      "br",
      "col",
      "embed",
      "hr",
      "img",
      "input",
      "link",
      "meta",
      "source",
      "track",
      "wbr"
    ];
    return selfClosingTags.includes(type.toLowerCase());
  }
  const createTextNode = function(tag) {
    if (!this.current.children) {
      this.current.children = [];
    }
    this.current.children.push({
      type: "#text",
      props: [tag]
    });
  };
  const findParentByTagName = function(type) {
    for (let i = this.stack.length - 1; i >= 0; i--) {
      if (this.stack[i].type === type) {
        return this.stack[i];
      }
    }
    return void 0;
  };
  const commitText = function() {
    if (typeof this.text !== "undefined") {
      if (this.ignoreDepth === 0) {
        let text = this.text;
        if (text) {
          createTextNode.call(this, { name: "textContent", value: text });
        }
      }
      delete this.text;
    }
  };
  const commitComments = function() {
    if (typeof this.comments !== "undefined") {
      if (this.ignoreDepth === 0) {
        let comments = this.comments;
        if (comments) {
          comments = comments.replace("<!--", "").replace("-->", "");
          if (!this.current.children) {
            this.current.children = [];
          }
          this.current.children.push({
            type: "#comments",
            props: [{ name: "text", value: comments }]
          });
        }
      }
      delete this.comments;
    }
  };
  const commitAttribute = function() {
    if (this.tag.attributeName) {
      if (!this.tag.props) {
        this.tag.props = [];
      }
      let k = this.tag.attributeName;
      let v = this.tag.attributeValue;
      if (typeof v === "undefined") {
        v = k;
      }
      this.tag.props.push({
        name: k,
        value: v
      });
      delete this.tag.attributeName;
      delete this.tag.attributeValue;
      if (this.tag.attributeIsReadyToClose) {
        delete this.tag.attributeIsReadyToClose;
      }
    }
  };
  const actions = function(control2, char) {
    const method = control2.action || "text";
    if (typeof actions[method] === "function") {
      actions[method].call(control2, char);
    }
  };
  actions.processTag = function(char) {
    commitText.call(this);
    if (char === "<") {
      this.tag = {
        type: ""
      };
    } else if (char.match(/[a-zA-Z0-9-:]/)) {
      this.tag.type += char;
    } else {
      this.action = "attributeName";
    }
  };
  actions.closeTag = function(char) {
    commitAttribute.call(this);
    if (char === ">") {
      const shouldIgnore = ignoreTags.has(this.tag.type.toLowerCase());
      if (this.tag.closingTag && shouldIgnore && this.ignoreDepth > 0) {
        this.ignoreDepth--;
        this.tag = null;
        this.action = "text";
        return;
      }
      if (this.ignoreDepth > 0) {
        this.tag = null;
        this.action = "text";
        return;
      }
      if (shouldIgnore && !this.tag.closingTag) {
        this.ignoreDepth++;
        this.tag = null;
        this.action = "text";
        return;
      }
      if (isSelfClosing(this.tag.type)) {
        if (!this.current.children) {
          this.current.children = [];
        }
        this.current.children.push(this.tag);
      } else if (this.tag.closingTag) {
        const parentNode = findParentByTagName.call(this, this.tag.type);
        if (parentNode) {
          while (this.stack.length > 0 && this.stack[this.stack.length - 1] !== parentNode) {
            this.stack.pop();
          }
          if (this.stack.length > 0) {
            this.stack.pop();
          }
          this.current = this.stack.length > 0 ? this.stack[this.stack.length - 1] : this.root;
        }
      } else {
        if (this.tag.closing) {
          if (!this.current.children) {
            this.current.children = [];
          }
          this.current.children.push(this.tag);
        } else {
          if (!this.current.children) {
            this.current.children = [];
          }
          this.current.children.push(this.tag);
          this.stack.push(this.tag);
          this.current = this.tag;
        }
      }
      delete this.tag.insideQuote;
      delete this.tag.closingTag;
      delete this.tag.closing;
      delete this.tag.locked;
      this.tag = null;
      this.action = "text";
    } else if (!this.tag.locked) {
      if (char === "/") {
        if (!this.tag.type) {
          this.tag.closingTag = true;
        }
        this.tag.closing = true;
      } else if (char.match(/[a-zA-Z0-9-:]/)) {
        if (this.tag.closingTag) {
          this.tag.type += char;
        }
      } else {
        if (this.tag.type) {
          this.tag.locked = true;
        }
      }
    }
  };
  actions.attributeName = function(char) {
    if (this.tag.attributeIsReadyToClose) {
      commitAttribute.call(this);
    }
    if (char.match(/[a-zA-Z0-9-:]/)) {
      if (!this.tag.attributeName) {
        this.tag.attributeName = "";
      }
      this.tag.attributeName += char;
    } else if (char === "=") {
      if (this.tag.attributeName) {
        this.action = "attributeValue";
        delete this.tag.attributeIsReadyToClose;
      }
    } else if (char.match(/\s/)) {
      if (this.tag.attributeName) {
        this.tag.attributeIsReadyToClose = true;
      }
    }
  };
  actions.attributeValue = function(char) {
    if (!this.tag.attributeValue) {
      this.tag.attributeValue = "";
    }
    if (char === '"' || char === "'") {
      if (this.tag.insideQuote) {
        if (this.tag.insideQuote === char) {
          this.tag.insideQuote = false;
        } else {
          this.tag.attributeValue += char;
        }
      } else {
        this.tag.insideQuote = char;
      }
    } else {
      if (this.tag.insideQuote) {
        if (this.tag.attributeValue) {
          this.tag.attributeValue += char;
        } else {
          this.tag.attributeValue = char;
        }
      } else if (typeof char === "string" && char.match(/\s/)) {
        if (this.tag.attributeValue) {
          this.action = "attributeName";
        }
        this.tag.attributeIsReadyToClose = true;
      } else {
        if (this.tag.attributeIsReadyToClose) {
          this.action = "attributeName";
          actions.attributeName.call(this, char);
        } else {
          if (this.tag.attributeValue) {
            this.tag.attributeValue += char;
          } else {
            this.tag.attributeValue = char;
          }
        }
      }
    }
  };
  actions.text = function(char) {
    if (!this.text) {
      this.text = "";
    }
    this.text += char;
  };
  actions.comments = function(char) {
    if (!this.comments) {
      this.comments = "";
    }
    this.comments += char;
    if (this.comments.endsWith("-->")) {
      commitComments.call(this);
      this.action = "text";
    }
  };
  const result = { type: "template" };
  const control = {
    root: result,
    current: result,
    stack: [],
    // Stack to track open tags
    action: "text",
    ignoreDepth: 0
    // Track depth of ignored tags
  };
  if (typeof html !== "string") {
    throw new TypeError("HTML input must be a string");
  }
  for (let i = 0; i < html.length; i++) {
    let char = html[i];
    if (control.action === "text" && char === "<" && html[i + 1] === "!" && html[i + 2] === "-" && html[i + 3] === "-") {
      control.action = "comments";
    }
    if (control.action !== "comments") {
      if (control.tag) {
        if (char === ">" || char === "/") {
          if (!control.tag.insideQuote) {
            control.action = "closeTag";
          }
        }
      } else {
        if (char === "<") {
          control.action = "processTag";
        }
      }
    }
    actions(control, char);
  }
  commitText.call(control);
  let finalResult;
  if (result.children) {
    if (result.children.length === 1) {
      finalResult = result.children[0];
    } else if (result.children.length > 1) {
      finalResult = result;
    }
  }
  return finalResult;
}

// src/parser/odsCore.js
function parseNumberFormat(node) {
  if (!node || !node.children) return null;
  const attrs = parseAttributes(node);
  const styleName = attrs["style:name"] || attrs["number:style-name"];
  let decimalPlaces = 0;
  let minDecimalPlaces = 0;
  let minIntegerDigits = 1;
  let isPercentage = node.type === "number:percentage-style";
  let isCurrency = node.type === "number:currency-style";
  let currencySymbol = "";
  const findNumberProps = (n) => {
    if (!n) return;
    if (Array.isArray(n)) {
      n.forEach(findNumberProps);
      return;
    }
    if (typeof n === "object") {
      if (n.type === "number:number") {
        const nattrs = parseAttributes(n);
        if (nattrs["number:decimal-places"]) {
          decimalPlaces = parseInt(nattrs["number:decimal-places"]) || 0;
        }
        if (nattrs["number:min-decimal-places"]) {
          minDecimalPlaces = parseInt(nattrs["number:min-decimal-places"]) || 0;
        }
        if (nattrs["number:min-integer-digits"]) {
          minIntegerDigits = parseInt(nattrs["number:min-integer-digits"]) || 1;
        }
      } else if (n.type === "number:currency-symbol") {
        currencySymbol = getTextContent(n) || "$";
      }
      if (n.children) findNumberProps(n.children);
    }
  };
  findNumberProps(node);
  let mask = "";
  if (isCurrency) {
    mask = currencySymbol + "#,##0";
    if (decimalPlaces > 0) {
      mask += "." + "0".repeat(decimalPlaces);
    }
  } else if (isPercentage) {
    mask = "0";
    if (decimalPlaces > 0) {
      mask += "." + "0".repeat(decimalPlaces);
    }
    mask += "%";
  } else {
    if (decimalPlaces > 0 || minDecimalPlaces > 0) {
      const decimals = Math.max(decimalPlaces, minDecimalPlaces);
      mask = "0." + "0".repeat(decimals);
    } else {
      mask = "0";
    }
  }
  return mask;
}
function parseNumberFormats(xml) {
  const formats = {};
  if (!xml) return formats;
  const parsed = parser(xml);
  const findFormats = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach(findFormats);
      return;
    }
    if (typeof node === "object") {
      if (node.type === "number:number-style" || node.type === "number:percentage-style" || node.type === "number:currency-style" || node.type === "number:date-style" || node.type === "number:time-style") {
        const attrs = parseAttributes(node);
        const name = attrs["style:name"];
        if (name) {
          const mask = parseNumberFormat(node);
          if (mask) {
            formats[name] = mask;
          }
        }
      }
      if (node.children) {
        findFormats(node.children);
      }
    }
  };
  findFormats(parsed);
  return formats;
}
function parseStyles2(stylesXml) {
  const styles = {};
  if (!stylesXml) return styles;
  const parsed = parser(stylesXml);
  const findStyles = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach(findStyles);
      return;
    }
    if (typeof node === "object") {
      if (node.type === "style:style") {
        const attrs = parseAttributes(node);
        const styleName = attrs["style:name"];
        if (styleName) {
          styles[styleName] = {
            family: attrs["style:family"],
            parentStyleName: attrs["style:parent-style-name"],
            dataStyleName: attrs["style:data-style-name"]
            // Store number format reference
          };
          if (node.children) {
            node.children.forEach((child) => {
              if (child.type === "style:text-properties") {
                styles[styleName].text = parseAttributes(child);
              } else if (child.type === "style:paragraph-properties") {
                styles[styleName].paragraph = parseAttributes(child);
              } else if (child.type === "style:table-cell-properties") {
                styles[styleName].cell = parseAttributes(child);
              } else if (child.type === "style:table-column-properties") {
                styles[styleName].column = parseAttributes(child);
              } else if (child.type === "style:table-row-properties") {
                styles[styleName].row = parseAttributes(child);
              }
            });
          }
        }
      }
      if (node.children) {
        findStyles(node.children);
      }
    }
  };
  findStyles(parsed);
  return styles;
}
function parseTableStyles(xml) {
  const tableStyles = {};
  if (!xml) return tableStyles;
  const parsed = parser(xml);
  const findTableStyles = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach(findTableStyles);
      return;
    }
    if (typeof node === "object") {
      if (node.type === "style:style") {
        const attrs = parseAttributes(node);
        if (attrs["style:family"] === "table") {
          const styleName = attrs["style:name"];
          tableStyles[styleName] = { name: styleName };
          if (node.children) {
            node.children.forEach((child) => {
              if (child.type === "style:table-properties") {
                const props = parseAttributes(child);
                tableStyles[styleName].display = props["table:display"];
              }
            });
          }
        }
      }
      if (node.children) {
        findTableStyles(node.children);
      }
    }
  };
  findTableStyles(parsed);
  return tableStyles;
}
function styleToCSS(styleName, styles) {
  if (!styleName || !styles[styleName]) return "";
  const style = styles[styleName];
  const css = [];
  if (style.text) {
    const t = style.text;
    if (t["fo:font-family"]) css.push(`font-family: ${t["fo:font-family"]}`);
    if (t["fo:font-size"]) css.push(`font-size: ${t["fo:font-size"]}`);
    if (t["fo:font-weight"]) css.push(`font-weight: ${t["fo:font-weight"]}`);
    if (t["fo:font-style"]) css.push(`font-style: ${t["fo:font-style"]}`);
    if (t["style:text-underline-style"]) css.push("text-decoration: underline");
    if (t["fo:color"]) css.push(`color: ${t["fo:color"]}`);
  }
  const normalizeBorder = (border) => {
    if (!border) return border;
    return border.replace(/\bthin\b/g, "1px").replace(/\bmedium\b/g, "2px").replace(/\bthick\b/g, "3px");
  };
  if (style.cell) {
    const c = style.cell;
    if (c["fo:background-color"]) css.push(`background-color: ${c["fo:background-color"]}`);
    if (c["fo:border"]) {
      const normalizedBorder = normalizeBorder(c["fo:border"]);
      css.push(`border-left: ${normalizedBorder}`);
      css.push(`border-right: ${normalizedBorder}`);
      css.push(`border-top: ${normalizedBorder}`);
      css.push(`border-bottom: ${normalizedBorder}`);
    }
    if (c["fo:border-left"] && c["fo:border-left"] !== "none") {
      css.push(`border-left: ${normalizeBorder(c["fo:border-left"])}`);
    }
    if (c["fo:border-right"] && c["fo:border-right"] !== "none") {
      css.push(`border-right: ${normalizeBorder(c["fo:border-right"])}`);
    }
    if (c["fo:border-top"] && c["fo:border-top"] !== "none") {
      css.push(`border-top: ${normalizeBorder(c["fo:border-top"])}`);
    }
    if (c["fo:border-bottom"] && c["fo:border-bottom"] !== "none") {
      css.push(`border-bottom: ${normalizeBorder(c["fo:border-bottom"])}`);
    }
  }
  if (style.paragraph) {
    const p = style.paragraph;
    if (p["fo:text-align"]) css.push(`text-align: ${p["fo:text-align"]}`);
  }
  return css.join("; ");
}
function parseRichText(node) {
  if (!node) return null;
  const segments = [];
  let hasFormatting = false;
  const processTextNode = (n, inFormattedNode = false) => {
    if (!n) return;
    if (Array.isArray(n)) {
      n.forEach((child) => processTextNode(child, inFormattedNode));
      return;
    }
    if (typeof n === "object") {
      if (n.type === "text:span") {
        hasFormatting = true;
        const attrs = parseAttributes(n);
        const styleName = attrs["text:style-name"];
        const text = getTextContent(n);
        if (text) {
          segments.push({
            text,
            style: styleName || null
          });
        }
        return;
      } else if (n.type === "text:a") {
        hasFormatting = true;
        const attrs = parseAttributes(n);
        const text = getTextContent(n);
        if (text) {
          segments.push({
            text,
            link: attrs["xlink:href"] || null
          });
        }
        return;
      } else if (n.type === "text:p") {
        if (n.children) {
          processTextNode(n.children, inFormattedNode);
        }
        return;
      } else if (n.type === "#text" && !inFormattedNode) {
        const text = getTextContent(n);
        if (text) {
          segments.push({ text });
        }
      }
      if (n.children && n.type !== "text:span" && n.type !== "text:a") {
        processTextNode(n.children, inFormattedNode);
      }
    }
  };
  processTextNode(node);
  return hasFormatting && segments.length > 0 ? segments : null;
}
function parseImage(frameNode, rowIndex, colIndex) {
  const attrs = parseAttributes(frameNode);
  const image = {
    type: "image",
    anchorType: attrs["text:anchor-type"] || "cell",
    x: attrs["svg:x"] || "0",
    y: attrs["svg:y"] || "0",
    width: attrs["svg:width"] || "0",
    height: attrs["svg:height"] || "0",
    row: rowIndex,
    col: colIndex,
    zIndex: parseInt(attrs["draw:z-index"]) || 0
  };
  const findImage = (node) => {
    if (!node) return null;
    if (Array.isArray(node)) {
      for (let n of node) {
        const result = findImage(n);
        if (result) return result;
      }
      return null;
    }
    if (typeof node === "object") {
      if (node.type === "draw:image") {
        const imgAttrs = parseAttributes(node);
        return {
          href: imgAttrs["xlink:href"],
          type: imgAttrs["xlink:type"]
        };
      }
      if (node.children) {
        return findImage(node.children);
      }
    }
    return null;
  };
  const imageData = findImage(frameNode);
  if (imageData) {
    image.href = imageData.href;
  }
  return image;
}
function parseChart(objectNode, rowIndex, colIndex) {
  const attrs = parseAttributes(objectNode);
  return {
    type: "chart",
    href: attrs["xlink:href"],
    x: attrs["svg:x"] || "0",
    y: attrs["svg:y"] || "0",
    width: attrs["svg:width"] || "0",
    height: attrs["svg:height"] || "0",
    row: rowIndex,
    col: colIndex,
    zIndex: parseInt(attrs["draw:z-index"]) || 0
  };
}
function parseComment(annotationNode, rowIndex, colIndex) {
  const attrs = parseAttributes(annotationNode);
  const comment = {
    author: attrs["office:author"] || "",
    date: attrs["office:date"] || "",
    text: getTextContent(annotationNode)
  };
  if (attrs["office:display"] === "true") {
    comment.visible = true;
  }
  return comment;
}
function parseTable(tableNode, styles, numberFormats, tableStyles = {}) {
  const result = {
    name: "",
    data: [],
    columns: [],
    rows: {},
    cells: {},
    style: {},
    mergeCells: {},
    media: [],
    // Unified array for images, charts, and shapes
    comments: {}
  };
  if (!tableNode) return result;
  const attrs = parseAttributes(tableNode);
  result.name = attrs["table:name"] || "Sheet1";
  if (attrs["table:display"] === "false") {
    result.gridline = false;
  }
  const tableStyleName = attrs["table:style-name"];
  const tableStyle = tableStyles[tableStyleName];
  if (attrs["table:display"] === "false" || tableStyle?.display === "false") {
    result.worksheetState = "hidden";
  }
  if (attrs["table:protected"] === "true") {
    result.locked = true;
    if (attrs["table:protection-key"]) {
      result.locked = { password: true };
    }
  }
  if (tableStyleName && styles[tableStyleName] && styles[tableStyleName].table) {
    const tableProps = styles[tableStyleName].table;
    const defaultWidth = tableProps["table:default-column-width"];
    if (defaultWidth) {
      const match = defaultWidth.match(/^([\d.]+)(cm|in|pt|px)?$/);
      if (match) {
        const value = parseFloat(match[1]);
        const unit = match[2] || "cm";
        let widthPx;
        if (unit === "cm") {
          widthPx = Math.round(value * 37.795);
        } else if (unit === "in") {
          widthPx = Math.round(value * 96);
        } else if (unit === "pt") {
          widthPx = convertWidthToPixels(value, "pt");
        } else if (unit === "px") {
          widthPx = Math.round(value);
        }
        result.defaultColWidth = widthPx + "px";
      }
    }
  }
  let rowIndex = 0;
  let maxCol = 0;
  const processNode = (node) => {
    if (!node) return;
    if (Array.isArray(node)) {
      node.forEach(processNode);
      return;
    }
    if (node.type === "table:table-column") {
      const colAttrs = parseAttributes(node);
      let repeat = parseInt(colAttrs["table:number-columns-repeated"]) || 1;
      const styleName = colAttrs["table:style-name"];
      const visibility = colAttrs["table:visibility"];
      if (!styleName && repeat > 100) {
        repeat = 0;
      }
      for (let i = 0; i < repeat; i++) {
        let width = 100;
        if (styleName && styles[styleName] && styles[styleName].column) {
          const colProps = styles[styleName].column;
          const styleWidth = colProps["style:column-width"];
          if (styleWidth) {
            const match = styleWidth.match(/^([\d.]+)(cm|in|pt|px)?$/);
            if (match) {
              const value = parseFloat(match[1]);
              const unit = match[2] || "cm";
              if (unit === "cm") {
                width = Math.round(value * 37.795);
              } else if (unit === "in") {
                width = Math.round(value * 96);
              } else if (unit === "pt") {
                width = convertWidthToPixels(value, "pt");
              } else if (unit === "px") {
                width = Math.round(value);
              }
            }
          }
        }
        const col = {
          width,
          title: getColumnName(result.columns.length),
          type: "text"
        };
        if (styleName) {
          col.style = styleName;
        }
        if (visibility === "collapse" || visibility === "filter") {
          col.visible = false;
        }
        result.columns.push(col);
      }
    }
    if (node.type === "table:table-row") {
      const rowAttrs = parseAttributes(node);
      let rowRepeat = parseInt(rowAttrs["table:number-rows-repeated"]) || 1;
      const visibility = rowAttrs["table:visibility"];
      let isEmptyRow = true;
      if (node.children) {
        for (let cellNode of node.children) {
          if (cellNode.type === "table:table-cell" || cellNode.type === "table:covered-table-cell") {
            const cellAttrs = parseAttributes(cellNode);
            const valueType = cellAttrs["office:value-type"];
            const textContent = getTextContent(cellNode);
            if (valueType || textContent || cellAttrs["table:formula"]) {
              isEmptyRow = false;
              break;
            }
          }
        }
      }
      if (isEmptyRow && rowRepeat > 100) {
        rowRepeat = 0;
      }
      for (let r = 0; r < rowRepeat; r++) {
        const currentRow = rowIndex + r;
        result.data[currentRow] = [];
        if (visibility === "collapse" || visibility === "filter") {
          if (!result.rows[currentRow]) result.rows[currentRow] = {};
          result.rows[currentRow].visible = false;
        }
        let colIndex = 0;
        if (node.children) {
          node.children.forEach((cellNode) => {
            if (cellNode.type === "table:table-cell" || cellNode.type === "table:covered-table-cell") {
              const cellAttrs = parseAttributes(cellNode);
              let colRepeat = parseInt(cellAttrs["table:number-columns-repeated"]) || 1;
              const colspan = parseInt(cellAttrs["table:number-columns-spanned"]) || 1;
              const rowspan = parseInt(cellAttrs["table:number-rows-spanned"]) || 1;
              const valueType = cellAttrs["office:value-type"];
              const cellStyleName = cellAttrs["table:style-name"];
              let value = "";
              if (valueType === "string") {
                value = getTextContent(cellNode);
              } else if (valueType === "float") {
                value = parseFloat(cellAttrs["office:value"]) || 0;
              } else if (valueType === "percentage") {
                value = parseFloat(cellAttrs["office:value"]) || 0;
              } else if (valueType === "currency") {
                value = parseFloat(cellAttrs["office:value"]) || 0;
              } else if (valueType === "date") {
                value = cellAttrs["office:date-value"] || "";
              } else if (valueType === "time") {
                value = cellAttrs["office:time-value"] || "";
              } else if (valueType === "boolean") {
                value = cellAttrs["office:boolean-value"] === "true";
              } else {
                value = getTextContent(cellNode);
              }
              const formula = cellAttrs["table:formula"];
              const cellProtected = cellAttrs["table:protected"] === "true";
              let richText = null;
              let hyperlink = null;
              if (cellNode.children) {
                const findHyperlink = (n) => {
                  if (!n) return null;
                  if (Array.isArray(n)) {
                    for (let item of n) {
                      const result2 = findHyperlink(item);
                      if (result2) return result2;
                    }
                    return null;
                  }
                  if (typeof n === "object") {
                    if (n.type === "text:a") {
                      const attrs2 = parseAttributes(n);
                      return attrs2["xlink:href"];
                    }
                    if (n.children) return findHyperlink(n.children);
                  }
                  return null;
                };
                hyperlink = findHyperlink(cellNode.children);
                richText = parseRichText(cellNode);
                const findDrawElements = (n) => {
                  if (!n) return;
                  if (Array.isArray(n)) {
                    n.forEach(findDrawElements);
                    return;
                  }
                  if (typeof n === "object") {
                    if (n.type === "draw:frame") {
                      const image = parseImage(n, currentRow, colIndex);
                      if (image) {
                        image.type = "image";
                        result.media.push(image);
                      }
                    } else if (n.type === "draw:object") {
                      const chart = parseChart(n, currentRow, colIndex);
                      if (chart) {
                        chart.type = "chart";
                        result.media.push(chart);
                      }
                    } else if (n.type === "office:annotation") {
                      const cellAddr = getCellNameFromCoords(colIndex, currentRow);
                      result.comments[cellAddr] = parseComment(n, currentRow, colIndex);
                    }
                    if (n.children) findDrawElements(n.children);
                  }
                };
                findDrawElements(cellNode.children);
              }
              if (!valueType && !value && !formula && colRepeat > 100) {
                colRepeat = 0;
              }
              for (let c = 0; c < colRepeat; c++) {
                const col = colIndex + c;
                if (formula && c === 0) {
                  let cleanFormula = formula.replace(/^of:/, "");
                  cleanFormula = cleanFormula.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
                  cleanFormula = cleanFormula.replace(/\[\.(\$?)([A-Z]+)(\$?)(\d+):\.(\$?)([A-Z]+)(\$?)(\d+)\]/g, "$1$2$3$4:$5$6$7$8");
                  cleanFormula = cleanFormula.replace(/\[\.(\$?)([A-Z]+)(\$?)(\d+)\]/g, "$1$2$3$4");
                  cleanFormula = cleanFormula.replace(/;/g, ",");
                  result.data[currentRow][col] = cleanFormula;
                } else {
                  result.data[currentRow][col] = value;
                }
                if (col > maxCol) maxCol = col;
                if (cellStyleName) {
                  const cellAddr = getCellNameFromCoords(col, currentRow);
                  const cssStyle = styleToCSS(cellStyleName, styles);
                  if (cssStyle) {
                    result.style[cellAddr] = cssStyle;
                  }
                  const isNumericType = valueType === "float" || valueType === "percentage" || valueType === "currency" || valueType === "date" || valueType === "time";
                  if (isNumericType && styles[cellStyleName] && styles[cellStyleName].dataStyleName) {
                    const dataStyleName = styles[cellStyleName].dataStyleName;
                    if (numberFormats && numberFormats[dataStyleName]) {
                      const format = numberFormats[dataStyleName];
                      const isDefaultFormat = format === "0" || format === "General";
                      const hasDecimal = value && typeof value === "number" && value % 1 !== 0;
                      if (!isDefaultFormat || !hasDecimal) {
                        if (!result.cells[cellAddr]) result.cells[cellAddr] = {};
                        if (!result.cells[cellAddr].format) {
                          result.cells[cellAddr].format = format;
                        }
                      }
                    }
                  }
                }
                if (hyperlink && c === 0) {
                  const cellAddr = getCellNameFromCoords(col, currentRow);
                  if (!result.cells[cellAddr]) result.cells[cellAddr] = {};
                  result.cells[cellAddr].url = hyperlink;
                }
                if ((colspan > 1 || rowspan > 1) && c === 0) {
                  const cellAddr = getCellNameFromCoords(col, currentRow);
                  result.mergeCells[cellAddr] = [colspan, rowspan];
                }
              }
              colIndex += colRepeat;
            }
          });
        }
      }
      rowIndex += rowRepeat;
    }
    if (node.children) {
      processNode(node.children);
    }
  };
  if (tableNode.children) {
    processNode(tableNode.children);
  }
  for (let r = 0; r < result.data.length; r++) {
    if (!result.data[r]) result.data[r] = [];
    for (let c = 0; c <= maxCol; c++) {
      if (result.data[r][c] === void 0) {
        result.data[r][c] = "";
      }
    }
  }
  if (result.columns.length > maxCol + 1) {
    result.columns = result.columns.slice(0, maxCol + 1);
  }
  while (result.columns.length <= maxCol) {
    result.columns.push({
      width: 100,
      title: getColumnName(result.columns.length),
      type: "text"
    });
  }
  result.minDimensions = [maxCol + 1, result.data.length];
  return result;
}
async function parseODSWith(unzipArchive, input, options = {}) {
  return parse(async (inp) => {
    const buffer = await loadAsBuffer(inp);
    const files2 = await unzipArchive(buffer, {
      filter: (name) => name === "content.xml" || name === "styles.xml" || name === "settings.xml"
    });
    const contentBytes = files2["content.xml"];
    if (!contentBytes) {
      throw new Error("content.xml not found in ODS file");
    }
    const contentXml = new TextDecoder("utf-8").decode(contentBytes);
    const content = parser(contentXml);
    let styles = {};
    const stylesBytes = files2["styles.xml"];
    let stylesXml = "";
    if (stylesBytes) {
      stylesXml = new TextDecoder("utf-8").decode(stylesBytes);
      styles = parseStyles2(stylesXml);
    }
    const contentStyles = parseStyles2(contentXml);
    styles = { ...styles, ...contentStyles };
    let tableStyles = {};
    if (stylesBytes) {
      tableStyles = parseTableStyles(stylesXml);
    }
    const contentTableStyles = parseTableStyles(contentXml);
    tableStyles = { ...tableStyles, ...contentTableStyles };
    let numberFormats = {};
    if (stylesBytes) {
      numberFormats = parseNumberFormats(stylesXml);
    }
    const contentFormats = parseNumberFormats(contentXml);
    numberFormats = { ...numberFormats, ...contentFormats };
    const definedNames = {};
    const parseDocumentFeatures = (node) => {
      if (!node) return;
      if (Array.isArray(node)) {
        node.forEach(parseDocumentFeatures);
        return;
      }
      if (typeof node === "object") {
        if (node.type === "table:named-range") {
          const attrs = parseAttributes(node);
          const name = attrs["table:name"];
          const range = attrs["table:cell-range-address"];
          if (name && range) {
            definedNames[name] = range;
          }
        }
        if (node.children) {
          parseDocumentFeatures(node.children);
        }
      }
    };
    parseDocumentFeatures(content);
    let frozenPanes = {};
    const settingsBytes = files2["settings.xml"];
    if (settingsBytes) {
      const settingsXml = new TextDecoder("utf-8").decode(settingsBytes);
      const settings = parser(settingsXml);
      const findConfig = (node, sheetName) => {
        if (!node) return null;
        if (Array.isArray(node)) {
          for (let n of node) {
            const result2 = findConfig(n, sheetName);
            if (result2) return result2;
          }
          return null;
        }
        if (typeof node === "object") {
          if (node.type === "config:config-item-map-entry") {
            const attrs = parseAttributes(node);
            const name = attrs["config:name"];
            if (name && node.children) {
              const panes = {};
              const findPaneSettings = (n) => {
                if (!n) return;
                if (Array.isArray(n)) {
                  n.forEach(findPaneSettings);
                  return;
                }
                if (typeof n === "object" && n.type === "config:config-item") {
                  const itemAttrs = parseAttributes(n);
                  const itemName = itemAttrs["config:name"];
                  const value = getTextContent(n);
                  if (itemName === "HorizontalSplitMode" && value === "2") {
                    panes.horizontalSplit = true;
                  } else if (itemName === "VerticalSplitMode" && value === "2") {
                    panes.verticalSplit = true;
                  } else if (itemName === "PositionRight") {
                    panes.frozenColumns = parseInt(value) || 0;
                  } else if (itemName === "PositionBottom") {
                    panes.frozenRows = parseInt(value) || 0;
                  }
                  if (n.children) findPaneSettings(n.children);
                } else if (n.children) {
                  findPaneSettings(n.children);
                }
              };
              findPaneSettings(node.children);
              if (panes.frozenRows || panes.frozenColumns) {
                frozenPanes[name] = panes;
              }
            }
          }
          if (node.children) {
            const result2 = findConfig(node.children, sheetName);
            if (result2) return result2;
          }
        }
        return null;
      };
      findConfig(settings);
    }
    const worksheets = [];
    const allValidations = [];
    const findTables = (node) => {
      if (!node) return;
      if (Array.isArray(node)) {
        node.forEach(findTables);
        return;
      }
      if (typeof node === "object") {
        if (node.type === "table:table") {
          const tableData = parseTable(node, styles, numberFormats, tableStyles);
          const sheetFrozenPanes = frozenPanes[tableData.name] || {};
          const worksheet = {
            worksheetName: tableData.name,
            data: tableData.data,
            columns: tableData.columns,
            rows: tableData.rows,
            cells: tableData.cells,
            style: tableData.style,
            mergeCells: tableData.mergeCells,
            media: tableData.media,
            comments: tableData.comments,
            minDimensions: tableData.minDimensions,
            freezeRows: sheetFrozenPanes.frozenRows || 0,
            freezeColumns: sheetFrozenPanes.frozenColumns || 0
          };
          if (tableData.gridline !== void 0) {
            worksheet.gridline = tableData.gridline;
          }
          if (tableData.locked !== void 0) {
            worksheet.locked = tableData.locked;
          }
          if (tableData.defaultColWidth !== void 0) {
            worksheet.defaultColWidth = tableData.defaultColWidth;
          }
          if (tableData.worksheetState !== void 0) {
            worksheet.worksheetState = tableData.worksheetState;
          }
          if (tableData.validations) {
            tableData.validations.forEach((validation) => {
              if (!validation.range.includes("!")) {
                validation.range = `${tableData.name}!${validation.range}`;
              }
              allValidations.push(validation);
            });
          }
          worksheets.push(worksheet);
          return;
        }
        if (node.children) {
          findTables(node.children);
        }
      }
    };
    findTables(content);
    const styleNames = Object.keys(styles);
    const globalStyles = styleNames.map((name) => styleToCSS(name, styles));
    const styleIndexMap = {};
    styleNames.forEach((name, index) => {
      styleIndexMap[name] = index;
    });
    worksheets.forEach((worksheet) => {
      worksheet.columns.forEach((col) => {
        if (col.style && styleIndexMap[col.style] !== void 0) {
          col.style = styleIndexMap[col.style];
        } else if (col.style) {
          delete col.style;
        }
      });
      const cellsWithoutTopBorder = /* @__PURE__ */ new Set();
      const cellsWithoutLeftBorder = /* @__PURE__ */ new Set();
      Object.keys(worksheet.style).forEach((cellAddr) => {
        const cssString = worksheet.style[cellAddr];
        if (cssString.includes("border-right:")) {
          const coords = getCoordsFromCellName(cellAddr);
          if (coords && coords[0] !== null && coords[1] !== null) {
            const mergeInfo = worksheet.mergeCells?.[cellAddr];
            if (mergeInfo && Array.isArray(mergeInfo)) {
              const [colSpan, rowSpan] = mergeInfo;
              for (let rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
                const rightCell = getCellNameFromCoords(coords[0] + colSpan, coords[1] + rowOffset);
                cellsWithoutLeftBorder.add(rightCell);
              }
            } else {
              const rightCell = getCellNameFromCoords(coords[0] + 1, coords[1]);
              cellsWithoutLeftBorder.add(rightCell);
            }
          }
        }
        if (cssString.includes("border-bottom:")) {
          const coords = getCoordsFromCellName(cellAddr);
          if (coords && coords[0] !== null && coords[1] !== null) {
            const mergeInfo = worksheet.mergeCells?.[cellAddr];
            if (mergeInfo && Array.isArray(mergeInfo)) {
              const [colSpan, rowSpan] = mergeInfo;
              for (let colOffset = 0; colOffset < colSpan; colOffset++) {
                const belowCell = getCellNameFromCoords(coords[0] + colOffset, coords[1] + rowSpan);
                cellsWithoutTopBorder.add(belowCell);
              }
            } else {
              const belowCell = getCellNameFromCoords(coords[0], coords[1] + 1);
              cellsWithoutTopBorder.add(belowCell);
            }
          }
        }
      });
      Object.keys(worksheet.style).forEach((cellAddr) => {
        let cssString = worksheet.style[cellAddr];
        if (cellsWithoutTopBorder.has(cellAddr)) {
          cssString = cssString.replace(/border-top:\s*[^;]+;?\s*/g, "");
        }
        if (cellsWithoutLeftBorder.has(cellAddr)) {
          cssString = cssString.replace(/border-left:\s*[^;]+;?\s*/g, "");
        }
        worksheet.style[cellAddr] = cssString;
      });
      const newStyleObj = {};
      Object.keys(worksheet.style).forEach((cellAddr) => {
        const cssString = worksheet.style[cellAddr];
        const matchingIndex = globalStyles.findIndex((css) => css === cssString);
        if (matchingIndex !== -1) {
          newStyleObj[cellAddr] = matchingIndex;
        } else {
          globalStyles.push(cssString);
          newStyleObj[cellAddr] = globalStyles.length - 1;
        }
      });
      worksheet.style = newStyleObj;
    });
    const result = { worksheets };
    if (globalStyles.length > 0) {
      result.style = globalStyles;
    }
    if (Object.keys(definedNames).length > 0) {
      result.definedNames = definedNames;
    }
    if (allValidations.length > 0) {
      result.validations = allValidations;
    }
    return result;
  }, input, options);
}

// src/browser-sync.js
var SUPPORTED = "xlsx, xls and ods";
async function toBytes(input) {
  if (input instanceof Uint8Array) {
    return input;
  }
  if (typeof ArrayBuffer !== "undefined" && input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  if (typeof Blob !== "undefined" && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }
  throw new Error("Invalid input: expected a File, Blob, Uint8Array or ArrayBuffer");
}
function detect(bytes) {
  if (bytes.length < 4) {
    return null;
  }
  if (bytes[0] === 80 && bytes[1] === 75 && (bytes[2] === 3 || bytes[2] === 5 || bytes[2] === 7)) {
    const names = zipEntryNamesSync(bytes);
    if (names.some((n) => n === "[Content_Types].xml" || n.startsWith("xl/"))) {
      return "xlsx";
    }
    if (names.includes("mimetype") || names.includes("content.xml")) {
      return "ods";
    }
    return "xlsx";
  }
  if (bytes[0] === 208 && bytes[1] === 207 && bytes[2] === 17 && bytes[3] === 224) {
    return "xls";
  }
  return null;
}
async function parseTabular(file, options = {}) {
  if (!file) {
    throw new Error("Invalid file");
  }
  const bytes = await toBytes(file);
  if (options.maxFileSize && bytes.length > options.maxFileSize) {
    throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
  }
  let ext = null;
  if (options.extension) {
    ext = String(options.extension).toLowerCase().replace(/^\./, "");
  } else if (options.format && options.format !== "tabular" && options.format !== "rows") {
    ext = String(options.format).toLowerCase().replace(/^\./, "");
  } else if (typeof file.name === "string" && file.name.lastIndexOf(".") > -1) {
    ext = file.name.toLowerCase().slice(file.name.lastIndexOf(".") + 1);
  }
  if (!ext || !["xlsx", "xls", "ods"].includes(ext)) {
    const detected = detect(bytes);
    if (!detected && ext) {
      throw new Error(`Unsupported file type: ${ext} (spreadsheet-utils/browser-sync reads ${SUPPORTED})`);
    }
    if (!detected) {
      throw new Error(`Unsupported file type (spreadsheet-utils/browser-sync reads ${SUPPORTED})`);
    }
    ext = detected;
  }
  if (ext === "xlsx") {
    return parseXLSXSync(bytes, options);
  }
  if (ext === "xls") {
    const result = parseXLSBuffer(bytes);
    if (options.onload) {
      options.onload(result);
    }
    return result;
  }
  return parseODSWith(unzipArchiveSyncAsPromise, bytes, options);
}
async function readSpreadsheet(file, options = {}) {
  return shapeReadResult(await parseTabular(file, options), options);
}
function readSheet(input, sheet, options) {
  if (!options && sheet && typeof sheet !== "number" && typeof sheet !== "string") {
    options = sheet;
    sheet = void 0;
  }
  return readSpreadsheet(input, { ...options, sheet });
}
var readXlsxFile = readSpreadsheet;
var writeXlsxFile = writeXlsxFileSync;
export {
  InvalidInputError,
  InvalidSpreadsheetError,
  SheetNotFoundError,
  convertRowsToObjects,
  readSpreadsheet as default,
  parseTabular,
  readSheet,
  readSpreadsheet,
  readXlsxFile,
  writeXlsxFile
};
/**
 * Custom Compound File Binary (CFB) / OLE2 Parser
 *
 * Lightweight MIT-licensed implementation for reading Microsoft Compound File Binary format.
 * Based on [MS-CFB] Microsoft Compound File Binary File Format specification.
 *
 * Only implements READ functionality - no write support needed.
 * Focused on extracting streams from XLS files.
 *
 * @license MIT
 */
//# sourceMappingURL=browser-sync.mjs.map
