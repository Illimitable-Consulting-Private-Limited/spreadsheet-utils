var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/write/index.js
var write_exports = {};
__export(write_exports, {
  CSVTransformStream: () => CSVTransformStream,
  JSONLTransformStream: () => JSONLTransformStream,
  THEMES: () => THEMES,
  applyTheme: () => applyTheme,
  calculateAutoWidths: () => calculateAutoWidths,
  createDBF: () => createDBF,
  createHTML: () => createHTML,
  createODS: () => createODS,
  createXLS: () => createXLS,
  default: () => write_default,
  getSheetData: () => getSheetData,
  getTheme: () => getTheme,
  normalizeSchemaToColumns: () => normalizeSchemaToColumns,
  normalizeSchemaToDbfFields: () => normalizeSchemaToDbfFields,
  streamCSV: () => streamCSV,
  streamJSONL: () => streamJSONL,
  streamTSV: () => streamTSV,
  writeCSV: () => writeCSV,
  writeDBF: () => writeDBF,
  writeHTML: () => writeHTML,
  writeJSON: () => writeJSON,
  writeODS: () => writeODS,
  writeSpreadsheet: () => writeSpreadsheet,
  writeTSV: () => writeTSV,
  writeXLS: () => writeXLS,
  writeXlsxFile: () => writeXlsxFile2
});
module.exports = __toCommonJS(write_exports);

// src/write/export/writeXlsxFileNode.js
var import_node_fs = __toESM(require("node:fs"), 1);

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
    var transform4 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform4 && transform4.insert) {
      var insertedContent = transform4.insert(sheetOptions, properties);
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
    var transform4 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform4 && transform4.transform) {
      content = transform4.transform(content, sheetOptions, properties);
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
    var transform4 = feature.files && feature.files.transform && feature.files.transform[fileName];
    if (transform4 && transform4.transformElementAttributes) {
      attributes = transform4.transformElementAttributes(tagName, attributes || NO_ATTRIBUTES, index, sheetOptionsOrSheetsOptions, properties);
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
var objectConstructor = {}.constructor;
function isObject(object) {
  return object !== void 0 && object !== null && object.constructor === objectConstructor;
}

// src/write/xlsx/helpers/isCellObject.js
function isCellObject(cell) {
  return isObject(cell);
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
        var read = _ref3.read;
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

// src/write/xlsx/getFeatures.js
function getFeatures(customFeatures) {
  var features = [];
  features.push(conditionalFormatting_default);
  features.push(images_default);
  features.push(stickyRowsOrColumns_default);
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
  var read = _ref.read;
  var writtenFiles = {};
  for (var _iterator = _createForOfIteratorHelperLoose13(features), _step; !(_step = _iterator()).done; ) {
    var feature = _step.value;
    if (feature.files && feature.files.write) {
      if (feature.files.write.files) {
        var files2 = feature.files.write.files(sheetsOptions, {
          read
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
    } else if (isObject(arg1[0])) {
      if (isObject(arg2) && Array.isArray(arg2.schema)) {
        throw new Error("`schema` parameter was removed, use `columns` parameter instead");
      }
      if (isObject(arg2) && Array.isArray(arg2.columns)) {
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

// src/write/export/convertFilesContentToUint8Arrays.js
var import_fflate = require("fflate");
function convertFilesContentToUint8Arrays(files2, convertFileContentToUint8Array2) {
  var convertedFiles = {};
  return Promise.all(Object.keys(files2).map(function(key) {
    if (files2[key] instanceof Uint8Array) {
      convertedFiles[key] = files2[key];
    } else if (typeof files2[key] === "string") {
      convertedFiles[key] = convertStringToUint8Array(files2[key]);
    } else {
      return convertFileContentToUint8Array2(files2[key]).then(function(uint8Array) {
        convertedFiles[key] = uint8Array;
      });
    }
  })).then(function() {
    return convertedFiles;
  });
}
function convertStringToUint8Array(string) {
  return (0, import_fflate.strToU8)(string);
}

// src/write/export/convertFileContentToUint8ArrayNode.js
var import_node_stream = __toESM(require("node:stream"), 1);
var import_node_buffer = require("node:buffer");
function convertFileContentToUint8Array(fileContent) {
  if (fileContent instanceof import_node_stream.default) {
    return streamToUint8Array(fileContent);
  }
  if (fileContent instanceof import_node_buffer.Blob) {
    return blobToUint8Array(fileContent);
  }
  if (fileContent instanceof import_node_buffer.Buffer) {
    return Promise.resolve(fileContent);
  }
  throw new Error("Unsupported file content type. Expected a `Stream`, a `Blob` or a `Buffer`");
}
function streamToUint8Array(stream) {
  return new Promise(function(resolve, reject) {
    var chunks = [];
    stream.on("data", function(chunk) {
      chunks.push(chunk);
    });
    stream.on("end", function() {
      resolve(import_node_buffer.Buffer.concat(chunks));
    });
    stream.on("error", reject);
  });
}
function blobToUint8Array(blob) {
  return blob.arrayBuffer().then(function(arrayBuffer) {
    return new Uint8Array(arrayBuffer);
  });
}

// src/write/zip/zipToStream.js
var import_fflate2 = require("fflate");
var import_node_buffer2 = require("node:buffer");
var import_node_stream2 = require("node:stream");
var COMPRESS_FILES_IN_PARALLEL = false;
var PASS_FILE_CONTENT_AS_NODE_BUFFER = true;
function zipToStream(files2) {
  var stream = new import_node_stream2.PassThrough();
  var fileDataSources = {};
  var compressionOptions = void 0;
  var errored = false;
  var onError = function onError2(error) {
    if (!errored) {
      errored = true;
      stream.destroy(error);
    }
  };
  var start = function start2(zip2) {
    for (var _i = 0, _Object$keys = Object.keys(fileDataSources); _i < _Object$keys.length; _i++) {
      var filePath2 = _Object$keys[_i];
      var fileData = PASS_FILE_CONTENT_AS_NODE_BUFFER ? import_node_buffer2.Buffer.from(files2[filePath2]) : (0, import_fflate2.strToU8)(files2[filePath2]);
      fileDataSources[filePath2].push(fileData, true);
    }
    zip2.end();
  };
  var zipperCallback = function zipperCallback2(error, dataChunk, isEndOfArchive) {
    if (errored) {
      return;
    }
    if (error) {
      onError(error);
      return;
    }
    if (isEndOfArchive) {
      stream.end(dataChunk);
    } else {
      stream.write(dataChunk);
    }
  };
  var zip = new import_fflate2.Zip(zipperCallback);
  for (var _i2 = 0, _Object$keys2 = Object.keys(files2); _i2 < _Object$keys2.length; _i2++) {
    var filePath = _Object$keys2[_i2];
    var fileDataSource = COMPRESS_FILES_IN_PARALLEL ? new import_fflate2.AsyncZipDeflate(filePath, compressionOptions) : new import_fflate2.ZipDeflate(filePath, compressionOptions);
    fileDataSources[filePath] = fileDataSource;
    zip.add(fileDataSource);
  }
  start(zip);
  return stream;
}

// src/write/export/writeXlsxFileNode.js
function writeXlsxFile(arg1, arg2, arg3) {
  var files2 = generateXlsxFileContents(arg1, arg2, arg3);
  var createReadableStream = function createReadableStream2() {
    return convertFilesContentToUint8Arrays(files2, convertFileContentToUint8Array).then(function(files3) {
      return zipToStream(files3);
    });
  };
  return {
    toBuffer: function toBuffer() {
      return createReadableStream().then(convertStreamToBuffer);
    },
    toBlob: function toBlob() {
      return createReadableStream().then(convertStreamToBuffer).then(function(buffer) {
        return new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });
      });
    },
    toStream: function toStream(writableStream) {
      if (writableStream) {
        return createReadableStream().then(function(readableStream) {
          return pipe(readableStream, writableStream);
        });
      }
      return createReadableStream();
    },
    toFile: function toFile(filePath) {
      return createReadableStream().then(function(readableStream) {
        return pipe(readableStream, import_node_fs.default.createWriteStream(filePath));
      });
    }
  };
}
function convertStreamToBuffer(stream) {
  return new Promise(function(resolve, reject) {
    var chunks = [];
    stream.on("data", function(chunk) {
      chunks.push(chunk);
    });
    stream.on("end", function() {
      resolve(Buffer.concat(chunks));
    });
    stream.on("error", reject);
  });
}
function pipe(readableStream, writableStream) {
  return new Promise(function(resolve, reject) {
    readableStream.pipe(writableStream).on("finish", resolve).on("error", reject);
  });
}

// src/write/dbf.js
var import_node_fs2 = __toESM(require("node:fs"), 1);
function encodeString(str, encoding = "utf-8") {
  if (str === null || str === void 0) {
    return new Uint8Array(0);
  }
  const enc = String(encoding || "utf-8").toLowerCase().replace("-", "");
  if (typeof Buffer !== "undefined") {
    const nodeEnc = enc === "latin1" || enc === "iso88591" ? "latin1" : enc === "ascii" ? "ascii" : "utf8";
    return Buffer.from(String(str), nodeEnc);
  }
  if (enc === "latin1" || enc === "iso88591" || enc === "ascii") {
    const s = String(str);
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) {
      bytes[i] = s.charCodeAt(i) & 255;
    }
    return bytes;
  }
  return new TextEncoder().encode(String(str));
}
function sanitizeDbfFieldName(rawName, index = 0, existingNames = /* @__PURE__ */ new Set()) {
  let name = String(rawName || "").trim();
  if (!name) {
    name = `FIELD_${index + 1}`;
  } else {
    name = name.replace(/[\s\-\.]+/g, "_");
    name = name.replace(/[^a-zA-Z0-9_]/g, "");
    name = name.toUpperCase();
    if (/^[0-9]/.test(name)) {
      name = `F_${name}`;
    }
    if (!name) {
      name = `FIELD_${index + 1}`;
    }
  }
  let base = name.slice(0, 10);
  let candidate = base;
  let counter = 2;
  while (existingNames.has(candidate)) {
    const suffix = `_${counter}`;
    const maxBaseLen = 10 - suffix.length;
    candidate = `${base.slice(0, Math.max(1, maxBaseLen))}${suffix}`;
    counter++;
  }
  existingNames.add(candidate);
  return candidate;
}
function extractDecimalsFromFormat(format) {
  if (!format || typeof format !== "string") return void 0;
  const match = format.match(/\.([0#]+)/);
  if (match) {
    return match[1].length;
  }
  if (format === "0" || format === "#,##0") {
    return 0;
  }
  return void 0;
}
function mapTypeToDbf(type, options = {}) {
  if (!type) return null;
  if (type === String || type === "string" || type === "text" || type === "char" || type === "character" || type === "C") {
    return "C";
  }
  if (type === Number || type === "number" || type === "numeric" || type === "N") {
    return "N";
  }
  if (type === "integer" || type === "int" || type === "I") {
    return options.fileVersion === 48 || type === "I" ? "I" : "N";
  }
  if (type === "float" || type === "F") {
    return "F";
  }
  if (type === "double" || type === "B" || type === "O") {
    return "B";
  }
  if (type === Boolean || type === "boolean" || type === "bool" || type === "logical" || type === "L") {
    return "L";
  }
  if (type === Date || type === "date" || type === "D") {
    return "D";
  }
  if (type === "datetime" || type === "timestamp" || type === "T") {
    return "T";
  }
  if (type === "currency" || type === "Y") {
    return "Y";
  }
  if (type === "memo" || type === "M") {
    return "M";
  }
  return "C";
}
function getDefaultFieldSize(type, decimalPlaces = 0) {
  switch (type) {
    case "C":
      return 30;
    case "N":
      return decimalPlaces > 0 ? 12 : 10;
    case "F":
      return 16;
    case "I":
      return 4;
    case "L":
      return 1;
    case "D":
      return 8;
    case "T":
      return 8;
    case "B":
      return 8;
    case "Y":
      return 8;
    case "M":
      return 10;
    default:
      return 30;
  }
}
function extractRowsAndHeaders(input) {
  if (!input) return { rows: [], headers: [], is2D: false };
  if (input.worksheets && Array.isArray(input.worksheets) && input.worksheets[0]) {
    const ws = input.worksheets[0];
    const data = ws.data || [];
    if (data.length > 0) {
      const hasExplicitColumns = Array.isArray(ws.columns) && ws.columns.length > 0;
      const headers = hasExplicitColumns ? ws.columns.map((c) => c.name || c.title || "") : data[0].map((c) => String(c && typeof c === "object" && "value" in c ? c.value : c));
      const rows = hasExplicitColumns ? data : data.slice(1);
      return { rows, headers, is2D: true };
    }
    return { rows: [], headers: [], is2D: true };
  }
  if (!Array.isArray(input)) return { rows: [], headers: [], is2D: false };
  if (input.length === 0) return { rows: [], headers: [], is2D: false };
  if (Array.isArray(input[0])) {
    const headers = input[0].map((c) => String(c && typeof c === "object" && "value" in c ? c.value : c));
    const rows = input.slice(1);
    return { rows, headers, is2D: true };
  }
  if (typeof input[0] === "object" && input[0] !== null) {
    const keySet = /* @__PURE__ */ new Set();
    for (const obj of input) {
      if (obj && typeof obj === "object") {
        for (const k of Object.keys(obj)) {
          if (k !== "_deleted" && k !== "__deleted") {
            keySet.add(k);
          }
        }
      }
    }
    const headers = Array.from(keySet);
    return { rows: input, headers, is2D: false };
  }
  return { rows: [], headers: [], is2D: false };
}
function inferFieldTypeFromValues(values) {
  let hasString = false;
  let hasNumber = false;
  let hasDate = false;
  let hasBoolean = false;
  for (const raw of values) {
    const val = raw && typeof raw === "object" && "value" in raw && !(raw instanceof Date) ? raw.value : raw;
    if (val === null || val === void 0) continue;
    if (typeof val === "boolean") {
      hasBoolean = true;
    } else if (val instanceof Date) {
      hasDate = true;
    } else if (typeof val === "number" && !isNaN(val)) {
      hasNumber = true;
    } else if (typeof val === "string") {
      if (/^\d{4}-\d{2}-\d{2}$/.test(val.trim())) {
        hasDate = true;
      } else {
        hasString = true;
      }
    } else {
      hasString = true;
    }
  }
  if (hasString) return "C";
  if (hasNumber) return "N";
  if (hasDate) return "D";
  if (hasBoolean) return "L";
  return "C";
}
function normalizeSchemaToDbfFields(schema, data = [], options = {}) {
  const { rows, headers, is2D } = extractRowsAndHeaders(data);
  const existingNames = /* @__PURE__ */ new Set();
  const rawSchema = schema || options.schema || options.fieldDescriptors || options.columns;
  let fieldConfigs = [];
  if (Array.isArray(rawSchema)) {
    fieldConfigs = rawSchema.map((colDef, index) => {
      const rawHeader = colDef.name || colDef.column || colDef.header || colDef.title || colDef.prop || `FIELD_${index + 1}`;
      const prop = colDef.prop || colDef.name || colDef.column;
      const accessor = typeof colDef.value === "function" ? colDef.value : (row) => {
        if (is2D && Array.isArray(row)) {
          return row[index];
        }
        if (row && typeof row === "object") {
          if (prop && prop in row) return row[prop];
          if (rawHeader in row) return row[rawHeader];
        }
        return null;
      };
      return {
        rawName: rawHeader,
        prop,
        type: colDef.type,
        size: colDef.size || colDef.length || colDef.width,
        decimalPlaces: colDef.decimalPlaces !== void 0 ? colDef.decimalPlaces : colDef.decimals !== void 0 ? colDef.decimals : colDef.decimalCount,
        format: colDef.format,
        value: accessor,
        index
      };
    });
  } else if (rawSchema && typeof rawSchema === "object") {
    fieldConfigs = Object.entries(rawSchema).map(([propKey, colDef], index) => {
      const rawHeader = colDef.column || colDef.header || colDef.name || colDef.title || propKey;
      const accessor = typeof colDef.value === "function" ? colDef.value : (row) => {
        if (is2D && Array.isArray(row)) {
          return row[index];
        }
        if (row && typeof row === "object") {
          if (propKey in row) return row[propKey];
          if (rawHeader in row) return row[rawHeader];
        }
        return null;
      };
      return {
        rawName: rawHeader,
        prop: propKey,
        type: colDef.type,
        size: colDef.size || colDef.length || colDef.width,
        decimalPlaces: colDef.decimalPlaces !== void 0 ? colDef.decimalPlaces : colDef.decimals !== void 0 ? colDef.decimals : colDef.decimalCount,
        format: colDef.format,
        value: accessor,
        index
      };
    });
  } else {
    fieldConfigs = headers.map((header, index) => {
      const accessor = (row) => {
        if (is2D && Array.isArray(row)) {
          return row[index];
        }
        if (row && typeof row === "object") {
          return row[header];
        }
        return null;
      };
      return {
        rawName: header || `FIELD_${index + 1}`,
        prop: header,
        type: null,
        size: null,
        decimalPlaces: null,
        format: null,
        value: accessor,
        index
      };
    });
  }
  return fieldConfigs.map((cfg, index) => {
    const sanitizedName = sanitizeDbfFieldName(cfg.rawName, index, existingNames);
    const sampleValues = [];
    for (let r = 0; r < Math.min(rows.length, 1e3); r++) {
      const rowVal = cfg.value(rows[r], r);
      const val = rowVal && typeof rowVal === "object" && "value" in rowVal && !(rowVal instanceof Date) ? rowVal.value : rowVal;
      if (val !== null && val !== void 0) {
        sampleValues.push(val);
      }
    }
    let dbfType = mapTypeToDbf(cfg.type, options);
    if (!dbfType) {
      dbfType = inferFieldTypeFromValues(sampleValues);
    }
    let decPlaces = cfg.decimalPlaces;
    if (decPlaces === null || decPlaces === void 0) {
      const fmtDec = extractDecimalsFromFormat(cfg.format);
      if (fmtDec !== void 0) {
        decPlaces = fmtDec;
      } else if (dbfType === "N" || dbfType === "F") {
        let maxFoundDecimals = 0;
        let hasFloats = false;
        for (const val of sampleValues) {
          if (typeof val === "number" && !isNaN(val) && val % 1 !== 0) {
            hasFloats = true;
            const decStr = String(val).split(".")[1] || "";
            if (decStr.length > maxFoundDecimals) {
              maxFoundDecimals = Math.min(decStr.length, 15);
            }
          }
        }
        decPlaces = hasFloats ? Math.max(2, maxFoundDecimals) : 0;
      } else if (dbfType === "Y") {
        decPlaces = 4;
      } else {
        decPlaces = 0;
      }
    }
    let fieldSize = cfg.size;
    if (!fieldSize || fieldSize <= 0) {
      if (dbfType === "C") {
        let maxByteLen = 0;
        for (const val of sampleValues) {
          const str = String(val);
          const len = typeof Buffer !== "undefined" ? Buffer.byteLength(str, "utf-8") : new TextEncoder().encode(str).length;
          if (len > maxByteLen) maxByteLen = len;
        }
        fieldSize = Math.min(254, Math.max(10, maxByteLen || 30));
      } else if (dbfType === "N" || dbfType === "F") {
        let maxIntDigits = 1;
        let hasNeg = false;
        for (const val of sampleValues) {
          if (typeof val === "number" && !isNaN(val)) {
            if (val < 0) hasNeg = true;
            const intDigits = Math.floor(Math.abs(val)).toString().length;
            if (intDigits > maxIntDigits) maxIntDigits = intDigits;
          }
        }
        const neededSize = maxIntDigits + (decPlaces > 0 ? decPlaces + 1 : 0) + (hasNeg ? 1 : 0);
        fieldSize = Math.min(20, Math.max(decPlaces > 0 ? 12 : 10, neededSize));
      } else {
        fieldSize = getDefaultFieldSize(dbfType, decPlaces);
      }
    }
    if (dbfType === "C") {
      fieldSize = Math.min(254, Math.max(1, fieldSize));
    } else if (dbfType === "N" || dbfType === "F") {
      fieldSize = Math.min(20, Math.max(1, fieldSize));
    } else if (dbfType === "I") {
      fieldSize = 4;
      decPlaces = 0;
    } else if (dbfType === "L") {
      fieldSize = 1;
      decPlaces = 0;
    } else if (dbfType === "D" || dbfType === "T" || dbfType === "B" || dbfType === "Y") {
      fieldSize = 8;
      if (dbfType !== "Y") decPlaces = 0;
    } else if (dbfType === "M") {
      fieldSize = 10;
      decPlaces = 0;
    }
    return {
      name: sanitizedName,
      prop: cfg.prop,
      type: dbfType,
      size: fieldSize,
      decimalPlaces: decPlaces,
      value: cfg.value
    };
  });
}
function writeFieldValue(buffer, view, offset, rawVal, field, encoding) {
  let val = rawVal;
  if (val && typeof val === "object" && "value" in val && !(val instanceof Date)) {
    val = val.value;
  }
  switch (field.type) {
    case "C": {
      if (val === null || val === void 0) {
        buffer.fill(32, offset, offset + field.size);
      } else {
        const bytes = encodeString(String(val), encoding);
        const copyLen = Math.min(bytes.length, field.size);
        buffer.set(bytes.subarray(0, copyLen), offset);
        if (copyLen < field.size) {
          buffer.fill(32, offset + copyLen, offset + field.size);
        }
      }
      break;
    }
    case "N":
    case "F": {
      if (val === null || val === void 0 || typeof val === "number" && isNaN(val)) {
        buffer.fill(32, offset, offset + field.size);
      } else {
        const num = Number(val);
        let str = "";
        if (isNaN(num)) {
          str = "";
        } else if (field.decimalPlaces > 0) {
          str = num.toFixed(field.decimalPlaces);
        } else {
          str = Math.round(num).toString();
        }
        if (str.length > field.size) {
          buffer.fill(42, offset, offset + field.size);
        } else {
          const padded = str.padStart(field.size, " ");
          for (let i = 0; i < field.size; i++) {
            buffer[offset + i] = padded.charCodeAt(i);
          }
        }
      }
      break;
    }
    case "L": {
      if (val === true || val === "T" || val === "t" || val === "Y" || val === "y" || val === 1 || val === "1") {
        buffer[offset] = 84;
      } else if (val === false || val === "F" || val === "f" || val === "N" || val === "n" || val === 0 || val === "0") {
        buffer[offset] = 70;
      } else {
        buffer[offset] = 63;
      }
      break;
    }
    case "D": {
      let dateStr = null;
      if (val instanceof Date && !isNaN(val.getTime())) {
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, "0");
        const d = String(val.getDate()).padStart(2, "0");
        dateStr = `${String(y).padStart(4, "0")}${m}${d}`;
      } else if (typeof val === "string" && val.trim().length > 0) {
        const clean = val.replace(/[^0-9]/g, "");
        if (clean.length >= 8) {
          dateStr = clean.slice(0, 8);
        }
      } else if (typeof val === "number" && !isNaN(val)) {
        const dt = new Date(val);
        if (!isNaN(dt.getTime())) {
          const y = dt.getFullYear();
          const m = String(dt.getMonth() + 1).padStart(2, "0");
          const d = String(dt.getDate()).padStart(2, "0");
          dateStr = `${String(y).padStart(4, "0")}${m}${d}`;
        }
      }
      if (dateStr && dateStr.length === 8) {
        for (let i = 0; i < 8; i++) {
          buffer[offset + i] = dateStr.charCodeAt(i);
        }
      } else {
        buffer.fill(32, offset, offset + 8);
      }
      break;
    }
    case "I": {
      const intVal = val === null || val === void 0 ? 0 : Math.round(Number(val) || 0);
      view.setInt32(offset, intVal, true);
      break;
    }
    case "B": {
      const dblVal = val === null || val === void 0 ? 0 : Number(val) || 0;
      view.setFloat64(offset, dblVal, true);
      break;
    }
    case "T": {
      if (val) {
        const dt = val instanceof Date ? val : new Date(val);
        if (!isNaN(dt.getTime())) {
          const epochMs = dt.getTime();
          const julian = Math.floor(epochMs / 864e5) + 2440588;
          const ms = (epochMs % 864e5 + 864e5) % 864e5;
          view.setUint32(offset, julian, true);
          view.setUint32(offset + 4, ms, true);
        } else {
          buffer.fill(0, offset, offset + 8);
        }
      } else {
        buffer.fill(0, offset, offset + 8);
      }
      break;
    }
    case "Y": {
      const curNum = val === null || val === void 0 ? 0 : Number(val) || 0;
      const scaled = BigInt(Math.round(curNum * 1e4));
      view.setBigInt64(offset, scaled, true);
      break;
    }
    case "M": {
      const memoStr = val === null || val === void 0 ? "" : String(val);
      const padded = memoStr.padStart(10, " ").slice(-10);
      for (let i = 0; i < 10; i++) {
        buffer[offset + i] = padded.charCodeAt(i);
      }
      break;
    }
    default: {
      const strVal = val === null || val === void 0 ? "" : String(val);
      const bytes = encodeString(strVal, encoding);
      const copyLen = Math.min(bytes.length, field.size);
      buffer.set(bytes.subarray(0, copyLen), offset);
      if (copyLen < field.size) {
        buffer.fill(32, offset + copyLen, offset + field.size);
      }
      break;
    }
  }
}
function createDBF(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const { rows } = extractRowsAndHeaders(data);
  const fields = normalizeSchemaToDbfFields(
    options.schema || options.fieldDescriptors || options.columns,
    data,
    options
  );
  if (fields.length === 0) {
    throw new Error("Cannot export DBF: No columns or fields defined.");
  }
  const encoding = options.encoding || "utf-8";
  const fileVersion = typeof options.fileVersion === "number" ? options.fileVersion : options.fileVersion === "Visual FoxPro" ? 48 : 3;
  const headerLength = 32 + fields.length * 32 + 1;
  const recordLength = 1 + fields.reduce((sum, f) => sum + f.size, 0);
  const totalLength = headerLength + rows.length * recordLength + 1;
  const buffer = new Uint8Array(totalLength);
  const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
  buffer[0] = fileVersion & 255;
  const lastUpdate = options.lastUpdate instanceof Date ? options.lastUpdate : /* @__PURE__ */ new Date();
  buffer[1] = lastUpdate.getFullYear() - 1900 & 255;
  buffer[2] = lastUpdate.getMonth() + 1 & 255;
  buffer[3] = lastUpdate.getDate() & 255;
  view.setUint32(4, rows.length, true);
  view.setUint16(8, headerLength, true);
  view.setUint16(10, recordLength, true);
  const encLower = encoding.toLowerCase();
  const defaultLangDriver = encLower.includes("latin") || encLower.includes("1252") || encLower.includes("iso") ? 3 : 0;
  buffer[29] = options.languageDriver !== void 0 ? options.languageDriver : defaultLangDriver;
  for (let i = 0; i < fields.length; i++) {
    const field = fields[i];
    const fieldOffset = 32 + i * 32;
    const nameBytes = encodeString(field.name, "ascii");
    const nameLen = Math.min(nameBytes.length, 10);
    buffer.set(nameBytes.subarray(0, nameLen), fieldOffset);
    buffer.fill(0, fieldOffset + nameLen, fieldOffset + 11);
    buffer[fieldOffset + 11] = field.type.charCodeAt(0);
    buffer[fieldOffset + 16] = field.size & 255;
    buffer[fieldOffset + 17] = (field.decimalPlaces || 0) & 255;
  }
  buffer[32 + fields.length * 32] = 13;
  let recOffset = headerLength;
  const deletedSet = Array.isArray(options.deletedRows) ? new Set(options.deletedRows) : null;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const isDeleted = Boolean(row && (row._deleted === true || row.__deleted === true)) || typeof options.isRowDeleted === "function" && options.isRowDeleted(row, r) || deletedSet && deletedSet.has(r);
    buffer[recOffset] = isDeleted ? 42 : 32;
    let colOffset = recOffset + 1;
    for (let f = 0; f < fields.length; f++) {
      const field = fields[f];
      const rawVal = field.value(row, r);
      writeFieldValue(buffer, view, colOffset, rawVal, field, encoding);
      colOffset += field.size;
    }
    recOffset += recordLength;
  }
  buffer[totalLength - 1] = 26;
  return buffer;
}
async function writeDBF(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const uint8Array = createDBF(data, options);
  if (options.filePath) {
    await import_node_fs2.default.promises.writeFile(options.filePath, uint8Array);
    return options.filePath;
  }
  if (options.uint8Array) {
    return uint8Array;
  }
  if (typeof Buffer !== "undefined") {
    return Buffer.from(uint8Array.buffer, uint8Array.byteOffset, uint8Array.byteLength);
  }
  return uint8Array;
}

// src/write/schema.js
var STYLE_PROPERTIES = [
  "fontWeight",
  "fontStyle",
  "fontSize",
  "fontFamily",
  "color",
  "textColor",
  "backgroundColor",
  "align",
  "verticalAlign",
  "wrap",
  "borderColor",
  "borderStyle",
  "leftBorderStyle",
  "rightBorderStyle",
  "topBorderStyle",
  "bottomBorderStyle",
  "leftBorderColor",
  "rightBorderColor",
  "topBorderColor",
  "bottomBorderColor",
  "columnSpan",
  "rowSpan",
  "span"
];
function extractStyles(def) {
  if (!def || typeof def !== "object") return {};
  const styles = {};
  for (const key of STYLE_PROPERTIES) {
    if (def[key] !== void 0) {
      styles[key] = def[key];
    }
  }
  if (def.style && typeof def.style === "object") {
    Object.assign(styles, def.style);
  }
  return styles;
}
function buildHeaderCell(colDef, fallbackName, options = {}) {
  const rawHeader = colDef.header !== void 0 ? colDef.header : colDef.column !== void 0 ? colDef.column : colDef.name !== void 0 ? colDef.name : fallbackName;
  const headerStyles = {
    ...options.headerStyle || {},
    ...colDef.headerStyle || {}
  };
  if (rawHeader && typeof rawHeader === "object" && !(rawHeader instanceof Date)) {
    return {
      value: rawHeader.value !== void 0 ? rawHeader.value : fallbackName,
      ...headerStyles,
      ...rawHeader
    };
  }
  if (Object.keys(headerStyles).length > 0) {
    return {
      value: rawHeader !== void 0 && rawHeader !== null ? rawHeader : fallbackName,
      ...headerStyles
    };
  }
  return rawHeader !== void 0 && rawHeader !== null ? rawHeader : fallbackName;
}
function normalizeSchemaToColumns(schema, options = {}) {
  if (Array.isArray(schema)) {
    return schema.map((col) => {
      const baseStyles = extractStyles(col);
      const header = buildHeaderCell(col, "", options);
      return {
        header,
        cell: (row, index) => {
          if (typeof col.cell === "function") {
            const userCell = col.cell(row, index);
            if (userCell && typeof userCell === "object" && !(userCell instanceof Date)) {
              return {
                type: col.type,
                format: col.format,
                ...baseStyles,
                ...userCell
              };
            }
            return {
              value: userCell !== void 0 ? userCell : null,
              type: col.type,
              format: col.format,
              ...baseStyles
            };
          }
          const val = typeof col.value === "function" ? col.value(row, index) : col.prop ? row[col.prop] : col.column ? row[col.column] : col.name ? row[col.name] : null;
          const dynamicStyles = typeof col.style === "function" ? col.style(val, row, index) : {};
          if (val && typeof val === "object" && !(val instanceof Date) && "value" in val) {
            return {
              type: col.type,
              format: col.format,
              ...baseStyles,
              ...dynamicStyles,
              ...val
            };
          }
          return {
            value: val !== void 0 ? val : null,
            type: col.type,
            format: col.format,
            ...baseStyles,
            ...dynamicStyles
          };
        },
        width: col.width
      };
    });
  }
  if (schema && typeof schema === "object") {
    return Object.entries(schema).map(([propKey, colDef]) => {
      const baseStyles = extractStyles(colDef);
      const header = buildHeaderCell(colDef, propKey, options);
      return {
        header,
        cell: (row, index) => {
          if (typeof colDef.cell === "function") {
            const userCell = colDef.cell(row, index);
            if (userCell && typeof userCell === "object" && !(userCell instanceof Date)) {
              return {
                type: colDef.type,
                format: colDef.format,
                ...baseStyles,
                ...userCell
              };
            }
            return {
              value: userCell !== void 0 ? userCell : null,
              type: colDef.type,
              format: colDef.format,
              ...baseStyles
            };
          }
          const val = typeof colDef.value === "function" ? colDef.value(row, index) : row ? row[propKey] : null;
          const dynamicStyles = typeof colDef.style === "function" ? colDef.style(val, row, index) : {};
          if (val && typeof val === "object" && !(val instanceof Date) && "value" in val) {
            return {
              type: colDef.type,
              format: colDef.format,
              ...baseStyles,
              ...dynamicStyles,
              ...val
            };
          }
          return {
            value: val !== void 0 ? val : null,
            type: colDef.type,
            format: colDef.format,
            ...baseStyles,
            ...dynamicStyles
          };
        },
        width: colDef.width
      };
    });
  }
  return [];
}
function getCellLength(cell) {
  if (cell == null) return 0;
  if (typeof cell === "object") {
    if (cell instanceof Date) {
      return 10;
    }
    if (cell.value !== void 0 && cell.value !== null) {
      return getCellLength(cell.value);
    }
    return 0;
  }
  if (typeof cell === "number") {
    return String(cell).length;
  }
  if (typeof cell === "boolean") {
    return cell ? 4 : 5;
  }
  const str = String(cell);
  if (str.includes("\n")) {
    const lines = str.split("\n");
    let maxL = 0;
    for (const line of lines) {
      if (line.length > maxL) maxL = line.length;
    }
    return maxL;
  }
  return str.length;
}
function calculateAutoWidths(rows, options = {}, existingColumns = null) {
  if (!Array.isArray(rows) || rows.length === 0) {
    return existingColumns || [];
  }
  const autoWidthConfig = typeof options.autoWidth === "object" && options.autoWidth !== null ? options.autoWidth : {};
  const minWidth = autoWidthConfig.minWidth ?? options.minWidth ?? 8;
  const maxWidth = autoWidthConfig.maxWidth ?? options.maxWidth ?? 40;
  const padding = autoWidthConfig.padding ?? options.padding ?? 2;
  const force = autoWidthConfig.force === true;
  let numCols = 0;
  if (Array.isArray(existingColumns)) {
    numCols = existingColumns.length;
  }
  for (const row of rows) {
    if (Array.isArray(row) && row.length > numCols) {
      numCols = row.length;
    }
  }
  if (numCols === 0) return existingColumns || [];
  const colWidths = new Array(numCols).fill(0);
  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const len = getCellLength(row[c]);
      if (len > colWidths[c]) {
        colWidths[c] = len;
      }
    }
  }
  const result = [];
  for (let c = 0; c < numCols; c++) {
    const existingCol = existingColumns && existingColumns[c] ? { ...existingColumns[c] } : {};
    if (!force && existingCol.width != null) {
      result.push(existingCol);
    } else {
      const calculated = Math.max(minWidth, Math.min(maxWidth, colWidths[c] + padding));
      existingCol.width = calculated;
      result.push(existingCol);
    }
  }
  return result;
}

// src/write/themes.js
var THEMES = {
  ocean: {
    name: "ocean",
    header: {
      backgroundColor: "#1B365D",
      color: "#FFFFFF",
      fontWeight: "bold",
      borderColor: "#CBD5E1",
      borderStyle: "thin"
    },
    alternateRow: {
      backgroundColor: "#F0F4F8"
    },
    borderColor: "#E2E8F0"
  },
  forest: {
    name: "forest",
    header: {
      backgroundColor: "#1B4332",
      color: "#FFFFFF",
      fontWeight: "bold",
      borderColor: "#D8F3DC",
      borderStyle: "thin"
    },
    alternateRow: {
      backgroundColor: "#F2F9F4"
    },
    borderColor: "#E0EFE6"
  },
  slate: {
    name: "slate",
    header: {
      backgroundColor: "#334155",
      color: "#FFFFFF",
      fontWeight: "bold",
      borderColor: "#CBD5E1",
      borderStyle: "thin"
    },
    alternateRow: {
      backgroundColor: "#F8FAFC"
    },
    borderColor: "#E2E8F0"
  },
  sunset: {
    name: "sunset",
    header: {
      backgroundColor: "#9C4123",
      color: "#FFFFFF",
      fontWeight: "bold",
      borderColor: "#FED7AA",
      borderStyle: "thin"
    },
    alternateRow: {
      backgroundColor: "#FFF7ED"
    },
    borderColor: "#FEE2E2"
  },
  striped: {
    name: "striped",
    header: {
      backgroundColor: "#475569",
      color: "#FFFFFF",
      fontWeight: "bold",
      borderColor: "#CBD5E1",
      borderStyle: "thin"
    },
    alternateRow: {
      backgroundColor: "#F1F5F9"
    },
    borderColor: "#E2E8F0"
  }
};
function getTheme(theme) {
  if (!theme) return null;
  if (typeof theme === "string") {
    const key = theme.toLowerCase().trim();
    return THEMES[key] || null;
  }
  if (typeof theme === "object") {
    return theme;
  }
  return null;
}
function applyTheme(rows, themeInput, options = {}) {
  const theme = getTheme(themeInput);
  if (!theme || !Array.isArray(rows) || rows.length === 0) {
    return rows;
  }
  const hasHeader = options.hasHeader !== false;
  const headerStyle = theme.header || {};
  const alternateStyle = theme.alternateRow || {};
  const zebra = options.zebra !== false && Boolean(alternateStyle.backgroundColor);
  return rows.map((row, rIdx) => {
    if (!Array.isArray(row)) return row;
    const isHeader = hasHeader && rIdx === 0;
    const isAlternate = !isHeader && zebra && (hasHeader ? rIdx % 2 === 0 : rIdx % 2 === 1);
    return row.map((cell) => {
      if (isHeader) {
        if (cell && typeof cell === "object" && !(cell instanceof Date)) {
          return { ...headerStyle, ...cell };
        }
        return { value: cell !== void 0 ? cell : null, ...headerStyle };
      }
      if (isAlternate) {
        if (cell && typeof cell === "object" && !(cell instanceof Date)) {
          return {
            backgroundColor: alternateStyle.backgroundColor,
            ...cell
          };
        }
        return {
          value: cell !== void 0 ? cell : null,
          backgroundColor: alternateStyle.backgroundColor
        };
      }
      return cell;
    });
  });
}

// src/write/formats.js
var import_node_fs5 = __toESM(require("node:fs"), 1);

// src/write/ods.js
var import_node_fs3 = __toESM(require("node:fs"), 1);
var import_fflate3 = require("fflate");
function escapeXml(str) {
  if (str === null || str === void 0) return "";
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function normalizeWorkbookSheets(data, options = {}) {
  let sheets = [];
  if (data && Array.isArray(data.worksheets)) {
    sheets = data.worksheets.map((ws, i) => {
      const sheetName = ws.worksheetName || ws.sheet || ws.name || `Sheet${i + 1}`;
      let rows = ws.data || [];
      let columns = ws.columns;
      if (ws.columns && Array.isArray(ws.columns) && rows.length > 0 && !Array.isArray(rows[0])) {
        columns = normalizeSchemaToColumns(ws.columns, { ...options, ...ws });
        rows = getSheetData(rows, columns);
      }
      return { name: sheetName, rows, columns };
    });
  } else if (Array.isArray(data) && data.length > 0 && typeof data[0] === "object" && data[0] !== null && !Array.isArray(data[0]) && (data[0].sheet || data[0].data)) {
    sheets = data.map((s, i) => {
      const sheetName = s.sheet || s.name || `Sheet${i + 1}`;
      let rows = s.data || [];
      const schema = s.schema || s.columns || options.schema;
      let columns = s.columns;
      if (schema && rows.length > 0 && !Array.isArray(rows[0])) {
        columns = normalizeSchemaToColumns(schema, { ...options, ...s });
        rows = getSheetData(rows, columns);
      }
      return { name: sheetName, rows, columns };
    });
  } else if (Array.isArray(data) && data.length > 0 && !Array.isArray(data[0])) {
    const schema = options.schema || options.columns;
    if (schema) {
      const columns = normalizeSchemaToColumns(schema, options);
      const grid = getSheetData(data, columns);
      sheets = [{
        name: options.sheet || "Sheet1",
        rows: grid,
        columns
      }];
    } else {
      const keys = Object.keys(data[0] || {});
      const headerStyle = options.headerStyle || { fontWeight: "bold" };
      const headerRow = keys.map((k) => ({ value: k, ...headerStyle }));
      const dataRows = data.map((obj) => keys.map((k) => ({ value: obj[k] })));
      sheets = [{
        name: options.sheet || "Sheet1",
        rows: [headerRow, ...dataRows]
      }];
    }
  } else if (Array.isArray(data)) {
    sheets = [{
      name: options.sheet || "Sheet1",
      rows: data,
      columns: options.columns
    }];
  } else {
    sheets = [{
      name: options.sheet || "Sheet1",
      rows: []
    }];
  }
  sheets.forEach((sheet) => {
    if (options.theme) {
      sheet.rows = applyTheme(sheet.rows, options.theme, options);
    }
    if (options.autoWidth) {
      sheet.columns = calculateAutoWidths(sheet.rows, options, sheet.columns);
    }
  });
  return sheets;
}
function formatBorderWidth(b) {
  if (!b) return "0.06pt";
  const s = String(b).toLowerCase();
  if (s === "medium") return "1.5pt";
  if (s === "thick") return "2.5pt";
  if (s.includes("px") || s.includes("pt")) return s.split(" ")[0];
  return "0.06pt";
}
function formatBorderStyle(b) {
  if (!b) return "solid";
  const s = String(b).toLowerCase();
  if (s.includes("dashed")) return "dashed";
  if (s.includes("dotted")) return "dotted";
  if (s.includes("double")) return "double";
  return "solid";
}
function extractCellStyle(cell, isDate) {
  if (!cell || typeof cell !== "object" || cell instanceof Date) {
    return isDate ? { isDate: true } : null;
  }
  const style = {};
  if (isDate) style.isDate = true;
  if (cell.backgroundColor) style.backgroundColor = cell.backgroundColor;
  if (cell.textColor || cell.color) style.color = cell.textColor || cell.color;
  if (cell.fontWeight === "bold" || cell.bold) style.bold = true;
  if (cell.fontStyle === "italic" || cell.italic) style.italic = true;
  if (cell.fontSize) style.fontSize = cell.fontSize;
  if (cell.fontFamily) style.fontFamily = cell.fontFamily;
  if (cell.align) style.align = cell.align;
  if (cell.borderStyle || cell.borderColor) {
    style.border = `${formatBorderWidth(cell.borderStyle)} ${formatBorderStyle(cell.borderStyle)} ${cell.borderColor || "#000000"}`;
  }
  if (cell.leftBorderStyle || cell.leftBorderColor) {
    style.borderLeft = `${formatBorderWidth(cell.leftBorderStyle)} ${formatBorderStyle(cell.leftBorderStyle)} ${cell.leftBorderColor || "#000000"}`;
  }
  if (cell.rightBorderStyle || cell.rightBorderColor) {
    style.borderRight = `${formatBorderWidth(cell.rightBorderStyle)} ${formatBorderStyle(cell.rightBorderStyle)} ${cell.rightBorderColor || "#000000"}`;
  }
  if (cell.topBorderStyle || cell.topBorderColor) {
    style.borderTop = `${formatBorderWidth(cell.topBorderStyle)} ${formatBorderStyle(cell.topBorderStyle)} ${cell.topBorderColor || "#000000"}`;
  }
  if (cell.bottomBorderStyle || cell.bottomBorderColor) {
    style.borderBottom = `${formatBorderWidth(cell.bottomBorderStyle)} ${formatBorderStyle(cell.bottomBorderStyle)} ${cell.bottomBorderColor || "#000000"}`;
  }
  return Object.keys(style).length > 0 ? style : null;
}
function serializeOdsStyle(styleName, style) {
  let xml = `    <style:style style:name="${escapeXml(styleName)}" style:family="table-cell"`;
  if (style.isDate) {
    xml += ` style:data-style-name="N_DATE"`;
  }
  xml += `>
`;
  const cellProps = [];
  if (style.backgroundColor) {
    cellProps.push(`fo:background-color="${escapeXml(style.backgroundColor)}"`);
  }
  if (style.border) {
    cellProps.push(`fo:border="${escapeXml(style.border)}"`);
  }
  if (style.borderLeft) {
    cellProps.push(`fo:border-left="${escapeXml(style.borderLeft)}"`);
  }
  if (style.borderRight) {
    cellProps.push(`fo:border-right="${escapeXml(style.borderRight)}"`);
  }
  if (style.borderTop) {
    cellProps.push(`fo:border-top="${escapeXml(style.borderTop)}"`);
  }
  if (style.borderBottom) {
    cellProps.push(`fo:border-bottom="${escapeXml(style.borderBottom)}"`);
  }
  if (cellProps.length > 0) {
    xml += `      <style:table-cell-properties ${cellProps.join(" ")}/>
`;
  }
  const textProps = [];
  if (style.color) {
    textProps.push(`fo:color="${escapeXml(style.color)}"`);
  }
  if (style.bold) {
    textProps.push(`fo:font-weight="bold"`);
  }
  if (style.italic) {
    textProps.push(`fo:font-style="italic"`);
  }
  if (style.fontSize) {
    const sz = typeof style.fontSize === "number" ? `${style.fontSize}pt` : style.fontSize;
    textProps.push(`fo:font-size="${escapeXml(sz)}"`);
  }
  if (style.fontFamily) {
    textProps.push(`fo:font-family="${escapeXml(style.fontFamily)}"`);
  }
  if (textProps.length > 0) {
    xml += `      <style:text-properties ${textProps.join(" ")}/>
`;
  }
  if (style.align) {
    const alignVal = style.align === "center" ? "center" : style.align === "right" ? "end" : "start";
    xml += `      <style:paragraph-properties fo:text-align="${alignVal}"/>
`;
  }
  xml += `    </style:style>
`;
  return xml;
}
function buildContentXml(sheets) {
  let xml = `<?xml version="1.0" encoding="UTF-8"?>
`;
  xml += `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2">
`;
  const stylesMap = /* @__PURE__ */ new Map();
  let nextStyleId = 1;
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        let val = cell;
        let isDate = false;
        if (cell && typeof cell === "object") {
          val = cell.value;
          if (val instanceof Date || cell.type === Date) isDate = true;
        } else if (cell instanceof Date) {
          isDate = true;
        }
        const st = extractCellStyle(cell, isDate);
        if (st) {
          const key = JSON.stringify(st);
          if (!stylesMap.has(key)) {
            stylesMap.set(key, { name: `s_${nextStyleId++}`, style: st });
          }
        }
      }
    }
  }
  xml += `  <office:automatic-styles>
`;
  xml += `    <number:date-style style:name="N_DATE">
`;
  xml += `      <number:year number:style="long"/>
`;
  xml += `      <number:text>-</number:text>
`;
  xml += `      <number:month number:style="long"/>
`;
  xml += `      <number:text>-</number:text>
`;
  xml += `      <number:day number:style="long"/>
`;
  xml += `    </number:date-style>
`;
  const colStylesMap = /* @__PURE__ */ new Map();
  let nextColStyleId = 1;
  for (const sheet of sheets) {
    if (sheet.columns && Array.isArray(sheet.columns)) {
      sheet.columns.forEach((col) => {
        if (col && typeof col.width === "number") {
          const widthCm = (col.width * 0.2).toFixed(2);
          if (!colStylesMap.has(widthCm)) {
            colStylesMap.set(widthCm, `co_${nextColStyleId++}`);
          }
        }
      });
    }
  }
  for (const [widthCm, styleName] of colStylesMap.entries()) {
    xml += `    <style:style style:name="${styleName}" style:family="table-column">
`;
    xml += `      <style:table-column-properties style:column-width="${widthCm}cm"/>
`;
    xml += `    </style:style>
`;
  }
  for (const { name, style } of stylesMap.values()) {
    xml += serializeOdsStyle(name, style);
  }
  xml += `  </office:automatic-styles>
`;
  xml += `  <office:body>
`;
  xml += `    <office:spreadsheet>
`;
  for (const sheet of sheets) {
    xml += `      <table:table table:name="${escapeXml(sheet.name)}">
`;
    if (sheet.columns && Array.isArray(sheet.columns)) {
      sheet.columns.forEach((col) => {
        if (col && typeof col.width === "number") {
          const widthCm = (col.width * 0.2).toFixed(2);
          const styleName = colStylesMap.get(widthCm);
          xml += `        <table:table-column table:style-name="${styleName}"/>
`;
        } else {
          xml += `        <table:table-column/>
`;
        }
      });
    }
    const coveredCells = /* @__PURE__ */ new Set();
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      xml += `        <table:table-row>
`;
      if (Array.isArray(row)) {
        for (let c = 0; c < row.length; c++) {
          const cellKey = `${r},${c}`;
          if (coveredCells.has(cellKey)) {
            xml += `          <table:covered-table-cell/>
`;
            continue;
          }
          const cell = row[c];
          let val = cell;
          let type = null;
          let isDate = false;
          let colSpan = 1;
          let rowSpan = 1;
          if (cell && typeof cell === "object" && !(cell instanceof Date)) {
            val = cell.value;
            type = cell.type;
            if (val instanceof Date || type === Date) isDate = true;
            colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
            rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);
          } else if (cell instanceof Date) {
            isDate = true;
          }
          let spanAttr = "";
          if (colSpan > 1) spanAttr += ` table:number-columns-spanned="${colSpan}"`;
          if (rowSpan > 1) spanAttr += ` table:number-rows-spanned="${rowSpan}"`;
          if (colSpan > 1 || rowSpan > 1) {
            for (let dr = 0; dr < rowSpan; dr++) {
              for (let dc = 0; dc < colSpan; dc++) {
                if (dr > 0 || dc > 0) {
                  coveredCells.add(`${r + dr},${c + dc}`);
                }
              }
            }
          }
          let styleAttr = "";
          const st = extractCellStyle(cell, isDate);
          if (st) {
            const styleEntry = stylesMap.get(JSON.stringify(st));
            if (styleEntry) {
              styleAttr = ` table:style-name="${styleEntry.name}"`;
            }
          }
          if (val === null || val === void 0 || val === "") {
            xml += `          <table:table-cell${spanAttr}${styleAttr}/>
`;
          } else if (isDate) {
            const d = val instanceof Date ? val : new Date(val);
            const isoDate = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : "";
            xml += `          <table:table-cell office:value-type="date" office:date-value="${isoDate}"${spanAttr}${styleAttr}><text:p>${isoDate}</text:p></table:table-cell>
`;
          } else if (typeof val === "number" || type === Number) {
            const num = Number(val);
            xml += `          <table:table-cell office:value-type="float" office:value="${num}"${spanAttr}${styleAttr}><text:p>${num}</text:p></table:table-cell>
`;
          } else if (typeof val === "boolean" || type === Boolean) {
            const boolVal = Boolean(val);
            xml += `          <table:table-cell office:value-type="boolean" office:boolean-value="${boolVal}"${spanAttr}${styleAttr}><text:p>${boolVal ? "TRUE" : "FALSE"}</text:p></table:table-cell>
`;
          } else {
            const strVal = String(val);
            xml += `          <table:table-cell office:value-type="string"${spanAttr}${styleAttr}><text:p>${escapeXml(strVal)}</text:p></table:table-cell>
`;
          }
        }
      }
      xml += `        </table:table-row>
`;
    }
    xml += `      </table:table>
`;
  }
  xml += `    </office:spreadsheet>
`;
  xml += `  </office:body>
`;
  xml += `</office:document-content>`;
  return xml;
}
function createODS(data, options = {}) {
  const sheets = normalizeWorkbookSheets(data, options);
  const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">
  <office:styles>
    <style:default-style style:family="table-cell">
      <style:paragraph-properties fo:text-align="start"/>
    </style:default-style>
  </office:styles>
</office:document-styles>`;
  const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2">
  <office:meta>
    <meta:generator>spreadsheet-utils</meta:generator>
    <meta:creation-date>${(/* @__PURE__ */ new Date()).toISOString()}</meta:creation-date>
  </office:meta>
</office:document-meta>`;
  const contentXml = buildContentXml(sheets);
  const files2 = {
    "mimetype": [(0, import_fflate3.strToU8)("application/vnd.oasis.opendocument.spreadsheet"), { level: 0 }],
    "META-INF/manifest.xml": (0, import_fflate3.strToU8)(manifestXml),
    "content.xml": (0, import_fflate3.strToU8)(contentXml),
    "styles.xml": (0, import_fflate3.strToU8)(stylesXml),
    "meta.xml": (0, import_fflate3.strToU8)(metaXml)
  };
  return (0, import_fflate3.zipSync)(files2);
}
async function writeODS(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const uint8Array = createODS(data, options);
  if (options.filePath) {
    await import_node_fs3.default.promises.writeFile(options.filePath, uint8Array);
    return options.filePath;
  }
  if (options.uint8Array) {
    return uint8Array;
  }
  if (typeof Buffer !== "undefined") {
    return Buffer.from(uint8Array.buffer, uint8Array.byteOffset, uint8Array.byteLength);
  }
  return uint8Array;
}

// src/write/xls.js
var import_node_fs4 = __toESM(require("node:fs"), 1);
var BIFF = {
  BOF: 2057,
  EOF: 10,
  DIMENSIONS: 512,
  BLANK: 513,
  NUMBER: 515,
  BOOLERR: 517,
  STRING: 519,
  BOUNDSHEET: 133,
  SST: 252,
  LABELSST: 253,
  FONT: 49,
  XF: 224,
  STYLE: 659,
  CODEPAGE: 66,
  WINDOW1: 61,
  WINDOW2: 574,
  MERGECELLS: 229,
  COLINFO: 125
};
var PALETTE = [
  { idx: 8, r: 0, g: 0, b: 0 },
  { idx: 9, r: 255, g: 255, b: 255 },
  { idx: 10, r: 255, g: 0, b: 0 },
  { idx: 11, r: 0, g: 255, b: 0 },
  { idx: 12, r: 0, g: 0, b: 255 },
  { idx: 13, r: 255, g: 255, b: 0 },
  { idx: 14, r: 255, g: 0, b: 255 },
  { idx: 15, r: 0, g: 255, b: 255 },
  { idx: 16, r: 128, g: 0, b: 0 },
  { idx: 17, r: 0, g: 128, b: 0 },
  { idx: 18, r: 0, g: 0, b: 128 },
  { idx: 19, r: 128, g: 128, b: 0 },
  { idx: 20, r: 128, g: 0, b: 128 },
  { idx: 21, r: 0, g: 128, b: 128 },
  { idx: 22, r: 192, g: 192, b: 192 },
  { idx: 23, r: 128, g: 128, b: 128 }
];
function hexToColorIndex(hex) {
  if (!hex) return 32767;
  const clean = String(hex).replace("#", "").trim();
  if (clean.length < 6) return 32767;
  const r = parseInt(clean.slice(0, 2), 16) || 0;
  const g = parseInt(clean.slice(2, 4), 16) || 0;
  const b = parseInt(clean.slice(4, 6), 16) || 0;
  let bestDist = Infinity;
  let bestIdx = 8;
  for (const p of PALETTE) {
    const d = (p.r - r) ** 2 + (p.g - g) ** 2 + (p.b - b) ** 2;
    if (d < bestDist) {
      bestDist = d;
      bestIdx = p.idx;
    }
  }
  return bestIdx;
}
function getBiffBorderStyle(style) {
  if (!style) return 0;
  const s = String(style).toLowerCase();
  if (s.includes("double")) return 6;
  if (s.includes("thick")) return 5;
  if (s.includes("dashed")) return 3;
  if (s.includes("dotted")) return 4;
  if (s.includes("medium")) return 2;
  return 1;
}
function createRecord(type, data) {
  const payload = data instanceof Uint8Array ? data : new Uint8Array(data || 0);
  const len = payload.length;
  const rec = new Uint8Array(4 + len);
  const view = new DataView(rec.buffer, rec.byteOffset, rec.byteLength);
  view.setUint16(0, type, true);
  view.setUint16(2, len, true);
  rec.set(payload, 4);
  return rec;
}
function concatBuffers(buffers) {
  let totalLen = 0;
  for (const b of buffers) totalLen += b.length;
  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (const b of buffers) {
    result.set(b, offset);
    offset += b.length;
  }
  return result;
}
function makeFontRecord(height, isBold, colorIndex = 32767) {
  const name = "Arial";
  const data = new Uint8Array(14 + 2 + name.length);
  const view = new DataView(data.buffer);
  view.setUint16(0, height, true);
  view.setUint16(2, 0, true);
  view.setUint16(4, colorIndex, true);
  view.setUint16(6, isBold ? 700 : 400, true);
  view.setUint16(8, 0, true);
  data[10] = 0;
  data[11] = 0;
  data[12] = 0;
  data[13] = 0;
  data[14] = name.length;
  data[15] = 0;
  for (let i = 0; i < name.length; i++) {
    data[16 + i] = name.charCodeAt(i);
  }
  return createRecord(BIFF.FONT, data);
}
function createCellXFRecord(fontIndex, align, border, bgIndex) {
  const xfData = new Uint8Array(20);
  const view = new DataView(xfData.buffer);
  view.setUint16(0, fontIndex, true);
  view.setUint16(2, 0, true);
  view.setUint16(4, 1, true);
  let hAlign = 0;
  if (align === "left") hAlign = 0;
  else if (align === "center") hAlign = 1;
  else if (align === "right") hAlign = 2;
  xfData[6] = hAlign & 7;
  if (border) {
    let bVal = 0;
    bVal |= border.left & 15;
    bVal |= (border.right & 15) << 4;
    bVal |= (border.top & 15) << 8;
    bVal |= (border.bottom & 15) << 12;
    view.setUint32(10, bVal, true);
  }
  if (bgIndex !== void 0 && bgIndex !== null && bgIndex !== 32767) {
    view.setUint32(14, 1 << 26, true);
    view.setUint16(18, bgIndex & 127, true);
  }
  return createRecord(BIFF.XF, xfData);
}
function buildFontRecords() {
  return [
    makeFontRecord(200, false),
    // Font 0: Arial 10pt regular
    makeFontRecord(200, true),
    // Font 1: Arial 10pt bold
    makeFontRecord(200, false),
    // Font 2: Arial 10pt regular
    makeFontRecord(200, false)
    // Font 3: Arial 10pt regular
  ];
}
function buildXFRecords() {
  const records = [];
  for (let i = 0; i < 15; i++) {
    const xfData = new Uint8Array(20);
    const view = new DataView(xfData.buffer);
    view.setUint16(0, 0, true);
    view.setUint16(2, 0, true);
    view.setUint16(4, 65520, true);
    records.push(createRecord(BIFF.XF, xfData));
  }
  {
    const xfData = new Uint8Array(20);
    const view = new DataView(xfData.buffer);
    view.setUint16(0, 0, true);
    view.setUint16(2, 0, true);
    view.setUint16(4, 1, true);
    records.push(createRecord(BIFF.XF, xfData));
  }
  {
    const xfData = new Uint8Array(20);
    const view = new DataView(xfData.buffer);
    view.setUint16(0, 1, true);
    view.setUint16(2, 0, true);
    view.setUint16(4, 1, true);
    records.push(createRecord(BIFF.XF, xfData));
  }
  return records;
}
function packageCFB(workbookStream) {
  const SECTOR_SIZE = 512;
  const streamSize = workbookStream.length;
  const numStreamSectors = Math.max(1, Math.ceil(streamSize / SECTOR_SIZE));
  const dirSector = 0;
  const fatSector = numStreamSectors + 1;
  const totalSectors = numStreamSectors + 2;
  const totalFileSize = 512 + totalSectors * SECTOR_SIZE;
  const file = new Uint8Array(totalFileSize);
  const view = new DataView(file.buffer);
  const sig = [208, 207, 17, 224, 161, 177, 26, 225];
  for (let i = 0; i < 8; i++) file[i] = sig[i];
  view.setUint16(24, 62, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 65534, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true);
  view.setUint32(48, dirSector, true);
  view.setUint32(56, 4096, true);
  view.setUint32(60, 4294967294, true);
  view.setUint32(64, 0, true);
  view.setUint32(68, 4294967294, true);
  view.setUint32(72, 0, true);
  view.setUint32(76, fatSector, true);
  for (let i = 1; i < 109; i++) {
    view.setUint32(76 + i * 4, 4294967295, true);
  }
  const dirOffset = 512;
  const rootName = "Root Entry";
  for (let i = 0; i < rootName.length; i++) {
    view.setUint16(dirOffset + i * 2, rootName.charCodeAt(i), true);
  }
  view.setUint16(dirOffset + 64, (rootName.length + 1) * 2, true);
  file[dirOffset + 66] = 5;
  file[dirOffset + 67] = 1;
  view.setUint32(dirOffset + 68, 4294967295, true);
  view.setUint32(dirOffset + 72, 4294967295, true);
  view.setUint32(dirOffset + 76, 1, true);
  view.setUint32(dirOffset + 116, 4294967294, true);
  view.setUint32(dirOffset + 120, 0, true);
  const wbEntryOffset = dirOffset + 128;
  const wbName = "Workbook";
  for (let i = 0; i < wbName.length; i++) {
    view.setUint16(wbEntryOffset + i * 2, wbName.charCodeAt(i), true);
  }
  view.setUint16(wbEntryOffset + 64, (wbName.length + 1) * 2, true);
  file[wbEntryOffset + 66] = 2;
  file[wbEntryOffset + 67] = 1;
  view.setUint32(wbEntryOffset + 68, 4294967295, true);
  view.setUint32(wbEntryOffset + 72, 4294967295, true);
  view.setUint32(wbEntryOffset + 76, 4294967295, true);
  view.setUint32(wbEntryOffset + 116, 1, true);
  view.setUint32(wbEntryOffset + 120, streamSize, true);
  const streamStartOffset = 512 + 1 * SECTOR_SIZE;
  file.set(workbookStream, streamStartOffset);
  const fatOffset = 512 + fatSector * SECTOR_SIZE;
  view.setUint32(fatOffset + 0, 4294967294, true);
  for (let i = 1; i < numStreamSectors; i++) {
    view.setUint32(fatOffset + i * 4, i + 1, true);
  }
  view.setUint32(fatOffset + numStreamSectors * 4, 4294967294, true);
  view.setUint32(fatOffset + fatSector * 4, 4294967293, true);
  for (let i = fatSector + 1; i < 128; i++) {
    view.setUint32(fatOffset + i * 4, 4294967295, true);
  }
  return file;
}
function createXLS(data, options = {}) {
  const sheets = normalizeWorkbookSheets(data, options);
  const sstMap = /* @__PURE__ */ new Map();
  const sstList = [];
  let totalStringRefs = 0;
  const getStringIndex = (str) => {
    totalStringRefs++;
    if (sstMap.has(str)) {
      return sstMap.get(str);
    }
    const idx = sstList.length;
    sstMap.set(str, idx);
    sstList.push(str);
    return idx;
  };
  const fontRecords = buildFontRecords();
  const xfRecords = buildXFRecords();
  const fontCache = /* @__PURE__ */ new Map();
  const xfCache = /* @__PURE__ */ new Map();
  const getFontIndex = (isBold, textColor) => {
    const colorIdx = textColor ? hexToColorIndex(textColor) : 32767;
    if (!textColor || colorIdx === 32767) {
      return isBold ? 1 : 0;
    }
    const key = `${isBold}:${colorIdx}`;
    if (fontCache.has(key)) return fontCache.get(key);
    const idx = fontRecords.length;
    fontRecords.push(makeFontRecord(200, isBold, colorIdx));
    fontCache.set(key, idx);
    return idx;
  };
  const getXFIndex = (cell) => {
    if (!cell || typeof cell !== "object" || cell instanceof Date) {
      return 15;
    }
    const isBold = Boolean(cell.fontWeight === "bold" || cell.bold);
    const textColor = cell.textColor || cell.color;
    const align = cell.align;
    const backgroundColor = cell.backgroundColor;
    const bgIdx = backgroundColor ? hexToColorIndex(backgroundColor) : 32767;
    let border = null;
    if (cell.borderStyle || cell.borderColor || cell.leftBorderStyle || cell.rightBorderStyle || cell.topBorderStyle || cell.bottomBorderStyle) {
      border = {
        left: getBiffBorderStyle(cell.leftBorderStyle || cell.borderStyle),
        right: getBiffBorderStyle(cell.rightBorderStyle || cell.borderStyle),
        top: getBiffBorderStyle(cell.topBorderStyle || cell.borderStyle),
        bottom: getBiffBorderStyle(cell.bottomBorderStyle || cell.borderStyle)
      };
    }
    if (!textColor && !align && !backgroundColor && !border) {
      return isBold ? 16 : 15;
    }
    const fontIdx = getFontIndex(isBold, textColor);
    const key = `${fontIdx}:${align || ""}:${bgIdx}:${border ? JSON.stringify(border) : ""}`;
    if (xfCache.has(key)) return xfCache.get(key);
    const xfIdx = xfRecords.length;
    xfRecords.push(createCellXFRecord(fontIdx, align, border, bgIdx));
    xfCache.set(key, xfIdx);
    return xfIdx;
  };
  for (const sheet of sheets) {
    for (const row of sheet.rows) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        let val = cell;
        let type = null;
        if (cell && typeof cell === "object" && !(cell instanceof Date)) {
          val = cell.value;
          type = cell.type;
          getXFIndex(cell);
        }
        if (val !== null && val !== void 0 && val !== "") {
          if (val instanceof Date || type === Date) {
            const d = val instanceof Date ? val : new Date(val);
            const dateStr = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : String(val);
            getStringIndex(dateStr);
          } else if (typeof val !== "number" && type !== Number && typeof val !== "boolean" && type !== Boolean) {
            getStringIndex(String(val));
          }
        }
      }
    }
  }
  const globalsRecords = [];
  const bofGlobalsData = new Uint8Array([
    0,
    6,
    // BIFF8
    5,
    0,
    // Workbook Globals
    187,
    13,
    // build
    204,
    7,
    // year
    0,
    0,
    0,
    0,
    6,
    0,
    0,
    0
  ]);
  globalsRecords.push(createRecord(BIFF.BOF, bofGlobalsData));
  globalsRecords.push(createRecord(BIFF.CODEPAGE, new Uint8Array([228, 4])));
  globalsRecords.push(createRecord(BIFF.WINDOW1, new Uint8Array([
    0,
    0,
    0,
    0,
    0,
    63,
    0,
    36,
    56,
    0,
    0,
    0,
    0,
    0,
    1,
    0,
    88,
    2
  ])));
  globalsRecords.push(...fontRecords);
  globalsRecords.push(...xfRecords);
  const boundsheetRecords = [];
  for (const sheet of sheets) {
    const name = sheet.name || "Sheet1";
    const data2 = new Uint8Array(8 + name.length);
    const view = new DataView(data2.buffer);
    view.setUint32(0, 0, true);
    data2[4] = 0;
    data2[5] = 0;
    data2[6] = name.length;
    data2[7] = 0;
    for (let i = 0; i < name.length; i++) {
      data2[8 + i] = name.charCodeAt(i);
    }
    const rec = createRecord(BIFF.BOUNDSHEET, data2);
    boundsheetRecords.push(rec);
    globalsRecords.push(rec);
  }
  let sstByteLength = 8;
  for (const str of sstList) {
    sstByteLength += 3 + str.length;
  }
  const sstData = new Uint8Array(sstByteLength);
  const sstView = new DataView(sstData.buffer);
  sstView.setUint32(0, totalStringRefs, true);
  sstView.setUint32(4, sstList.length, true);
  let sstOffset = 8;
  for (const str of sstList) {
    sstView.setUint16(sstOffset, str.length, true);
    sstData[sstOffset + 2] = 0;
    for (let i = 0; i < str.length; i++) {
      sstData[sstOffset + 3 + i] = str.charCodeAt(i) & 255;
    }
    sstOffset += 3 + str.length;
  }
  globalsRecords.push(createRecord(BIFF.SST, sstData));
  globalsRecords.push(createRecord(BIFF.EOF, new Uint8Array(0)));
  let currentStreamOffset = globalsRecords.reduce((sum, r) => sum + r.length, 0);
  const worksheetStreams = [];
  const sheetBofOffsets = [];
  for (let sIdx = 0; sIdx < sheets.length; sIdx++) {
    sheetBofOffsets.push(currentStreamOffset);
    const sheet = sheets[sIdx];
    const sheetRecords = [];
    const bofSheetData = new Uint8Array([
      0,
      6,
      // BIFF8
      16,
      0,
      // Worksheet
      187,
      13,
      204,
      7,
      0,
      0,
      0,
      0,
      6,
      0,
      0,
      0
    ]);
    sheetRecords.push(createRecord(BIFF.BOF, bofSheetData));
    let maxRow = 0;
    let maxCol = 0;
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (Array.isArray(row)) {
        if (r > maxRow) maxRow = r;
        if (row.length > maxCol) maxCol = row.length;
      }
    }
    const dimData = new Uint8Array(14);
    const dimView = new DataView(dimData.buffer);
    dimView.setUint32(0, 0, true);
    dimView.setUint32(4, Math.max(1, maxRow + 1), true);
    dimView.setUint16(8, 0, true);
    dimView.setUint16(10, Math.max(1, maxCol), true);
    sheetRecords.push(createRecord(BIFF.DIMENSIONS, dimData));
    if (sheet.columns && Array.isArray(sheet.columns)) {
      sheet.columns.forEach((col, cIdx) => {
        if (col && typeof col.width === "number") {
          const colInfoData = new Uint8Array(12);
          const view = new DataView(colInfoData.buffer);
          view.setUint16(0, cIdx, true);
          view.setUint16(2, cIdx, true);
          view.setUint16(4, Math.round(col.width * 256), true);
          view.setUint16(6, 15, true);
          view.setUint16(8, 0, true);
          view.setUint16(10, 0, true);
          sheetRecords.push(createRecord(BIFF.COLINFO, colInfoData));
        }
      });
    }
    const win2Data = new Uint8Array(18);
    new DataView(win2Data.buffer).setUint16(0, 694, true);
    sheetRecords.push(createRecord(BIFF.WINDOW2, win2Data));
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (!Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        let val = cell;
        let type = null;
        let isBold = false;
        if (cell && typeof cell === "object" && !(cell instanceof Date)) {
          val = cell.value;
          type = cell.type;
          if (cell.fontWeight === "bold" || cell.bold) {
            isBold = true;
          }
        }
        if (val === null || val === void 0 || val === "") {
          continue;
        }
        const xfIndex = getXFIndex(cell);
        if (typeof val === "number" || type === Number) {
          const numData = new Uint8Array(14);
          const view = new DataView(numData.buffer);
          view.setUint16(0, r, true);
          view.setUint16(2, c, true);
          view.setUint16(4, xfIndex, true);
          view.setFloat64(6, Number(val), true);
          sheetRecords.push(createRecord(BIFF.NUMBER, numData));
        } else if (typeof val === "boolean" || type === Boolean) {
          const boolData = new Uint8Array(8);
          const view = new DataView(boolData.buffer);
          view.setUint16(0, r, true);
          view.setUint16(2, c, true);
          view.setUint16(4, xfIndex, true);
          boolData[6] = val ? 1 : 0;
          boolData[7] = 0;
          sheetRecords.push(createRecord(BIFF.BOOLERR, boolData));
        } else if (val instanceof Date || type === Date) {
          const d = val instanceof Date ? val : new Date(val);
          const dateStr = !isNaN(d.getTime()) ? d.toISOString().slice(0, 10) : String(val);
          const sstIdx = sstMap.get(dateStr) !== void 0 ? sstMap.get(dateStr) : getStringIndex(dateStr);
          const labelData = new Uint8Array(10);
          const view = new DataView(labelData.buffer);
          view.setUint16(0, r, true);
          view.setUint16(2, c, true);
          view.setUint16(4, xfIndex, true);
          view.setUint32(6, sstIdx, true);
          sheetRecords.push(createRecord(BIFF.LABELSST, labelData));
        } else {
          const strVal = String(val);
          const sstIdx = sstMap.get(strVal) !== void 0 ? sstMap.get(strVal) : getStringIndex(strVal);
          const labelData = new Uint8Array(10);
          const view = new DataView(labelData.buffer);
          view.setUint16(0, r, true);
          view.setUint16(2, c, true);
          view.setUint16(4, xfIndex, true);
          view.setUint32(6, sstIdx, true);
          sheetRecords.push(createRecord(BIFF.LABELSST, labelData));
        }
      }
    }
    const mergedCellsList = [];
    for (let r = 0; r < sheet.rows.length; r++) {
      const row = sheet.rows[r];
      if (!Array.isArray(row)) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (cell && typeof cell === "object") {
          const colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
          const rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);
          if (colSpan > 1 || rowSpan > 1) {
            mergedCellsList.push({
              rowFirst: r,
              rowLast: r + rowSpan - 1,
              colFirst: c,
              colLast: c + colSpan - 1
            });
          }
        }
      }
    }
    if (mergedCellsList.length > 0) {
      const count = mergedCellsList.length;
      const mergeData = new Uint8Array(2 + count * 8);
      const view = new DataView(mergeData.buffer);
      view.setUint16(0, count, true);
      for (let i = 0; i < count; i++) {
        const m = mergedCellsList[i];
        const offset = 2 + i * 8;
        view.setUint16(offset, m.rowFirst, true);
        view.setUint16(offset + 2, m.rowLast, true);
        view.setUint16(offset + 4, m.colFirst, true);
        view.setUint16(offset + 6, m.colLast, true);
      }
      sheetRecords.push(createRecord(BIFF.MERGECELLS, mergeData));
    }
    sheetRecords.push(createRecord(BIFF.EOF, new Uint8Array(0)));
    const sheetBytes = concatBuffers(sheetRecords);
    worksheetStreams.push(sheetBytes);
    currentStreamOffset += sheetBytes.length;
  }
  for (let sIdx = 0; sIdx < boundsheetRecords.length; sIdx++) {
    const bsRec = boundsheetRecords[sIdx];
    const offset = sheetBofOffsets[sIdx];
    new DataView(bsRec.buffer, bsRec.byteOffset + 4).setUint32(0, offset, true);
  }
  const workbookStream = concatBuffers([
    ...globalsRecords,
    ...worksheetStreams
  ]);
  return packageCFB(workbookStream);
}
async function writeXLS(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const uint8Array = createXLS(data, options);
  if (options.filePath) {
    await import_node_fs4.default.promises.writeFile(options.filePath, uint8Array);
    return options.filePath;
  }
  if (options.uint8Array) {
    return uint8Array;
  }
  if (typeof Buffer !== "undefined") {
    return Buffer.from(uint8Array.buffer, uint8Array.byteOffset, uint8Array.byteLength);
  }
  return uint8Array;
}

// src/write/formats.js
function normalizeTo2DArray(input) {
  if (!input) return [];
  if (input.worksheets && Array.isArray(input.worksheets) && input.worksheets[0]) {
    return input.worksheets[0].data || [];
  }
  if (!Array.isArray(input)) return [];
  if (input.length === 0) return [];
  if (Array.isArray(input[0])) {
    return input.map((row) => row.map((cell) => cell && typeof cell === "object" && "value" in cell ? cell.value : cell));
  }
  if (typeof input[0] === "object" && input[0] !== null) {
    const keys = Object.keys(input[0]);
    const headerRow = [...keys];
    const rows = input.map((obj) => keys.map((k) => obj[k]));
    return [headerRow, ...rows];
  }
  return [];
}
function escapeCSVCell(value, delimiter = ",") {
  if (value === null || value === void 0) {
    return "";
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  const str = String(value);
  if (str.includes(delimiter) || str.includes('"') || str.includes("\n") || str.includes("\r")) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}
async function writeCSV(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const delimiter = options.delimiter || ",";
  const grid = normalizeTo2DArray(data);
  const lines = grid.map((row) => row.map((cell) => escapeCSVCell(cell, delimiter)).join(delimiter));
  const csvContent = lines.join("\n");
  if (options.filePath) {
    await import_node_fs5.default.promises.writeFile(options.filePath, csvContent, "utf-8");
    return options.filePath;
  }
  if (options.buffer) {
    return Buffer.from(csvContent, "utf-8");
  }
  return csvContent;
}
async function writeTSV(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  return writeCSV(data, { ...options, delimiter: "	" });
}
function escapeHtml(str) {
  if (str === null || str === void 0) return "";
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
function formatCellStyles(cell) {
  if (!cell || typeof cell !== "object" || cell instanceof Date) return "";
  const styles = [];
  if (cell.fontWeight) styles.push(`font-weight:${cell.fontWeight}`);
  if (cell.fontStyle) styles.push(`font-style:${cell.fontStyle}`);
  if (cell.fontSize) {
    const sz = typeof cell.fontSize === "number" ? `${cell.fontSize}pt` : cell.fontSize;
    styles.push(`font-size:${sz}`);
  }
  if (cell.fontFamily) styles.push(`font-family:${cell.fontFamily}`);
  if (cell.backgroundColor) styles.push(`background-color:${cell.backgroundColor}`);
  if (cell.textColor || cell.color) styles.push(`color:${cell.textColor || cell.color}`);
  if (cell.align) styles.push(`text-align:${cell.align}`);
  if (cell.verticalAlign) {
    const va = cell.verticalAlign === "center" ? "middle" : cell.verticalAlign;
    styles.push(`vertical-align:${va}`);
  }
  if (cell.wrap) styles.push("white-space:pre-wrap;word-break:break-word");
  if (cell.border) styles.push(`border:${cell.border}`);
  if (cell.borderStyle || cell.borderColor) {
    styles.push(`border:1px ${cell.borderStyle || "solid"} ${cell.borderColor || "#ccc"}`);
  }
  if (cell.leftBorderStyle || cell.leftBorderColor) {
    styles.push(`border-left:1px ${cell.leftBorderStyle || "solid"} ${cell.leftBorderColor || "#ccc"}`);
  }
  if (cell.rightBorderStyle || cell.rightBorderColor) {
    styles.push(`border-right:1px ${cell.rightBorderStyle || "solid"} ${cell.rightBorderColor || "#ccc"}`);
  }
  if (cell.topBorderStyle || cell.topBorderColor) {
    styles.push(`border-top:1px ${cell.topBorderStyle || "solid"} ${cell.topBorderColor || "#ccc"}`);
  }
  if (cell.bottomBorderStyle || cell.bottomBorderColor) {
    styles.push(`border-bottom:1px ${cell.bottomBorderStyle || "solid"} ${cell.bottomBorderColor || "#ccc"}`);
  }
  return styles.length > 0 ? ` style="${styles.join(";")}"` : "";
}
function createHTML(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  let rows = [];
  let columns = options.columns;
  if (data && Array.isArray(data.worksheets) && data.worksheets[0]) {
    const ws = data.worksheets[0];
    rows = ws.data || [];
    if (ws.columns && Array.isArray(ws.columns) && rows.length > 0 && !Array.isArray(rows[0])) {
      columns = normalizeSchemaToColumns(ws.columns, ws);
      rows = getSheetData(rows, columns);
    }
  } else if (Array.isArray(data) && data.length > 0 && typeof data[0] === "object" && data[0] !== null && !Array.isArray(data[0])) {
    const schema = options.schema || options.columns;
    if (schema) {
      columns = normalizeSchemaToColumns(schema, options);
      rows = getSheetData(data, columns);
    } else {
      const keys = Object.keys(data[0]);
      const headerStyle = options.headerStyle || { fontWeight: "bold" };
      const headerRow = keys.map((k) => ({ value: k, ...headerStyle }));
      const dataRows = data.map((obj) => keys.map((k) => ({ value: obj[k] })));
      rows = [headerRow, ...dataRows];
    }
  } else if (Array.isArray(data)) {
    rows = data;
  }
  if (options.theme) {
    rows = applyTheme(rows, options.theme, options);
  }
  if (options.autoWidth) {
    columns = calculateAutoWidths(rows, options, columns);
  }
  const tableClass = options.tableClass || "spreadsheet-table";
  let html = `<table class="${tableClass}">
`;
  if (columns && Array.isArray(columns) && columns.some((c) => c && c.width)) {
    html += "  <colgroup>\n";
    columns.forEach((c) => {
      if (c && c.width) {
        html += `    <col style="width:${c.width}ch;">
`;
      } else {
        html += `    <col>
`;
      }
    });
    html += "  </colgroup>\n";
  }
  const coveredCells = /* @__PURE__ */ new Set();
  const hasHeader = options.hasHeader !== false && rows.length > 0;
  const headerRowIndex = hasHeader ? 0 : -1;
  let inThead = false;
  let inTbody = false;
  rows.forEach((row, rowIndex) => {
    if (!Array.isArray(row)) return;
    const isHeader = rowIndex === headerRowIndex;
    if (isHeader) {
      html += "  <thead>\n";
      inThead = true;
    } else if (!inTbody) {
      if (inThead) {
        html += "  </thead>\n";
        inThead = false;
      }
      html += "  <tbody>\n";
      inTbody = true;
    }
    const tag = isHeader ? "th" : "td";
    html += "    <tr>\n";
    row.forEach((cell, colIndex) => {
      const cellKey = `${rowIndex},${colIndex}`;
      if (coveredCells.has(cellKey)) {
        return;
      }
      let val = cell;
      let spanAttrs = "";
      let style = "";
      if (cell && typeof cell === "object" && !(cell instanceof Date)) {
        val = cell.value !== void 0 ? cell.value : "";
        const colSpan = Math.max(1, parseInt(cell.columnSpan || cell.colspan || cell.span) || 1);
        const rowSpan = Math.max(1, parseInt(cell.rowSpan || cell.rowspan) || 1);
        if (colSpan > 1) spanAttrs += ` colspan="${colSpan}"`;
        if (rowSpan > 1) spanAttrs += ` rowspan="${rowSpan}"`;
        if (colSpan > 1 || rowSpan > 1) {
          for (let dr = 0; dr < rowSpan; dr++) {
            for (let dc = 0; dc < colSpan; dc++) {
              if (dr > 0 || dc > 0) {
                coveredCells.add(`${rowIndex + dr},${colIndex + dc}`);
              }
            }
          }
        }
        style = formatCellStyles(cell);
      } else if (cell instanceof Date) {
        val = !isNaN(cell.getTime()) ? cell.toISOString().slice(0, 10) : "";
      }
      if (val === null || val === void 0) val = "";
      const safeVal = escapeHtml(val);
      html += `      <${tag}${spanAttrs}${style}>${safeVal}</${tag}>
`;
    });
    html += "    </tr>\n";
  });
  if (inThead) {
    html += "  </thead>\n";
  }
  if (inTbody) {
    html += "  </tbody>\n";
  }
  html += "</table>";
  if (options.fullDocument) {
    const title = escapeHtml(options.title || "Spreadsheet Export");
    html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${title}</title>
<style>
table.${tableClass} { border-collapse: collapse; width: 100%; font-family: sans-serif; }
table.${tableClass} th, table.${tableClass} td { border: 1px solid #ddd; padding: 8px; text-align: left; }
table.${tableClass} th { background-color: #f4f4f4; font-weight: bold; }
</style>
</head>
<body>
${html}
</body>
</html>`;
  }
  return html;
}
async function writeHTML(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const html = createHTML(data, options);
  if (options.filePath) {
    await import_node_fs5.default.promises.writeFile(options.filePath, html, "utf-8");
    return options.filePath;
  }
  if (options.buffer) {
    return Buffer.from(html, "utf-8");
  }
  return html;
}
async function writeJSON(data, options = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  let result = data;
  if (Array.isArray(data) && Array.isArray(data[0])) {
    if (options.asObjects && data.length > 1) {
      const keys = data[0].map((k) => String(k && typeof k === "object" ? k.value : k));
      result = data.slice(1).map((row) => {
        const obj = {};
        keys.forEach((key, i) => {
          const cell = row[i];
          obj[key] = cell && typeof cell === "object" && "value" in cell ? cell.value : cell;
        });
        return obj;
      });
    } else {
      result = data.map((r) => r.map((c) => c && typeof c === "object" && "value" in c ? c.value : c));
    }
  }
  const spaces = options.spaces !== void 0 ? options.spaces : 2;
  const jsonString = JSON.stringify(result, null, spaces);
  if (options.filePath) {
    await import_node_fs5.default.promises.writeFile(options.filePath, jsonString, "utf-8");
    return options.filePath;
  }
  if (options.buffer) {
    return Buffer.from(jsonString, "utf-8");
  }
  return jsonString;
}

// src/write/streaming.js
var import_node_fs6 = __toESM(require("node:fs"), 1);
var import_node_stream3 = require("node:stream");
function formatCSVLine(cells, delimiter = ",") {
  if (!Array.isArray(cells)) return "";
  return cells.map((cell) => {
    const val = cell && typeof cell === "object" && !(cell instanceof Date) && "value" in cell ? cell.value : cell;
    if (val === null || val === void 0) return "";
    if (val instanceof Date) return val.toISOString();
    const str = String(val);
    if (str.includes(delimiter) || str.includes('"') || str.includes("\n") || str.includes("\r")) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }).join(delimiter);
}
var CSVTransformStream = class extends import_node_stream3.Transform {
  constructor(options = {}) {
    super({ objectMode: true });
    this.delimiter = options.delimiter || ",";
    this.hasHeader = options.hasHeader ?? true;
    this.headers = options.headers || null;
    this.headerEmitted = false;
  }
  _transform(chunk, encoding, callback) {
    try {
      if (Array.isArray(chunk)) {
        this.push(formatCSVLine(chunk, this.delimiter) + "\n");
      } else if (chunk && typeof chunk === "object") {
        if (!this.headerEmitted && this.hasHeader) {
          const keys2 = this.headers || Object.keys(chunk);
          this.push(formatCSVLine(keys2, this.delimiter) + "\n");
          this.headerEmitted = true;
        }
        const keys = this.headers || Object.keys(chunk);
        const values = keys.map((k) => chunk[k]);
        this.push(formatCSVLine(values, this.delimiter) + "\n");
      }
      callback();
    } catch (err) {
      callback(err);
    }
  }
};
var JSONLTransformStream = class extends import_node_stream3.Transform {
  constructor(options = {}) {
    super({ objectMode: true });
  }
  _transform(chunk, encoding, callback) {
    try {
      this.push(JSON.stringify(chunk) + "\n");
      callback();
    } catch (err) {
      callback(err);
    }
  }
};
function runPipeline(readable, transform4, writable) {
  return new Promise((resolve, reject) => {
    (0, import_node_stream3.pipeline)(readable, transform4, writable, (err) => {
      if (err) reject(err);
      else resolve();
    });
  });
}
async function streamCSV(source, destination, options = {}) {
  const readable = source instanceof import_node_stream3.Readable ? source : import_node_stream3.Readable.from(source);
  const transform4 = new CSVTransformStream(options);
  const isPath = typeof destination === "string";
  const writable = isPath ? import_node_fs6.default.createWriteStream(destination, { encoding: "utf-8" }) : destination;
  await runPipeline(readable, transform4, writable);
  return isPath ? destination : writable;
}
async function streamTSV(source, destination, options = {}) {
  return streamCSV(source, destination, { ...options, delimiter: "	" });
}
async function streamJSONL(source, destination, options = {}) {
  const readable = source instanceof import_node_stream3.Readable ? source : import_node_stream3.Readable.from(source);
  const transform4 = new JSONLTransformStream(options);
  const isPath = typeof destination === "string";
  const writable = isPath ? import_node_fs6.default.createWriteStream(destination, { encoding: "utf-8" }) : destination;
  await runPipeline(readable, transform4, writable);
  return isPath ? destination : writable;
}

// src/core/registry.js
var registeredWriters = /* @__PURE__ */ new Map();
var KNOWN_PLUGIN_MAPPINGS = {
  csv: { plugin: "textPlugin", importPath: "spreadsheet-utils/plugins" },
  tsv: { plugin: "textPlugin", importPath: "spreadsheet-utils/plugins" },
  tab: { plugin: "textPlugin", importPath: "spreadsheet-utils/plugins" },
  txt: { plugin: "textPlugin", importPath: "spreadsheet-utils/plugins" },
  xls: { plugin: "xlsPlugin", importPath: "spreadsheet-utils/plugins" },
  ods: { plugin: "markupPlugin", importPath: "spreadsheet-utils/plugins" },
  html: { plugin: "markupPlugin", importPath: "spreadsheet-utils/plugins" },
  htm: { plugin: "markupPlugin", importPath: "spreadsheet-utils/plugins" },
  xml: { plugin: "markupPlugin", importPath: "spreadsheet-utils/plugins" },
  dbf: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  dif: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  slk: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  sylk: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  wks: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  wk1: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  wk3: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  wk4: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  "123": { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" },
  numbers: { plugin: "legacyPlugin", importPath: "spreadsheet-utils/plugins" }
};
function getWriterForExtension(ext, localRegistry) {
  const key = (ext || "").toLowerCase().replace(/^\./, "");
  if (localRegistry && localRegistry.has(key)) {
    return localRegistry.get(key);
  }
  return registeredWriters.get(key) || null;
}
function createMissingPluginError(ext, fileName) {
  const key = (ext || "").toLowerCase().replace(/^\./, "");
  const mapping = KNOWN_PLUGIN_MAPPINGS[key];
  if (mapping) {
    const err2 = new Error(
      `File format "${key}" requires the "${mapping.plugin}". Register it before reading or writing:

  import { ${mapping.plugin} } from '${mapping.importPath}';
  registerPlugin(${mapping.plugin});
`
    );
    err2.code = "ERR_PLUGIN_REQUIRED";
    err2.plugin = mapping.plugin;
    err2.format = key;
    return err2;
  }
  const err = new Error(`Unsupported file type: ${ext || fileName || "(unknown)"}`);
  err.code = "ERR_UNSUPPORTED_FORMAT";
  return err;
}

// src/write/index.js
async function writeSpreadsheet(data, options = {}, extraOptions = {}) {
  if (typeof options === "string") {
    options = { filePath: options };
  }
  const format = options.format ? String(options.format).toLowerCase().replace(/^\./, "") : options.filePath ? options.filePath.split(".").pop().toLowerCase() : "xlsx";
  if (format === "csv") {
    return writeCSV(data, options);
  }
  if (format === "tsv") {
    return writeTSV(data, options);
  }
  if (format === "html" || format === "htm") {
    return writeHTML(data, options);
  }
  if (format === "json") {
    return writeJSON(data, options);
  }
  if (format === "dbf") {
    const writer2 = getWriterForExtension("dbf");
    if (writer2) {
      return writer2("dbf", data, options, extraOptions);
    }
    if (KNOWN_PLUGIN_MAPPINGS["dbf"]) {
      throw createMissingPluginError("dbf", options.filePath);
    }
    return writeDBF(data, options);
  }
  if (format === "xls") {
    const writer2 = getWriterForExtension("xls");
    if (writer2) {
      return writer2("xls", data, options, extraOptions);
    }
    if (KNOWN_PLUGIN_MAPPINGS["xls"]) {
      throw createMissingPluginError("xls", options.filePath);
    }
    return writeXLS(data, options);
  }
  if (format === "ods") {
    const writer2 = getWriterForExtension("ods");
    if (writer2) {
      return writer2("ods", data, options, extraOptions);
    }
    if (KNOWN_PLUGIN_MAPPINGS["ods"]) {
      throw createMissingPluginError("ods", options.filePath);
    }
    return writeODS(data, options);
  }
  const customWriter = getWriterForExtension(format);
  if (customWriter) {
    return customWriter(format, data, options, extraOptions);
  }
  let inputData = data;
  let sheetOptions = { ...options };
  if (inputData && Array.isArray(inputData.worksheets)) {
    inputData = inputData.worksheets.map((ws, i) => ({
      data: ws.data || [],
      sheet: ws.worksheetName || ws.sheet || `Sheet${i + 1}`,
      columns: ws.columns
    }));
  }
  if (Array.isArray(inputData) && inputData.length > 0 && !Array.isArray(inputData[0])) {
    if (sheetOptions.schema) {
      const columns = normalizeSchemaToColumns(sheetOptions.schema, sheetOptions);
      inputData = getSheetData(inputData, columns);
      sheetOptions.columns = columns;
      delete sheetOptions.schema;
    } else if (sheetOptions.columns && typeof sheetOptions.columns[0]?.cell !== "function") {
      const columns = normalizeSchemaToColumns(sheetOptions.columns, sheetOptions);
      inputData = getSheetData(inputData, columns);
      sheetOptions.columns = columns;
    }
  }
  if (Array.isArray(inputData) && inputData.length > 0 && typeof inputData[0] === "object" && inputData[0] !== null && !Array.isArray(inputData[0]) && inputData[0].data) {
    inputData = inputData.map((s) => {
      let sData = s.data || [];
      if (sheetOptions.theme) {
        sData = applyTheme(sData, sheetOptions.theme, sheetOptions);
      }
      let sCols = s.columns;
      if (sheetOptions.autoWidth) {
        sCols = calculateAutoWidths(sData, sheetOptions, sCols);
      }
      return { ...s, data: sData, columns: sCols };
    });
  } else if (Array.isArray(inputData)) {
    if (sheetOptions.theme) {
      inputData = applyTheme(inputData, sheetOptions.theme, sheetOptions);
    }
    if (sheetOptions.autoWidth) {
      sheetOptions.columns = calculateAutoWidths(inputData, sheetOptions, sheetOptions.columns);
    }
  }
  const writer = writeXlsxFile(inputData, sheetOptions, extraOptions);
  if (sheetOptions.filePath) {
    if (typeof writer.toFile === "function") {
      await writer.toFile(sheetOptions.filePath);
      return sheetOptions.filePath;
    }
    throw new Error("filePath option is only supported in Node.js environments");
  }
  if (sheetOptions.buffer) {
    return await writer.toBuffer();
  }
  if (sheetOptions.blob) {
    return await writer.toBlob();
  }
  return await writer.toBuffer();
}
var writeXlsxFile2 = writeXlsxFile;
var write_default = writeSpreadsheet;
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  CSVTransformStream,
  JSONLTransformStream,
  THEMES,
  applyTheme,
  calculateAutoWidths,
  createDBF,
  createHTML,
  createODS,
  createXLS,
  getSheetData,
  getTheme,
  normalizeSchemaToColumns,
  normalizeSchemaToDbfFields,
  streamCSV,
  streamJSONL,
  streamTSV,
  writeCSV,
  writeDBF,
  writeHTML,
  writeJSON,
  writeODS,
  writeSpreadsheet,
  writeTSV,
  writeXLS,
  writeXlsxFile
});
//# sourceMappingURL=write.cjs.map
