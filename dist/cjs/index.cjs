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

// src/index.js
var src_exports = {};
__export(src_exports, {
  Email: () => Email,
  Integer: () => Integer,
  InvalidInputError: () => InvalidInputError,
  InvalidSpreadsheetError: () => InvalidSpreadsheetError,
  SheetNotFoundError: () => SheetNotFoundError,
  URL: () => URL,
  convertRowsToObjects: () => convertRowsToObjects,
  createReader: () => createReader,
  default: () => readSpreadsheet,
  detectFormat: () => detectFormat,
  getWriterForExtension: () => getWriterForExtension,
  isPluginRegistered: () => isPluginRegistered,
  iterateSpreadsheet: () => iterateSpreadsheet,
  parseSheetData: () => parseSheetData,
  parseTabular: () => parseTabular,
  parseXLSX: () => parseXLSX,
  parser: () => parser,
  readSheet: () => readSheet,
  readSpreadsheet: () => readSpreadsheet,
  readXlsxFile: () => readXlsxFile,
  registerPlugin: () => registerPlugin,
  tabularjs: () => tabularjs,
  unregisterPlugin: () => unregisterPlugin
});
module.exports = __toCommonJS(src_exports);

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
  const daysBeforeUnixEpoch = JANUARY_0TH_1900_DAY + ERRONEOUS_FEBRUARY_29_1990_DAY + (1970 - 1900) * DAYS_IN_YEAR + NUMBER_OF_LEAP_YEARS_BETWEEN_1900_AND_1970;
  return Math.floor((excelSerialDate - daysBeforeUnixEpoch) * DAY);
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
var EMPTY_CELL_VALUE = null;
var EMPTY_CELL = [null, EMPTY_CELL_VALUE];
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
var EMPTY_CELL_VALUE2 = null;
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
        row.push(EMPTY_CELL_VALUE2);
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
      } else if (cell.value !== EMPTY_CELL_VALUE2) {
        while (cell.column > state2.row.length + 1) {
          state2.row.push(EMPTY_CELL_VALUE2);
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
      state2.vText ||= "";
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
var import_fflate = require("fflate");

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
function resetCheckpoint() {
  latestCheckpointTimestamp = void 0;
}

// src/utils/convertValuesFromUint8ArraysToStrings.js
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
    return (0, import_fflate.strFromU8)(data);
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
      parse(xml);
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
        leftoverXml = parse(leftoverXml + xml, true) || "";
      }
      return this;
    };
    this["end"] = function() {
      if (!streaming) {
        resetState();
      }
      streaming = false;
      if (!returnError) {
        parse(leftoverXml);
      }
      leftoverXml = "";
      getContext = noopGetContext;
      parseStop = false;
      return returnError;
    };
    this["stop"] = function() {
      parseStop = true;
    };
    function parse(xml, streaming2 = false) {
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

// src/zip/unzipFromArrayBuffer.js
var import_fflate2 = require("fflate");

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

// src/zip/unzipFromArrayBuffer.js
function unzipFromArrayBuffer(input, options) {
  return unzipFromArrayBufferUsingFunction(input, options, unzipAsync, true);
}
function unzipFromArrayBufferUsingFunction(input, options = {}, unzip3, isAsync) {
  const filter = options.filter;
  const maxBytes = options.maxUncompressedSize || options.maxBytes || 256 * 1024 * 1024;
  let totalUncompressed = 0;
  return unzip3(new Uint8Array(input), {
    // Ignore certain types of files and protect against zip bombs
    filter: (file) => {
      if (filter) {
        if (!filter({ path: file.name })) {
          return false;
        }
      }
      if (file.originalSize) {
        totalUncompressed += file.originalSize;
        if (totalUncompressed > maxBytes) {
          throw new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`);
        }
      }
      return true;
    }
  }).then(
    (result) => result,
    (error) => {
      if (isFlateError(error)) {
        throw createUnzipError(error);
      } else {
        throw error;
      }
    }
  );
}
function unzipAsync(archive) {
  return new Promise((resolve, reject) => {
    (0, import_fflate2.unzip)(archive, (error, files) => {
      if (error) {
        reject(error);
      } else {
        resolve(files);
      }
    });
  });
}
function isFlateError(error) {
  return typeof error.code === "number";
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

// src/export/filterZipArchiveEntry.js
function filterZipArchiveEntry({ path }) {
  return path.endsWith(".xml") || path.endsWith(".xml.rels");
}

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

// src/export/unpackXlsxFileUniversal.js
function unpackXlsxFile(input, options) {
  resetCheckpoint();
  checkpoint("unpack files");
  return getArrayBuffer(input).then((arrayBuffer) => {
    validateLeadingBytes(new Uint8Array(arrayBuffer));
    return unzipFromArrayBuffer(arrayBuffer, { filter: filterZipArchiveEntry, ...options }).then(
      (result) => result,
      (error) => {
        if (error instanceof UnzipError) {
          throw new InvalidInputError("INVALID_ZIP", error.cause);
        } else {
          throw error;
        }
      }
    );
  });
}
function getArrayBuffer(input) {
  if (input instanceof Blob) {
    return input.arrayBuffer();
  }
  if (input instanceof ArrayBuffer) {
    return Promise.resolve(input);
  }
  throw new InvalidInputError("INPUT_TYPE_NOT_SUPPORTED");
}

// src/zip/unzipFromStream.unzipper.js
var import_unzipper_esm = require("unzipper-esm");
var PROMISE_RESOLVE_VALUE = void 0;
function unzipFromStream_(stream, onFile, onFileData, onFileDataEnd) {
  return new Promise((resolve, reject) => {
    const promises = [];
    let errored = false;
    const unzipperStream = (0, import_unzipper_esm.Parse)();
    const onError = (error) => {
      if (error instanceof import_unzipper_esm.ZipFileError) {
        error = createUnzipError(error);
      }
      if (!errored) {
        errored = true;
        try {
          if (typeof unzipperStream.destroy === "function") {
            unzipperStream.destroy(error);
          }
        } catch {
        }
        try {
          if (typeof stream.destroy === "function") {
            stream.destroy(error);
          }
        } catch {
        }
        reject(error);
      }
    };
    stream.on("error", onError).pipe(unzipperStream).on("error", (err) => {
      if (!errored) onError(err);
    }).on("finish", () => {
      if (!errored) {
        Promise.all(promises).then(() => {
          resolve(PROMISE_RESOLVE_VALUE);
        }, onError);
      }
    }).on("entry", (entry) => {
      if (errored) {
        if (typeof entry.autodrain === "function") {
          entry.autodrain().on("error", () => {
          });
        }
        return;
      }
      let ignore = false;
      if (entry.type === "Directory") {
        ignore = true;
      }
      const uncompressedFileSizeIsKnown = !(entry.vars.flags & 8) || entry.vars.compressedSize > 0;
      const uncompressedFileSize = uncompressedFileSizeIsKnown ? entry.vars.uncompressedSize : void 0;
      try {
        if (onFile(entry.path, uncompressedFileSize) === false) {
          ignore = true;
        }
      } catch (err) {
        onError(err);
        return;
      }
      if (ignore) {
        entry.autodrain().on("error", (err) => {
          if (!errored) onError(err);
        });
        return;
      }
      promises.push(new Promise((resolve2) => {
        entry.on("data", (data) => {
          if (errored) return;
          try {
            onFileData(entry.path, data);
          } catch (err) {
            try {
              if (typeof entry.destroy === "function") {
                entry.destroy(err);
              } else if (typeof entry.autodrain === "function") {
                entry.autodrain().on("error", () => {
                });
              }
            } catch {
            }
            onError(err);
            resolve2();
          }
        }).on("error", (error) => {
          if (!errored) onError(error);
          resolve2();
        }).on("finish", () => {
          if (!errored) {
            try {
              onFileDataEnd(entry.path);
            } catch (err) {
              onError(err);
            }
          }
          resolve2();
        });
      }));
    });
  });
}

// src/zip/unzipFromStream.js
function unzipFromStream(stream, options = {}) {
  const filter = options.filter;
  const maxBytes = options.maxUncompressedSize || options.maxBytes || 256 * 1024 * 1024;
  let totalUncompressed = 0;
  const files = {};
  const filesData = {};
  const onFile = (filePath, fileSize) => {
    if (filter && !filter({ path: filePath })) {
      return false;
    }
    filesData[filePath] = {
      size: fileSize,
      chunks: [],
      buffer: fileSize === void 0 ? void 0 : Buffer.alloc(fileSize),
      position: 0
    };
  };
  const onFileData = (filePath, chunk) => {
    totalUncompressed += chunk.length;
    if (totalUncompressed > maxBytes) {
      throw new Error(`Zip decompression limit of ${maxBytes} bytes exceeded (possible zip bomb)`);
    }
    if (filesData[filePath].size === void 0) {
      filesData[filePath].chunks.push(chunk);
    } else {
      filesData[filePath].position += chunk.copy(filesData[filePath].buffer, filesData[filePath].position);
    }
  };
  const onFileDataEnd = (filePath) => {
    if (filesData[filePath].size === void 0) {
      files[filePath] = concatChunksToNonPooledBuffer(filesData[filePath].chunks);
    } else {
      files[filePath] = filesData[filePath].buffer;
    }
    delete filesData[filePath];
  };
  return unzipFromStream_(stream, onFile, onFileData, onFileDataEnd).then(() => files);
}
function concatChunksToNonPooledBuffer(chunks) {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const buffer = Buffer.allocUnsafeSlow(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    chunk.copy(buffer, offset);
    offset += chunk.length;
  }
  return buffer;
}

// src/export/convertInputToNodeStream.js
var import_node_fs = __toESM(require("node:fs"), 1);
var import_node_buffer = require("node:buffer");
var import_node_stream = __toESM(require("node:stream"), 1);
function convertInputToNodeStream(input) {
  return input instanceof import_node_stream.default ? input : input instanceof Buffer ? createReadableStreamFromBuffer(input) : input instanceof import_node_buffer.Blob ? createReadableStreamFromBlob(input) : import_node_fs.default.createReadStream(input);
}
function createReadableStreamFromBuffer(buffer) {
  return import_node_stream.Readable.from(buffer);
}
function createReadableStreamFromBlob(blob) {
  return import_node_stream.Readable.fromWeb(blob.stream());
}

// src/xlsx/file/InputValidationStream.js
var import_node_stream2 = require("node:stream");
var InputValidationStream = class extends import_node_stream2.Transform {
  constructor() {
    super();
    this.d = createFileTypeDetector();
    this.i = 0;
    this.chunks = [];
  }
  _transform(chunk, encoding, callback) {
    if (this.d) {
      this.chunks.push(chunk);
      for (const byte of chunk) {
        try {
          if (validateByte(byte, this.d)) {
            this.d = void 0;
            chunk = Buffer.concat(this.chunks);
            this.chunks = void 0;
            break;
          }
        } catch (error) {
          return callback(error);
        }
        this.i++;
      }
      if (this.d) {
        return callback();
      }
    }
    this.push(chunk);
    callback();
  }
  // `_flush()` is called when there will be no more input, before the "end" event is emitted.
  _flush(callback) {
    if (this.d) {
      try {
        noFileTypeCouldBeDetermined(this.i);
      } catch (error) {
        return callback(error);
      }
    }
    callback();
  }
};

// src/export/unpackXlsxFileNode.js
function unpackXlsxFile2(input, options) {
  resetCheckpoint();
  checkpoint("unpack files");
  const stream = convertInputToNodeStream(input);
  return new Promise((resolve, reject) => {
    const streamWithInputValidation = stream.on("error", reject).pipe(new InputValidationStream());
    unzipFromStream(streamWithInputValidation, { filter: filterZipArchiveEntry, ...options }).then(
      resolve,
      (error) => {
        if (error instanceof UnzipError) {
          reject(new InvalidInputError("INVALID_ZIP", error.cause));
        } else {
          reject(error);
        }
      }
    );
  });
}

// src/parser/xlsx.js
async function parseXLSX(input, options = {}) {
  const isNode = typeof process !== "undefined" && process.versions && process.versions.node;
  const isNodeSpecific = typeof input === "string" || typeof Buffer !== "undefined" && Buffer.isBuffer(input);
  let unpacker = unpackXlsxFile;
  if (isNode && isNodeSpecific) {
    unpacker = unpackXlsxFile2;
  }
  const contents = await unpacker(input, options);
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
      columns.push({
        title: getColumnName(c),
        width: 100
      });
    }
    return {
      worksheetName: s.sheet || s.name || "Sheet1",
      data,
      columns
    };
  });
  const result = { worksheets };
  if (options.onload) {
    options.onload(result);
  }
  return result;
}

// src/utils/loader.js
var fs2 = null;
async function loadAsBuffer(input) {
  if (input instanceof Uint8Array) return input;
  if (typeof ArrayBuffer !== "undefined" && input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  if (typeof input === "string") {
    if (!fs2) fs2 = await import("fs");
    if (input.length < 4096 && !input.includes("\n") && !input.includes("\r")) {
      try {
        if (fs2.existsSync(input)) {
          return new Uint8Array(fs2.readFileSync(input));
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

// src/zip/zip-utils.js
var import_fflate3 = require("fflate");
var DEFAULT_MAX_ENTRY_BYTES = 256 * 1024 * 1024;
function toUint8Array(input) {
  if (input instanceof Uint8Array) {
    return input;
  }
  if (typeof ArrayBuffer !== "undefined" && input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  throw new TypeError("Expected Uint8Array, ArrayBuffer, or Buffer");
}
function getZipEntryNames(buffer) {
  const bytes = toUint8Array(buffer);
  const names = [];
  try {
    (0, import_fflate3.unzipSync)(bytes, {
      filter: (file) => {
        names.push(file.name);
        return false;
      }
    });
  } catch {
  }
  return names;
}

// src/core/registry.js
var registeredParsers = /* @__PURE__ */ new Map();
var registeredWriters = /* @__PURE__ */ new Map();
var registeredSniffers = /* @__PURE__ */ new Set();
var pluginNames = /* @__PURE__ */ new Set();
var pluginsMap = /* @__PURE__ */ new Map();
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
function registerPlugin(...plugins) {
  for (const plugin of plugins) {
    if (!plugin || typeof plugin !== "object") continue;
    const name = plugin.name || `plugin_${Math.random().toString(36).slice(2)}`;
    pluginNames.add(name);
    pluginsMap.set(name, plugin);
    if (Array.isArray(plugin.extensions) && typeof plugin.parse === "function") {
      for (const ext of plugin.extensions) {
        registeredParsers.set(ext.toLowerCase(), plugin.parse);
      }
    }
    if (typeof plugin.detect === "function") {
      registeredSniffers.add(plugin.detect);
    }
    if (typeof plugin.write === "function") {
      const writeExts = Array.isArray(plugin.writeExtensions) ? plugin.writeExtensions : Array.isArray(plugin.extensions) ? plugin.extensions : [];
      for (const ext of writeExts) {
        registeredWriters.set(ext.toLowerCase(), plugin.write);
      }
    }
    if (plugin.writers && typeof plugin.writers === "object") {
      for (const [ext, writerFn] of Object.entries(plugin.writers)) {
        if (typeof writerFn === "function") {
          registeredWriters.set(ext.toLowerCase(), writerFn);
        }
      }
    }
  }
}
function unregisterPlugin(target) {
  if (!target) {
    registeredParsers.clear();
    registeredWriters.clear();
    registeredSniffers.clear();
    pluginNames.clear();
    pluginsMap.clear();
    return;
  }
  const name = typeof target === "string" ? target : target.name;
  const plugin = typeof target === "string" ? pluginsMap.get(target) : target;
  if (name) {
    pluginNames.delete(name);
    pluginsMap.delete(name);
  }
  if (plugin && plugin.extensions) {
    for (const ext of plugin.extensions) {
      registeredParsers.delete(ext.toLowerCase());
      registeredWriters.delete(ext.toLowerCase());
    }
  }
  if (plugin && plugin.writeExtensions) {
    for (const ext of plugin.writeExtensions) {
      registeredWriters.delete(ext.toLowerCase());
    }
  }
  if (plugin && plugin.writers) {
    for (const ext of Object.keys(plugin.writers)) {
      registeredWriters.delete(ext.toLowerCase());
    }
  }
  if (plugin && plugin.detect) {
    registeredSniffers.delete(plugin.detect);
  }
}
function getParserForExtension(ext, localRegistry) {
  const key = (ext || "").toLowerCase().replace(/^\./, "");
  if (localRegistry && localRegistry.has(key)) {
    return localRegistry.get(key);
  }
  return registeredParsers.get(key) || null;
}
function getWriterForExtension(ext, localRegistry) {
  const key = (ext || "").toLowerCase().replace(/^\./, "");
  if (localRegistry && localRegistry.has(key)) {
    return localRegistry.get(key);
  }
  return registeredWriters.get(key) || null;
}
async function runSniffers(buffer, localSniffers) {
  const sniffers = localSniffers || registeredSniffers;
  for (const sniffer of sniffers) {
    try {
      const detected = await sniffer(buffer);
      if (detected) return detected;
    } catch {
    }
  }
  return null;
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
function isPluginRegistered(name) {
  return pluginNames.has(name);
}

// src/core/detector.js
async function detectFormat(buffer, localSniffers) {
  if (!buffer || buffer.length < 4) {
    return null;
  }
  if (buffer[0] === 80 && buffer[1] === 75 && (buffer[2] === 3 || buffer[2] === 5 || buffer[2] === 7)) {
    try {
      const names = getZipEntryNames(buffer);
      if (names.some((n) => n === "[Content_Types].xml" || n.startsWith("xl/"))) {
        return "xlsx";
      }
      if (names.some((n) => n === "Index.zip" || n.startsWith("Index/") || n.endsWith(".iwa"))) {
        return "numbers";
      }
      if (names.includes("mimetype") || names.includes("content.xml")) {
        return "ods";
      }
    } catch {
    }
    return "xlsx";
  }
  const pluginDetected = await runSniffers(buffer, localSniffers);
  if (pluginDetected) {
    return pluginDetected;
  }
  if (buffer[0] === 208 && buffer[1] === 207 && buffer[2] === 17 && buffer[3] === 224) {
    return "xls";
  }
  if (buffer[0] === 0 && buffer[1] === 0 && (buffer[2] === 2 || buffer[2] === 26) && buffer[3] === 0) {
    return "wk1";
  }
  if ([2, 3, 48, 49, 131, 139, 245].includes(buffer[0]) && buffer.length > 32) {
    const month = buffer[2];
    const day = buffer[3];
    const headerLength = buffer[8] | buffer[9] << 8;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && headerLength > 32) {
      return "dbf";
    }
  }
  const sample = new TextDecoder("utf-8", { fatal: false }).decode(buffer.slice(0, 4096));
  if (sample.includes("\0")) {
    return null;
  }
  const text = sample.replace(/^\uFEFF/, "").trimStart();
  const lower = text.toLowerCase();
  if (text.startsWith("ID;")) {
    return "slk";
  }
  if (text.startsWith("TABLE") && text.includes("VECTORS")) {
    return "dif";
  }
  if (lower.startsWith("<?xml")) {
    if (lower.includes("urn:schemas-microsoft-com:office:spreadsheet") || lower.includes("<workbook")) {
      return "xml";
    }
    if (lower.includes("<html") || lower.includes("<table")) {
      return "html";
    }
    return "xml";
  }
  if (lower.startsWith("<!doctype html") || lower.includes("<html") || lower.includes("<table")) {
    return "html";
  }
  if (text.length) {
    const firstLine = text.split(/\r?\n/, 1)[0];
    if (firstLine.includes("	") && !firstLine.includes(",")) {
      return "tsv";
    }
    return "csv";
  }
  return null;
}

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

// src/schema/parseSheetData.js
var EMPTY_CELL_VALUE3 = null;
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
  if (cellValue === EMPTY_CELL_VALUE3) {
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
  if (value === EMPTY_CELL_VALUE3) {
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
  if (value === EMPTY_CELL_VALUE3) {
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
      return { value: EMPTY_CELL_VALUE3 };
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

// src/core/dispatcher.js
function dispatch(ext, file, options = {}, context = {}) {
  const key = (ext || "").toLowerCase().replace(/^\./, "");
  if (key === "xlsx") {
    return parseXLSX(file, options);
  }
  const parser2 = getParserForExtension(key, context.localParsers);
  if (parser2) {
    return parser2(key, file, options);
  }
  if (KNOWN_PLUGIN_MAPPINGS[key]) {
    throw createMissingPluginError(key);
  }
  return null;
}
async function parseTabular(file, options = {}, context = {}) {
  if (!file) {
    throw new Error("Invalid file");
  }
  if (options.maxFileSize) {
    if (typeof file === "object" && typeof file.size === "number" && file.size > options.maxFileSize) {
      throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
    }
    if (file instanceof Uint8Array && file.length > options.maxFileSize) {
      throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
    }
  }
  if (typeof ArrayBuffer !== "undefined" && file instanceof ArrayBuffer) {
    file = new Uint8Array(file);
  }
  const isBinaryInput = file instanceof Uint8Array || typeof Buffer !== "undefined" && Buffer.isBuffer(file) || typeof Blob !== "undefined" && file instanceof Blob;
  let fileName = "";
  if (typeof file === "string") {
    fileName = file;
  } else if (file.name && typeof file.name === "string") {
    fileName = file.name;
  } else if (!isBinaryInput && !(file && typeof file.pipe === "function")) {
    throw new Error("Invalid file: must be a file path, File, Blob, Uint8Array, ArrayBuffer, or Buffer");
  }
  let ext = null;
  if (options.extension) {
    ext = String(options.extension).toLowerCase().replace(/^\./, "");
  } else if (options.format && options.format !== "tabular" && options.format !== "rows") {
    ext = String(options.format).toLowerCase().replace(/^\./, "");
  } else {
    const dot = fileName.lastIndexOf(".");
    if (dot > -1) {
      ext = fileName.toLowerCase().slice(dot + 1);
    }
  }
  if (ext) {
    const result = dispatch(ext, file, options, context);
    if (result) {
      return await result;
    }
    if (!KNOWN_PLUGIN_MAPPINGS[ext]) {
      throw new Error(`Unsupported file type: ${ext || fileName}`);
    }
  }
  if (!ext && typeof file === "string" && !file.includes("\n") && !file.includes("<")) {
    throw new Error(`Unsupported file type: ${fileName || "(none)"}`);
  }
  let buffer = null;
  let detected = null;
  try {
    buffer = await loadAsBuffer(file);
    if (options.maxFileSize && buffer.length > options.maxFileSize) {
      throw new Error(`File size limit of ${options.maxFileSize} bytes exceeded`);
    }
    detected = await detectFormat(buffer, context.localSniffers);
  } catch (err) {
    if (err && err.message && err.message.includes("File size limit")) {
      throw err;
    }
  }
  if (detected) {
    const result = dispatch(detected, buffer || file, options, context);
    if (result) {
      return await result;
    }
  }
  if (ext && KNOWN_PLUGIN_MAPPINGS[ext]) {
    throw createMissingPluginError(ext, fileName);
  }
  if (detected && KNOWN_PLUGIN_MAPPINGS[detected]) {
    throw createMissingPluginError(detected, fileName);
  }
  throw new Error(`Unsupported file type: ${ext || detected || fileName || "(buffer)"}`);
}
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
async function readSpreadsheet(file, options = {}, context = {}) {
  const tabularResult = await parseTabular(file, options, context);
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
async function* iterateSpreadsheet(file, options = {}, context = {}) {
  const rowQueue = [];
  let resolveNext = null;
  let finished = false;
  let parseError = null;
  const asObjects = options.asObjects === true || options.header === true;
  let headers = null;
  const streamingOptions = {
    ...options,
    // Override asObjects in the inner read so onRow receives raw rows
    asObjects: false,
    header: false,
    onRow: (row, rowIndex) => {
      if (asObjects) {
        if (headers === null) {
          const seen = /* @__PURE__ */ new Map();
          headers = (row || []).map((h, i) => {
            let base = h != null ? String(h).trim() : "";
            if (!base) base = `column${i + 1}`;
            const count = seen.get(base) || 0;
            seen.set(base, count + 1);
            return count > 0 ? `${base}_${count + 1}` : base;
          });
          return;
        }
        const obj = {};
        for (let c = 0; c < headers.length; c++) {
          obj[headers[c]] = row[c] !== void 0 ? row[c] : null;
        }
        rowQueue.push(obj);
      } else {
        rowQueue.push(row);
      }
      if (resolveNext) {
        const cb = resolveNext;
        resolveNext = null;
        cb();
      }
    }
  };
  const parsePromise = readSpreadsheet(file, streamingOptions, context).then(() => {
    finished = true;
    if (resolveNext) {
      const cb = resolveNext;
      resolveNext = null;
      cb();
    }
  }).catch((err) => {
    parseError = err;
    finished = true;
    if (resolveNext) {
      const cb = resolveNext;
      resolveNext = null;
      cb();
    }
  });
  while (true) {
    if (rowQueue.length > 0) {
      yield rowQueue.shift();
    } else if (finished) {
      if (parseError) {
        throw parseError;
      }
      break;
    } else {
      await new Promise((resolve) => {
        resolveNext = resolve;
      });
    }
  }
  await parsePromise;
}
function createReader(plugins = []) {
  const localParsers = /* @__PURE__ */ new Map();
  const localSniffers = /* @__PURE__ */ new Set();
  const pluginList = Array.isArray(plugins) ? plugins : [plugins];
  function addPlugins(list) {
    for (const plugin of list) {
      if (!plugin || typeof plugin !== "object") continue;
      if (Array.isArray(plugin.extensions) && typeof plugin.parse === "function") {
        for (const ext of plugin.extensions) {
          localParsers.set(ext.toLowerCase(), plugin.parse);
        }
      }
      if (typeof plugin.detect === "function") {
        localSniffers.add(plugin.detect);
      }
    }
  }
  addPlugins(pluginList);
  const context = { localParsers, localSniffers };
  const reader = (file, options = {}) => readSpreadsheet(file, options, context);
  reader.tabular = (file, options = {}) => parseTabular(file, options, context);
  reader.iterate = (file, options = {}) => iterateSpreadsheet(file, options, context);
  reader.register = (...extra) => {
    addPlugins(extra);
    return reader;
  };
  return reader;
}

// src/schema/types/additional/Integer.js
function Integer(value) {
  value = NumberType(value);
  if (!isInteger(value)) {
    throw new InvalidError("not_an_integer");
  }
  return value;
}
function isInteger(x) {
  return (x | 0) === x;
}

// src/schema/types/additional/Email.js
function Email(value) {
  if (typeof value === "string") {
    if (isEmail(value)) {
      return value;
    }
    throw new InvalidError("not_an_email");
  }
  throw new InvalidError("not_a_string");
}
var EMAIL_REGEXP = /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i;
function isEmail(value) {
  return EMAIL_REGEXP.test(value);
}

// src/schema/types/additional/URL.js
function URL(value) {
  if (typeof value === "string") {
    if (isURL(value)) {
      return value;
    }
    throw new InvalidError("not_a_url");
  }
  throw new InvalidError("not_a_string");
}
var URL_REGEXP = /^(?:(?:(?:https?|ftp):)?\/\/)(?:(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z0-9\u00a1-\uffff][a-z0-9\u00a1-\uffff_-]{0,62})?[a-z0-9\u00a1-\uffff]\.)*(?:[a-z\u00a1-\uffff]{2,}))(?::\d{2,5})?(?:[/?#]\S*)?$/i;
function isURL(value) {
  return URL_REGEXP.test(value);
}

// src/index.js
function readSheet(input, sheet, options) {
  if (!options && sheet && typeof sheet !== "number" && typeof sheet !== "string") {
    options = sheet;
    sheet = void 0;
  }
  return readSpreadsheet(input, { ...options, sheet });
}
var readXlsxFile = readSpreadsheet;
var tabularjs = parseTabular;
var parser = parseTabular;
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  Email,
  Integer,
  InvalidInputError,
  InvalidSpreadsheetError,
  SheetNotFoundError,
  URL,
  convertRowsToObjects,
  createReader,
  detectFormat,
  getWriterForExtension,
  isPluginRegistered,
  iterateSpreadsheet,
  parseSheetData,
  parseTabular,
  parseXLSX,
  parser,
  readSheet,
  readSpreadsheet,
  readXlsxFile,
  registerPlugin,
  tabularjs,
  unregisterPlugin
});
//# sourceMappingURL=index.cjs.map
