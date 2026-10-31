// Local addition (not in write-excel-file): the `rowBreaks` sheet option, manual page breaks.
//
// `rowBreaks: number[]` holds 1-based row numbers; a page break is written after each of them
// (the same meaning as ExcelJS's `worksheet.getRow(r).addPageBreak()`). The element is written the way Excel and
// ExcelJS write it: `<rowBreaks count="n" manualBreakCount="n"><brk id="r" max="16383" man="1"/></rowBreaks>`.
import insertElementMarkupAccordingToOrderOfSiblings from '../../xml/insertElementMarkupAccordingToOrderOfSiblings.js';
import getOrderOfSiblings from '../helpers/getOrderOfSiblings.js';

// 16383 is the last column index (XFD), the usual `max` of a row break.
var LAST_COLUMN_INDEX = 16383;

export default {
  files: {
    transform: {
      'xl/worksheets/sheet{id}.xml': {
        // Placed by the schema order of `<worksheet/>` children (after `<pageSetup/>` and `<headerFooter/>`,
        // before `<drawing/>`), and `<drawing/>` is already in the XML by the time `transform()` runs.
        transform: function transform(xml, sheetOptions) {
          var rows = normalizeRowBreaks(sheetOptions.rowBreaks);
          if (rows.length === 0) {
            return xml;
          }
          var breaksXml = rows.map(function (row) {
            return '<brk id="' + row + '" max="' + LAST_COLUMN_INDEX + '" man="1"/>';
          }).join('');
          var elementXml = '<rowBreaks count="' + rows.length + '" manualBreakCount="' + rows.length + '">' + breaksXml + '</rowBreaks>';
          return insertElementMarkupAccordingToOrderOfSiblings(xml, elementXml, getOrderOfSiblings('xl/worksheets/sheet{id}.xml', 'worksheet'), 'worksheet');
        }
      }
    }
  }
};

/**
 * Validates `rowBreaks`, then sorts it and drops duplicates.
 * @param {number[]} [rowBreaks]
 * @returns {number[]}
 */
export function normalizeRowBreaks(rowBreaks) {
  if (rowBreaks === undefined || rowBreaks === null) {
    return [];
  }
  if (!Array.isArray(rowBreaks)) {
    throw new Error('`rowBreaks` must be an array of 1-based row numbers');
  }
  rowBreaks.forEach(function (row) {
    if (typeof row !== 'number' || !Number.isInteger(row) || row < 1) {
      throw new Error('`rowBreaks` must hold positive integers (1-based row numbers), got: ' + String(row));
    }
  });
  return rowBreaks.filter(function (row, i) {
    return rowBreaks.indexOf(row) === i;
  }).sort(function (a, b) {
    return a - b;
  });
}
