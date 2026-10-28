import { createConverter } from './convert-base.js';
import { writeSpreadsheet as browserWriter } from '../write/browser.js';

export { createConverter };
export const convertSpreadsheet = createConverter(browserWriter);
export default convertSpreadsheet;
