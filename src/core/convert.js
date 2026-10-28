import { createConverter } from './convert-base.js';
import { writeSpreadsheet as defaultNodeWriter } from '../write/index.js';

export { createConverter };
export const convertSpreadsheet = createConverter(defaultNodeWriter);
export default convertSpreadsheet;
