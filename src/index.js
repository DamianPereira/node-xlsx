
import XLSX from 'xlsx';
import bufferFrom from 'buffer-from';
import {buildSheetFromMatrix, isString} from './helpers';
import Workbook from './workbook';

export function parse(mixed, options = {}) {
  const workSheet = XLSX[isString(mixed) ? 'readFile' : 'read'](mixed, options);
  return Object.keys(workSheet.Sheets).map((name) => {
    const sheet = workSheet.Sheets[name];
    return {name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: options.raw !== false})};
  });
}

export function build(worksheets, workbookOptions = {}) {
  const defaults = {
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'
  };
  const workBook = new Workbook();
  worksheets.forEach((worksheet) => {
    const { name = 'Sheet', data = [], ...worksheetOptions } = worksheet;
    const worksheetData = buildSheetFromMatrix(data, worksheetOptions);
    workBook.SheetNames.push(name);
    workBook.Sheets[name] = worksheetData;
  });
  const excelData = XLSX.write(workBook, workbookOptions);
  return excelData instanceof Buffer ? excelData : bufferFrom(excelData, 'binary');
}

export default {parse, build};
