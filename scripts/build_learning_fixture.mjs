// Small, original synthetic source for inexpensive model-comparison trials.
import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const root = path.resolve(import.meta.dirname, '..');
const output = path.join(root, 'fixtures', 'customer_cube_learning.xlsx');
const qa = path.join(root, 'outputs', 'workbench', 'learning-fixture-qa');
await fs.mkdir(qa, {recursive:true});
const records = [
  ['Sample Alpha','2025-01-01',100,'Core'],
  ['Sample Alpha','2025-02-01',100,'Core'],
  ['Sample Alpha','2025-03-01',100,'Core'],
  ['Sample Beta','2025-03-01',50,'Core'],
  ['Sample Alpha','2025-04-01',120,'Core'],
  ['Sample Alpha','2025-04-01',30,'Add-on'],
  ['Sample Alpha','2025-05-01',80,'Core'],
  ['Sample Alpha','2025-05-01',30,'Add-on'],
  ['Sample Alpha','2025-06-01',80,'Core'],
  ['Sample Alpha','2025-06-01',30,'Add-on'],
  ['Sample Gamma','2025-06-01',40,'Core'],
  ['Sample Gamma','2025-06-01',10,'Core'],
];
const workbook = Workbook.create();
const sheet = workbook.worksheets.add('Raw Data');
sheet.showGridLines = false;
sheet.getRange('A1:D1').values = [['Customer name','Month','Amount','Product']];
sheet.getRange('A2:D13').values = records.map(([customer,month,amount,product]) => [customer,new Date(month+'T00:00:00Z'),amount,product]);
sheet.getRange('A1:D13').format.font = {name:'Arial',size:11};
sheet.getRange('A1:D1').format = {fill:'#217c66',font:{name:'Arial',size:11,bold:true,color:'#ffffff'},rowHeight:26};
sheet.getRange('A2:D13').format.rowHeight = 23;
sheet.getRange('C2:C13').setNumberFormat('#,##0;(#,##0);-');
sheet.getRange('B2:B13').setNumberFormat('mmm-yy');
sheet.getRange('A1:A13').format.columnWidth=24;
sheet.getRange('B1:B13').format.columnWidth=17;
sheet.getRange('C1:C13').format.columnWidth=16;
sheet.getRange('D1:D13').format.columnWidth=18;
sheet.freezePanes.freezeRows(1);
workbook.recalculate();
const inspection = await workbook.inspect({kind:'table',range:'Raw Data!A1:D13',include:'values',tableMaxRows:14,tableMaxCols:4});
await fs.writeFile(path.join(qa,'inspection.txt'),inspection.ndjson);
const preview = await workbook.render({sheetName:'Raw Data',range:'A1:D13',scale:1.5,format:'png'});
await fs.writeFile(path.join(qa,'source.png'),new Uint8Array(await preview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(workbook)).save(output);
console.log('Created 12-row original synthetic learning fixture. Missing customer/product months are intentionally zero.');
