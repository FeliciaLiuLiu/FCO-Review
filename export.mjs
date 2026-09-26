import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

let input = '';
for await (const chunk of process.stdin) input += chunk;
const { headers, rows, output, preview } = JSON.parse(input);
const workbook = Workbook.create();
const sheet = workbook.worksheets.add('Transactions');
sheet.showGridLines = false;
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(1);
sheet.getRange('A1:AS1').values = [headers];
sheet.getRange('A1:AS1').format = {
  fill: '#203864', font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' },
  wrapText: true, rowHeight: 64, columnWidth: 25,
  horizontalAlignment: 'center', verticalAlignment: 'center',
};
sheet.getRange('C1').format.columnWidth = 36;
for (let offset = 0; offset < rows.length; offset += 5000) {
  const block = rows.slice(offset, offset + 5000).map(row => row.map(value =>
    typeof value === 'string' && value.startsWith('=') ? "'" + value : value));
  const range = sheet.getRangeByIndexes(offset + 1, 0, block.length, headers.length);
  range.setNumberFormat('@');
  range.values = block;
  range.format.font = { name: 'Arial', size: 10 };
  sheet.getRangeByIndexes(offset + 1, 3, block.length, 1).setNumberFormat('0.00#############');
}
workbook.recalculate();
if (sheet.getRange('A1:AS1').values[0].some((v, i) => v !== headers[i])) {
  throw new Error('Output header validation failed');
}
await fs.mkdir(path.dirname(output), { recursive: true });
if (preview) {
  for (const [index, range] of ['A1:K2', 'L1:U2', 'V1:AF2', 'AG1:AS2'].entries()) {
    const image = await workbook.render({ sheetName: 'Transactions', range, scale: 1, format: 'png' });
    await fs.writeFile(path.join(path.dirname(output), `preview-${index + 1}.png`),
      new Uint8Array(await image.arrayBuffer()));
  }
}
const temp = output + '.tmp.xlsx';
try {
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(temp);
  await fs.rename(temp, output);
} finally {
  await fs.rm(temp, { force: true });
}
