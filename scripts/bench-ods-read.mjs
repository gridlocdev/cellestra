// Measures reading a large generated .ods: wall time and the process's peak resident memory.
//
//   node scripts/bench-ods-read.mjs generate <file> [rows] [columns]
//   node scripts/bench-ods-read.mjs read <file>
//
// Run `read` in a fresh process per measurement so the peak is the read's alone.
import { readFileSync, writeFileSync } from 'node:fs';
import { CellStyles, CellValue, OdsReader, OdsWriter, Workbook } from '../packages/core/src/index.ts';

const [mode, file, rowsArg, columnsArg] = process.argv.slice(2);
if (!file || (mode !== 'generate' && mode !== 'read')) {
  console.error('usage: bench-ods-read.mjs generate|read <file> [rows] [columns]');
  process.exit(2);
}

if (mode === 'generate') {
  const rows = Number(rowsArg ?? 100_000);
  const columns = Number(columnsArg ?? 10);
  const workbook = new Workbook();
  const sheet = workbook.addSheet('Data');
  const bold = workbook.styles.intern({ ...CellStyles.default, bold: true });
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < columns; c++) {
      const cell = sheet.getOrCreate(r, c);
      if (c === 0) cell.value = CellValue.text(`Item ${r}`);
      else if (c === columns - 1 && r > 0) {
        cell.formula = `SUM(B${r + 1}:${String.fromCharCode(65 + columns - 2)}${r + 1})`;
        cell.value = CellValue.number(r * c);
      } else cell.value = CellValue.number(r * columns + c + 0.25);
      if (r % 7 === 0) cell.styleId = bold;
    }
  }
  const bytes = OdsWriter.write(workbook);
  writeFileSync(file, bytes);
  console.log(`wrote ${file}: ${rows} x ${columns} cells, ${(bytes.length / 2 ** 20).toFixed(1)} MiB`);
} else {
  const bytes = new Uint8Array(readFileSync(file));
  const start = process.hrtime.bigint();
  const workbook = OdsReader.read(bytes);
  const ms = Number(process.hrtime.bigint() - start) / 1e6;
  const peak = process.resourceUsage().maxRSS / 1024;
  const sheet = workbook.worksheets[0];
  console.log(
    `read ${file}: ${ms.toFixed(0)} ms, peak RSS ${peak.toFixed(0)} MiB, ${sheet?.cellCount ?? '?'} cells`,
  );
}
