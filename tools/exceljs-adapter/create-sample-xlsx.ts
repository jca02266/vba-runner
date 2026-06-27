/**
 * テスト用 input.xlsx を生成するスクリプト
 * Usage: npx tsx tools/exceljs-adapter/create-sample-xlsx.ts
 */
import ExcelJS from 'exceljs';
import * as path from 'path';

const outPath = path.join(import.meta.dirname, 'input.xlsx');

const wb = new ExcelJS.Workbook();

// Sales シート（VBA マクロの読み込み対象）
const sales = wb.addWorksheet('Sales');
sales.addRow(['担当者', '商品名', '売上金額']);
sales.addRow(['Alice', 'Widget A', 12000]);
sales.addRow(['Bob',   'Widget B',  8500]);
sales.addRow(['Carol', 'Widget A', 21000]);
sales.addRow(['Dave',  'Widget C',  6500]);
sales.addRow(['Eve',   'Widget B', 14500]);
sales.addRow(['Frank', 'Widget C',  9800]);
sales.addRow(['Grace', 'Widget A', 17200]);

// Summary シート（VBA マクロが書き出す先）
wb.addWorksheet('Summary');

await wb.xlsx.writeFile(outPath);
console.log(`Created: ${outPath}`);
