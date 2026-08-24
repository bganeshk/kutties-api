import ExcelJS from 'exceljs';
import path from 'path';
import { v4 as uuidv4 } from 'uuid';
import { Row, QueryOptions } from '../types';

const DB_PATH = path.resolve(__dirname, '../../data/database.xlsx');

async function loadWorkbook(): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.readFile(DB_PATH);
  } catch {
    // file doesn't exist yet — return empty workbook
  }
  return wb;
}

async function saveWorkbook(wb: ExcelJS.Workbook): Promise<void> {
  await wb.xlsx.writeFile(DB_PATH);
}

function getOrCreateSheet(wb: ExcelJS.Workbook, sheet: string): ExcelJS.Worksheet {
  return wb.getWorksheet(sheet) ?? wb.addWorksheet(sheet);
}

function sheetToRows(ws: ExcelJS.Worksheet): Row[] {
  const rows: Row[] = [];
  const headerRow = ws.getRow(1);
  const headers: string[] = [];

  headerRow.eachCell((cell) => {
    headers.push(String(cell.value ?? ''));
  });

  if (headers.length === 0) return rows;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj: Row = { id: '' };
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const key = headers[colNumber - 1];
      if (key) obj[key] = cell.value ?? null;
    });
    if (obj['id']) rows.push(obj);
  });

  return rows;
}

function ensureHeaders(ws: ExcelJS.Worksheet, data: Record<string, unknown>): void {
  const headerRow = ws.getRow(1);
  const existing: string[] = [];
  headerRow.eachCell((cell) => existing.push(String(cell.value ?? '')));

  const allKeys = ['id', ...Object.keys(data).filter((k) => k !== 'id')];
  allKeys.forEach((key, i) => {
    if (!existing.includes(key)) {
      ws.getRow(1).getCell(existing.length + i + 1).value = key;
    }
  });
  ws.getRow(1).commit();
}

function getHeaders(ws: ExcelJS.Worksheet): string[] {
  const headers: string[] = [];
  ws.getRow(1).eachCell((cell) => headers.push(String(cell.value ?? '')));
  return headers;
}

export async function listRows(sheet: string, opts: QueryOptions = {}): Promise<{ rows: Row[]; total: number }> {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(sheet);
  if (!ws) return { rows: [], total: 0 };

  let rows = sheetToRows(ws);

  if (opts.filter) {
    for (const [key, value] of Object.entries(opts.filter)) {
      rows = rows.filter((r) => String(r[key] ?? '').toLowerCase().includes(value.toLowerCase()));
    }
  }

  const total = rows.length;
  const offset = opts.offset ?? 0;
  const limit = opts.limit ?? total;
  return { rows: rows.slice(offset, offset + limit), total };
}

export async function getRow(sheet: string, id: string): Promise<Row | null> {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(sheet);
  if (!ws) return null;
  return sheetToRows(ws).find((r) => r.id === id) ?? null;
}

export async function insertRow(sheet: string, data: Record<string, unknown>): Promise<Row> {
  const wb = await loadWorkbook();
  const ws = getOrCreateSheet(wb, sheet);

  ensureHeaders(ws, data);
  const headers = getHeaders(ws);

  const newRow: Row = { id: uuidv4(), ...data };
  const rowValues = headers.map((h) => newRow[h] ?? null);
  ws.addRow(rowValues).commit();

  await saveWorkbook(wb);
  return newRow;
}

export async function updateRow(sheet: string, id: string, data: Record<string, unknown>): Promise<Row | null> {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(sheet);
  if (!ws) return null;

  const headers = getHeaders(ws);
  const idCol = headers.indexOf('id') + 1;
  if (idCol === 0) return null;

  let found: Row | null = null;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(idCol).value) === id) {
      const updated: Row = { id };
      headers.forEach((h, i) => {
        if (h === 'id') return;
        const newVal = data[h] !== undefined ? data[h] : row.getCell(i + 1).value;
        row.getCell(i + 1).value = newVal as ExcelJS.CellValue;
        updated[h] = newVal;
      });
      row.commit();
      found = updated;
    }
  });

  if (found) await saveWorkbook(wb);
  return found;
}

export async function deleteRow(sheet: string, id: string): Promise<boolean> {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(sheet);
  if (!ws) return false;

  const headers = getHeaders(ws);
  const idCol = headers.indexOf('id') + 1;
  if (idCol === 0) return false;

  let deletedRowNumber = -1;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(idCol).value) === id) {
      deletedRowNumber = rowNumber;
    }
  });

  if (deletedRowNumber === -1) return false;

  ws.spliceRows(deletedRowNumber, 1);
  await saveWorkbook(wb);
  return true;
}

export async function listSheets(): Promise<string[]> {
  const wb = await loadWorkbook();
  return wb.worksheets.map((ws) => ws.name);
}
