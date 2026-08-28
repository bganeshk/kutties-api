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

/** Flatten any ExcelJS cell value to a plain JS primitive. */
function cellValue(raw: ExcelJS.CellValue): string | number | boolean | null {
  if (raw === null || raw === undefined) return null;
  // Rich-text object: { richText: Array<{text: string}> }
  if (typeof raw === 'object' && 'richText' in (raw as object)) {
    return (raw as any).richText.map((r: any) => r.text ?? '').join('');
  }
  // Hyperlink object: { text: string, hyperlink: string }
  if (typeof raw === 'object' && 'text' in (raw as object)) {
    return String((raw as any).text ?? '');
  }
  // Date object
  if (raw instanceof Date) return raw.toISOString();
  // Formula result: { formula: string, result: ... }
  if (typeof raw === 'object' && 'result' in (raw as object)) {
    return cellValue((raw as any).result);
  }
  if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') return raw;
  return String(raw);
}

function sheetToRows(ws: ExcelJS.Worksheet): Row[] {
  const rows: Row[] = [];
  const headerRow = ws.getRow(1);
  const headers: string[] = [];

  headerRow.eachCell((cell) => {
    headers.push(String(cell.value ?? ''));
  });

  if (headers.length === 0) return rows;

  // detect which column acts as the primary key
  const idCol = headers.includes('id') ? 'id'
    : headers.includes('txid') ? 'txid'
    : null;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj: Row = { id: '' };
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const key = headers[colNumber - 1];
      if (key) obj[key] = cellValue(cell.value);
    });
    // synthesise a stable id from the pk column or fall back to row number
    if (idCol) {
      obj.id = String(obj[idCol] ?? '');
    } else {
      obj.id = String(rowNumber);
    }
    if (obj.id) rows.push(obj);
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

  // Case-insensitive id column lookup (handles 'id', 'Id', 'ID', etc.)
  const idColIndex = headers.findIndex((h) => h.toLowerCase() === 'id');
  if (idColIndex === -1) return null;
  const idCol = idColIndex + 1;

  // Build a case-insensitive map from lowercased payload key → actual value
  // so header 'FullName' matches payload key 'fullName'
  const dataLower: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(data)) {
    dataLower[k.toLowerCase()] = v;
  }

  let found: Row | null = null;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(idCol).value) === id) {
      const updated: Row = { id };
      headers.forEach((h, i) => {
        if (h.toLowerCase() === 'id') return;
        const newVal = dataLower[h.toLowerCase()] !== undefined
          ? dataLower[h.toLowerCase()]
          : row.getCell(i + 1).value;
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
  const idColIndex = headers.findIndex((h) => h.toLowerCase() === 'id');
  if (idColIndex === -1) return false;
  const idCol = idColIndex + 1;

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
