import { Router, Request, Response } from 'express';
import * as excel from '../services/excel.service';
import { ApiResponse, Row } from '../types';

const router = Router();

// GET /api/sheets — list all sheets
router.get('/sheets', async (_req: Request, res: Response) => {
  const sheets = await excel.listSheets();
  res.json({ success: true, data: sheets } as ApiResponse<string[]>);
});

// GET /api/:sheet — list rows (supports ?filter[field]=value&limit=N&offset=N)
router.get('/:sheet', async (req: Request, res: Response) => {
  const { sheet } = req.params;
  const { limit, offset, sinceDate, ...rest } = req.query as Record<string, string>;

  const filter: Record<string, string> = {};
  for (const [key, value] of Object.entries(rest)) {
    const match = key.match(/^filter\[(.+)]$/);
    if (match) filter[match[1]] = value;
  }

  const { rows, total } = await excel.listRows(sheet, {
    filter: Object.keys(filter).length ? filter : undefined,
    limit: limit ? parseInt(limit) : undefined,
    offset: offset ? parseInt(offset) : undefined,
    sinceDate: sinceDate || undefined,
  });

  res.json({ success: true, data: rows, total } as ApiResponse<Row[]>);
});

// GET /api/:sheet/:id — get single row
router.get('/:sheet/:id', async (req: Request, res: Response) => {
  const { sheet, id } = req.params;
  const row = await excel.getRow(sheet, id);
  if (!row) {
    res.status(404).json({ success: false, error: 'Row not found' } as ApiResponse<null>);
    return;
  }
  res.json({ success: true, data: row } as ApiResponse<Row>);
});

// POST /api/:sheet — insert row
router.post('/:sheet', async (req: Request, res: Response) => {
  const { sheet } = req.params;
  const body = req.body as Record<string, unknown>;
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    res.status(400).json({ success: false, error: 'Request body must be a JSON object' });
    return;
  }
  const row = await excel.insertRow(sheet, body);
  res.status(201).json({ success: true, data: row } as ApiResponse<Row>);
});

// PUT /api/:sheet/:id — update row
router.put('/:sheet/:id', async (req: Request, res: Response) => {
  const { sheet, id } = req.params;
  const body = req.body as Record<string, unknown>;
  const updated = await excel.updateRow(sheet, id, body);
  if (!updated) {
    res.status(404).json({ success: false, error: 'Row not found' } as ApiResponse<null>);
    return;
  }
  res.json({ success: true, data: updated } as ApiResponse<Row>);
});

// DELETE /api/:sheet/:id — delete row
router.delete('/:sheet/:id', async (req: Request, res: Response) => {
  const { sheet, id } = req.params;
  const deleted = await excel.deleteRow(sheet, id);
  if (!deleted) {
    res.status(404).json({ success: false, error: 'Row not found' } as ApiResponse<null>);
    return;
  }
  res.json({ success: true, data: { id } } as ApiResponse<{ id: string }>);
});

export default router;
