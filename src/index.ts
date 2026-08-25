import express from 'express';
import cors from 'cors';
import path from 'path';
import crudRouter from './routes/crud';

const app = express();
const PORT = process.env.PORT ?? 3000;

// Serve dashboard images from the Expo app's assets folder
const ASSETS_DIR = process.env.ASSETS_DIR
  ?? path.resolve(__dirname, '../../kutties-app/assets');

app.use(cors());
app.use(express.json());
app.use('/assets', express.static(ASSETS_DIR));

app.get('/health', (_req, res) => res.json({ status: 'ok' }));

app.use('/api', crudRouter);

app.use((err: Error, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  console.error(err);
  res.status(500).json({ success: false, error: err.message });
});

app.listen(PORT, () => {
  console.log(`Excel API running on http://localhost:${PORT}`);
  console.log(`Database: data/database.xlsx`);
});
