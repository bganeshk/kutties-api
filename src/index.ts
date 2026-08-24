import express from 'express';
import cors from 'cors';
import crudRouter from './routes/crud';

const app = express();
const PORT = process.env.PORT ?? 3000;

app.use(cors());
app.use(express.json());

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
