// server.ts
import express, { Request, Response } from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import cors from 'cors';
import path from 'path';

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json());

// Налаштування multer для завантаження файлів
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    if (ext === '.xlsx' || ext === '.xls') {
      cb(null, true);
    } else {
      cb(new Error('Тільки Excel файли дозволені'));
    }
  },
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB
});

// Типи
interface ExcelData {
  headers: string[];
  rows: any[][];
  fileName: string;
  rowCount: number;
}

interface ParsedSheet {
  sheetName: string;
  data: ExcelData;
}

// Маршрут для завантаження та парсингу Excel
app.post('/api/upload', upload.single('file'), (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не завантажено' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
   
    console.info('Check parsed data:', jsonData);
    
    if (jsonData.length === 0) {
      return res.status(400).json({ error: 'Файл порожній' });
    }

    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    const response: ExcelData = {
      headers,
      rows,
      fileName: req.file.originalname,
      rowCount: rows.length
    };

    res.json(response);
  } catch (error) {
    console.error('Помилка парсингу:', error);
    res.status(500).json({ error: 'Помилка обробки файлу' });
  }
});

// Маршрут для завантаження всіх листів
app.post('/api/upload/all-sheets', upload.single('file'), (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не завантажено' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheets: ParsedSheet[] = [];

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (jsonData.length > 0) {
        const headers = jsonData[0];
        const rows = jsonData.slice(1);

        sheets.push({
          sheetName,
          data: {
            headers,
            rows,
            fileName: req.file!.originalname,
            rowCount: rows.length
          }
        });
      }
    });

    res.json({ sheets, totalSheets: sheets.length });
  } catch (error) {
    console.error('Помилка парсингу:', error);
    res.status(500).json({ error: 'Помилка обробки файлу' });
  }
});

// Маршрут для пошуку в даних
app.post('/api/search', express.json(), (req: Request, res: Response) => {
  try {
    const { data, searchTerm } = req.body;

    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: 'Невірний формат даних' });
    }

    const filtered = data.filter((row: any[]) =>
      row.some((cell) =>
        cell?.toString().toLowerCase().includes(searchTerm.toLowerCase())
      )
    );

    res.json({ results: filtered, count: filtered.length });
  } catch (error) {
    console.error('Помилка пошуку:', error);
    res.status(500).json({ error: 'Помилка пошуку' });
  }
});

// Health check
app.get('/api/health', (req: Request, res: Response) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Запуск сервера
app.listen(PORT, () => {
  console.log(`🚀 Сервер запущено на http://localhost:${PORT}`);
  console.log(`📊 API доступне на http://localhost:${PORT}/api`);
});

export default app;