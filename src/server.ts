// server.ts
import express, { Request, Response } from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import cors from 'cors';
import path from 'path';
import admin from 'firebase-admin';

// Ініціалізація Firebase Admin
// ВАЖЛИВО: Завантажте свій serviceAccountKey.json з Firebase Console
const serviceAccount = require('../serviceAccountKey.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

const app = express();
const PORT = process.env.PORT || 3007;

// Middleware
app.use(cors());
app.use(express.json());

// Налаштування multer
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

// Функція для збереження в Firestore
async function saveToFirestore(data: ExcelData, documentId?: string) {
  try {
    const collectionRef = db.collection('excel_data');
    console.info(`saveToFirestore: ${data.fileName}, documentId: ${documentId}`);
    const docData = {
      fileName: data.fileName,
      headers: data.headers,
      rows: data.rows,
      rowCount: data.rowCount,
      uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    };
    console.info(`saveToFirestore:2  headers: ${docData.headers.length}, rows: ${docData.rows.length}`);

    let docRef;
    if (documentId) {
      console.info(`saveToFirestore:3 updating document ${documentId} start`);
      // Оновлення існуючого документа
      docRef = collectionRef.doc(documentId);
      await docRef.update({
        ...docData,
        updatedAt: admin.firestore.FieldValue.serverTimestamp()
      });
      console.info(`saveToFirestore:3 updating document ${documentId} end`);

    } else {
      // Створення нового документа
      console.info(`saveToFirestore:4 adding new document start ${JSON.stringify(docData)}`);

      docRef = await collectionRef.add(docData);
      console.info(`saveToFirestore:4 adding new document end`);
    }

    return {
      id: docRef.id,
      success: true,
      message: 'Дані успішно збережено в Firestore'
    };
  } catch (error) {
    console.error('Помилка збереження в Firestore:', error);
    throw error;
  }
}

// Маршрут для завантаження та збереження Excel
app.post('/api/upload', upload.single('file'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не завантажено' });
    }

    // Парсинг Excel
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length === 0) {
      return res.status(400).json({ error: 'Файл порожній' });
    }

    const headers = jsonData[0];
    const rows = jsonData.slice(1);

    const excelData: ExcelData = {
      headers,
      rows,
      fileName: req.file.originalname,
      rowCount: rows.length
    };

    // Збереження в Firestore
    const firestoreResult = await saveToFirestore(excelData);// "test_excel_db"

    res.json({
      ...excelData,
      firestore: firestoreResult
    });
  } catch (error) {
    console.error('Помилка обробки:', error);
    res.status(500).json({
      error: 'Помилка обробки файлу',
      details: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

// Маршрут для отримання списку всіх завантажених файлів
app.get('/api/files', async (req: Request, res: Response) => {
  try {
    const snapshot = await db.collection('excel_data')
      .orderBy('uploadedAt', 'desc')
      .get();

    const files = snapshot.docs.map(doc => ({
      id: doc.id,
      fileName: doc.data().fileName,
      rowCount: doc.data().rowCount,
      uploadedAt: doc.data().uploadedAt,
      updatedAt: doc.data().updatedAt
    }));

    res.json({ files, count: files.length });
  } catch (error) {
    console.error('Помилка отримання файлів:', error);
    res.status(500).json({ error: 'Помилка отримання даних' });
  }
});

// Маршрут для отримання конкретного файлу
app.get('/api/files/:id', async (req: Request, res: Response) => {
  try {
    const docRef = db.collection('excel_data').doc(req.params.id);
    const doc = await docRef.get();

    if (!doc.exists) {
      return res.status(404).json({ error: 'Файл не знайдено' });
    }

    res.json({
      id: doc.id,
      ...doc.data()
    });
  } catch (error) {
    console.error('Помилка отримання файлу:', error);
    res.status(500).json({ error: 'Помилка отримання даних' });
  }
});

// Маршрут для видалення файлу
app.delete('/api/files/:id', async (req: Request, res: Response) => {
  try {
    await db.collection('excel_data').doc(req.params.id).delete();
    res.json({ success: true, message: 'Файл видалено' });
  } catch (error) {
    console.error('Помилка видалення:', error);
    res.status(500).json({ error: 'Помилка видалення файлу' });
  }
});

// Маршрут для пошуку в Firestore
app.post('/api/search', async (req: Request, res: Response) => {
  try {
    const { searchTerm } = req.body;

    if (!searchTerm) {
      return res.status(400).json({ error: 'Пошуковий запит відсутній' });
    }

    // Отримати всі документи (в реальному проекті краще використовувати індексацію)
    const snapshot = await db.collection('excel_data').get();
    const results: any[] = [];

    snapshot.docs.forEach(doc => {
      const data = doc.data();
      const matchingRows = data.rows.filter((row: any[]) =>
        row.some((cell: any) =>
          cell?.toString().toLowerCase().includes(searchTerm.toLowerCase())
        )
      );

      if (matchingRows.length > 0) {
        results.push({
          id: doc.id,
          fileName: data.fileName,
          matchingRows,
          matchCount: matchingRows.length
        });
      }
    });

    res.json({ results, totalMatches: results.length });
  } catch (error) {
    console.error('Помилка пошуку:', error);
    res.status(500).json({ error: 'Помилка пошуку' });
  }
});

// Health check
app.get('/api/health', (req: Request, res: Response) => {
  res.json({
    status: 'OK',
    timestamp: new Date().toISOString(),
    firebase: 'Connected'
  });
});

// // Запуск сервера
// app.listen(PORT, () => {
//   console.log(`🚀 Сервер запущено на http://localhost:${PORT}`);
//   console.log(`📊 API доступне на http://localhost:${PORT}/api`);
//   console.log(`🔥 Firebase Firestore підключено`);
// });

// Замініть app.listen на:
const startServer = async (port: number) => {
  try {
    app.listen(port, () => {
      console.log(`🚀 Сервер запущено на http://localhost:${port}`);
      console.log(`📊 API доступне на http://localhost:${port}/api`);
      console.log(`🔥 Firebase Firestore підключено`);
    });
  } catch (error) {
    if ((error as any).code === 'EADDRINUSE') {
      console.log(`⚠️ Порт ${port} зайнятий, пробую ${port + 1}...`);
      startServer(port + 1);
    } else {
      throw error;
    }
  }
};

startServer(Number(PORT));


export default app;