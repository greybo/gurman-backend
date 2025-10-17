// server.ts
import express, { Request, Response } from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import cors from 'cors';
import path from 'path';
import admin from 'firebase-admin';

// Ініціалізація Firebase Admin
// ВАЖЛИВО: Завантажте свій serviceAccountKey.json з Firebase Console
const serviceAccount = require('./serviceAccountKey.json');

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount)
});

const db = admin.firestore();

const app = express();
const PORT = process.env.PORT || 3001;

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

// Функція для конвертації rows в Firestore-сумісний формат
function convertRowsForFirestore(headers: string[], rows: any[][]) {
  return rows.map((row, index) => {
    const rowObject: any = {
      rowIndex: index,
    };

    headers.forEach((header, headerIndex) => {
      // Очищаємо назву поля від спецсимволів
      const fieldName = `col_${headerIndex}`;
      rowObject[fieldName] = row[headerIndex]?.toString() || '';
    });

    return rowObject;
  });
}

// Функція для генерації безпечного ID з імені файлу
function generateDocumentId(fileName: string): string {
  // Видаляємо розширення і спецсимволи
  const baseName = fileName
    .replace(/\.(xlsx|xls)$/i, '') // Видаляємо розширення
    .replace(/[^a-zA-Z0-9_-]/g, '_') // Замінюємо спецсимволи на _
    .toLowerCase();

  // Додаємо timestamp для унікальності
  const timestamp = Date.now();

  return `${baseName}_${timestamp}`;
}

// Функція для збереження в Firestore
async function saveToFirestore(data: ExcelData, documentId?: string) {
  try {
    const collectionRef = db.collection('excel_data');

    // Конвертуємо вкладені масиви в об'єкти
    const convertedRows = convertRowsForFirestore(data.headers, data.rows);

    const docData = {
      fileName: data.fileName,
      headers: data.headers,
      rowsData: convertedRows,
      rowCount: data.rowCount,
      uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    let docRef;
    let finalDocId: string;

    if (documentId) {
      // Використовуємо переданий ID
      finalDocId = documentId;
      docRef = collectionRef.doc(documentId);
      await docRef.set(docData, { merge: true });
    } else {
      // Генеруємо ID з імені файлу
      finalDocId = generateDocumentId(data.fileName);
      docRef = collectionRef.doc(finalDocId);
      await docRef.set(docData);
    }

    return {
      id: finalDocId,
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

    // ОтримуємоCustom ID з body (опціонально)
    const customId = req.body.documentId;

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
      fileName: customId,//req.file.originalname,
      rowCount: rows.length
    };

    // Збереження в Firestore з custom ID (якщо переданий)
    const firestoreResult = await saveToFirestore(excelData, customId);

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

// Функція для конвертації назад в масиви (при читанні)
function convertFirestoreToRows(headers: string[], rowsData: any[]) {
  return rowsData.map(rowObj => {
    const row: any[] = [];
    headers.forEach((header, index) => {
      row.push(rowObj[`col_${index}`] || '');
    });
    return row;
  });
}

// Маршрут для отримання конкретного файлу
app.get('/api/files/:id', async (req: Request, res: Response) => {
  try {
    const docRef = db.collection('excel_data').doc(req.params.id);
    const doc = await docRef.get();

    if (!doc.exists) {
      return res.status(404).json({ error: 'Файл не знайдено' });
    }

    const data = doc.data();

    // Конвертуємо назад в масиви для клієнта
    const rows = convertFirestoreToRows(data!.headers, data!.rowsData);

    res.json({
      id: doc.id,
      fileName: data!.fileName,
      headers: data!.headers,
      rows: rows,
      rowCount: data!.rowCount,
      uploadedAt: data!.uploadedAt,
      updatedAt: data!.updatedAt
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

    const snapshot = await db.collection('excel_data').get();
    const results: any[] = [];

    snapshot.docs.forEach(doc => {
      const data = doc.data();

      // Конвертуємо назад в масиви для пошуку
      const rows = convertFirestoreToRows(data.headers, data.rowsData);

      const matchingRows = rows.filter((row: any[]) =>
        row.some((cell: any) =>
          cell?.toString().toLowerCase().includes(searchTerm.toLowerCase())
        )
      );

      if (matchingRows.length > 0) {
        results.push({
          id: doc.id,
          fileName: data.fileName,
          headers: data.headers,
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

// Запуск сервера
app.listen(PORT, () => {
  console.log(`🚀 Сервер запущено на http://localhost:${PORT}`);
  console.log(`📊 API доступне на http://localhost:${PORT}/api`);
  console.log(`🔥 Firebase Firestore підключено`);
});

export default app;