// server.ts
import express, { Request, Response } from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import cors from 'cors';
import path from 'path';
import admin from 'firebase-admin';

// Ініціалізація Firebase Admin
try {
  if (process.env.FIREBASE_SERVICE_ACCOUNT) {
    // Production - використовуємо закодований в base64 JSON
    const serviceAccountJson = Buffer.from(
      process.env.FIREBASE_SERVICE_ACCOUNT,
      'base64'
    ).toString('utf-8');
    
    const serviceAccount = JSON.parse(serviceAccountJson);
    
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount)
    });
  } else if (process.env.FIREBASE_PROJECT_ID && process.env.FIREBASE_PRIVATE_KEY) {
    // Альтернатива: окремі змінні з правильною обробкою newline
    admin.initializeApp({
      credential: admin.credential.cert({
        projectId: process.env.FIREBASE_PROJECT_ID,
        privateKey: process.env.FIREBASE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
      })
    });
  } else {
    // Development - файл
    const serviceAccount = require("./serviceAccountKey.json");
    admin.initializeApp({
      credential: admin.credential.cert(serviceAccount)
    });
  }
} catch (error) {
  console.error('Firebase initialization error:', error);
  process.exit(1);
}

const db = admin.firestore();
const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors({
  origin: [
    'http://localhost:5173',
    'http://localhost:3000',
    'https://gurman-admin.vercel.app', 
    process.env.FRONTEND_URL || '*'
  ],
  credentials: true
}));
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
  limits: { fileSize: 10 * 1024 * 1024 }
});

// Типи
interface ExcelData {
  headers: string[];
  rows: any[][];
  fileName: string;
  rowCount: number;
}

function convertRowsForFirestore(headers: string[], rows: any[][]) {
  return rows.map((row, index) => {
    const rowObject: any = {
      rowIndex: index,
    };

    headers.forEach((header, headerIndex) => {
      const fieldName = `col_${headerIndex}`;
      rowObject[fieldName] = row[headerIndex]?.toString() || '';
    });

    return rowObject;
  });
}

function generateDocumentId(fileName: string): string {
  const baseName = fileName
    .replace(/\.(xlsx|xls)$/i, '')
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .toLowerCase();

  const timestamp = Date.now();
  return `${baseName}_${timestamp}`;
}

async function saveToFirestore(data: ExcelData, documentId?: string) {
  try {
    const collectionRef = db.collection('excel_data');
    const convertedRows = convertRowsForFirestore(data.headers, data.rows);

    const docData = {
      fileName: data.fileName,
      headers: data.headers,
      rowsData: convertedRows,
      rowCount: data.rowCount,
      uploadedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    };

    let finalDocId: string;

    if (documentId) {
      finalDocId = documentId;
      await collectionRef.doc(documentId).set(docData, { merge: true });
    } else {
      finalDocId = generateDocumentId(data.fileName);
      await collectionRef.doc(finalDocId).set(docData);
    }

    return {
      id: finalDocId,
      success: true,
      message: 'Дані успішно збережено в Firestore'
    };
  } catch (error) {
    console.error('Firebase save error:', error);
    throw error;
  }
}

app.post('/api/upload', upload.single('file'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не завантажено' });
    }

    const customId = req.body.documentId;

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
      fileName: customId,
      rowCount: rows.length
    };

    const firestoreResult = await saveToFirestore(excelData, customId);

    res.json({
      ...excelData,
      firestore: firestoreResult
    });
  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({
      error: 'Помилка обробки файлу',
      details: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

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
    console.error('Get files error:', error);
    res.status(500).json({ error: 'Помилка отримання даних' });
  }
});

function convertFirestoreToRows(headers: string[], rowsData: any[]) {
  return rowsData.map(rowObj => {
    const row: any[] = [];
    headers.forEach((header, index) => {
      row.push(rowObj[`col_${index}`] || '');
    });
    return row;
  });
}

app.get('/api/files/:id', async (req: Request, res: Response) => {
  try {
    const docRef = db.collection('excel_data').doc(req.params.id);
    const doc = await docRef.get();

    if (!doc.exists) {
      return res.status(404).json({ error: 'Файл не знайдено' });
    }

    const data = doc.data();
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
    console.error('Get file error:', error);
    res.status(500).json({ error: 'Помилка отримання даних' });
  }
});

app.delete('/api/files/:id', async (req: Request, res: Response) => {
  try {
    await db.collection('excel_data').doc(req.params.id).delete();
    res.json({ success: true, message: 'Файл видалено' });
  } catch (error) {
    console.error('Delete error:', error);
    res.status(500).json({ error: 'Помилка видалення файлу' });
  }
});

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
    console.error('Search error:', error);
    res.status(500).json({ error: 'Помилка пошуку' });
  }
});

app.get('/api/health', (req: Request, res: Response) => {
  res.json({
    status: 'OK',
    timestamp: new Date().toISOString(),
    firebase: 'Connected'
  });
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`📊 API available at http://localhost:${PORT}/api`);
  console.log(`🔥 Firebase Firestore connected`);
});

export default app;