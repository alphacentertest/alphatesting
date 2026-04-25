// Імпорт необхідних модулів
require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const { MongoClient, ObjectId } = require('mongodb');
const bcrypt = require('bcrypt');
const fs = require('fs');
const multer = require('multer');
const nodemailer = require('nodemailer');
const { body, validationResult } = require('express-validator');
const jwt = require('jsonwebtoken');
const winston = require('winston');
const session = require('express-session');
const MongoStore = require('connect-mongo');


// === ВИПРАВЛЕННЯ ЧАСОВОГО ПОЯСУ ДЛЯ VERCEL ===
process.env.TZ = 'Europe/Kiev';   // або 'Europe/Kyiv' (залежить від версії Node)
console.log('Серверний часовий пояс встановлено:', process.env.TZ);

// Ініціалізація Express-додатку
const app = express();

// Увімкнення довіри до проксі
app.set('trust proxy', 1);

// Налаштування логування
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.File({ filename: 'error.log', level: 'error' }),
    new winston.transports.File({ filename: 'combined.log' }),
    new winston.transports.Console()
  ]
});

// Налаштування multer
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 4 * 1024 * 1024 }
});

// Налаштування nodemailer
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
    pass: process.env.EMAIL_PASS || 'xfcd cvkl xiii qhtl'
  }
});

// Функція для відправки email
const sendSuspiciousActivityEmail = async (user, activityDetails) => {
  try {
    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: process.env.EMAIL_USER,
      subject: 'Підозріла активність',
      text: `
        Користувач: ${user}
        Час поза вкладкою: ${activityDetails.timeAwayPercent}%
        Переключення вкладок: ${activityDetails.switchCount}
        Середній час відповіді (сек): ${activityDetails.avgResponseTime}
        Загальна кількість дій: ${activityDetails.totalActivityCount}
      `
    };
    await transporter.sendMail(mailOptions);
    logger.info(`Email відправлено для ${user}`);
  } catch (error) {
    logger.error('Помилка відправки email', { message: error.message, stack: error.stack });
  }
};

// Конфігурація
const config = {
  suspiciousActivity: {
    timeAwayThreshold: 50,
    switchCountThreshold: 5
  }
};

// Налаштування MongoDB
const MONGODB_URI = process.env.MONGODB_URI;
const client = new MongoClient(MONGODB_URI, {
  connectTimeoutMS: 30000,          // 30 секунд
  serverSelectionTimeoutMS: 30000,  // 30 секунд
  socketTimeoutMS: 60000,           // 60 секунд на сокет
  retryWrites: true,
  w: 'majority',
  family: 4                         // примусово IPv4, іноді допомагає
});
let db;

// Клас CacheManager
class CacheManager {
  static cache = {};

  static async getOrFetch(key, testNumber, fetchFn) {
    const cacheKey = `${key}:${testNumber}`;
    if (this.cache[cacheKey]) {
      logger.info(`Кеш для ${cacheKey}`);
      return this.cache[cacheKey];
    }
    logger.info(`Промах кешу для ${cacheKey}`);
    const startTime = Date.now();
    const data = await fetchFn();
    this.cache[cacheKey] = data;
    logger.info(`Оновлено кеш ${key} за ${Date.now() - startTime} мс`);
    return data;
  }

  static async invalidateCache(key, testNumber) {
    const cacheKey = `${key}:${testNumber}`;
    delete this.cache[cacheKey];
    logger.info(`Інвалідовано кеш ${cacheKey}`);
  }

  static async getQuestions(testNumber) {
    return await this.getOrFetch('questions', testNumber, async () => {
      return await db.collection('questions').find({ testNumber }).sort({ order: 1 }).toArray();
    });
  }

  static async getAllQuestions() {
    return await this.getOrFetch('allQuestions', 'all', async () => {
      return await db.collection('questions').find({}).sort({ order: 1 }).toArray();
    });
  }
}

// Кеш
let userCache = [];
const questionsCache = {};

let isInitialized = false;
let initializationError = null;
let testNames = {};

// Підключення до MongoDB
const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    logger.info(`Спроба підключення до MongoDB (${attempt}/${maxAttempts})`, {
      uriSnippet: MONGODB_URI.substring(0, 50) + '...', // без пароля
      attempt
    });

    const startTime = Date.now();
    await client.connect();

    db = client.db('alpha');

    // Перевірка, чи база дійсно існує і доступна
    const dbName = db.databaseName;
    const collections = await db.listCollections().toArray();
    const collectionNames = collections.map(c => c.name);

    logger.info(`Підключено до MongoDB за ${Date.now() - startTime} мс`, {
      databaseName: dbName,
      cluster: MONGODB_URI.split('@')[1]?.split('.')[0] || 'невідомо',
      collectionsCount: collectionNames.length,
      collections: collectionNames.slice(0, 10).join(', ') + (collectionNames.length > 10 ? '...' : ''),
      hasTestResults: collectionNames.includes('test_results')
    });

    // Якщо колекції test_results немає — попереджаємо
    if (!collectionNames.includes('test_results')) {
      logger.warn('Колекція test_results відсутня в базі alpha — результати тестів не будуть відображатися!');
    }

  } catch (error) {
    logger.error('Помилка підключення до MongoDB', {
      message: error.message,
      stack: error.stack,
      attempt,
      uriSnippet: MONGODB_URI.substring(0, 50) + '...'
    });

    if (attempt < maxAttempts) {
      logger.info(`Повторна спроба через 5 секунд...`);
      await new Promise(resolve => setTimeout(resolve, 5000));
      return connectToMongoDB(attempt + 1, maxAttempts);
    }

    throw error;
  }
};

// Завантаження тестів
const loadTestsFromMongoDB = async () => {
  try {
    const tests = await db.collection('tests').find({}).toArray();
    testNames = {};
    tests.forEach(test => {
      testNames[test.testNumber] = {
        name: test.name,
        timeLimit: test.timeLimit,
        randomQuestions: test.randomQuestions,
        randomAnswers: test.randomAnswers,
        questionLimit: test.questionLimit,
        attemptLimit: test.attemptLimit,
        isQuickTest: test.isQuickTest || false,
        timePerQuestion: test.timePerQuestion || null
      };
    });
    logger.info(`Завантажено ${tests.length} тестів`);
  } catch (error) {
    logger.error('Помилка завантаження тестів', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Збереження тесту
const saveTestToMongoDB = async (testNumber, testData) => {
  try {
    await db.collection('tests').updateOne(
      { testNumber },
      { $set: { 
        testNumber,
        name: testData.name,
        timeLimit: testData.timeLimit,
        randomQuestions: testData.randomQuestions,
        randomAnswers: testData.randomAnswers,
        questionLimit: testData.questionLimit,
        attemptLimit: testData.attemptLimit,
        isQuickTest: testData.isQuickTest || false,
        timePerQuestion: testData.timePerQuestion || null
      }},
      { upsert: true }
    );
    logger.info('Тест збережено', { testNumber });
  } catch (error) {
    logger.error('Помилка збереження тесту', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Видалення тесту
const deleteTestFromMongoDB = async (testNumber) => {
  try {
    await db.collection('tests').deleteOne({ testNumber });
    logger.info('Тест видалено', { testNumber });
  } catch (error) {
    logger.error('Помилка видалення тесту', { testNumber, message: error.message, stack: error.stack });
    throw error;
  }
};

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Налаштування сесій
app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    client: client,
    dbName: 'alpha',
    collectionName: 'sessions',
    ttl: 24 * 60 * 60
  }).on('error', (error) => {
    logger.error('Помилка MongoStore', { message: error.message, stack: error.stack });
  }),
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
    maxAge: 24 * 60 * 60 * 1000
  }
}));

app.use((req, res, next) => {
  logger.info('Отримано cookie', { cookies: req.cookies, sessionID: req.sessionID || 'unknown' });
  next();
});

app.use((req, res, next) => {
  logger.info('Запит отримано', { url: req.url, method: req.method, userRole: req.userRole, timestamp: new Date().toISOString() });
  next();
});

// Ініціалізація res.locals
app.use((req, res, next) => {
  if (!res.locals) {
    res.locals = {};
    logger.info('Ініціалізовано res.locals', { url: req.url });
  }
  next();
});

// Генерація CSRF-токена (залишаємо для інших маршрутів)
app.use((req, res, next) => {
  if (!req.session) {
    logger.error('Сесія відсутня', { url: req.url });
    return res.status(500).json({ success: false, message: 'Помилка сесії' });
  }
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = require('crypto').randomBytes(32).toString('hex');
    logger.info('Згенеровано CSRF-секрет', { secret: req.session.csrfSecret });
  }
  const token = require('crypto').createHmac('sha256', req.session.csrfSecret).update(req.sessionID).digest('hex');
  res.locals._csrf = token;
  res.cookie('XSRF-TOKEN', token, { 
    httpOnly: false, 
    secure: process.env.NODE_ENV === 'production', 
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax' 
  });
  logger.info('Згенеровано CSRF-токен', { token, url: req.url });
  next();
});

// CSRF-валідація (залишаємо для маршрутів, крім імпорту)
app.use((req, res, next) => {
  if (['POST', 'PUT', 'DELETE'].includes(req.method) && 
      !req.url.startsWith('/admin/import-users') && 
      !req.url.startsWith('/admin/import-questions')) {
    const token = req.body._csrf || req.headers['x-csrf-token'] || req.headers['xsrf-token'];
    if (!token) {
      logger.error('Відсутній CSRF-токен', { method: req.method, url: req.url, body: req.body, headers: req.headers });
      return res.status(403).json({ success: false, message: 'CSRF-токен відсутній' });
    }
    if (!req.session || !req.session.csrfSecret) {
      logger.error('Відсутній CSRF-секрет', { sessionId: req.sessionID || 'unknown', url: req.url });
      return res.status(403).json({ success: false, message: 'Помилка сесії' });
    }
    const computedToken = require('crypto').createHmac('sha256', req.session.csrfSecret).update(req.sessionID).digest('hex');
    if (token !== computedToken) {
      logger.error('Недійсний CSRF-токен', { expected: computedToken, received: token, url: req.url });
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }
    logger.info('CSRF-токен валідовано', { token, url: req.url });
  }
  next();
});

// Запобігання кешуванню
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Обробка помилок MongoDB
app.use((err, req, res, next) => {
  if (err.name === 'MongoNetworkError' || err.name === 'MongoServerError') {
    logger.error('Помилка MongoDB', { message: err.message, stack: err.stack });
    res.status(503).json({ success: false, message: 'Помилка бази даних' });
  } else {
    next(err);
  }
});

// Водяний знак
app.use((req, res, next) => {
  const originalSend = res.send;
  res.send = function (body) {
    if (typeof body === 'string' && body.includes('</body>') && req.user) {
      const watermarkScript = `
        <style>
          .watermark {
            position: fixed;
            top: 10px;
            right: 10px;
            color: rgba(255, 0, 0, 0.3);
            font-size: 24px;
            pointer-events: none;
            z-index: 10000;
          }
        </style>
        <div class="watermark">Користувач: ${req.user}</div>
        <script>
          document.addEventListener('keydown', (e) => {
            if (
              e.key === 'PrintScreen' ||
              (e.ctrlKey && ['p', 'P', 's', 'S'].includes(e.key)) ||
              (e.metaKey && ['p', 'P', 's', 'S'].includes(e.key)) ||
              (e.altKey && e.key === 'PrintScreen') ||
              (e.metaKey && e.shiftKey && ['3', '4'].includes(e.key))
            ) {
              e.preventDefault();
            }
          });
          document.addEventListener('contextmenu', (e) => e.preventDefault());
          document.addEventListener('selectstart', (e) => e.preventDefault());
          document.addEventListener('copy', (e) => e.preventDefault());
          document.addEventListener('visibilitychange', () => {
            if (document.hidden) console.log('Вкладка невидима');
          });
        </script>
      `;
      body = body.replace('</body>', `${watermarkScript}</body>`);
    }
    return originalSend.call(this, body);
  };
  next();
});

// Імпорт користувачів
const importUsersToMongoDB = async (buffer) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    let sheet = workbook.getWorksheet('Users') || workbook.getWorksheet('Sheet1');
    if (!sheet) throw new Error('Лист "Users" або "Sheet1" не знайдено');
    const users = [];
    const saltRounds = 10;
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {
      const row = sheet.getRow(rowNumber);
      const username = String(row.getCell(1).value || '').trim();
      const password = String(row.getCell(2).value || '').trim();
      const role = String(row.getCell(3).value || '').trim().toLowerCase();
      if (username && password) {
        const hashedPassword = await bcrypt.hash(password, saltRounds);
        const userRole = role === 'admin' ? 'admin' : role === 'instructor' ? 'instructor' : 'user';
        users.push({ username, password: hashedPassword, role: userRole });
      }
    }
    if (!users.length) throw new Error('Не знайдено користувачів');
    await db.collection('users').deleteMany({});
    await db.collection('users').insertMany(users);
    logger.info(`Імпортовано ${users.length} користувачів`);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    return users.length;
  } catch (error) {
    logger.error('Помилка імпорту користувачів', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Імпорт питань — покращена версія з діагностикою
const importQuestionsToMongoDB = async (buffer, testNumber) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // === ДІАГНОСТИКА ЛИСТІВ ===
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    logger.info(`Імпорт питань: доступні листи: ${sheetNames.join(', ')}`);

    let sheet = workbook.getWorksheet('Questions') || 
                workbook.getWorksheet('questions') || 
                workbook.getWorksheet('Питання') || 
                workbook.getWorksheet('Sheet1') || 
                workbook.worksheets[0];

    if (!sheet) {
      throw new Error(`Лист з питаннями не знайдено. Доступні: ${sheetNames.join(', ')}`);
    }

    logger.info(`Використовуємо лист "${sheet.name}", рядків: ${sheet.rowCount}`);

    const MAX_ROWS = 1000;
    if (sheet.rowCount > MAX_ROWS + 1) {
      throw new Error(`Занадто багато рядків (${sheet.rowCount - 1}). Максимум: ${MAX_ROWS}`);
    }

    const questions = [];
    let errorRows = [];

    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= 1) return; // пропускаємо заголовок

      try {
        const rowValues = row.values.slice(1);
        let questionText = rowValues[1];
        if (typeof questionText === 'object' && questionText) {
          questionText = questionText.text || questionText.value || '[Невірний текст]';
        }
        questionText = String(questionText || '').trim();

        if (!questionText) throw new Error('Текст питання відсутній');

        const picture = String(rowValues[0] || '').trim();
        let options = rowValues.slice(2, 14).filter(Boolean).map(val => String(val).trim());
        const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(val => String(val).trim());
        const type = String(rowValues[26] || 'multiple').toLowerCase().trim();
        const points = Number(rowValues[27]) || 1;
        const variant = String(rowValues[28] || '').trim();

        if (type === 'truefalse') options = ["Правда", "Неправда"];

        const normalizedPicture = picture ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '') : null;

        let questionData = {
          testNumber,
          picture: null,
          originalPicture: normalizedPicture,
          text: questionText,
          options,
          correctAnswers,
          type,
          points,
          variant,
          order: rowNumber - 1
        };

        // Обробка зображення
        if (normalizedPicture) {
          // ... (залишаємо твій код обробки картинок без змін)
          const pictureMatch = normalizedPicture.match(/^Picture(\d+)$/i);
          if (pictureMatch) {
            const pictureNumber = parseInt(pictureMatch[1], 10);
            const targetFileNameBase = `Picture${pictureNumber}`;
            const extensions = ['.png', '.jpg', '.jpeg', '.gif'];
            const imageDir = path.join(__dirname, 'public', 'images');
            const filesInDir = fs.existsSync(imageDir) ? fs.readdirSync(imageDir) : [];
            for (const ext of extensions) {
              const expectedFileName = `${targetFileNameBase}${ext}`;
              const matchedFile = filesInDir.find(file => file.toLowerCase() === expectedFileName.toLowerCase());
              if (matchedFile) {
                questionData.picture = `/images/${matchedFile}`;
                break;
              }
            }
          }
        }

        // Matching
        if (type === 'matching') {
          questionData.pairs = options.map((opt, idx) => ({
            left: opt || '',
            right: questionData.correctAnswers[idx] || ''
          })).filter(pair => pair.left && pair.right);
          if (!questionData.pairs.length) throw new Error('Для Matching потрібні пари');
          questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
        }

        // Fillblank — тільки перевірка структури
        if (type === 'fillblank') {
          questionData.text = questionText.replace(/\s*___/g, '___');
          const blankCount = (questionData.text.match(/___/g) || []).length;
          if (blankCount === 0 || blankCount !== correctAnswers.length) {
            throw new Error(`Пропусків (${blankCount}) не відповідає відповідям (${correctAnswers.length})`);
          }
          questionData.blankCount = blankCount;
        }

        // Singlechoice
        if (type === 'singlechoice') {
          if (correctAnswers.length !== 1 || options.length < 2) {
            throw new Error('Single Choice: потрібна 1 відповідь і ≥2 варіанти');
          }
          questionData.correctAnswer = correctAnswers[0];
        }

        // Input
        if (type === 'input') {
          if (correctAnswers.length !== 1) throw new Error('Input: потрібна 1 відповідь');
          const correctAnswer = correctAnswers[0];
          if (correctAnswer.includes('-')) {
            const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
            if (isNaN(min) || isNaN(max) || min > max) throw new Error('Невірний діапазон');
          }
        }

        questions.push(questionData);

      } catch (rowError) {
        errorRows.push({ row: rowNumber, error: rowError.message });
        logger.warn(`Помилка в рядку ${rowNumber}: ${rowError.message}`);
      }
    });

    if (errorRows.length > 0) {
      logger.error(`Імпорт завершено з помилками в ${errorRows.length} рядках`, errorRows);
    }

    if (!questions.length) {
      throw new Error('Не знайдено жодного валідного питання');
    }

    await db.collection('questions').deleteMany({ testNumber });
    await db.collection('questions').insertMany(questions);

    await CacheManager.invalidateCache('questions', testNumber);

    logger.info(`Успішно імпортовано ${questions.length} питань для тесту ${testNumber}`);
    return questions.length;

  } catch (error) {
    logger.error('Критична помилка імпорту питань', { 
      message: error.message, 
      stack: error.stack,
      testNumber 
    });
    throw error;
  }
};

// Перемішування масиву
const shuffleArray = (array) => {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
};

// Завантаження користувачів
const loadUsersToCache = async () => {
  try {
    const startTime = Date.now();
    userCache = await db.collection('users').find({}).toArray();
    logger.info(`Оновлено кеш користувачів: ${userCache.length} за ${Date.now() - startTime} мс`);
  } catch (error) {
    logger.error('Помилка кешу користувачів', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Завантаження питань
const loadQuestions = async (testNumber) => {
  try {
    const startTime = Date.now();
    if (questionsCache[testNumber]) {
      logger.info(`Завантажено ${questionsCache[testNumber].length} питань з кешу`);
      return questionsCache[testNumber];
    }
    const questions = await db.collection('questions').find({ testNumber: testNumber.toString() }).sort({ order: 1 }).toArray();
    if (!questions.length) throw new Error(`Не знайдено питань для тесту ${testNumber}`);
    questionsCache[testNumber] = questions;
    logger.info(`Завантажено ${questions.length} питань за ${Date.now() - startTime} мс`);
    return questions;
  } catch (error) {
    logger.error(`Помилка loadQuestions`, { message: error.message, stack: error.stack });
    throw error;
  }
};

// Перевірка ініціалізації
const ensureInitialized = (req, res, next) => {
  if (!isInitialized) {
    if (initializationError) {
      logger.error('Сервер не ініціалізовано', { message: initializationError.message, stack: initializationError.stack });
      return res.status(500).json({ success: false, message: `Помилка ініціалізації: ${initializationError.message}` });
    }
    logger.warn('Сервер ініціалізується');
    return res.status(503).json({ success: false, message: 'Сервер ініціалізується' });
  }
  next();
};

// Оновлення паролів
const updateUserPasswords = async () => {
  const startTime = Date.now();
  const users = await db.collection('users').find({}).toArray();
  const saltRounds = 10;
  for (const user of users) {
    if (!user.password.startsWith('$2b$')) {
      const hashedPassword = await bcrypt.hash(user.password, saltRounds);
      await db.collection('users').updateOne(
        { _id: user._id },
        { $set: { password: hashedPassword } }
      );
    }
  }
  logger.info(`Оновлено паролі за ${Date.now() - startTime} мс`);
  await CacheManager.invalidateCache('users', null);
};

// Ініціалізація сервера
const initializeServer = async () => {
  try {
    await connectToMongoDB();
    await db.collection('users').createIndex({ username: 1 }, { unique: true });
    await db.collection('questions').createIndex({ testNumber: 1, variant: 1 });
    await db.collection('test_results').createIndex({ user: 1, testNumber: 1, endTime: -1 });
    await db.collection('activity_log').createIndex({ user: 1, timestamp: -1 });
    await db.collection('test_attempts').createIndex({ user: 1, testNumber: 1, attemptDate: 1 });
    await db.collection('login_attempts').createIndex({ ipAddress: 1, lastAttempt: 1 });
    await db.collection('tests').createIndex({ testNumber: 1 }, { unique: true });
    await db.collection('active_tests').createIndex({ user: 1 }, { unique: true });
    logger.info('Індекси створено');

    const userCount = await db.collection('users').countDocuments();
    if (userCount > 0) {
      await db.collection('users').updateMany(
        { role: { $exists: false }, username: "admin" },
        { $set: { role: "admin" } }
      );
      await db.collection('users').updateMany(
        { role: { $exists: false }, username: "Instructor" },
        { $set: { role: "instructor" } }
      );
      await db.collection('users').updateMany(
        { role: { $exists: false }, username: { $nin: ["admin", "Instructor"] } },
        { $set: { role: "user" } }
      );
      logger.info('Міграція ролей завершена');
    }

    const testCount = await db.collection('tests').countDocuments();
    if (!testCount) {
      const defaultTests = {
        "1": { name: "Тест 1", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "2": { name: "Тест 2", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "3": { name: "Тест 3", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 }
      };
      for (const [testNumber, testData] of Object.entries(defaultTests)) {
        await saveTestToMongoDB(testNumber, testData);
      }
      logger.info('Міграція тестів завершена', { count: Object.keys(defaultTests).length });
    }

    await updateUserPasswords();
    await loadUsersToCache();
    await loadTestsFromMongoDB();
    await CacheManager.invalidateCache('questions', null);
    isInitialized = true;
    initializationError = null;
  } catch (error) {
    logger.error('Помилка ініціалізації', { message: error.message, stack: error.stack });
    initializationError = error;
    throw error;
  }
};

// Очищення активності
const cleanupActivityLog = async () => {
  try {
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const result = await db.collection('activity_log').deleteMany({
      timestamp: { $lt: thirtyDaysAgo.toISOString() }
    });
    logger.info('Очищено активність', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення активності', { message: error.message, stack: error.stack });
  }
};

// Очищення тестів
const cleanupActiveTests = async () => {
  try {
    const twentyFourHoursAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const result = await db.collection('active_tests').deleteMany({
      startTime: { $lt: twentyFourHoursAgo.getTime() }
    });
    logger.info('Очищено тести', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення тестів', { message: error.message, stack: error.stack });
  }
};

// Періодичне очищення
setInterval(cleanupActivityLog, 24 * 60 * 60 * 1000);
setInterval(cleanupActiveTests, 24 * 60 * 60 * 1000);

// Ініціалізація
(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
    await cleanupActivityLog();
    await cleanupActiveTests();
  } catch (error) {
    logger.error('Помилка запуску', { message: error.message, stack: error.stack });
    process.exit(1);
  }
})();

// Обмеження спроб входу
const MAX_LOGIN_ATTEMPTS = 30;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;

const checkLoginAttempts = async (ipAddress, reset = false) => {
  const now = Date.now();
  const startOfDay = new Date(now).setHours(0, 0, 0, 0);
  const endOfDay = startOfDay + ONE_DAY_MS;

  const attempts = await db.collection('login_attempts').findOne({
    ipAddress,
    lastAttempt: { $gte: startOfDay, $lt: endOfDay }
  });

  if (reset) {
    await db.collection('login_attempts').updateOne(
      { ipAddress, lastAttempt: { $gte: startOfDay, $lt: endOfDay } },
      { $set: { count: 0, lastAttempt: now } },
      { upsert: true }
    );
    return;
  }

  if (!attempts) {
    await db.collection('login_attempts').insertOne({
      ipAddress,
      count: 0,
      lastAttempt: now
    });
  } else if (attempts.count >= MAX_LOGIN_ATTEMPTS) {
    throw new Error('Перевищено ліміт спроб входу');
  }

  await db.collection('login_attempts').updateOne(
    { ipAddress, lastAttempt: { $gte: startOfDay, $lt: endOfDay } },
    { $inc: { count: 1 }, $set: { lastAttempt: now } },
    { upsert: true }
  );
};

// Логування активності
const logActivity = async (user, action, ipAddress, additionalInfo = {}, session = null) => {
  try {
    const timestamp = new Date();
    await db.collection('activity_log').insertOne({
      user,
      action,
      ipAddress,
      timestamp: timestamp.toISOString(),
      additionalInfo
    }, { session });
    logger.info(`Активність: ${user} - ${action} о ${timestamp}`);
  } catch (error) {
    logger.error('Помилка логування', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Тест MongoDB
app.get('/test-mongo', async (req, res) => {
  try {
    if (!db) throw new Error('MongoDB не підключено');
    await db.collection('users').findOne();
    res.json({ success: true, message: 'MongoDB підключено' });
  } catch (error) {
    logger.error('Тест MongoDB провалився', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка MongoDB' });
  }
});

// Тест API
app.get('/api/test', (req, res) => {
  logger.info('Запит /api/test');
  res.json({ success: true, message: 'API працює' });
});

// Favicon
app.get('/favicon.ico', (req, res) => {
  res.status(204).end();
});

// Головна сторінка з формою авторизації
app.get('/', (req, res) => {
  logger.info('Відображення index.html');
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Авторизація</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f5f5f5; margin: 0; }
          h1 { font-size: 36px; margin-bottom: 20px; }
          label { display: block; font-size: 18px; margin-bottom: 5px; }
          input[type="text"], input[type="password"] { padding: 10px; font-size: 18px; width: 200px; margin-bottom: 10px; }
          button { padding: 10px 20px; font-size: 18px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
          button:hover { background-color: #45a049; }
          button:disabled { background-color: #cccccc; cursor: not-allowed; }
          .error { color: red; margin-top: 10px; }
          .checkbox-container { margin-bottom: 10px; }
          .checkbox-container input[type="checkbox"] { vertical-align: middle; margin: 0 5px 0 0; }
          .checkbox-container label { display: inline; font-size: 16px; margin: 0; vertical-align: middle; }
          @media (max-width: 600px) {
            h1 { font-size: 28px; }
            label { font-size: 16px; }
            input[type="text"], input[type="password"], button { font-size: 16px; width: 90%; padding: 15px; }
          }
        </style>
      </head>
      <body>
        <h1>Авторизація</h1>
        <form id="login-form" method="POST" action="/login">
          <input type="hidden" name="_csrf" value="${res.locals._csrf}">
          <label for="username">Користувач:</label>
          <input type="text" id="username" name="username" placeholder="Логін" required><br>
          <label for="password">Пароль:</label>
          <input type="password" id="password" name="password" placeholder="Пароль" required><br>
          <div class="checkbox-container">
            <input type="checkbox" id="show-password" onclick="togglePassword()">
            <label for="show-password">Показати пароль</label>
          </div>
          <button type="submit" id="login-button">Увійти</button>
        </form>
        <div id="error-message" class="error"></div>
        <script>
          document.getElementById('login-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const errorMessage = document.getElementById('error-message');
            const loginButton = document.getElementById('login-button');

            loginButton.disabled = true;
            loginButton.textContent = 'Завантаження...';

            const formData = new URLSearchParams();
            formData.append('username', username);
            formData.append('password', password);
            const csrfToken = document.querySelector('input[name="_csrf"]').value;
            console.log('Відправка CSRF-токена:', csrfToken);
            formData.append('_csrf', csrfToken);

            try {
              const response = await fetch('/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: formData
              });

              const result = await response.json();
              console.log('Відповідь на вхід:', result);

              if (result.success) {
                window.location.href = result.redirect + '?nocache=' + Date.now();
              } else {
                if (response.status === 429) {
                  errorMessage.textContent = result.message || 'Перевищено ліміт спроб входу. Спробуйте знову завтра.';
                } else if (response.status === 400) {
                  errorMessage.textContent = result.message || 'Некоректні дані. Перевірте логін та пароль.';
                } else if (response.status === 401) {
                  errorMessage.textContent = result.message || 'Невірний логін або пароль.';
                } else {
                  errorMessage.textContent = result.message || 'Помилка входу.';
                }
              }
            } catch (error) {
              console.error('Помилка під час входу:', error);
              errorMessage.textContent = 'Не вдалося підключитися до сервера. Перевірте ваше з’єднання з Інтернетом.';
            } finally {
              loginButton.disabled = false;
              loginButton.textContent = 'Увійти';
            }
          });

          function togglePassword() {
            const passwordField = document.getElementById('password');
            const showPasswordCheckbox = document.getElementById('show-password');
            passwordField.type = showPasswordCheckbox.checked ? 'text' : 'password';
          }
        </script>
      </body>
    </html>
  `);
});

// Обробка входу користувача
app.post('/login', [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Логін має бути від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Логін може містити лише літери та цифри'),
  body('password')
    .isLength({ min: 6, max: 100 }).withMessage('Пароль має бути від 6 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await checkLoginAttempts(ipAddress);

    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, message: errors.array()[0].msg });
    }

    const { username, password } = req.body;
    logger.info('Отримано дані для входу', { username });

    if (!username || !password) {
      logger.warn('Логін або пароль не вказано');
      return res.status(400).json({ success: false, message: 'Логін або пароль не вказано' });
    }

    if (userCache.length === 0) {
      logger.warn('Кеш користувачів порожній, повторне завантаження з MongoDB');
      await loadUsersToCache();
      if (userCache.length === 0) {
        logger.error('Користувачів не знайдено в MongoDB після повторного завантаження');
        throw new Error('Не вдалося завантажити користувачів з бази даних');
      }
    }

    const foundUser = userCache.find(user => user.username === username);
    logger.info('Користувача знайдено в кеші', { username, cachedPassword: foundUser?.password });

    if (!foundUser) {
      logger.warn('Користувача не знайдено', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    const passwordMatch = await bcrypt.compare(password, foundUser.password);
    if (!passwordMatch) {
      logger.warn('Невірний пароль для користувача', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    await checkLoginAttempts(ipAddress, true);

    const token = jwt.sign(
      { username: foundUser.username, role: foundUser.role },
      process.env.JWT_SECRET,
      { expiresIn: '24h' }
    );

    // Встановлення httpOnly cookie для серверної авторизації
    res.cookie('token', token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    // Встановлення не-httpOnly cookie для клієнтського доступу
    res.cookie('auth_token', token, {
      httpOnly: false,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    await logActivity(foundUser.username, 'увійшов на сайт', ipAddress);

    if (foundUser.role === 'admin') {
      res.json({ success: true, redirect: '/admin' });
    } else {
      res.json({ success: true, redirect: '/select-test' });
    }
  } catch (error) {
    logger.error('Помилка в /login', { message: error.message, stack: error.stack });
    res.status(error.message.includes('Перевищено ліміт') ? 429 : 500).json({ success: false, message: error.message || 'Помилка сервера' });
  } finally {
    logger.info('Маршрут /login виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Middleware для перевірки авторизації через JWT
const checkAuth = (req, res, next) => {
  const token = req.headers['authorization']?.split(' ')[1] || req.cookies.token;
  if (!token) {
    return res.redirect('/');
  }

  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    req.user = decoded.username;
    req.userRole = decoded.role;
    next();
  } catch (error) {
    logger.error('Помилка перевірки JWT', { message: error.message, stack: error.stack });
    res.redirect('/');
  }
};

// Middleware для перевірки ролі адміністратора
const checkAdmin = (req, res, next) => {
  if (req.userRole !== 'admin') {
    return res.status(403).send('Доступно тільки для адміністратора (403 Forbidden)');
  }
  next();
};

// Сторінка вибору тесту
app.get('/select-test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    // Перевірка кешу тестів
    if (Object.keys(testNames).length === 0) {
      logger.warn('testNames порожній, повторне завантаження з MongoDB');
      await loadTestsFromMongoDB();
      if (Object.keys(testNames).length === 0) {
        logger.error('Тести не знайдено в MongoDB після повторного завантаження');
        throw new Error('Не вдалося завантажити тести з бази даних');
      }
    }
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Вибір тесту</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              text-align: center; 
              padding: 20px; 
              padding-bottom: 120px; 
              margin: 0; 
              background-color: #f5f5f5;
            }
            h1 { 
              font-size: 24px; 
              margin-bottom: 20px; 
              color: #333;
            }
            .test-buttons {
              display: grid;
              grid-template-columns: 1fr;
              gap: 12px;
              max-width: 400px;
              margin: 0 auto;
            }

            @media (min-width: 900px) {
              .test-buttons {
                grid-template-columns: repeat(3, 1fr);
                max-width: 1000px;
              }
            }

            button, .instructions-btn, .feedback-btn, .results-btn {
              padding: 16px 24px;
              font-size: 18px;
              font-weight: bold;
              cursor: pointer;
              border: none;
              border-radius: 8px;
              text-align: center;
              text-decoration: none;
              min-height: 70px;          /* однакова висота по найбільшій */
              display: flex;
              align-items: center;
              justify-content: center;
              box-sizing: border-box;
            }

            .test-btn {
              background-color: #4CAF50;
              color: white;
            }

            .test-btn:hover { background-color: #45a049; }

            .instructions-btn {
              background-color: #ffeb3b;
              color: #333;
            }

            .instructions-btn:hover { background-color: #ffd700; }

            .feedback-btn, .results-btn {
              background-color: #ffeb3b;
              color: #333;
            }

            .feedback-btn:hover, .results-btn:hover { background-color: #ffd700; }

            @media (max-width: 600px) {
              button, .instructions-btn, .feedback-btn, .results-btn {
                font-size: 16px;
                padding: 14px 20px;
                min-height: 60px;
              }
            }

            /* Стилі для кнопки Вийти — червона, фіксована внизу по центру */
            #logout {
              position: fixed;
              bottom: 20px;
              left: 50%;
              transform: translateX(-50%);
              background: #ef5350;           /* червоний колір */
              color: white;
              padding: 16px 32px;
              font-size: 18px;
              font-weight: bold;
              border: none;
              border-radius: 10px;
              cursor: pointer;
              min-height: 60px;
              min-width: 180px;
              box-shadow: 0 4px 10px rgba(0,0,0,0.15);
              z-index: 100;
            }

            #logout:hover {
              background: #d32f2f;
              transform: translateX(-50%) translateY(-2px);
            }
          </style>
        </head>
        <body>
          <h1>Виберіть тест</h1>
          <div class="test-buttons">
            <a href="/instructions" class="instructions-btn">Інструкція до тестів</a>
            ${Object.entries(testNames).length > 0
              ? Object.entries(testNames).map(([num, data]) => `
                  <button class="test-btn" onclick="window.location.href='/test?test=${num}'">${data.name.replace(/"/g, '\\"')}</button>
                `).join('')
              : '<p class="no-tests">Немає доступних тестів</p>'
            }
            ${req.userRole === 'instructor' ? `
              <button class="results-btn" onclick="window.location.href='/admin/results'">Переглянути результати</button>
            ` : ''}
            <a href="/feedback" class="feedback-btn">Зворотний зв’язок</a>
          </div>

          <!-- Кнопка Вийти — фіксована внизу по центру -->
          <button id="logout" onclick="logout()">Вийти</button>

          <script>
            async function logout() {
              console.log('Спроба вийти, CSRF-токен:', '${res.locals._csrf}');
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                console.log('Статус відповіді на вихід:', response.status);
                if (!response.ok) {
                  throw new Error('HTTP error! status: ' + response.status);
                }
                const result = await response.json();
                console.log('Відповідь на вихід:', result);
                if (result.success) {
                  window.location.href = '/';
                } else {
                  throw new Error('Вихід не вдався: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка під час виходу:', error);
                alert('Не вдалося вийти. Перевірте консоль браузера для деталей.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /select-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки вибору тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /select-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка виходу користувача
app.post('/logout', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Отримано запит на вихід', { user: req.user });

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await logActivity(req.user, 'покинув сайт', ipAddress);

    // Очищаємо cookie
    res.clearCookie('token');
    res.clearCookie('auth_token');

    // Знищуємо сесію
    if (req.session) {
      req.session.destroy((err) => {
        if (err) {
          logger.error('Помилка знищення сесії', { error: err.message });
        }
      });
    }

    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка при виході', { message: error.message, stack: error.stack });
    // Навіть при помилці очищаємо cookie і повертаємо успіх клієнту
    res.clearCookie('token');
    res.clearCookie('auth_token');
    res.json({ success: true, message: 'Вихід виконано з помилкою на сервері, але сесія очищена' });
  } finally {
    logger.info('Маршрут /logout виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Збереження результатів тесту (оновлено: перераховуємо все перед збереженням)
const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress, testSessionId) => {
  const startTimeLog = Date.now();

  logger.info('[SAVE-RESULT] Початок збереження (спрощена версія без транзакції)', {
    user,
    testNumber,
    testSessionId,
    score,
    percentage: percentage.toFixed(1) + '%',
    totalQuestions,
    answersCount: Object.keys(answers).length,
    variant: variant || '(немає)'
  });

  try {
    // Завантажуємо питання
    let questions = await db.collection('questions')
      .find({ testNumber })
      .sort({ order: 1 })
      .toArray();

    logger.info('[SAVE-RESULT] Знайдено питань у базі', { count: questions.length });

    questions = questions.filter(q =>
      !q.variant || q.variant === '' || q.variant === variant
    );

    logger.info('[SAVE-RESULT] Після фільтрації за варіантом', { count: questions.length });

    if (questions.length === 0) {
      logger.warn('[SAVE-RESULT] Немає питань після фільтрації — пропускаємо перерахунок');
    }

    // Перерахунок (якщо питань немає — використовуємо передані значення)
    let actualScore = score;
    let actualTotalPoints = totalPoints;
    let actualPercentage = percentage;
    let actualTotalQuestions = totalQuestions;

    if (questions.length > 0) {
      const actualScoresPerQuestion = questions.map((q, index) => {
        const userAnswer = answers[index];
        return calculateQuestionScore(q, userAnswer);
      });

      actualScore = actualScoresPerQuestion.reduce((sum, s) => sum + s, 0);
      actualTotalPoints = questions.reduce((sum, q) => sum + (q.points || 0), 0);
      actualPercentage = actualTotalPoints > 0 ? (actualScore / actualTotalPoints) * 100 : 0;
      actualTotalQuestions = questions.length;

      logger.info('[SAVE-RESULT] Перераховано значення', {
        actualScore: actualScore.toFixed(2),
        actualTotalPoints: actualTotalPoints.toFixed(2),
        actualPercentage: actualPercentage.toFixed(1) + '%',
        actualTotalQuestions
      });
    } else {
      logger.warn('[SAVE-RESULT] Використовуємо передані значення (питань не знайдено)');
    }

    const duration = Math.round((endTime - startTime) / 1000);

    const result = {
      user,
      testNumber,
      score: actualScore,
      totalPoints: actualTotalPoints,
      totalClicks,
      correctClicks,
      totalQuestions: actualTotalQuestions,
      percentage: actualPercentage,
      startTime: new Date(startTime).toISOString(),
      endTime: new Date(endTime).toISOString(),
      duration,
      answers: Object.fromEntries(Object.entries(answers).sort((a, b) => parseInt(a[0]) - parseInt(b[0]))),
      suspiciousActivity,
      variant: variant ? `Variant ${variant}` : 'Немає',
      testSessionId,
      createdAt: new Date(),
      ipAddress
    };

    logger.info('[SAVE-RESULT] Готовий документ для вставки', { testSessionId, size: JSON.stringify(result).length });

    const insertResult = await db.collection('test_results').insertOne(result);

    logger.info('[SAVE-RESULT] Успішно вставлено документ', {
      insertedId: insertResult.insertedId.toString(),
      testSessionId
    });

    await logActivity(
      user,
      `завершив тест ${testNames[testNumber]?.name || 'Тест'} з результатом ${Math.round(actualPercentage)}%`,
      ipAddress,
      { percentage: Math.round(actualPercentage), testSessionId }
    );

  } catch (error) {
    logger.error('[SAVE-RESULT] Критична помилка при збереженні', {
      message: error.message,
      stack: error.stack,
      testSessionId
    });
    throw error; // не ковтаємо помилку — нехай маршрут /result її побачить
  } finally {
    const duration = Date.now() - startTimeLog;
    logger.info('[SAVE-RESULT] Завершено', { duration: `${duration} мс` });
  }
};

// Перевірка кількості спроб проходження тесту
const checkTestAttempts = async (user, testNumber) => {
  try {
    const now = new Date();
    const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const endOfDay = new Date(startOfDay.getTime() + 24 * 60 * 60 * 1000);

    const attemptLimit = testNames[testNumber]?.attemptLimit || 1;

    const attempts = await db.collection('test_attempts').countDocuments({
      user,
      testNumber,
      attemptDate: {
        $gte: startOfDay.toISOString(),
        $lt: endOfDay.toISOString()
      }
    });

    logger.info(`Користувач ${user} має ${attemptLimit - attempts} спроб для тесту ${testNumber} сьогодні`);

    if (attempts >= attemptLimit) {
      return false;
    }

    await db.collection('test_attempts').insertOne({
      user,
      testNumber,
      attemptDate: now.toISOString()
    });
    return true;
  } catch (error) {
    logger.error('Помилка перевірки спроб тесту', { message: error.message, stack: error.stack });
    throw error;
  }
};

app.get('/feedback', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Зворотний зв’язок</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 20px;
              background-color: #f5f5f5;
              text-align: center;
            }
            .container {
              max-width: 600px;
              margin: 0 auto;
              background-color: white;
              padding: 20px;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            h1 {
              font-size: 24px;
              margin-bottom: 20px;
              color: #333;
            }
            label {
              display: block;
              font-size: 16px;
              margin-bottom: 5px;
              text-align: left;
            }
            textarea {
              width: 100%;
              height: 150px;
              padding: 10px;
              font-size: 16px;
              border: 1px solid #ccc;
              border-radius: 5px;
              margin-bottom: 10px;
              box-sizing: border-box;
            }
            button {
              padding: 10px 20px;
              font-size: 16px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #4CAF50;
              color: white;
            }
            button:hover {
              background-color: #45a049;
            }
            button:disabled {
              background-color: #cccccc;
              cursor: not-allowed;
            }
            .error {
              color: red;
              margin-top: 10px;
              font-size: 14px;
            }
            .back-btn {
              background-color: #007bff;
              margin-top: 10px;
            }
            .back-btn:hover {
              background-color: #0056b3;
            }
            @media (max-width: 600px) {
              .container {
                padding: 15px;
              }
              h1 {
                font-size: 20px;
              }
              textarea {
                font-size: 14px;
              }
              button {
                width: 100%;
                font-size: 14px;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Зворотний зв’язок</h1>
            <form id="feedback-form" method="POST" action="/feedback">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <label for="message">Ваше повідомлення:</label>
              <textarea id="message" name="message" placeholder="Введіть ваше повідомлення, пропозицію або повідомте про проблему" required></textarea>
              <button type="submit" id="submit-btn">Надіслати</button>
            </form>
            <div id="error-message" class="error"></div>
            <button class="back-btn" onclick="window.location.href='/select-test'">Назад до вибору тесту</button>
          </div>
          <script>
            document.getElementById('feedback-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const message = document.getElementById('message').value;
              const errorMessage = document.getElementById('error-message');
              const submitBtn = document.getElementById('submit-btn');

              submitBtn.disabled = true;
              submitBtn.textContent = 'Надсилання...';

              const formData = new URLSearchParams();
              formData.append('message', message);
              formData.append('_csrf', document.querySelector('input[name="_csrf"]').value);

              try {
                const response = await fetch('/feedback', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                const result = await response.json();
                if (result.success) {
                  errorMessage.style.color = 'green';
                  errorMessage.textContent = 'Повідомлення успішно надіслано!';
                  document.getElementById('message').value = '';
                } else {
                  errorMessage.textContent = result.message || 'Помилка при надсиланні повідомлення.';
                }
              } catch (error) {
                console.error('Помилка надсилання зворотного зв’язку:', error);
                errorMessage.textContent = 'Не вдалося підключитися до сервера. Перевірте ваше з’єднання з Інтернетом.';
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Надіслати';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /feedback', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми зворотного зв’язку');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /feedback виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/feedback', checkAuth, [
  body('message')
    .isLength({ min: 5, max: 1000 }).withMessage('Повідомлення має бути від 5 до 1000 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, message: errors.array()[0].msg });
    }

    const { message } = req.body;
    const user = req.user;
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    const timestamp = new Date().toISOString();

    // Збереження повідомлення в MongoDB
    await db.collection('feedback').insertOne({
      user,
      message,
      timestamp,
      ipAddress,
      read: false
    });

    logger.info('Зворотний зв’язок збережено', { user, message });

    // Надсилання email адміністратору
    try {
      const mailOptions = {
        from: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
        to: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
        subject: 'Нове повідомлення зворотного зв’язку',
        text: `
          Користувач: ${user}
          Повідомлення: ${message}
          Час: ${new Date(timestamp).toLocaleString('uk-UA')}
          IP-адреса: ${ipAddress}
        `
      };
      await transporter.sendMail(mailOptions);
      logger.info('Email зворотного зв’язку надіслано', { user, email: process.env.EMAIL_USER });
    } catch (emailError) {
      logger.error('Помилка відправки email зворотного зв’язку', { 
        message: emailError.message, 
        stack: emailError.stack, 
        emailUser: process.env.EMAIL_USER 
      });
    }

    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка в /feedback (POST)', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при надсиланні зворотного зв’язку' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /feedback (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/feedback/delete/:id', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const feedbackId = req.params.id;
    if (!ObjectId.isValid(feedbackId)) {
      return res.status(400).json({ success: false, message: 'Невірний ID повідомлення' });
    }

    const result = await db.collection('feedback').deleteOne({ _id: new ObjectId(feedbackId) });
    if (result.deletedCount === 0) {
      return res.status(404).json({ success: false, message: 'Повідомлення не знайдено' });
    }

    logger.info('Видалено повідомлення зворотного зв’язку', { feedbackId, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення повідомлення', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні повідомлення' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/feedback/delete/:id виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/feedback/delete-all', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const result = await db.collection('feedback').deleteMany({});
    logger.info('Видалено всі повідомлення зворотного зв’язку', { deletedCount: result.deletedCount, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення всіх повідомлень', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні всіх повідомлень' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/feedback/delete-all виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.get('/admin/feedback', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;

    const feedback = await db.collection('feedback')
      .find({})
      .sort({ timestamp: -1 })
      .skip(skip)
      .limit(limit)
      .toArray();

    const totalFeedback = await db.collection('feedback').countDocuments();
    const totalPages = Math.ceil(totalFeedback / limit);

    // Позначити всі повідомлення як прочитані
    await db.collection('feedback').updateMany({ read: false }, { $set: { read: true } });

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Зворотний зв’язок</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
            }
            .container {
              max-width: 900px;
              margin: 0 auto;
              background-color: white;
              padding: 20px;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            h1 {
              font-size: 24px;
              text-align: center;
              margin-bottom: 20px;
            }
            table {
              border-collapse: collapse;
              width: 100%;
              margin-top: 20px;
            }
            th, td {
              border: 1px solid #ddd;
              padding: 8px;
              text-align: left;
            }
            th {
              background-color: #f2f2f2;
            }
            .message {
              white-space: pre-wrap;
              max-width: 400px;
              word-wrap: break-word;
            }
            .nav-btn, .delete-btn, .delete-all-btn {
              padding: 8px 16px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              font-size: 14px;
              margin: 5px;
            }
            .nav-btn {
              background-color: #007bff;
              color: white;
            }
            .nav-btn:hover {
              background-color: #0056b3;
            }
            .delete-btn {
              background-color: #ef5350;
              color: white;
            }
            .delete-btn:hover {
              background-color: #d32f2f;
            }
            .delete-all-btn {
              background-color: #d32f2f;
              color: white;
            }
            .delete-all-btn:hover {
              background-color: #b71c1c;
            }
            .pagination {
              margin-top: 20px;
              text-align: center;
            }
            .pagination a {
              margin: 0 5px;
              padding: 5px 10px;
              background-color: #007bff;
              color: white;
              text-decoration: none;
              border-radius: 5px;
            }
            .pagination a:hover {
              background-color: #0056b3;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 20px;
              }
              table {
                font-size: 14px;
              }
              .message {
                max-width: 200px;
              }
              .nav-btn, .delete-btn, .delete-all-btn {
                width: 100%;
                box-sizing: border-box;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Зворотний зв’язок від користувачів</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="delete-all-btn" onclick="deleteAllFeedback()">Видалити всі повідомлення</button>
            <table>
              <tr>
                <th>Користувач</th>
                <th>Повідомлення</th>
                <th>Час</th>
                <th>IP-адреса</th>
                <th>Дії</th>
              </tr>
              ${feedback.length > 0 ? feedback.map(f => `
                <tr>
                  <td>${f.user}</td>
                  <td class="message">${f.message.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>
                  <td>${new Date(f.timestamp).toLocaleString('uk-UA')}</td>
                  <td>${f.ipAddress}</td>
                  <td>
                    <button class="delete-btn" onclick="deleteFeedback('${f._id}')">Видалити</button>
                  </td>
                </tr>
              `).join('') : '<tr><td colspan="5">Немає повідомлень</td></tr>'}
            </table>
            <div class="pagination">
              ${page > 1 ? `<a href="/admin/feedback?page=${page - 1}">Попередня</a>` : ''}
              <span>Сторінка ${page} з ${totalPages}</span>
              ${page < totalPages ? `<a href="/admin/feedback?page=${page + 1}">Наступна</a>` : ''}
            </div>
          </div>
          <script>
            async function deleteFeedback(id) {
              if (!confirm('Ви впевнені, що хочете видалити це повідомлення?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/feedback/delete/' + id, {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  window.location.reload();
                } else {
                  alert('Помилка видалення: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка видалення:', error);
                alert('Не вдалося видалити повідомлення.');
              }
            }

            async function deleteAllFeedback() {
              if (!confirm('Ви впевнені, що хочете видалити ВСІ повідомлення?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/feedback/delete-all', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  window.location.reload();
                } else {
                  alert('Помилка видалення: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка видалення всіх повідомлень:', error);
                alert('Не вдалося видалити всі повідомлення.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/feedback', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні зворотного зв’язку');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/feedback виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Початок тесту
app.get('/test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');
    const testNumber = req.query.test;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).send('Номер тесту не вказано або тест не знайдено');
    }

    const canAttemptTest = await checkTestAttempts(req.user, testNumber);
    if (!canAttemptTest) {
      return res.send(`
        <!DOCTYPE html>
        <html lang="uk">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Помилка</title>
            <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f5f5f5; margin: 0; }
              #modal { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border: 2px solid black; z-index: 1000; box-shadow: 0 0 10px rgba(0,0,0,0.3); border-radius: 10px; }
              button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; transition: background-color 0.3s; }
              button:hover { background-color: #45a049; }
              .overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 999; }
              h2 { margin-bottom: 20px; font-size: 24px; color: #333; }
            </style>
          </head>
          <body>
            <div class="overlay"></div>
            <div id="modal">
              <h2>Ви вже проходили сьогодні цей тест</h2>
              <button onclick="window.location.href='/select-test'">Повернутися до вибору тесту</button>
            </div>
          </body>
        </html>
      `);
    }

    let questions = await loadQuestions(testNumber);
    const userVariant = Math.floor(Math.random() * 3) + 1;
    logger.info(`Призначено варіант користувачу ${req.user} для тесту ${testNumber}: Variant ${userVariant}`);

    questions = questions.filter(q => !q.variant || q.variant === '' || q.variant === `Variant ${userVariant}`);
    logger.info(`Відфільтровано питання для тесту ${testNumber}, варіант ${userVariant}: знайдено ${questions.length} питань`);

    if (questions.length === 0) {
      return res.status(400).send(`Немає питань для варіанту ${userVariant} у тесті ${testNumber}`);
    }

    const questionLimit = testNames[testNumber].questionLimit;
    if (questionLimit && questions.length > questionLimit) {
      questions = shuffleArray([...questions]).slice(0, questionLimit);
    }

    if (testNames[testNumber].randomQuestions) {
      questions = shuffleArray([...questions]);
    }

    if (testNames[testNumber].randomAnswers) {
      questions = questions.map(q => {
        if (q.options && q.options.length > 0 && q.type !== 'ordering' && q.type !== 'matching') {
          const shuffledOptions = shuffleArray([...q.options]);
          return { ...q, options: shuffledOptions };
        } else if (q.type === 'matching' && q.pairs) {
          const shuffledPairs = shuffleArray([...q.pairs]);
          return { ...q, pairs: shuffledPairs };
        }
        return q;
      });
    }

    const testStartTime = Date.now();
    const testSessionId = `${req.user}_${testNumber}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    // Збереження стану тесту в MongoDB
    const testData = {
      user: req.user,
      testNumber,
      questions,
      answers: {},
      currentQuestion: 0,
      startTime: testStartTime,
      timeLimit: testNames[testNumber].timeLimit * 1000,
      variant: userVariant,
      isQuickTest: testNames[testNumber].isQuickTest,
      timePerQuestion: testNames[testNumber].timePerQuestion,
      testSessionId: testSessionId,
      isSavingResult: false,
      answerTimestamps: {},
      questionStartTime: {},
      suspiciousActivity: { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] }
    };

    await db.collection('active_tests').updateOne(
      { user: req.user },
      { $set: testData },
      { upsert: true }
    );

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await logActivity(req.user, `розпочав тест ${testNames[testNumber].name.replace(/"/g, '\\"')}`, ipAddress);
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    logger.error('Помилка в /test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.get('/instructions', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Інструкція з проходження тестів</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 20px;
              background-color: #f5f5f5;
              line-height: 1.6;
              color: #333;
            }
            .container {
              max-width: 800px;
              margin: 0 auto;
              background-color: white;
              padding: 30px;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            h1 {
              font-size: 28px;
              text-align: center;
              color: #2c3e50;
              margin-bottom: 20px;
            }
            h2 {
              font-size: 22px;
              color: #34495e;
              margin-top: 30px;
              margin-bottom: 10px;
              border-bottom: 1px solid #eee;
              padding-bottom: 5px;
            }
            p, li {
              font-size: 16px;
              margin-bottom: 10px;
            }
            ul {
              list-style-type: disc;
              padding-left: 25px;
              margin: 10px 0;
            }
            li {
              margin-bottom: 8px;
            }
            strong {
              color: #2c3e50;
            }
            .nav-btn {
              display: inline-block;
              padding: 12px 24px;
              margin-top: 30px;
              cursor: pointer;
              border: none;
              border-radius: 6px;
              background-color: #4CAF50;
              color: white;
              text-decoration: none;
              font-size: 17px;
              text-align: center;
              transition: background 0.2s;
            }
            .nav-btn:hover {
              background-color: #45a049;
            }
            @media (max-width: 600px) {
              .container {
                padding: 20px 15px;
              }
              h1 { font-size: 24px; }
              h2 { font-size: 20px; }
              p, li { font-size: 15px; }
              .nav-btn {
                width: 100%;
                box-sizing: border-box;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Інструкція з проходження тестів</h1>
            <p>Вітаємо Вас у Центрі тестування! Щоб отримати максимально точні результати та зручно пройти тест, будь ласка, дотримуйтесь цих простих рекомендацій:</p>

            <h2>1. Підготовка до тесту</h2>
            <ul>
              <li>Перевірте стабільне інтернет-з’єднання (Wi-Fi або мобільний інтернет).</li>
              <li>Використовуйте актуальну версію браузера: Google Chrome, Mozilla Firefox або Microsoft Edge.</li>
              <li>Закрийте зайві вкладки та програми, щоб тест працював швидко.</li>
            </ul>

            <h2>2. Початок тесту</h2>
            <ul>
              <li>Оберіть потрібний тест зі списку.</li>
              <li>Уважно прочитайте назву та інструкцію до тесту перед початком.</li>
              <li>Пам’ятайте: кількість спроб на день може бути обмежена (залежить від правил Вашого курсу/групи). Якщо спроби закінчилися — дочекайтеся наступного дня або зверніться до адміністратора.</li>
            </ul>

            <h2>3. Процес проходження тесту</h2>
            <ul>
              <li><strong>Типи питань та як відповідати:</strong>
                <ul>
                  <li><strong>Одинарний вибір</strong> — оберіть лише одну правильну відповідь. В цих типах питань лише одна правильна відповідь.</li>
                  <li><strong>Множинний вибір</strong> — оберіть усі правильні відповіді. В цих типах питань може бути кілька правильних відповідей.</li>
                  <li><strong>Введення тексту</strong> — введіть відповідь вручну (враховується регістр і пробіли).</li>
                  <li><strong>Заповнення пропусків</strong> — заповніть усі пропуски в реченні.</li>
                  <li><strong>Впорядкування</strong> — перетягніть варіанти у правильну послідовність.</li>
                  <li><strong>Встановлення відповідностей (matching)</strong> — зіставте елементи лівої колонки до відповідних елементів правої.</li>
                </ul>
                Завжди читайте інструкцію під текстом питання — вона підказує, скільки відповідей потрібно обрати.
              </li>

              <li><strong>Полоса прогресу (кружечки зверху)</strong><br>
                Кожен кружечок — це окреме питання.<br>
                <strong>Клікніть на будь-який кружечок</strong> — система збереже поточну відповідь і перенесе Вас на це питання. Це зручно, щоб швидко повернутися до пропущених або сумнівних питань.<br>
                <strong>Що означають кольори кружечків під час проходження:</strong>
                <ul>
                  <li><strong>Червоний</strong> — питання не отримано жодної відповіді (пропущене).</li>
                  <li><strong>Зелений</strong> — на питання дано відповідь.</li>
                </ul>
              </li>

              <li><strong>Таймер</strong><br>
                Слідкуйте за таймером у верхній частині екрана. Якщо час закінчиться — тест автоматично завершиться.
              </li>

              <li><strong>Перехід між питаннями</strong><br>
                Використовуйте кнопки «Назад» / «Далі» або клікайте на кружечки прогресу.
              </li>
            </ul>

            <h2>4. Підрахунок балів</h2>
            <ul>
              <li>Кожне питання має свою вагу (кількість балів).</li>
              <li>Правильна відповідь дає повну кількість балів за питання.</li>
              <li>Частково правильна відповідь (наприклад, обрано не всі варіанти) дає частину балів або 0 — залежить від типу питання.</li>
              <li>Неправильна відповідь або пропущене питання = 0 балів.</li>
              <li>Загальний відсоток = (набрані бали / максимально можливі бали) × 100%.</li>
            </ul>

            <h2>5. Завершення тесту</h2>
            <ul>
              <li>До самого завершення тесту Ви можете редагувати відповіді на питання.</li>
              <li>Натисніть «Завершити тест», коли закінчите (або дочекайтеся таймера).</li>
              <li>Після завершення Ви побачите сторінку результатів з відсотком, кількістю правильних відповідей, набраними балами та максимальною кількістю балів у тесті.</li>
              <li>На сторінці результатів кружечки питань будуть показувати правильність відповідей:
                <ul>
                  <li><strong>Зелений</strong> — відповідь вірна.</li>
                  <li><strong>Червоний</strong> — відповідь невірна.</li>
                  <li><strong>Жовтий</strong> — дано часткову відповідь (наприклад, вибрано не всі правильні варіанти або є суміш правильних і неправильних).</li>
                </ul>
              </li>
              <li>Натисніть «Експортувати в PDF», щоб зберегти результати у файл (ім’я файлу буде у форматі ваш_логін_результат.pdf).</li>
              <li>Кнопка «Вихід» поверне Вас на сторінку вибору тестів.</li>
            </ul>

            <h2>6. Важливі моменти</h2>
            <ul>
              <li>Не перемикайтеся часто між вкладками та не залишайте сторінку надовго — система може зафіксувати підозрілу активність.</li>
              <li>Якщо тест перервався (збій інтернету, закриття вкладки тощо) — спробуйте повернутися на сторінку тесту. Прогрес зазвичай зберігається.</li>
              <li>Якщо з’являється помилка — оновіть сторінку. Якщо не допомагає — зверніться до адміністратора.</li>
            </ul>

            <h2>7. Підтримка</h2>
            <p>Якщо у Вас є питання, щось не працює, питання тесту не завантажується або Ви не можете пройти тест — напишіть адміністратору (через пошту, Telegram або форму зворотного зв’язку).</p>

            <p style="text-align: center; font-size: 18px; margin: 30px 0 20px; color: #2c3e50;">
              Бажаємо успіхів у тестуванні! 💪<br>
              Ваш результат — це Ваш прогрес! 😊
            </p>

            <a href="/select-test" class="nav-btn">Назад до вибору тесту</a>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /instructions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні інструкції');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /instructions виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Відображення питання тесту — з правильним переносом тексту цілими словами
app.get('/test/question', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');

    let userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      return res.status(400).send('Тест не розпочато');
    }

    const { 
      questions, 
      testNumber, 
      answers, 
      currentQuestion, 
      startTime: testStartTime, 
      timeLimit, 
      isQuickTest, 
      timePerQuestion, 
      suspiciousActivity, 
      variant, 
      testSessionId 
    } = userTest;

    // Перевірка кешу тестів
    if (!testNames[testNumber]) {
      logger.info('Номер тесту не знайдено в кеші, повторне завантаження тестів', { testNumber });
      const tests = await db.collection('tests').find().toArray();
      testNames = tests.reduce((acc, test) => {
        acc[test.testNumber] = test;
        return acc;
      }, {});
      logger.info('Оновлено кеш testNames', { testCount: Object.keys(testNames).length });
    }

    // Перевірка доступності тесту
    if (!testNames[testNumber]) {
      let score = 0;
      const totalPoints = questions.reduce((sum, q) => sum + q.points, 0);
      const scoresPerQuestion = questions.map((q, index) => {
        const userAnswer = answers[index];
        let questionScore = 0;
        const normalizeAnswer = (answer) => {
          if (!answer) return '';
          return String(answer)
            .trim()
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(',', '.')
            .replace(/\\'/g, "'")
            .replace(/°/g, 'deg');
        };
        if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const isCorrect = correctAnswers.length === userAnswers.length &&
            correctAnswers.every(val => userAnswers.includes(val)) &&
            userAnswers.every(val => correctAnswers.includes(val));
          if (isCorrect) questionScore = q.points;
        } else if (q.type === 'input' && userAnswer) {
          const normalizedUserAnswer = normalizeAnswer(userAnswer);
          const normalizedCorrectAnswer = normalizeAnswer(q.correctAnswers[0]);
          if (normalizedCorrectAnswer.includes('-')) {
            const [min, max] = normalizedCorrectAnswer.split('-').map(val => parseFloat(val.trim()));
            const userValue = parseFloat(normalizedUserAnswer);
            const isCorrect = !isNaN(userValue) && userValue >= min && userValue <= max;
            if (isCorrect) questionScore = q.points;
          } else {
            const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
            if (isCorrect) questionScore = q.points;
          }
        } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
          if (isCorrect) questionScore = q.points;
        } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
          const userPairs = userAnswer.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const correctPairs = q.correctPairs.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const isCorrect = userPairs.length === correctPairs.length &&
            userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]));
          if (isCorrect) questionScore = q.points;
        } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const isCorrect = userAnswers.length === correctAnswers.length &&
            userAnswers.every((answer, idx) => {
              const correctAnswer = correctAnswers[idx];
              if (correctAnswer.includes('-')) {
                const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
                const userValue = parseFloat(answer);
                return !isNaN(userValue) && userValue >= min && userValue <= max;
              } else {
                return answer === correctAnswer;
              }
            });
          if (isCorrect) questionScore = q.points;
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswer = normalizeAnswer(q.correctAnswer);
          const isCorrect = userAnswers.length === 1 && userAnswers[0] === correctAnswer;
          if (isCorrect) questionScore = q.points;
        }
        return questionScore;
      });
      score = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
      const endTime = Date.now();
      const percentage = (score / totalPoints) * 100;
      const totalClicks = Object.keys(answers).length;
      const correctClicks = scoresPerQuestion.filter(s => s > 0).length;
      const totalQuestions = questions.length;
      const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
      const existingResult = await db.collection('test_results').findOne({ testSessionId });
      if (!existingResult && !userTest.isSavingResult) {
        await db.collection('active_tests').updateOne(
          { user: req.user },
          { $set: { isSavingResult: true } }
        );
        await saveResult(
          req.user,
          testNumber,
          score,
          totalPoints,
          testStartTime,
          endTime,
          totalClicks,
          correctClicks,
          totalQuestions,
          percentage,
          suspiciousActivity,
          answers,
          scoresPerQuestion,
          variant,
          ipAddress,
          testSessionId
        );
        logger.info(`Результат збережено для testSessionId: ${testSessionId} через недоступність тесту`);
      }
      await db.collection('active_tests').deleteOne({ user: req.user });
      return res.send(`
        <!DOCTYPE html>
        <html lang="uk">
          <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Помилка</title>
            <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f5f5f5; margin: 0; }
              h2 { font-size: 24px; margin-bottom: 20px; }
              button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
              button:hover { background-color: #45a049; }
            </style>
          </head>
          <body>
            <h2>Цей тест більше недоступний. Ваші відповіді збережено. Оберіть інший тест.</h2>
            <button onclick="window.location.href='/select-test'">Повернутися до вибору тестів</button>
          </body>
        </html>
      `);
    }

    const index = parseInt(req.query.index) || 0;
    if (index < 0 || index >= questions.length) {
      return res.status(400).send('Невірний номер питання');
    }

    const updateData = {
      currentQuestion: index,
      answerTimestamps: userTest.answerTimestamps || {},
      suspiciousActivity: {
        timeAway: userTest.suspiciousActivity?.timeAway || 0,
        switchCount: userTest.suspiciousActivity?.switchCount || 0,
        responseTimes: userTest.suspiciousActivity?.responseTimes || [],
        activityCounts: userTest.suspiciousActivity?.activityCounts || []
      }
    };
    updateData.answerTimestamps[index] = Date.now();
    await db.collection('active_tests').updateOne(
      { user: req.user },
      { $set: updateData }
    );

    const q = questions[index];
    const progress = Array.from({ length: questions.length }, (_, i) => ({
      number: i + 1,
      answered: answers[i] && (Array.isArray(answers[i]) ? answers[i].length > 0 : answers[i].trim() !== '')
    }));

    let totalTestTime = timeLimit / 1000;
    if (isQuickTest) {
      totalTestTime = questions.length * timePerQuestion;
    }
    const elapsedTime = Math.floor((Date.now() - testStartTime) / 1000);
    const remainingTime = Math.max(0, totalTestTime - elapsedTime);
    const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
    const seconds = (remainingTime % 60).toString().padStart(2, '0');

    const selectedOptions = answers[index] || [];
    const selectedOptionsString = JSON.stringify(selectedOptions);

    const questionStartTimeObj = userTest.questionStartTime || {};
    if (!questionStartTimeObj[index]) {
      questionStartTimeObj[index] = Date.now();
      await db.collection('active_tests').updateOne(
        { user: req.user },
        { $set: { questionStartTime: questionStartTimeObj } }
      );
    }

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${testNames[testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}</title>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js"></script>
          <style>
            body {
              font-family: Arial, sans-serif;
              margin: 0;
              padding: 20px;
              padding-bottom: 120px;
              background-color: #f0f0f0;
            }
            h1 {
              font-size: 28px;
              text-align: center;
              margin-bottom: 15px;
            }
            img {
              max-width: 100%;
              max-height: 400px;
              margin: 15px auto;
              display: block;
              border-radius: 8px;
              object-fit: contain;
            }
            .progress-bar {
              display: flex;
              flex-wrap: wrap;
              justify-content: center;
              gap: 8px;
              max-width: 95vw;
              margin: 0 auto 30px;
              padding: 0 10px;
            }
            .progress-circle {
              width: 38px;
              height: 38px;
              border-radius: 50%;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 14px;
              font-weight: bold;
              flex-shrink: 0;
              min-width: 38px;
              box-shadow: 0 2px 5px rgba(0,0,0,0.1);
              cursor: pointer;
              transition: all 0.18s ease;
            }
            .progress-circle:hover {
              transform: scale(1.15);
              box-shadow: 0 4px 12px rgba(0,0,0,0.25);
            }
            .progress-circle.unanswered { background: #ff4d4d; color: white; }
            .progress-circle.answered   { background: #4CAF50; color: white; }
            .progress-line {
              width: 6px;
              height: 3px;
              background: #ccc;
              align-self: center;
              flex-shrink: 0;
            }
            .progress-line.answered { background: #4CAF50; }
            .question-box {
              background: white;
              padding: 20px;
              border-radius: 10px;
              box-shadow: 0 4px 12px rgba(0,0,0,0.1);
              margin-bottom: 25px;
            }
            .instruction {
              font-style: italic;
              color: #555;
              margin: 10px 0 20px;
              font-size: 17px;
            }
            #question-container {
              max-width: 1100px;
              margin: 0 auto;
            }
            #answers .option-box,
            #answers .matching-item,
            #answers input[type="text"],
            #answers .blank-input {
              background: white;
              border: 2px solid #ddd;
              box-sizing: border-box;
            }
            .option-box {
              border: 2px solid #ddd;
              padding: 16px 20px;
              margin: 10px 0;
              border-radius: 10px;
              cursor: pointer;
              font-size: 17px;
              user-select: none;
              transition: all 0.2s;
              background: white;
              min-height: 64px;
              height: auto;
              display: flex;
              align-items: center;
              justify-content: flex-start;
              text-align: left;
              white-space: normal;
              overflow-wrap: break-word;
              word-break: normal;
              hyphens: none;
              line-height: 1.45;
              overflow: visible;
            }
            .option-box:hover {
              background: #f8f9fa;
              border-color: #bbb;
            }
            .option-box.selected {
              background: #d4edda !important;
              border-color: #28a745 !important;
              box-shadow: 0 0 0 3px rgba(40,167,69,0.3) !important;
            }
            @media (max-width: 600px) {
              .option-box {
                font-size: 15px;
                padding: 14px 18px;
                min-height: 60px;
              }
              .progress-circle {
                width: 32px;
                height: 32px;
                font-size: 12px;
                min-width: 32px;
              }
              .progress-line {
                width: 5px;
              }
            }
            .option-box.draggable { cursor: move; }
            .option-box.dragging { opacity: 0.6; box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
            .matching-container {
              display: grid;
              grid-template-columns: 1fr 1fr;
              gap: 12px;
              margin: 20px 0;
            }
            .matching-column {
              display: flex;
              flex-direction: column;
              gap: 10px;
            }
            .matching-item {
              border: 2px solid #ddd;
              padding: 16px 20px;
              border-radius: 10px;
              font-size: 17px;
              min-height: 70px;
              height: auto;
              overflow: visible;
              text-overflow: unset;
              white-space: normal;
              overflow-wrap: break-word;
              word-break: normal;
              hyphens: none;
              display: flex;
              align-items: center;
              justify-content: flex-start;
              background: white;
              transition: all 0.2s;
              box-sizing: border-box;
              line-height: 1.45;
            }
            .static { cursor: default; background: #f8f9fa; }
            .button-container {
              position: fixed;
              bottom: 20px;
              left: 20px;
              right: 20px;
              display: flex;
              justify-content: space-between;
              gap: 12px;
              z-index: 100;
              flex-wrap: nowrap;
            }
            @media (max-width: 600px) {
              .button-container {
                gap: 8px;
                padding: 0 10px;
                flex-wrap: nowrap !important;
                overflow-x: auto;
                box-sizing: border-box;
              }
              button {
                font-size: 15px;
                padding: 12px 16px;
                min-height: 55px;
              }
            }
            button {
              flex: 1;
              padding: 16px 24px;
              font-size: 18px;
              font-weight: bold;
              border: none;
              border-radius: 10px;
              cursor: pointer;
              min-height: 60px;
              display: flex;
              align-items: center;
              justify-content: center;
              box-shadow: 0 4px 10px rgba(0,0,0,0.15);
              transition: all 0.2s;
            }
            button:hover { transform: translateY(-2px); box-shadow: 0 6px 14px rgba(0,0,0,0.2); }
            .back-btn   { background: #dc3545; color: white; }
            .next-btn   { background: #007bff; color: white; }
            .finish-btn { background: #28a745; color: white; }
            button:disabled {
              background: #ccc !important;
              cursor: not-allowed;
              transform: none;
              box-shadow: none;
            }
            #timer {
              font-size: 26px;
              font-weight: bold;
              text-align: center;
              margin: 15px 0 25px;
              color: #333;
            }
            #question-timer {
              position: relative;
              width: 90px;
              height: 90px;
              margin: 0 auto 15px;
            }
            #question-timer svg { width: 100%; height: 100%; transform: rotate(-90deg); }
            #question-timer circle { fill: none; stroke-width: 10; }
            #question-timer .timer-circle-bg { stroke: #e0e0e0; }
            #question-timer .timer-circle { stroke: #ff4d4d; stroke-dasharray: 280; transition: stroke-dashoffset 0.3s linear; }
            #question-timer .timer-text {
              position: absolute;
              top: 50%;
              left: 50%;
              transform: translate(-50%, -50%);
              font-size: 28px;
              font-weight: bold;
              color: #333;
            }
            #confirm-modal {
              display: none;
              position: fixed;
              top: 50%;
              left: 50%;
              transform: translate(-50%, -50%);
              background: white;
              padding: 30px;
              border-radius: 12px;
              box-shadow: 0 10px 30px rgba(0,0,0,0.4);
              z-index: 1000;
              text-align: center;
              max-width: 90%;
            }
            #confirm-modal h2 { margin: 0 0 25px; font-size: 24px; }
            #confirm-modal .buttons {
              display: flex;
              justify-content: center;
              gap: 20px;
              margin-top: 20px;
            }
            #confirm-modal button {
              min-width: 120px;
              padding: 14px 30px;
              font-size: 18px;
            }
          </style>
        </head>
        <body>
          <h1>${testNames[testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}</h1>
          <div id="timer">Залишилось часу: ${minutes} хв ${seconds} с</div>

          <div class="progress-bar">
            ${progress.map((p, j) => `
              <div 
                class="progress-circle ${p.answered ? 'answered' : 'unanswered'}" 
                data-index="${j}"
                onclick="goToQuestion(${j})"
                title="Перейти до питання ${p.number}"
              >${p.number}</div>
              ${j < progress.length - 1 ? '<div class="progress-line ' + (p.answered ? 'answered' : '') + '"></div>' : ''}
            `).join('')}
          </div>

          <div id="question-container">
    `;

    if (isQuickTest) {
      html += `
        <div id="question-timer">
          <svg viewBox="0 0 80 80">
            <circle class="timer-circle-bg" cx="40" cy="40" r="36" />
            <circle class="timer-circle" cx="40" cy="40" r="36" />
          </svg>
          <div class="timer-text" id="timer-text">${timePerQuestion}</div>
        </div>
      `;
    }

    if (q.picture && q.picture.trim() !== '') {
      html += `
        <div id="image-container">
          <img src="${q.picture}" alt="Picture" style="max-height: 400px; object-fit: contain;" onerror="this.style.display='none'; document.getElementById('image-error').style.display='block';">
          <div id="image-error" class="image-error" style="display: none;">Зображення недоступне</div>
        </div>
      `;
    }

    const instructionText = q.type === 'multiple' ? 'Виберіть усі правильні відповіді' :
                           q.type === 'input' ? 'Введіть правильну відповідь' :
                           q.type === 'ordering' ? 'Розташуйте відповіді у правильній послідовності' :
                           q.type === 'matching' ? 'Складіть правильні пари, перетягуючи елементи' :
                           q.type === 'fillblank' ? 'Заповніть пропуски у реченні' :
                           q.type === 'singlechoice' ? 'Виберіть правильну відповідь' : '';

    html += `
            <div class="question-box">
              <h2 id="question-text">${index + 1}. `;
   
    if (q.type === 'fillblank') {
      const userAnswers = Array.isArray(answers[index]) ? answers[index] : [];
      const parts = q.text.split('___');
      let inputHtml = '';
      parts.forEach((part, i) => {
        inputHtml += `<span class="question-text">${part}</span>`;
        if (i < parts.length - 1) {
          const userAnswer = userAnswers[i] || '';
          inputHtml += `<input type="text" class="blank-input" id="blank_${i}" value="${userAnswer.replace(/"/g, '\\"')}" placeholder="Введіть відповідь" autocomplete="off">`;
        }
      });
      html += inputHtml;
    } else {
      html += q.text;
    }

    html += `
              </h2>
            </div>
            <p id="instruction" class="instruction">${instructionText}</p>
            <div id="answers">
    `;

        // ==================== MATCHING ====================
    if (q.type === 'matching' && q.pairs) {
      const leftItems = q.pairs.map(p => p.left);
      const rightItems = shuffleArray([...q.pairs.map(p => p.right)]);

      html += `
        <div class="matching-container" id="matching-${index}">
          <!-- Ліва колонка (сортується, але без перетягування в іншу колонку) -->
          <div class="matching-column" id="left-column-${index}">
            <h4 style="margin-bottom:15px;color:#333;text-align:center;">Терміни</h4>
            ${leftItems.map(item => {
              const escaped = item.replace(/'/g, "\\'").replace(/"/g, '\\"');
              return `<div class="matching-item draggable" data-value="${escaped}">${item}</div>`;
            }).join('')}
          </div>

          <!-- Права колонка (сортується тільки всередині себе) -->
          <div class="matching-column" id="right-column-${index}">
            <h4 style="margin-bottom:15px;color:#333;text-align:center;">Відповіді (сортуйте)</h4>
            <div id="sortable-right-${index}">
              ${rightItems.map(item => {
                const escaped = item.replace(/'/g, "\\'").replace(/"/g, '\\"');
                return `<div class="matching-item draggable" data-value="${escaped}">${item}</div>`;
              }).join('')}
            </div>
          </div>
        </div>

        <button onclick="resetMatching(${index})" style="margin-top:15px;">Скинути порядок</button>

        <script>
          // Сортування тільки всередині своєї колонки (без group — немає перетягування між колонками)
          new Sortable(document.getElementById('left-column-${index}'), {
            animation: 150,
            ghostClass: 'sortable-ghost',
            onEnd: () => saveMatchingAnswer(${index})
          });

          new Sortable(document.getElementById('sortable-right-${index}'), {
            animation: 150,
            ghostClass: 'sortable-ghost',
            onEnd: () => saveMatchingAnswer(${index})
          });

          function saveMatchingAnswer(idx) {
            const leftEls = document.querySelectorAll('#left-column-' + idx + ' .matching-item');
            const rightEls = document.querySelectorAll('#sortable-right-' + idx + ' .matching-item');
            
            const pairs = [];
            leftEls.forEach((leftEl, i) => {
              const rightEl = rightEls[i];
              if (leftEl && rightEl) {
                pairs.push([
                  leftEl.getAttribute('data-value'),
                  rightEl.getAttribute('data-value')
                ]);
              }
            });

            window.currentAnswers = window.currentAnswers || {};
            window.currentAnswers[idx] = pairs;

            if (typeof saveAnswer === 'function') {
              saveAnswer(idx, pairs);
            }
          }

          function resetMatching(idx) {
            if (confirm('Скинути порядок?')) location.reload();
          }

          // Вирівнювання висоти блоків
          function equalizeMatchingHeights() {
            const allItems = document.querySelectorAll('#matching-' + idx + ' .matching-item');
            if (allItems.length === 0) return;
            let maxHeight = 0;
            allItems.forEach(item => {
              item.style.height = 'auto';
              const height = item.getBoundingClientRect().height;
              if (height > maxHeight) maxHeight = height;
            });
            allItems.forEach(item => {
              item.style.height = maxHeight + 'px';
            });
          }

          window.addEventListener('load', equalizeMatchingHeights);
        </script>
      `;
    } else if (!q.options || q.options.length === 0) {
      if (q.type !== 'fillblank') {
        const userAnswer = answers[index] || '';
        html += `
          <input type="text" name="q${index}" id="q${index}_input" value="${userAnswer}" placeholder="Введіть відповідь" class="answer-option" autocomplete="off"><br>
        `;
      }
    } else {
      if (q.type === 'ordering') {
        html += `
          <div id="sortable-options">
            ${(answers[index] || q.options).map((option, optIndex) => {
              const escapedOption = option.replace(/'/g, "\\'").replace(/"/g, '\\"');
              return `
                <div class="option-box draggable" data-index="${optIndex}" data-value="${escapedOption}">
                  ${option}
                </div>
              `;
            }).join('')}
          </div>
        `;
      } else {
        q.options.forEach((option, optIndex) => {
          const selected = selectedOptions.includes(option) ? 'selected' : '';
          const escapedOption = option.replace(/'/g, "\\'").replace(/"/g, '\\"');
          html += `
            <div class="option-box ${selected}" data-value="${escapedOption}">
              ${option}
            </div>
          `;
        });
      }
    }

    html += `
            </div>
          </div>
          <div class="button-container">
            ${!isQuickTest ? `
              <button class="back-btn" ${index === 0 ? 'disabled' : ''} onclick="window.location.href='/test/question?index=${index - 1}'">Назад</button>
            ` : ''}
            <button id="submit-answer" class="next-btn" ${index === questions.length - 1 ? 'disabled' : ''} onclick="saveAndNext(${index})">Далі</button>
            <button class="finish-btn" onclick="showConfirm(${index})">Завершити тест</button>
          </div>
          <div id="confirm-modal">
            <h2>Ви дійсно бажаєте завершити тест?</h2>
            <div class="buttons">
              <button onclick="finishTest(${index})">Так</button>
              <button onclick="hideConfirm()">Ні</button>
            </div>
          </div>

          <script>
            const startTime = ${testStartTime};
            const timeLimit = ${timeLimit};
            const totalTestTime = ${totalTestTime};
            const timerElement = document.getElementById('timer');
            const isQuickTest = ${isQuickTest};
            const timePerQuestion = ${timePerQuestion || 0};
            const totalQuestions = ${questions.length};
            let timeAway = ${userTest.suspiciousActivity?.timeAway || 0};
            let lastBlurTime = 0;
            let switchCount = ${userTest.suspiciousActivity?.switchCount || 0};
            let lastActivityTime = Date.now();
            let activityCount = ${userTest.suspiciousActivity?.activityCounts?.[index] || 0};
            let lastMouseMoveTime = 0;
            let lastKeydownTime = 0;
            const debounceDelay = 100;
            const blurDebounceDelay = 200;
            let blurTimeout = null;
            let selectedOptions = ${selectedOptionsString};
            let matchingPairs = ${JSON.stringify(answers[index] || [])};
            let questionTimeRemaining = timePerQuestion;
            let currentQuestionIndex = ${index};
            let lastGlobalUpdateTime = Date.now();
            let isSaving = false;
            let hasMovedToNext = false;
            let questionStartTime = ${questionStartTimeObj[index]};

            // Функція переходу на інше питання
            function goToQuestion(targetIndex) {
              if (targetIndex === currentQuestionIndex) return;
              saveCurrentAnswer(currentQuestionIndex).then(() => {
                window.location.href = '/test/question?index=' + targetIndex;
              }).catch(err => {
                console.error('Помилка перед переходом:', err);
                window.location.href = '/test/question?index=' + targetIndex;
              });
            }

            // Автозбереження відповіді
            async function saveCurrentAnswer(index) {
              if (isSaving) return;
              isSaving = true;
              try {
                let answers = selectedOptions;
                if (document.querySelector('input[name="q' + index + '"]')) {
                  answers = document.getElementById('q' + index + '_input').value;
                } else if (document.getElementById('sortable-options')) {
                  answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
                } else if (document.getElementById('left-column')) {
                  answers = matchingPairs;
                } else if ('${q.type}' === 'fillblank') {
                  answers = [];
                  for (let i = 0; i < ${q.blankCount || 1}; i++) {
                    const input = document.getElementById('blank_' + i);
                    answers.push(input ? input.value.trim() : '');
                  }
                }
                const responseTime = (Date.now() - questionStartTime) / 1000;
                const formData = new URLSearchParams();
                formData.append('index', index);
                const safeAnswer = JSON.stringify(answers);
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');
                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = await response.json();
                if (!result.success) console.error('Помилка автозбереження:', result.error);
              } catch (error) {
                console.error('Помилка автозбереження:', error);
              } finally {
                isSaving = false;
              }
            }

            // Збереження + наступне питання
            async function saveAndNext(index) {
              if (isSaving) return;
              isSaving = true;
              try {
                hasMovedToNext = true;
                let answers = selectedOptions;
                if (document.querySelector('input[name="q' + index + '"]')) {
                  answers = document.getElementById('q' + index + '_input').value;
                } else if (document.getElementById('sortable-options')) {
                  answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
                } else if (document.getElementById('left-column')) {
                  answers = matchingPairs;
                } else if ('${q.type}' === 'fillblank') {
                  answers = [];
                  for (let i = 0; i < ${q.blankCount || 1}; i++) {
                    const input = document.getElementById('blank_' + i);
                    answers.push(input ? input.value.trim() : '');
                  }
                }
                const responseTime = (Date.now() - questionStartTime) / 1000;
                const formData = new URLSearchParams();
                formData.append('index', index);
                const safeAnswer = JSON.stringify(answers);
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');
                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = await response.json();
                if (result.success) {
                  const nextIndex = index + 1;
                  fetch('/set-question-start-time?index=' + nextIndex, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({ '_csrf': '${res.locals._csrf}' })
                  }).then(() => {
                    if (nextIndex < ${questions.length}) {
                      window.location.href = '/test/question?index=' + nextIndex;
                    } else {
                      setTimeout(() => window.location.href = '/result', 300);
                    }
                  });
                } else {
                  console.error('Помилка збереження:', result.error);
                  alert('Помилка збереження відповіді');
                }
              } catch (error) {
                console.error('Помилка в saveAndNext:', error);
                alert('Не вдалося зберегти відповідь');
              } finally {
                isSaving = false;
              }
            }

            // Завершення тесту
            async function finishTest(index) {
              if (isSaving) return;
              isSaving = true;
              try {
                await saveCurrentAnswer(index);
                let answers = selectedOptions;
                if (document.querySelector('input[name="q' + index + '"]')) {
                  answers = document.getElementById('q' + index + '_input').value;
                } else if (document.getElementById('sortable-options')) {
                  answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
                } else if (document.getElementById('left-column')) {
                  answers = matchingPairs;
                } else if ('${q.type}' === 'fillblank') {
                  answers = [];
                  for (let i = 0; i < ${q.blankCount || 1}; i++) {
                    const input = document.getElementById('blank_' + i);
                    answers.push(input ? input.value.trim() : '');
                  }
                }
                const responseTime = (Date.now() - questionStartTime) / 1000;
                const formData = new URLSearchParams();
                formData.append('index', index);
                const safeAnswer = JSON.stringify(answers);
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');
                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = await response.json();
                if (result.success) {
                  setTimeout(() => window.location.href = '/result', 300);
                } else {
                  console.error('Помилка завершення:', result.error);
                  alert('Помилка завершення тесту');
                }
              } catch (error) {
                console.error('Помилка в finishTest:', error);
                alert('Не вдалося завершити тест');
              } finally {
                isSaving = false;
              }
            }

            function showConfirm(index) {
              document.getElementById('confirm-modal').style.display = 'block';
            }

            function hideConfirm() {
              document.getElementById('confirm-modal').style.display = 'none';
            }

            function updateGlobalTimer() {
              const now = Date.now();
              const elapsedTime = Math.floor((now - startTime) / 1000);
              const remainingTime = Math.max(0, totalTestTime - elapsedTime);
              const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
              const seconds = (remainingTime % 60).toString().padStart(2, '0');
              timerElement.textContent = 'Залишилось часу: ' + minutes + ' хв ' + seconds + ' с';
              lastGlobalUpdateTime = now;
              if (remainingTime <= 0) {
                saveCurrentAnswer(currentQuestionIndex).then(() => {
                  setTimeout(() => window.location.href = '/result', 1500);
                });
              }
            }
            setInterval(updateGlobalTimer, 1000);

            if (isQuickTest) {
              function updateQuestionTimer() {
                const now = Date.now();
                const elapsedSinceQuestionStart = Math.floor((now - questionStartTime) / 1000);
                questionTimeRemaining = Math.max(0, timePerQuestion - elapsedSinceQuestionStart);
                const timerText = document.getElementById('timer-text');
                const timerCircle = document.querySelector('#question-timer .timer-circle');
                if (timerText && timerCircle) {
                  timerText.textContent = Math.round(questionTimeRemaining);
                  const circumference = 251;
                  const offset = (1 - questionTimeRemaining / timePerQuestion) * circumference;
                  timerCircle.style.strokeDashoffset = offset;
                }
                if (questionTimeRemaining <= 0 && currentQuestionIndex < totalQuestions - 1 && !hasMovedToNext) {
                  hasMovedToNext = true;
                  saveCurrentAnswer(currentQuestionIndex).then(() => {
                    saveAndNext(currentQuestionIndex);
                  });
                }
              }
              const questionTimerInterval = setInterval(updateQuestionTimer, 50);
            }

            window.addEventListener('blur', () => {
              if (!blurTimeout) {
                blurTimeout = setTimeout(() => {
                  if (lastBlurTime === 0) {
                    lastBlurTime = performance.now();
                    switchCount = Math.min(switchCount + 1, 1000);
                  }
                  blurTimeout = null;
                }, blurDebounceDelay);
              }
            });

            window.addEventListener('focus', () => {
              if (blurTimeout) {
                clearTimeout(blurTimeout);
                blurTimeout = null;
              }
              if (lastBlurTime > 0) {
                const now = performance.now();
                const awayDuration = Math.min((now - lastBlurTime) / 1000, 60);
                timeAway += awayDuration;
                lastBlurTime = 0;
                saveCurrentAnswer(currentQuestionIndex);
              }
            });

            function debounceMouseMove() {
              const now = Date.now();
              if (now - lastMouseMoveTime >= debounceDelay) {
                lastMouseMoveTime = now;
                lastActivityTime = now;
                activityCount++;
              }
            }

            function debounceKeydown() {
              const now = Date.now();
              if (now - lastKeydownTime >= debounceDelay) {
                lastKeydownTime = now;
                lastActivityTime = now;
                activityCount++;
              }
            }

            document.addEventListener('mousemove', debounceMouseMove);
            document.addEventListener('keydown', debounceKeydown);

            document.querySelectorAll('.option-box:not(.draggable)').forEach(box => {
              box.style.pointerEvents = 'auto';
              box.style.cursor = 'pointer';

              box.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();

                const questionType = '${q.type}';
                const option = box.getAttribute('data-value');

                if (['truefalse', 'multiple', 'singlechoice'].includes(questionType)) {
                  if (questionType === 'multiple') {
                    const idx = selectedOptions.indexOf(option);
                    if (idx === -1) {
                      selectedOptions.push(option);
                      box.classList.add('selected');
                    } else {
                      selectedOptions.splice(idx, 1);
                      box.classList.remove('selected');
                    }
                  } else {
                    if (selectedOptions[0] === option) {
                      // повторний клік — залишаємо
                    } else {
                      document.querySelectorAll('.option-box:not(.draggable)').forEach(b => {
                        b.classList.remove('selected');
                      });
                      selectedOptions = [option];
                      box.classList.add('selected');
                    }
                  }
                }
              });
            });

            const sortable = document.getElementById('sortable-options');
            if (sortable) {
              new Sortable(sortable, { animation: 150 });
            }

            updateGlobalTimer();
            if (isQuickTest) {
              const questionTimerInterval = setInterval(() => {
                const now = Date.now();
                const elapsed = Math.floor((now - questionStartTime) / 1000);
                questionTimeRemaining = Math.max(0, timePerQuestion - elapsed);
                const timerText = document.getElementById('timer-text');
                const timerCircle = document.querySelector('#question-timer .timer-circle');
                if (timerText && timerCircle) {
                  timerText.textContent = Math.round(questionTimeRemaining);
                  const circumference = 251;
                  const offset = (1 - questionTimeRemaining / timePerQuestion) * circumference;
                  timerCircle.style.strokeDashoffset = offset;
                }
                if (questionTimeRemaining <= 0 && currentQuestionIndex < totalQuestions - 1 && !hasMovedToNext) {
                  hasMovedToNext = true;
                  saveCurrentAnswer(currentQuestionIndex).then(() => {
                    saveAndNext(currentQuestionIndex);
                  });
                }
              }, 50);
            }

            document.addEventListener('visibilitychange', () => {
              if (!document.hidden) {
                updateGlobalTimer();
                if (isQuickTest) updateQuestionTimer();
              }
            });
          </script>
        </body>
      </html>
    `;

    res.send(html);
  } catch (error) {
    logger.error('Помилка в /test/question', { message: error.message, stack: error.stack });
    res.status(500).send('Внутрішня помилка сервера. Спробуйте ще раз або зверніться до адміністратора.');
  } finally {
    logger.info('Маршрут /test/question виконано', { 
      duration: (Date.now() - startTime) + ' мс' 
    });
  }
});

// Маршрут для оновлення часу початку питання
app.post('/set-question-start-time', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      return res.status(400).json({ success: false, error: 'Тест не розпочато' });
    }
    const index = parseInt(req.query.index);
    if (index >= 0 && index < userTest.questions.length) {
      const questionStartTime = userTest.questionStartTime || {};
      questionStartTime[index] = Date.now();
      await db.collection('active_tests').updateOne(
        { user: req.user },
        { $set: { questionStartTime } }
      );
      res.json({ success: true });
    } else {
      res.status(400).json({ success: false, error: 'Невірний номер питання' });
    }
  } catch (error) {
    logger.error('Помилка в /set-question-start-time', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /set-question-start-time виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для збереження відповіді
app.post('/answer', checkAuth, express.urlencoded({ extended: true }), async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');

    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;
    if (!index || answer === undefined) {
      logger.error('Відсутні параметри в /answer', { index, answer });
      return res.status(400).json({ success: false, error: 'Необхідно index та answer' });
    }

    let parsedAnswer;
    try {
      if (typeof answer === 'string') {
        if (answer.trim() === '') {
          parsedAnswer = [];
        } else {
          parsedAnswer = JSON.parse(answer);
          logger.info('[ANSWER] Успішно розпарсено рядок', { index, parsedType: typeof parsedAnswer, length: parsedAnswer.length });
        }
      } else if (Array.isArray(answer)) {
        parsedAnswer = answer;
      } else {
        parsedAnswer = [];
      }

      // Додаткове виправлення для matching (якщо прийшли об’єкти)
      const userTest = await db.collection('active_tests').findOne({ user: req.user });
      if (userTest?.questions?.[index]?.type === 'matching') {
        if (Array.isArray(parsedAnswer) && parsedAnswer.length > 0 && typeof parsedAnswer[0] === 'object' && !Array.isArray(parsedAnswer[0])) {
          logger.warn('[ANSWER] Matching: конвертуємо об’єкти в масиви пар');
          parsedAnswer = parsedAnswer.map(p => [p.left || '', p.right || '']);
        }
      }
    } catch (error) {
      logger.error('[ANSWER] Помилка парсингу відповіді', { answer, error: error.message });
      parsedAnswer = [];
    }

    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      logger.warn('[ANSWER] Активний тест не знайдено');
      return res.json({ success: true }); // не падаємо, просто ігноруємо
    }

    userTest.answers[index] = parsedAnswer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = Math.max(0, parseFloat(timeAway) || 0);
    userTest.suspiciousActivity.switchCount = Math.min(Math.max(0, parseInt(switchCount) || 0), 1000);
    userTest.suspiciousActivity.responseTimes[index] = Math.max(0, parseFloat(responseTime) || 0);
    userTest.suspiciousActivity.activityCounts[index] = parseInt(activityCount) || 0;

    await db.collection('active_tests').updateOne(
      { user: req.user },
      { $set: { answers: userTest.answers, suspiciousActivity: userTest.suspiciousActivity } }
    );

    logger.info('[ANSWER] Відповідь збережена', { index, type: userTest.questions?.[index]?.type, parsedAnswerLength: parsedAnswer.length });

    res.json({ success: true });
  } catch (error) {
    logger.error('[ANSWER] Критична помилка', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    logger.info('Маршрут /answer виконано', { duration: Date.now() - startTime });
  }
});

/**
 * Обчислює точний бал за одне питання
 * Підтримує введення чисел як через кому (3,14), так і через крапку (3.14)
 */
function calculateQuestionScore(question, userAnswer) {
  let score = 0;

  // Нормалізація тексту
  const normalize = (val) => {
    if (val === null || val === undefined) return '';
    return String(val)
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .replace(/[^a-z0-9а-яіїєґ\s.,-]/gi, '');
  };

  const normalizeNumber = (val) => {
    if (val === null || val === undefined) return NaN;
    let str = String(val).trim().replace(/,/g, '.').replace(/\s+/g, '');
    return parseFloat(str);
  };

  const normalizeDecimalForCompare = (val) => {
    if (val === null || val === undefined) return '';
    let str = String(val).trim()
      .replace(/,/g, '.')
      .replace(/\s+/g, '');

    const parts = str.split('.');
    if (parts.length > 2) {
      str = parts[0] + '.' + parts.slice(1).join('');
    }
    return str;
  };

  switch (question.type) {
    case 'multiple': {
      if (!Array.isArray(userAnswer)) return 0;
      const correctSet = new Set(question.correctAnswers.map(normalize));
      const userSet = new Set(userAnswer.map(normalize));
      const partial = question.points / correctSet.size;

      correctSet.forEach(c => { if (userSet.has(c)) score += partial; });
      userSet.forEach(u => { if (!correctSet.has(u)) score -= partial; });
      break;
    }

    case 'fillblank': {
      // Твій поточний код для fillblank — залишаємо без змін
      if (!Array.isArray(userAnswer)) return 0;

      const partial = question.points / question.correctAnswers.length;
      let correctCount = 0;

      question.correctAnswers.forEach((correctRaw, i) => {
        const userVal = userAnswer[i] || '';
        if (!userVal) return;

        const user = normalize(userVal);
        const correct = normalize(correctRaw);

        let isCorrect = false;

        if (correctRaw.includes('-')) {
          const [minStr, maxStr] = correctRaw.split('-').map(s => s.trim());
          const min = normalizeNumber(minStr);
          const max = normalizeNumber(maxStr);
          const val = normalizeNumber(userVal);

          if (!isNaN(val) && !isNaN(min) && !isNaN(max)) {
            isCorrect = val >= min && val <= max;
          }
        } else {
          if (user.includes(correct) || correct.includes(user)) {
            isCorrect = true;
          } else {
            const correctLen = correct.length;
            const userLen = user.length;
            if (correctLen > 0 && userLen > 0) {
              const minLen = Math.min(correctLen, userLen);
              let similarity = 0;
              for (let j = 0; j < minLen; j++) {
                if (correct[j] === user[j]) similarity++;
              }
              if (similarity / minLen >= 0.5) isCorrect = true;
            }
          }
        }

        if (isCorrect) correctCount++;
      });

      score = (correctCount / question.correctAnswers.length) * question.points;
      break;
    }

    case 'input': {
      // Твій код для input — залишаємо
      const correct = question.correctAnswers?.[0] || '';
      let isCorrect = false;

      if (correct.includes('-')) {
        const [minStr, maxStr] = correct.split('-').map(s => s.trim());
        const min = normalizeNumber(minStr);
        const max = normalizeNumber(maxStr);
        const val = normalizeNumber(userAnswer);
        isCorrect = !isNaN(val) && !isNaN(min) && !isNaN(max) && val >= min && val <= max;
      } else {
        const userNormalized = normalizeDecimalForCompare(userAnswer);
        const correctNormalized = normalizeDecimalForCompare(correct);
        isCorrect = userNormalized === correctNormalized;
      }

      score = isCorrect ? question.points : 0;
      break;
    }

    case 'matching': {
      // Твій поточний код — залишаємо
      if (!Array.isArray(userAnswer) || userAnswer.length === 0) return 0;

      let userPairs = userAnswer;
      if (userAnswer[0] && typeof userAnswer[0] === 'object' && !Array.isArray(userAnswer[0])) {
        userPairs = userAnswer.map(p => [p.left || '', p.right || '']);
      }

      const correctPairs = question.correctPairs || [];
      const pairCount = correctPairs.length || 1;
      const partial = question.points / pairCount;

      const correctSet = new Set(
        correctPairs.map(pair => `${normalize(pair[0])}|||${normalize(pair[1])}`)
      );

      userPairs.forEach(userPair => {
        if (!Array.isArray(userPair) || userPair.length !== 2) {
          score -= partial;
          return;
        }
        const pairKey = `${normalize(userPair[0])}|||${normalize(userPair[1])}`;
        if (correctSet.has(pairKey)) {
          score += partial;
        } else {
          score -= partial;
        }
      });
      break;
    }

    case 'ordering': {
      if (!Array.isArray(userAnswer)) return 0;
      const correctOrder = question.correctAnswers.map(normalize);
      const userOrder = userAnswer.map(normalize);
      const partial = question.points / correctOrder.length;

      correctOrder.forEach((correct, i) => {
        if (userOrder[i] === correct) score += partial;
        else score -= partial;
      });
      break;
    }

    // ====================== ВИПРАВЛЕНИЙ БЛОК ======================
    case 'singlechoice':
    case 'truefalse': {
      if (!userAnswer) return 0;

      // Приймаємо як рядок, так і масив з одним елементом
      const userVal = Array.isArray(userAnswer) ? userAnswer[0] : userAnswer;
      const correctVal = question.correctAnswer || (question.correctAnswers && question.correctAnswers[0]);

      const isCorrect = normalize(userVal) === normalize(correctVal);
      score = isCorrect ? question.points : 0;
      break;
    }

    default:
      score = 0;
  }

  return Math.max(0, score);
}

// Маршрут для відображення результатів тесту — ФІНАЛЬНЕ ВИПРАВЛЕННЯ (-3 години)
app.get('/result', checkAuth, async (req, res) => {
  const routeStartTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');

    logger.info('[RESULT] Початок обробки результату', { user: req.user });

    // 1. Знаходимо дані тесту
    let userTest = await db.collection('active_tests').findOne({ user: req.user });
    let testData;
    let dataSource = 'невідомо';

    if (userTest) {
      testData = userTest;
      dataSource = 'active_tests';
      logger.info('[RESULT] Знайдено активний тест', { testSessionId: userTest.testSessionId });
    } else {
      const recentResult = await db.collection('test_results').findOne(
        { user: req.user },
        { sort: { endTime: -1 } }
      );
      if (!recentResult) {
        logger.warn('[RESULT] Немає ні активного тесту, ні збережених результатів');
        return res.status(400).send('Тест не розпочато або перерваний. Розпочніть новий тест.');
      }
      testData = recentResult;
      dataSource = 'test_results (останній)';
    }

    // 2. Витягуємо поля
    const testNumber     = testData.testNumber;
    const answers        = testData.answers || {};
    let   startTimeMs    = testData.startTime || Date.now();
    const timeLimit      = testData.timeLimit || 3600000;
    const suspiciousActivity = testData.suspiciousActivity || {};
    let   variant        = testData.variant || '';
    const testSessionId  = testData.testSessionId || `fallback_${req.user}_${Date.now()}`;

    if (!testNumber) {
      logger.error('[RESULT] testNumber відсутній', { dataSource });
      return res.status(500).send('Помилка: не вдалося визначити номер тесту');
    }

    // Жорстке виправлення startTimeMs (якщо він прийшов в локальному часі)
    if (typeof startTimeMs === 'string') {
      startTimeMs = new Date(startTimeMs).getTime();
    }
    if (typeof startTimeMs !== 'number' || isNaN(startTimeMs)) {
      startTimeMs = Date.now() - timeLimit;
      logger.warn('[RESULT] startTimeMs невалідний, використовуємо fallback');
    }

    // Нормалізація варіанту
    if (variant) {
      variant = String(variant).trim().toLowerCase().replace(/\s+/g, ' ');
      if (variant.startsWith('variant ')) variant = variant.replace('variant ', '');
      if (variant.startsWith('варіант ')) variant = variant.replace('варіант ', '');
    }

    logger.info('[RESULT] Основні дані', { dataSource, testNumber, variant: variant || '(немає)' });

    // 3. Питання
    let allQuestions = await db.collection('questions')
      .find({ testNumber })
      .sort({ order: 1 })
      .toArray();

    let questions = allQuestions.filter(q => {
      if (!q.variant || q.variant === '') return true;
      const qVar = String(q.variant).trim().toLowerCase().replace(/\s+/g, ' ');
      return qVar === variant || qVar.includes(variant) || variant.includes(qVar);
    });

    if (questions.length === 0 && allQuestions.length > 0) {
      questions = [...allQuestions];
    }

    // 4. Бали
    const scoresPerQuestion = questions.map((q, index) => {
      const userAnswer = answers[index];
      return calculateQuestionScore(q, userAnswer);
    });

    const score = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
    const totalPoints = questions.reduce((sum, q) => sum + (q.points || 0), 0);
    const percentage = totalPoints > 0 ? (score / totalPoints) * 100 : 0;

    const totalQuestions = questions.length;
    const correctClicks = scoresPerQuestion.filter(s => s > 0).length;

    // 5. ЧАС — найжорсткіше виправлення
    let endTimeMs = Date.now();

    if (testData.endTime) {
      let parsed = testData.endTime;
      if (typeof parsed === 'string') parsed = new Date(parsed);
      if (parsed && !isNaN(parsed.getTime())) {
        endTimeMs = parsed.getTime();
      }
    }

    // Додаємо 3 години, якщо різниця негативна (тимчасовий "костиль", поки не виправимо збереження)
    const rawDuration = (endTimeMs - startTimeMs) / 1000;
    if (rawDuration < 0) {
      logger.warn('[RESULT] Виявлено негативну тривалість, додаємо 3 години (UTC+3)');
      endTimeMs += 3 * 60 * 60 * 1000;   // +3 години
    }

    const maxEndTimeMs = startTimeMs + timeLimit;
    if (endTimeMs > maxEndTimeMs) endTimeMs = maxEndTimeMs;

    const duration = Math.round((endTimeMs - startTimeMs) / 1000);

    const timeAway = suspiciousActivity.timeAway || 0;
    const correctedTimeAway = Math.min(timeAway, duration);
    const timeAwayPercent = duration > 0 ? Math.round((correctedTimeAway / duration) * 100) : 0;
    const switchCount = suspiciousActivity.switchCount || 0;

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;

    // 6. Збереження (якщо потрібно)
    const existingResult = await db.collection('test_results').findOne({ testSessionId });

    if (!existingResult) {
      logger.info('[RESULT] Зберігаємо новий результат', { testSessionId });

      if (userTest && !testData.isSaved) {
        await db.collection('active_tests').updateOne(
          { user: req.user },
          { $set: { isSavingResult: true } }
        );
      }

      await saveResult(
        req.user, testNumber, score, totalPoints, startTimeMs, endTimeMs,
        Object.keys(answers).length, correctClicks, totalQuestions, percentage,
        { timeAway: correctedTimeAway, switchCount, responseTimes: suspiciousActivity.responseTimes || [], activityCounts: suspiciousActivity.activityCounts || [] },
        answers, scoresPerQuestion, variant, ipAddress, testSessionId
      );
    }

    if (userTest) {
      await db.collection('active_tests').deleteOne({ user: req.user });
    }

    // 8. Форматування
    const endDateTime = new Date(endTimeMs);
    const formattedTime = endDateTime.toLocaleTimeString('uk-UA', { hour12: false });
    const formattedDate = endDateTime.toLocaleDateString('uk-UA');

    // 9. Зображення A.png
    const imagePath = path.join(__dirname, 'public', 'images', 'A.png');
    let imageBase64 = '';
    try {
      const imageBuffer = fs.readFileSync(imagePath);
      imageBase64 = imageBuffer.toString('base64');
    } catch (error) {
      logger.error('[RESULT] Помилка читання A.png', { message: error.message });
    }

    // 10. Повний HTML (без жодного скорочення)
    const resultHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Результати ${testNames[testNumber]?.name?.replace(/"/g, '\\"') || 'Тест'}</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 30px 20px;
              background-color: #f5f5f5;
              margin: 0;
            }
            h1 { color: #333; margin-bottom: 30px; font-size: 32px; }
            .result-section { margin: 0 auto 40px; max-width: 320px; }
            .result-container {
              position: relative;
              width: 180px;
              height: 180px;
              margin: 0 auto;
            }
            .result-circle-bg { stroke: #e0e0e0; stroke-width: 12; fill: none; }
            .result-circle { stroke: #4CAF50; stroke-width: 12; fill: none; stroke-dasharray: 530; stroke-dashoffset: 530; animation: fillCircle 1.8s ease-out forwards; }
            .result-text {
              font-size: 48px;
              font-weight: bold;
              fill: #333;
              pointer-events: none;
              text-anchor: middle;
              dominant-baseline: central;
              alignment-baseline: central;
              width: 100%;
              text-align: center;
              line-height: 1;
            }
            .progress-circles {
              display: flex;
              flex-wrap: wrap;
              justify-content: center;
              gap: 8px;
              max-width: 95vw;
              margin: 0 auto 40px;
              padding: 0 10px;
            }
            .progress-circle {
              width: 38px;
              height: 38px;
              border-radius: 50%;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 14px;
              font-weight: bold;
              color: white;
              box-shadow: 0 2px 6px rgba(0,0,0,0.15);
              min-width: 38px;
              flex-shrink: 0;
            }
            .correct   { background: #28a745; }
            .wrong     { background: #dc3545; }
            .partial   { background: #ffc107; }
            .summary-text {
              font-size: 20px;
              line-height: 1.6;
              margin-bottom: 40px;
              color: #444;
            }
            .buttons {
              margin-top: 30px;
            }
            button {
              padding: 14px 32px;
              margin: 10px;
              font-size: 18px;
              cursor: pointer;
              border: none;
              border-radius: 8px;
              min-width: 180px;
              transition: all 0.2s;
            }
            button:hover {
              transform: translateY(-2px);
              box-shadow: 0 6px 14px rgba(0,0,0,0.2);
            }
            #exportPDF { background: #ffeb3b; color: #333; }
            #restart   { background: #ef5350; color: white; }

            @keyframes fillCircle {
              to { stroke-dashoffset: ${(530 * (100 - percentage)) / 100}; }
            }

            @media (max-width: 600px) {
              h1 { font-size: 26px; }
              .result-container { width: 140px; height: 140px; }
              .result-text { font-size: 38px; }
              .progress-circle { width: 28px; height: 28px; font-size: 11px; min-width: 28px; }
              .progress-circles { gap: 6px; }
              button { padding: 12px 24px; font-size: 16px; min-width: 140px; }
            }
          </style>
          <script src="/pdfmake/pdfmake.min.js"></script>
          <script src="/pdfmake/vfs_fonts.js"></script>
        </head>
        <body>
          <h1>Результат тесту</h1>

          <div class="result-section">
            <div class="result-container">
              <svg width="100%" height="100%" viewBox="0 0 180 180" preserveAspectRatio="xMidYMid meet">
                <circle class="result-circle-bg" cx="90" cy="90" r="78" />
                <circle class="result-circle" cx="90" cy="90" r="78" />
                <text x="90" y="92" class="result-text" text-anchor="middle" dominant-baseline="middle" alignment-baseline="middle">
                  ${Math.round(percentage)}%
                </text>
              </svg>
            </div>
          </div>

          <div class="progress-circles">
            ${scoresPerQuestion.map((s, i) => {
              let colorClass = 'wrong';
              if (s === questions[i]?.points) colorClass = 'correct';
              else if (s > 0) colorClass = 'partial';
              return `<div class="progress-circle ${colorClass}">${i + 1}</div>`;
            }).join('')}
          </div>

          <div class="summary-text">
            Кількість питань: ${totalQuestions}<br>
            Правильних відповідей: ${correctClicks}<br>
            Набрано балів: ${Math.round(score)}<br>
            Максимально можлива кількість балів: ${Math.round(totalPoints)}<br>
          </div>

          <div class="buttons">
            <button id="exportPDF">Експортувати в PDF</button>
            <button id="restart">Вихід</button>
          </div>

          <script>
            const user = "${req.user.replace(/"/g, '\\"')}";
            const testName = "${testNames[testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}";
            const totalQuestionsVal = ${totalQuestions};
            const correctClicksVal = ${correctClicks};
            const scoreVal = ${Math.round(score)};
            const totalPointsVal = ${Math.round(totalPoints)};
            const percentageVal = ${Math.round(percentage)};
            const timeVal = "${formattedTime.replace(/"/g, '\\"')}";
            const dateVal = "${formattedDate.replace(/"/g, '\\"')}";
            const imageBase64Val = "${imageBase64.replace(/"/g, '\\"')}";

            document.addEventListener('DOMContentLoaded', () => {
              const exportBtn = document.getElementById('exportPDF');
              const restartBtn = document.getElementById('restart');

              if (exportBtn) {
                exportBtn.addEventListener('click', () => {
                  if (typeof pdfMake === 'undefined' || !pdfMake.createPdf) {
                    alert('PDF-генератор не завантажився. Оновіть сторінку або перевірте інтернет.');
                    return;
                  }

                  const docDefinition = {
                    content: [
                      imageBase64Val ? {
                        image: 'data:image/png;base64,' + imageBase64Val,
                        width: 50,
                        alignment: 'center',
                        margin: [0, 0, 0, 20]
                      } : { text: 'Логотип відсутній', alignment: 'center', margin: [0, 0, 0, 20] },
                      { text: 'Результат тесту користувача ' + user + ' з тесту ' + testName + ' складає ' + percentageVal + '%', style: 'header' },
                      { text: 'Кількість питань: ' + totalQuestionsVal, lineHeight: 2 },
                      { text: 'Правильних відповідей: ' + correctClicksVal, lineHeight: 2 },
                      { text: 'Набрано балів: ' + scoreVal, lineHeight: 2 },
                      { text: 'Максимально можлива кількість балів: ' + totalPointsVal, lineHeight: 2 },
                      {
                        columns: [
                          { text: 'Час: ' + timeVal, width: '50%', lineHeight: 2 },
                          { text: 'Дата: ' + dateVal, width: '50%', alignment: 'right', lineHeight: 2 }
                        ],
                        margin: [0, 10, 0, 0]
                      }
                    ],
                    styles: {
                      header: { fontSize: 14, bold: true, margin: [0, 0, 0, 10], lineHeight: 2 }
                    }
                  };

                  pdfMake.createPdf(docDefinition).download(user + '_результат.pdf');                  
                });
              }

              if (restartBtn) {
                restartBtn.addEventListener('click', () => {
                  window.location.href = '/select-test';
                });
              }
            });
          </script>
        </body>
      </html>
    `;

    res.send(resultHtml);

  } catch (error) {
    logger.error('[RESULT] Помилка', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні результатів');
  } finally {
    logger.info('Маршрут /result виконано', { duration: Date.now() - routeStartTime });
  }
});

// Маршрут для перегляду результатів користувача (з таблицею питань)
app.get('/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');

    let userTest = await db.collection('active_tests').findOne({ user: req.user });
    let testData;

    if (!userTest) {
      const recentResult = await db.collection('test_results').findOne(
        { user: req.user },
        { sort: { endTime: -1 } }
      );
      if (!recentResult) {
        return res.send('<h1>Немає завершених тестів</h1><a href="/select-test">Повернутися</a>');
      }
      testData = recentResult;
    } else {
      testData = userTest;
    }

    const { questions: rawQuestions, testNumber, answers, startTime: testStartTime, suspiciousActivity, variant, testSessionId, timeLimit } = testData;

    let questions = rawQuestions.filter(q => 
      !q.variant || q.variant === '' || q.variant === variant
    );

    const scoresPerQuestion = questions.map((q, displayIndex) => {
      const userAnswer = answers[displayIndex];
      return calculateQuestionScore(q, userAnswer);
    });

    const exactScore = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
    const roundedScore = Math.round(exactScore * 10) / 10;
    const totalPoints = questions.reduce((sum, q) => sum + q.points, 0);
    const percentage = totalPoints > 0 ? (exactScore / totalPoints) * 100 : 0;
    const roundedPercentage = Math.round(percentage * 10) / 10;

    const totalQuestions = questions.length;
    const correctClicks = scoresPerQuestion.filter(s => s > 0).length;

    let endTime = testData.endTime ? new Date(testData.endTime).getTime() : Date.now();
    const maxEndTime = testStartTime + timeLimit;
    if (endTime > maxEndTime) endTime = maxEndTime;

    const formattedTime = new Date(endTime).toLocaleTimeString('uk-UA', { hour12: false });
    const formattedDate = new Date(endTime).toLocaleDateString('uk-UA');

    let resultsHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Ваші результати</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 30px 20px; background: #f5f5f5; }
            .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
            h1 { text-align: center; color: #333; }
            table { border-collapse: collapse; width: 100%; margin: 20px 0; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
            th { background: #f2f2f2; }
            .summary { font-size: 20px; margin: 20px 0 40px; padding: 20px; background: #f8f9fa; border-radius: 8px; }
            .buttons { text-align: center; margin-top: 30px; }
            button { padding: 14px 28px; margin: 10px; border: none; border-radius: 8px; font-size: 18px; cursor: pointer; }
            #exportPDF { background: #ffeb3b; color: #333; }
            #exportPDF:hover { background: #ffe082; }
            #restart { background: #ef5350; color: white; }
            #restart:hover { background: #e53935; }
          </style>
          <script src="/pdfmake/pdfmake.min.js"></script>
          <script src="/pdfmake/vfs_fonts.js"></script>
        </head>
        <body>
          <div class="container">
            <h1>Ваші результати</h1>

            <div class="summary">
              <strong>Тест:</strong> ${testNames[testNumber]?.name?.replace(/"/g, '\\"') || 'Тест'}<br>
              <strong>Варіант:</strong> ${variant || 'Немає'}<br>
              <strong>Бали:</strong> ${roundedScore.toFixed(1)} з ${totalPoints}<br>
              <strong>Відсоток:</strong> ${roundedPercentage.toFixed(1)}%<br>
              <strong>Питань:</strong> ${totalQuestions}<br>
              <strong>Правильних:</strong> ${correctClicks}
            </div>

            <table>
              <tr>
                <th>Питання</th>
                <th>Ваша відповідь</th>
                <th>Правильна відповідь</th>
                <th>Бали</th>
              </tr>
    `;

    questions.forEach((q, index) => {
      const userAnswer = answers[index] || 'Не відповіли';
      const questionScore = scoresPerQuestion[index];

      let userAnswerDisplay = userAnswer;
      let correctAnswerDisplay = q.correctAnswers?.join(', ') || q.correctAnswer || '—';

      if (Array.isArray(userAnswer)) {
        userAnswerDisplay = userAnswer.join(', ');
      }

      resultsHtml += `
        <tr>
          <td>${q.text}</td>
          <td>${userAnswerDisplay}</td>
          <td>${correctAnswerDisplay}</td>
          <td>${questionScore.toFixed(3)} / ${q.points}</td>
        </tr>
      `;
    });

    resultsHtml += `
            </table>

            <div class="buttons">
              <button id="exportPDF">Експортувати в PDF</button>
              <button id="restart">Повернутися до тестів</button>
            </div>
          </div>

          <script>
            document.getElementById('exportPDF').addEventListener('click', () => {
              const docDefinition = {
                content: [
                  imageBase64 ? { image: 'data:image/png;base64,' + imageBase64, width: 60, alignment: 'center', margin: [0, 0, 0, 20] } : {},
                  { text: 'Результат тесту користувача ' + user + ' з тесту ' + testName + ' складає ' + percentage + '%', style: 'header' },
                  { text: 'Кількість питань: ' + totalQuestions, margin: [0, 10, 0, 0] },
                  { text: 'Правильних відповідей: ' + correctClicks, margin: [0, 5, 0, 0] },
                  { text: 'Набрано балів: ' + score, margin: [0, 5, 0, 0] },
                  { text: 'Максимально можлива кількість балів: ' + totalPoints, margin: [0, 5, 0, 0] },
                  { columns: [
                    { text: 'Час: ' + time, width: '50%', lineHeight: 1.8 },
                    { text: 'Дата: ' + date, width: '50%', alignment: 'right', lineHeight: 1.8 }
                  ], margin: [0, 20, 0, 0] }
                ],
                styles: {
                  header: { fontSize: 18, bold: true, margin: [0, 0, 0, 10] }
                }
              };
              pdfMake.createPdf(docDefinition).download(user + '_результат.pdf');
            });
          </script>
        </body>
      </html>
    `;

    res.send(resultsHtml);
  } catch (error) {
    logger.error('Помилка в /results', error);
    res.status(500).send('Помилка завантаження результатів');
  } finally {
    logger.info('Маршрут /results виконано');
  }
});

// Маршрут для адмін-панелі
// Маршрут для адмін-панелі
app.get('/admin', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    // Підрахунок непрочитаних повідомлень зворотного зв’язку
    const unreadFeedbackCount = await db.collection('feedback').countDocuments({ read: false });
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Адмін-панель</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              text-align: center; 
              padding: 50px; 
              font-size: 24px; 
              margin: 0; 
              background-color: #f5f5f5;
              min-height: 100vh;
              position: relative;
            }
            h1 { 
              font-size: 36px; 
              margin-bottom: 40px; 
            }
            .buttons-container {
              display: flex;
              flex-direction: column;
              align-items: center;
              gap: 20px;
              max-width: 400px;
              margin: 0 auto 100px; /* відступ знизу, щоб кнопка Вийти не перекривала */
            }
            button { 
              padding: 15px 30px; 
              margin: 10px 0; 
              font-size: 24px; 
              cursor: pointer; 
              width: 100%;
              max-width: 400px;
              border: none; 
              border-radius: 8px; 
              background-color: #4CAF50; 
              color: white; 
              position: relative;
              box-shadow: 0 4px 10px rgba(0,0,0,0.15);
              transition: all 0.2s;
            }
            button:hover { 
              background-color: #45a049; 
              transform: translateY(-2px);
              box-shadow: 0 6px 14px rgba(0,0,0,0.2);
            }
            #feedback-btn {
              background-color: ${unreadFeedbackCount > 0 ? '#ef5350' : '#4CAF50'};
            }
            #feedback-btn:hover {
              background-color: ${unreadFeedbackCount > 0 ? '#d32f2f' : '#45a049'};
            }
            .notification-badge {
              position: absolute;
              top: -10px;
              right: -10px;
              background-color: #ff9800;
              color: white;
              border-radius: 50%;
              padding: 5px 10px;
              font-size: 14px;
              min-width: 24px;
              height: 24px;
              display: flex;
              align-items: center;
              justify-content: center;
            }
            /* Кнопка Вийти — фіксована внизу по центру */
            #logout {
              position: fixed;
              bottom: 30px;
              left: 50%;
              transform: translateX(-50%);
              background: #ef5350;
              color: white;
              padding: 16px 40px;
              font-size: 20px;
              font-weight: bold;
              border: none;
              border-radius: 10px;
              cursor: pointer;
              min-height: 70px;
              min-width: 220px;
              box-shadow: 0 6px 15px rgba(0,0,0,0.2);
              z-index: 1000;
              transition: all 0.2s;
            }
            #logout:hover {
              background: #d32f2f;
              transform: translateX(-50%) translateY(-3px);
              box-shadow: 0 8px 20px rgba(0,0,0,0.25);
            }
            @media (max-width: 600px) {
              h1 { font-size: 28px; }
              button { font-size: 20px; padding: 14px 20px; }
              #logout { 
                padding: 14px 30px; 
                font-size: 18px; 
                min-height: 60px; 
                min-width: 180px; 
              }
            }
          </style>
        </head>
        <body>
          <h1>Адмін-панель</h1>
          <div class="buttons-container">
            <button onclick="window.location.href='/admin/users'">Керування користувачами</button>
            <button onclick="window.location.href='/admin/questions'">Керування питаннями</button>
            <button onclick="window.location.href='/admin/import-users'">Імпорт користувачів</button>
            <button onclick="window.location.href='/admin/import-questions'">Імпорт питань</button>
            <button onclick="window.location.href='/admin/results'">Перегляд результатів</button>
            <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button>
            <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button>
            <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button>
            <button id="feedback-btn" onclick="window.location.href='/admin/feedback'">
              Зворотний зв’язок
              ${unreadFeedbackCount > 0 ? `<span class="notification-badge">${unreadFeedbackCount}</span>` : ''}
            </button>
          </div>

          <!-- Кнопка Вийти — фіксована внизу по центру -->
          <button id="logout" onclick="logout()">Вийти</button>

          <script>
            async function logout() {
              console.log('Спроба виходу з адмін-панелі');
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                console.log('Статус відповіді:', response.status);
                if (!response.ok) {
                  throw new Error('HTTP помилка: ' + response.status);
                }
                const result = await response.json();
                console.log('Відповідь сервера:', result);
                if (result.success) {
                  window.location.href = '/';
                } else {
                  alert('Вихід не вдався: ' + (result.message || 'невідома помилка'));
                }
              } catch (error) {
                console.error('Помилка при виході:', error);
                alert('Не вдалося вийти. Перевірте консоль (F12) для деталей.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні адмін-панелі');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для керування користувачами
app.get('/admin/users', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const sortBy = req.query.sortBy || 'asc'; // 'asc' або 'desc'
    const search = req.query.search || '';

    let users = [];
    let errorMessage = '';

    try {
      const query = search ? { username: { $regex: search, $options: 'i' } } : {};
      users = await db.collection('users')
        .find(query)
        .sort({ username: sortBy === 'asc' ? 1 : -1 })
        .toArray();
      await CacheManager.invalidateCache('users', null);
    } catch (error) {
      logger.error('Помилка отримання користувачів із MongoDB', { message: error.message, stack: error.stack });
      errorMessage = `Помилка MongoDB: ${error.message}`;
    }

    let adminHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Керування користувачами</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .nav-btn, .action-btn, .sort-btn, .search-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .sort-btn { background-color: #6c757d; color: white; }
            .search-btn { background-color: #28a745; color: white; }
            input[type="text"] { padding: 8px; margin: 5px; width: 200px; }
          </style>
        </head>
        <body>
          <h1>Керування користувачами</h1>
          <div>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="nav-btn" onclick="window.location.href='/admin/add-user'">Додати користувача</button>
            <button class="sort-btn" onclick="window.location.href='/admin/users?sortBy=${sortBy === 'asc' ? 'desc' : 'asc'}&search=${encodeURIComponent(search)}'">Сортувати за алфавітом (${sortBy === 'asc' ? 'А-Я' : 'Я-А'})</button>
          </div>
          <div>
            <form id="search-form">
              <input type="text" id="search" name="search" placeholder="Пошук за логіном" value="${search}">
              <button type="submit" class="search-btn">Пошук</button>
            </form>
          </div>
    `;
    if (errorMessage) {
      adminHtml += `<p class="error">${errorMessage}</p>`;
    }
    adminHtml += `
          <table>
            <tr>
              <th>Ім'я користувача</th>
              <th>Дії</th>
            </tr>
    `;
    if (!users || users.length === 0) {
      adminHtml += '<tr><td colspan="2">Немає користувачів</td></tr>';
    } else {
      users.forEach(user => {
        adminHtml += `
          <tr>
            <td>${user.username}</td>
            <td>
              <button class="action-btn edit" onclick="window.location.href='/admin/edit-user?username=${user.username}'">Редагувати</button>
              <button class="action-btn delete" onclick="deleteUser('${user.username}')">Видалити</button>
            </td>
          </tr>
        `;
      });
    }
    adminHtml += `
          </table>
          <script>
            async function deleteUser(username) {
              if (confirm('Ви впевнені, що хочете видалити користувача ' + username + '?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('username', username);
                  formData.append('_csrf', '${res.locals._csrf}');
                  const response = await fetch('/admin/delete-user', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP-помилка! статус: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні користувача: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення користувача:', error);
                  alert('Не вдалося видалити користувача. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }

            document.getElementById('search-form').addEventListener('submit', (e) => {
              e.preventDefault();
              const search = document.getElementById('search').value;
              window.location.href = '/admin/users?sortBy=${sortBy}&search=' + encodeURIComponent(search);
            });
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml);
  } finally {
    logger.info('Маршрут /admin/users виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для додавання нового користувача
app.get('/admin/add-user', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Додати користувача</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { padding: 5px; width: 300px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Додати користувача</h1>
          <form method="POST" action="/admin/add-user" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <label for="username">Користувач:</label>
            <input type="text" id="username" name="username" required>
            <label for="password">Пароль:</label>
            <input type="text" id="password" name="password" required>
            <button type="submit" class="submit-btn">Додати</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
          <script>
            function validateForm() {
              const username = document.getElementById('username').value;
              const password = document.getElementById('password').value;
              const errorMessage = document.getElementById('error-message');
              if (username.length < 3 || username.length > 50) {
                errorMessage.textContent = 'Ім’я користувача має бути від 3 до 50 символів';
                return false;
              }
              if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
                errorMessage.textContent = 'Ім’я користувача може містити лише літери та цифри';
                return false;
              }
              if (password.length < 6 || password.length > 100) {
                errorMessage.textContent = 'Пароль має бути від 6 до 100 символів';
                return false;
              }
              return true;
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-user виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки додавання користувача
app.post('/admin/add-user', checkAuth, checkAdmin, [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Ім’я користувача має бути від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Ім’я користувача може містити лише літери та цифри'),
  body('password')
    .isLength({ min: 6, max: 100 }).withMessage('Пароль має бути від 6 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).send(errors.array()[0].msg);
    }

    const { username, password } = req.body;
    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser) {
      return res.status(400).send('Користувач із таким ім’ям уже існує');
    }
    const saltRounds = 10;
    const hashedPassword = await bcrypt.hash(password, saltRounds);
    const newUser = { username, password: hashedPassword, role: username === 'Instructor' ? 'instructor' : 'user' };
    await db.collection('users').insertOne(newUser);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш користувачів оновлено після додавання нового користувача');
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувача додано</title>
        </head>
        <body>
          <h1>Користувача ${username} успішно додано</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка додавання користувача', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні користувача');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-user (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для редагування користувача
app.get('/admin/edit-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username } = req.query;
    const user = await db.collection('users').findOne({ username });
    if (!user) {
      return res.status(404).send('Користувача не знайдено');
    }
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Редагувати користувача</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { padding: 5px; width: 300px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Редагувати користувача: ${username}</h1>
          <form method="POST" action="/admin/edit-user" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="hidden" name="oldUsername" value="${username}">
            <label for="username">Нове ім'я користувача:</label>
            <input type="text" id="username" name="username" value="${username}" required>
            <label for="password">Новий пароль (залиште порожнім, щоб не змінювати):</label>
            <input type="text" id="password" name="password" placeholder="Введіть новий пароль">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
          <script>
            function validateForm() {
              const username = document.getElementById('username').value;
              const password = document.getElementById('password').value;
              const errorMessage = document.getElementById('error-message');
              if (username.length < 3 || username.length > 50) {
                errorMessage.textContent = 'Ім’я користувача має бути від 3 до 50 символів';
                return false;
              }
              if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
                errorMessage.textContent = 'Ім’я користувача може містити лише літери та цифри';
                return false;
              }
              if (password && (password.length < 6 || password.length > 100)) {
                errorMessage.textContent = 'Пароль має бути від 6 до 100 символів';
                return false;
              }
              return true;
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-user виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки редагування користувача
app.post('/admin/edit-user', checkAuth, checkAdmin, [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Ім’я користувача має бути від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Ім’я користувача може містити лише літери та цифри'),
  body('password')
    .optional({ checkFalsy: true })
    .isLength({ min: 6, max: 100 }).withMessage('Пароль має бути від 6 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації в /admin/edit-user', { errors: errors.array() });
      return res.status(400).send(errors.array()[0].msg);
    }

    const { oldUsername, username, password } = req.body;
    logger.info('Отримано дані для оновлення користувача', { oldUsername, username, passwordProvided: !!password });

    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser && username !== oldUsername) {
      logger.warn('Ім’я користувача вже існує', { username });
      return res.status(400).send('Користувач із таким ім’ям уже існує');
    }

    const updateData = { username };
    if (password) {
      const saltRounds = 10;
      const hashedPassword = await bcrypt.hash(password, saltRounds);
      updateData.password = hashedPassword;
      logger.info('Пароль оновлено для користувача', { username });
    } else {
      logger.info('Пароль не надано, пропускаємо оновлення пароля', { username });
    }

    if (username === 'Instructor') {
      updateData.role = 'instructor';
    } else if (username === 'admin') {
      updateData.role = 'admin';
    } else {
      updateData.role = 'user';
    }

    const updateResult = await db.collection('users').updateOne(
      { username: oldUsername },
      { $set: updateData }
    );
    logger.info('Результат оновлення', { matchedCount: updateResult.matchedCount, modifiedCount: updateResult.modifiedCount });

    if (updateResult.matchedCount === 0) {
      logger.error('Не знайдено користувача для оновлення', { oldUsername });
      return res.status(404).send('Користувача не знайдено');
    }

    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш користувачів оновлено після редагування');

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувача оновлено</title>
        </head>
        <body>
          <h1>Користувача ${username} успішно оновлено</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка редагування користувача', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні користувача');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-user (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для видалення користувача
app.post('/admin/delete-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username } = req.body;
    await db.collection('users').deleteOne({ username });
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш користувачів оновлено після видалення');
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення користувача', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні користувача' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-user виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для керування питаннями
app.get('/admin/questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const sortBy = req.query.sortBy || 'order';
    const testNumber = req.query.testNumber || '';
    const limit = 50;
    const skip = (page - 1) * limit;

    let questions = [];
    let errorMessage = '';
    let totalQuestions = 0;
    let totalPages = 0;

    try {
      const query = testNumber ? { testNumber } : {};
      totalQuestions = await db.collection('questions').countDocuments(query);
      totalPages = Math.ceil(totalQuestions / limit);

      if (sortBy === 'testName') {
        questions = await db.collection('questions')
          .find(query)
          .skip(skip)
          .limit(limit)
          .toArray();

        questions.sort((a, b) => {
          const testNameA = testNames[a.testNumber]?.name || '';
          const testNameB = testNames[b.testNumber]?.name || '';
          return testNameA.localeCompare(testNameB, 'uk');
        });
      } else {
        questions = await db.collection('questions')
          .find(query)
          .sort({ order: 1 })
          .skip(skip)
          .limit(limit)
          .toArray();
      }

      await CacheManager.invalidateCache('questions', null);
    } catch (error) {
      logger.error('Помилка отримання питань із MongoDB', { message: error.message, stack: error.stack });
      errorMessage = `Помилка MongoDB: ${error.message}`;
    }

    let adminHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Керування питаннями</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .nav-btn, .action-btn, .sort-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .sort-btn { background-color: #6c757d; color: white; }
            select { padding: 8px; margin: 5px; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
          </style>
        </head>
        <body>
          <h1>Керування питаннями</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <button class="nav-btn" onclick="window.location.href='/admin/add-question'">Додати питання</button>
          <div>
            <form id="filter-form">
              <label for="testNumber">Фільтр за тестом:</label>
              <select id="testNumber" name="testNumber" onchange="this.form.submit()">
                <option value="">Усі тести</option>
                ${Object.keys(testNames).map(num => `<option value="${num}" ${num === testNumber ? 'selected' : ''}>${testNames[num].name.replace(/"/g, '\\"')}</option>`).join('')}
              </select>
            </form>
            <button class="sort-btn" onclick="window.location.href='/admin/questions?page=${page}&sortBy=order&testNumber=${encodeURIComponent(testNumber)}'">Сортувати за порядком</button>
            <button class="sort-btn" onclick="window.location.href='/admin/questions?page=${page}&sortBy=testName&testNumber=${encodeURIComponent(testNumber)}'">Сортувати за назвою тесту</button>
          </div>
    `;
    if (errorMessage) {
      adminHtml += `<p class="error">${errorMessage}</p>`;
    }
    adminHtml += `
          <table>
            <tr>
              <th>Тест</th>
              <th>Текст питання</th>
              <th>Тип</th>
              <th>Варіант</th>
              <th>Дії</th>
            </tr>
    `;
    if (!questions || questions.length === 0) {
      adminHtml += '<tr><td colspan="5">Немає питань</td></tr>';
    } else {
      questions.forEach(question => {
        adminHtml += `
          <tr>
            <td>${testNames[question.testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</td>
            <td>${question.text}</td>
            <td>${question.type}</td>
            <td>${question.variant || 'Немає'}</td>
            <td>
              <button class="action-btn edit" onclick="window.location.href='/admin/edit-question?id=${question._id}'">Редагувати</button>
              <button class="action-btn delete" onclick="deleteQuestion('${question._id}')">Видалити</button>
            </td>
          </tr>
        `;
      });
    }
    adminHtml += `
          </table>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/questions?page=${page - 1}&sortBy=${sortBy}&testNumber=${encodeURIComponent(testNumber)}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/questions?page=${page + 1}&sortBy=${sortBy}&testNumber=${encodeURIComponent(testNumber)}">Наступна</a>` : ''}
          </div>
          <script>
            async function deleteQuestion(id) {
              if (confirm('Ви впевнені, що хочете видалити це питання?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('id', id);
                  formData.append('_csrf', '${res.locals._csrf}');
                  const response = await fetch('/admin/delete-question', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP-помилка! статус: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні питання: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення питання:', error);
                  alert('Не вдалося видалити питання. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml);
  } finally {
    logger.info('Маршрут /admin/questions виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для додавання нового питання
app.get('/admin/add-question', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    if (!testNames || !Object.keys(testNames).length) {
      throw new Error('Список тестів недоступний');
    }

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Додати питання</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select, textarea { padding: 5px; width: 300px; margin-bottom: 10px; }
            textarea { width: 500px; height: 100px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            .note { color: blue; font-style: italic; }
          </style>
        </head>
        <body>
          <h1>Додати питання</h1>
          <form method="POST" action="/admin/add-question" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf || ''}">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}">${testNames[num].name.replace(/"/g, '\\"')}</option>`).join('')}
            </select>
            <label for="picture">Назва файлу зображення (опціонально, наприклад, Picture1.png):</label>
            <p class="note">Файл зображення має бути у папці public/images.</p>
            <input type="text" id="picture" name="picture" placeholder="Picture1.png">
            <label for="text">Текст питання:</label>
            <p class="note">Для типу Fillblank використовуйте ___ для позначення пропусків.</p>
            <textarea id="text" name="text" required placeholder="Введіть текст питання"></textarea>
            <label for="type">Тип питання:</label>
            <select id="type" name="type" required onchange="updateFormFields()">
              <option value="multiple">Multiple Choice</option>
              <option value="singlechoice">Single Choice</option>
              <option value="truefalse">True/False</option>
              <option value="input">Input</option>
              <option value="ordering">Ordering</option>
              <option value="matching">Matching</option>
              <option value="fillblank">Fill in the Blank</option>
            </select>
            <div id="options-container">
              <label for="options">Варіанти відповідей (через крапку з комою):</label>
              <textarea id="options" name="options" placeholder="Введіть варіанти через крапку з комою"></textarea>
            </div>
            <label for="correctAnswers">Правильні відповіді (через крапку з комою):</label>
            <p id="correctAnswersNote" class="note">Для типів Input і Fillblank можна вказати діапазон у форматі "число1-число2", наприклад, "12-14".</p>
            <textarea id="correctAnswers" name="correctAnswers" required placeholder="Введіть правильні відповіді через крапку з комою"></textarea>
            <label for="points">Бали за питання:</label>
            <input type="number" id="points" name="points" value="1" min="1" required>
            <label for="variant">Варіант (опціонально):</label>
            <input type="text" id="variant" name="variant" placeholder="Наприклад, Variant 1">
            <button type="submit" class="submit-btn">Додати</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
          <script>
            function updateFormFields() {
              const type = document.getElementById('type').value;
              const optionsContainer = document.getElementById('options-container');
              const correctAnswersNote = document.getElementById('correctAnswersNote');
              if (type === 'truefalse') {
                optionsContainer.style.display = 'none';
                document.getElementById('options').value = 'Правда; Неправда';
              } else if (type === 'input' || type === 'fillblank') {
                optionsContainer.style.display = 'none';
                correctAnswersNote.style.display = 'block';
              } else {
                optionsContainer.style.display = 'block';
                if (type !== 'input' && type !== 'fillblank') {
                  correctAnswersNote.style.display = 'none';
                }
              }
            }

            function validateForm() {
              const text = document.getElementById('text').value;
              const points = document.getElementById('points').value;
              const variant = document.getElementById('variant').value;
              const picture = document.getElementById('picture').value;
              const type = document.getElementById('type').value;
              const correctAnswers = document.getElementById('correctAnswers').value;
              const errorMessage = document.getElementById('error-message');

              if (text.length < 5 || text.length > 1000) {
                errorMessage.textContent = 'Текст питання має бути від 5 до 1000 символів';
                return false;
              }
              if (points < 1 || points > 100) {
                errorMessage.textContent = 'Бали мають бути числом від 1 до 100';
                return false;
              }
              if (variant && (variant.length < 1 || variant.length > 50)) {
                errorMessage.textContent = 'Варіант має бути від 1 до 50 символів';
                return false;
              }
              if (picture && !/\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Назва файлу зображення має закінчуватися на .jpeg, .jpg, .png або .gif';
                return false;
              }
              if (type === 'input' || type === 'fillblank') {
                const answersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (type === 'input' && answersArray.length !== 1) {
                  errorMessage.textContent = 'Для типу Input потрібна лише одна правильна відповідь';
                  return false;
                }
                if (type === 'fillblank') {
                  const blankCount = (text.match(/___/g) || []).length;
                  if (blankCount === 0 || blankCount !== answersArray.length) {
                    errorMessage.textContent = 'Кількість пропусків у тексті питання не відповідає кількості правильних відповідей';
                    return false;
                  }
                }
                for (let i = 0; i < answersArray.length; i++) {
                  const answer = answersArray[i];
                  if (answer.includes('-')) {
                    const [min, max] = answer.split('-').map(val => parseFloat(val.trim()));
                    if (isNaN(min) || isNaN(max) || min > max) {
                      errorMessage.textContent = \`Правильна відповідь \${i + 1} має невірний формат діапазону. Використовуйте "число1-число2", де число1 <= число2.\`;
                      return false;
                    }
                  } else {
                    const value = parseFloat(answer);
                    if (isNaN(value)) {
                      errorMessage.textContent = \`Правильна відповідь \${i + 1} для типу \${type} має бути числом або діапазоном у форматі "число1-число2".\`;
                      return false;
                    }
                  }
                }
              }
              if (type === 'singlechoice') {
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (correctAnswersArray.length !== 1) {
                  errorMessage.textContent = 'Для типу Single Choice потрібна одна правильна відповідь';
                  return false;
                }
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                if (options.length < 2) {
                  errorMessage.textContent = 'Для типу Single Choice потрібно мінімум 2 варіанти відповідей';
                  return false;
                }
              }
              if (type === 'matching') {
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean);
                if (options.length === 0 || options.length !== correctAnswersArray.length) {
                  errorMessage.textContent = 'Для типу Matching кількість варіантів має відповідати кількості правильних відповідей';
                  return false;
                }
              }
              return true;
            }

            updateFormFields();
          </script>
        </body>
      </html>
    `.trim();
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/add-question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-question виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки додавання питання
app.post('/admin/add-question', checkAuth, checkAdmin, [
  body('testNumber').notEmpty().withMessage('Номер тесту обов’язковий'),
  body('text')
    .isLength({ min: 5, max: 1000 }).withMessage('Текст питання має бути від 5 до 1000 символів'),
  body('type')
    .isIn(['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank']).withMessage('Невірний тип питання'),
  body('correctAnswers').notEmpty().withMessage('Правильні відповіді обов’язкові'),
  body('points')
    .isInt({ min: 1, max: 100 }).withMessage('Бали мають бути числом від 1 до 100'),
  body('variant')
    .optional({ checkFalsy: true })
    .isLength({ min: 1, max: 50 }).withMessage('Варіант має бути від 1 до 50 символів'),
  body('picture')
    .optional({ checkFalsy: true })
    .matches(/\.(jpeg|jpg|png|gif)$/i).withMessage('Назва файлу зображення має закінчуватися на .jpeg, .jpg, .png або .gif')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації в /admin/add-question', { errors: errors.array() });
      return res.status(400).send(errors.array()[0].msg);
    }

    const { testNumber, text, type, options, correctAnswers, points, variant, picture } = req.body;

    const normalizedPicture = picture
      ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '')
      : null;

    let questionData = {
      testNumber,
      picture: picture ? `/images/${normalizedPicture}` : null,
      originalPicture: normalizedPicture,
      text,
      type: type.toLowerCase(),
      options: options ? options.split(';').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean),
      points: Number(points),
      variant: variant || '',
      order: await db.collection('questions').countDocuments({ testNumber })
    };

    if (questionData.picture) {
      const imagePath = path.join(__dirname, 'public', questionData.picture);
      if (!fs.existsSync(imagePath)) {
        logger.warn(`Зображення не знайдено за шляхом: ${imagePath}`);
        questionData.picture = null;
      } else {
        logger.info(`Зображення знайдено: ${questionData.picture}`);
      }
    }

    if (type === 'truefalse') {
      questionData.options = ["Правда", "Неправда"];
    }

    if (type === 'matching') {
      questionData.pairs = questionData.options.map((opt, idx) => ({
        left: opt || '',
        right: questionData.correctAnswers[idx] || ''
      })).filter(pair => pair.left && pair.right);
      if (questionData.pairs.length === 0) {
        logger.warn('Для типу Matching потрібні пари', { testNumber, text });
        return res.status(400).send('Для типу Matching потрібні пари відповідей');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        logger.warn('Невідповідність між пропусками та правильними відповідями для fillblank', { blankCount, correctAnswersLength: questionData.correctAnswers.length });
        return res.status(400).send('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
      }
      questionData.blankCount = blankCount;

      questionData.correctAnswers.forEach((correctAnswer, idx) => {
        if (correctAnswer.includes('-')) {
          const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
          if (isNaN(min) || isNaN(max) || min > max) {
            return res.status(400).send(`Невірний формат діапазону для правильної відповіді ${idx + 1}. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.`);
          }
        } else {
          const value = parseFloat(correctAnswer);
          if (isNaN(value)) {
            return res.status(400).send(`Правильна відповідь ${idx + 1} для типу Fillblank має бути числом або діапазоном у форматі "число1-число2".`);
          }
        }
      });
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        logger.warn('Для типу Single Choice потрібна одна правильна відповідь та щонайменше 2 варіанти', {
          correctAnswersLength: questionData.correctAnswers.length,
          optionsLength: questionData.options.length
        });
        return res.status(400).send('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    if (type === 'input') {
      if (questionData.correctAnswers.length !== 1) {
        return res.status(400).send('Для типу Input потрібна одна правильна відповідь');
      }
      const correctAnswer = questionData.correctAnswers[0];
      if (correctAnswer.includes('-')) {
        const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
        if (isNaN(min) || isNaN(max) || min > max) {
          return res.status(400).send('Невірний формат діапазону для правильної відповіді. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.');
        }
      } else {
        const value = parseFloat(correctAnswer);
        if (isNaN(value)) {
          return res.status(400).send('Правильна відповідь для типу Input має бути числом або діапазоном у форматі "число1-число2".');
        }
      }
    }

    await db.collection('questions').insertOne(questionData);
    logger.info('Питання додано до MongoDB', { testNumber, text, type });

    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    logger.info('Кеш очищено після додавання питання', { testNumber });

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання додано</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Питання успішно додано</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка додавання питання в /admin/add-question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні питання: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-question (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для редагування питання
app.get('/admin/edit-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.query;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).send('Невірний ідентифікатор питання');
    }

    const question = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    if (!question) {
      return res.status(404).send('Питання не знайдено');
    }

    const pictureName = question.picture ? question.picture.replace('/images/', '') : '';
    const normalizedOriginalPicture = question.originalPicture
      ? question.originalPicture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '')
      : '';
    const warningMessage = question.picture === null && question.originalPicture && question.originalPicture.trim() !== ''
      ? `Попередження: зображення "${normalizedOriginalPicture}" не було знайдено під час імпорту. Перевірте, чи файл зображення є в папці public/images.`
      : '';

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редагувати питання</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select, textarea { padding: 5px; width: 300px; margin-bottom: 10px; }
            textarea { width: 500px; height: 100px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            .warning { color: orange; margin-bottom: 10px; }
            .note { color: blue; font-style: italic; }
            img#image-preview { max-width: 200px; margin-top: 10px; }
          </style>
        </head>
        <body>
          <h1>Редагувати питання</h1>
          <form method="POST" action="/admin/edit-question" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf || ''}">
            <input type="hidden" name="id" value="${id}">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}" ${num === question.testNumber ? 'selected' : ''}>${testNames[num].name.replace(/"/g, '\\"')}</option>`).join('')}
            </select>
            <label for="picture">Назва файлу зображення (опціонально, наприклад, Picture1.png):</label>
            <p class="note">Файл зображення має бути у папці public/images.</p>
            <input type="text" id="picture" name="picture" value="${pictureName}" placeholder="Picture1.png">
            ${warningMessage ? `<p class="warning">${warningMessage}</p>` : ''}
            ${pictureName ? `<img id="image-preview" src="/images/${pictureName}" alt="Зображення питання" onerror="this.onerror=null;this.src='';this.alt='Зображення недоступне';">` : ''}
            <label for="text">Текст питання:</label>
            <p class="note">Для типу Fillblank використовуйте ___ для позначення пропусків.</p>
            <textarea id="text" name="text" required placeholder="Введіть текст питання">${question.text}</textarea>
            <label for="type">Тип питання:</label>
            <select id="type" name="type" required onchange="updateFormFields()">
              <option value="multiple" ${question.type === 'multiple' ? 'selected' : ''}>Multiple Choice</option>
              <option value="singlechoice" ${question.type === 'singlechoice' ? 'selected' : ''}>Single Choice</option>
              <option value="truefalse" ${question.type === 'truefalse' ? 'selected' : ''}>True/False</option>
              <option value="input" ${question.type === 'input' ? 'selected' : ''}>Input</option>
              <option value="ordering" ${question.type === 'ordering' ? 'selected' : ''}>Ordering</option>
              <option value="matching" ${question.type === 'matching' ? 'selected' : ''}>Matching</option>
              <option value="fillblank" ${question.type === 'fillblank' ? 'selected' : ''}>Fill in the Blank</option>
            </select>
            <div id="options-container">
              <label for="options">Варіанти відповідей (через крапку з комою):</label>
              <textarea id="options" name="options" placeholder="Введіть варіанти через крапку з комою">${question.options.join('; ')}</textarea>
            </div>
            <label for="correctAnswers">Правильні відповіді (через крапку з комою):</label>
            <p id="correctAnswersNote" class="note">Для типів Input і Fillblank можна вказати діапазон у форматі "число1-число2", наприклад, "12-14".</p>
            <textarea id="correctAnswers" name="correctAnswers" required placeholder="Введіть правильні відповіді через крапку з комою">${question.correctAnswers.join('; ')}</textarea>
            <label for="points">Бали за питання:</label>
            <input type="number" id="points" name="points" value="${question.points}" min="1" required>
            <label for="variant">Варіант:</label>
            <input type="text" id="variant" name="variant" value="${question.variant}" placeholder="Наприклад, Variant 1">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
          <script>
            function updateFormFields() {
              const type = document.getElementById('type').value;
              const optionsContainer = document.getElementById('options-container');
              const correctAnswersNote = document.getElementById('correctAnswersNote');
              if (type === 'truefalse') {
                optionsContainer.style.display = 'none';
                document.getElementById('options').value = 'Правда; Неправда';
              } else if (type === 'input' || type === 'fillblank') {
                optionsContainer.style.display = 'none';
                correctAnswersNote.style.display = 'block';
              } else {
                optionsContainer.style.display = 'block';
                if (type !== 'input' && type !== 'fillblank') {
                  correctAnswersNote.style.display = 'none';
                }
              }
            }

            function validateForm() {
              const text = document.getElementById('text').value;
              const points = document.getElementById('points').value;
              const variant = document.getElementById('variant').value;
              const picture = document.getElementById('picture').value;
              const type = document.getElementById('type').value;
              const correctAnswers = document.getElementById('correctAnswers').value;
              const errorMessage = document.getElementById('error-message');

              if (text.length < 5 || text.length > 1000) {
                errorMessage.textContent = 'Текст питання має бути від 5 до 1000 символів';
                return false;
              }
              if (points < 1 || points > 100) {
                errorMessage.textContent = 'Бали мають бути числом від 1 до 100';
                return false;
              }
              if (variant && (variant.length < 1 || variant.length > 50)) {
                errorMessage.textContent = 'Варіант має бути від 1 до 50 символів';
                return false;
              }
              if (picture && !/\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Назва файлу зображення має закінчуватися на .jpeg, .jpg, .png або .gif';
                return false;
              }
              if (type === 'input' || type === 'fillblank') {
                const answersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (type === 'input' && answersArray.length !== 1) {
                  errorMessage.textContent = 'Для типу Input потрібна лише одна правильна відповідь';
                  return false;
                }
                if (type === 'fillblank') {
                  const blankCount = (text.match(/___/g) || []).length;
                  if (blankCount === 0 || blankCount !== answersArray.length) {
                    errorMessage.textContent = 'Кількість пропусків у тексті питання не відповідає кількості правильних відповідей';
                    return false;
                  }
                }
                for (let i = 0; i < answersArray.length; i++) {
                  const answer = answersArray[i];
                  if (answer.includes('-')) {
                    const [min, max] = answer.split('-').map(val => parseFloat(val.trim()));
                    if (isNaN(min) || isNaN(max) || min > max) {
                      errorMessage.textContent = \`Правильна відповідь \${i + 1} має невірний формат діапазону. Використовуйте "число1-число2", де число1 <= число2.\`;
                      return false;
                    }
                  } else {
                    const value = parseFloat(answer);
                    if (isNaN(value)) {
                      errorMessage.textContent = \`Правильна відповідь \${i + 1} для типу \${type} має бути числом або діапазоном у форматі "число1-число2".\`;
                      return false;
                    }
                  }
                }
              }
              if (type === 'singlechoice') {
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (correctAnswersArray.length !== 1) {
                  errorMessage.textContent = 'Для типу Single Choice потрібна одна правильна відповідь';
                  return false;
                }
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                if (options.length < 2) {
                  errorMessage.textContent = 'Для типу Single Choice потрібно мінімум 2 варіанти відповідей';
                  return false;
                }
              }
              if (type === 'matching') {
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean);
                if (options.length === 0 || options.length !== correctAnswersArray.length) {
                  errorMessage.textContent = 'Для типу Matching кількість варіантів має відповідати кількості правильних відповідей';
                  return false;
                }
              }
              return true;
            }

            document.getElementById('picture').addEventListener('input', (e) => {
              const pictureName = e.target.value;
              const preview = document.getElementById('image-preview');
              if (pictureName) {
                preview.src = '/images/' + pictureName;
                preview.onerror = () => {
                  preview.src = '';
                  preview.alt = 'Зображення недоступне';
                };
              } else {
                preview.src = '';
                preview.alt = '';
              }
            });

            updateFormFields();
          </script>
        </body>
      </html>
    `.trim();
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-question виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки редагування питання
app.post('/admin/edit-question', checkAuth, checkAdmin, [
  body('testNumber').notEmpty().withMessage('Номер тесту обов’язковий'),
  body('text')
    .isLength({ min: 5, max: 1000 }).withMessage('Текст питання має бути від 5 до 1000 символів'),
  body('type')
    .isIn(['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank']).withMessage('Невірний тип питання'),
  body('correctAnswers').notEmpty().withMessage('Правильні відповіді обов’язкові'),
  body('points')
    .isInt({ min: 1, max: 100 }).withMessage('Бали мають бути числом від 1 до 100'),
  body('variant')
    .optional({ checkFalsy: true })
    .isLength({ min: 1, max: 50 }).withMessage('Варіант має бути від 1 до 50 символів'),
  body('picture')
    .optional({ checkFalsy: true })
    .matches(/\.(jpeg|jpg|png|gif)$/i).withMessage('Назва файлу зображення має закінчуватися на .jpeg, .jpg, .png або .gif')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).send(errors.array()[0].msg);
    }

    const { id, testNumber, text, type, options, correctAnswers, points, variant, picture } = req.body;

    const oldQuestion = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    if (!oldQuestion) {
      return res.status(404).send('Питання не знайдено');
    }

    const normalizedPicture = picture
      ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '')
      : null;

    let questionData = {
      testNumber,
      picture: oldQuestion.picture,
      originalPicture: normalizedPicture || oldQuestion.originalPicture,
      text,
      type: type.toLowerCase(),
      options: options ? options.split(';').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean),
      points: Number(points),
      variant: variant || '',
      order: oldQuestion.order
    };

    if (picture && picture !== oldQuestion.picture?.replace('/images/', '')) {
      const imageDir = path.join(__dirname, 'public', 'images');
      const extensions = ['.png', '.jpg', '.jpeg', '.gif'];
      let found = false;

      logger.info(`Перевірка зображення для ${normalizedPicture} у ${imageDir}`);

      for (const ext of extensions) {
        const expectedFileName = `${normalizedPicture}${ext}`;
        const imagePath = path.join(imageDir, expectedFileName);
        if (fs.existsSync(imagePath)) {
          questionData.picture = `/images/${normalizedPicture}${ext.toLowerCase()}`;
          logger.info(`Зображення знайдено: ${questionData.picture}`);
          found = true;
          break;
        }
      }

      if (!found) {
        const filesInDir = fs.existsSync(imageDir) ? fs.readdirSync(imageDir) : [];
        logger.warn(`Зображення ${normalizedPicture} не знайдено в public/images під час редагування. Доступні файли: ${filesInDir.join(', ')}`);
        questionData.picture = null;
      }
    } else {
      logger.info(`Поле зображення не змінено, зберігаємо існуюче зображення: ${questionData.picture}`);
    }

    if (type === 'truefalse') {
      questionData.options = ["Правда", "Неправда"];
    }

    if (type === 'matching') {
      questionData.pairs = questionData.options.map((opt, idx) => ({
        left: opt || '',
        right: questionData.correctAnswers[idx] || ''
      })).filter(pair => pair.left && pair.right);
      if (questionData.pairs.length === 0) {
        logger.warn('Для типу Matching потрібні пари', { testNumber, text });
        return res.status(400).send('Для типу Matching потрібні пари відповідей');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        logger.warn('Невідповідність між пропусками та правильними відповідями для fillblank', { blankCount, correctAnswersLength: questionData.correctAnswers.length });
        return res.status(400).send('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
      }
      questionData.blankCount = blankCount;

      questionData.correctAnswers.forEach((correctAnswer, idx) => {
        if (correctAnswer.includes('-')) {
          const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
          if (isNaN(min) || isNaN(max) || min > max) {
            return res.status(400).send(`Невірний формат діапазону для правильної відповіді ${idx + 1}. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.`);
          }
        } else {
          const value = parseFloat(correctAnswer);
          if (isNaN(value)) {
            return res.status(400).send(`Правильна відповідь ${idx + 1} для типу Fillblank має бути числом або діапазоном у форматі "число1-число2".`);
          }
        }
      });
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        logger.warn('Для типу Single Choice потрібна одна правильна відповідь та щонайменше 2 варіанти', {
          correctAnswersLength: questionData.correctAnswers.length,
          optionsLength: questionData.options.length
        });
        return res.status(400).send('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    if (type === 'input') {
      if (questionData.correctAnswers.length !== 1) {
        return res.status(400).send('Для типу Input потрібна одна правильна відповідь');
      }
      const correctAnswer = questionData.correctAnswers[0];
      if (correctAnswer.includes('-')) {
        const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
        if (isNaN(min) || isNaN(max) || min > max) {
          return res.status(400).send('Невірний формат діапазону для правильної відповіді. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.');
        }
      } else {
        const value = parseFloat(correctAnswer);
        if (isNaN(value)) {
          return res.status(400).send('Правильна відповідь для типу Input має бути числом або діапазоном у форматі "число1-число2".');
        }
      }
    }

    await db.collection('questions').updateOne(
      { _id: new ObjectId(id) },
      { $set: questionData }
    );
    logger.info('Питання оновлено в MongoDB', { id, testNumber, text, type });

    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    logger.info('Кеш очищено після оновлення питання', { testNumber });

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання оновлено</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Питання успішно оновлено</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка оновлення питання в /admin/edit-question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні питання: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-question (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для видалення питання
app.post('/admin/delete-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.body;
    const question = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    if (!question) {
      return res.status(404).json({ success: false, message: 'Питання не знайдено' });
    }
    await db.collection('questions').deleteOne({ _id: new ObjectId(id) });
    await CacheManager.invalidateCache('questions', question.testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення питання', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні питання' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-question виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для імпорту користувачів
app.get('/admin/import-users', ensureInitialized, checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    if (!res.locals) {
      res.locals = {};
      logger.info('Ініціалізовано res.locals для /admin/import-users', { url: req.url });
    }
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Імпорт користувачів</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { padding: 5px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .submit-btn:disabled { background-color: #cccccc; cursor: not-allowed; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Імпорт користувачів із Excel</h1>
          <form id="import-form" enctype="multipart/form-data">
            <label for="file">Виберіть файл users.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn" id="submit-btn">Завантажити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            document.getElementById('import-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const fileInput = document.getElementById('file');
              const errorMessage = document.getElementById('error-message');
              const submitBtn = document.getElementById('submit-btn');
              
              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано.';
                console.error('Помилка: Файл не вибрано.');
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Завантаження...';

              const formData = new FormData();
              formData.append('file', fileInput.files[0]);

              // Отримання JWT із cookies
              const authToken = document.cookie.split('; ').find(row => row.startsWith('auth_token='))?.split('=')[1];
              console.log('auth_token:', authToken ? 'Присутній' : 'Відсутній');
              if (!authToken) {
                errorMessage.textContent = 'Токен авторизації відсутній. Увійдіть знову.';
                console.error('Помилка: auth_token відсутній у cookies:', document.cookie);
                submitBtn.disabled = false;
                submitBtn.textContent = 'Завантажити';
                return;
              }

              try {
                const response = await fetch('/admin/import-users', {
                  method: 'POST',
                  body: formData,
                  headers: {
                    'Authorization': 'Bearer ' + authToken
                  }
                });

                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'Помилка: ' + response.status);
                }

                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка імпорту:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Завантажити';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
    logger.info('Відображено форму імпорту користувачів', { url: req.url });
  } catch (error) {
    logger.error('Помилка в /admin/import-users (GET)', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Помилка завантаження форми</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-users (GET) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Обробка імпорту користувачів
app.post('/admin/import-users', ensureInitialized, checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    const token = req.headers['authorization']?.split(' ')[1] || req.cookies.token;
    logger.info('Отримано JWT для /admin/import-users', { token: token ? '[присутній]' : '[відсутній]' });

    if (!req.file) {
      logger.error('Файл не надано', { url: req.url });
      return res.status(400).send('Файл не надано');
    }

    const count = await importUsersToMongoDB(req.file.buffer);
    logger.info(`Імпортовано ${count} користувачів`);

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувачів імпортовано</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Успішно імпортовано ${count} користувачів</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка імпорту користувачів (POST)', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка імпорту</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Помилка імпорту користувачів</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-users'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-users (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для імпорту питань
app.get('/admin/import-questions', ensureInitialized, checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    if (!res.locals) {
      res.locals = {};
      logger.info('Ініціалізовано res.locals для /admin/import-questions', { url: req.url });
    }
    if (!testNames || !Object.keys(testNames).length) {
      logger.warn('Список тестів порожній, перезавантаження', { url: req.url });
      await loadTestsFromMongoDB();
      if (!Object.keys(testNames).length) {
        throw new Error('Не вдалося завантажити список тестів');
      }
    }
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Імпорт питань</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select { padding: 5px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .submit-btn:disabled { background-color: #cccccc; cursor: not-allowed; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Імпорт питань із Excel</h1>
          <form id="import-form" enctype="multipart/form-data">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}">${testNames[num].name.replace(/"/g, '\\"')}</option>`).join('')}
            </select>
            <label for="file">Виберіть файл questions.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn" id="submit-btn">Завантажити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            document.getElementById('import-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const testNumber = document.getElementById('testNumber').value;
              const fileInput = document.getElementById('file');
              const errorMessage = document.getElementById('error-message');
              const submitBtn = document.getElementById('submit-btn');

              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано.';
                console.error('Помилка: Файл не вибрано.');
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Завантаження...';

              const formData = new FormData();
              formData.append('testNumber', testNumber);
              formData.append('file', fileInput.files[0]);

              // Отримання JWT із cookies
              const authToken = document.cookie.split('; ').find(row => row.startsWith('auth_token='))?.split('=')[1];
              console.log('auth_token:', authToken ? 'Присутній' : 'Відсутній');
              if (!authToken) {
                errorMessage.textContent = 'Токен авторизації відсутній. Увійдіть знову.';
                console.error('Помилка: auth_token відсутній у cookies:', document.cookie);
                submitBtn.disabled = false;
                submitBtn.textContent = 'Завантажити';
                return;
              }

              try {
                const response = await fetch('/admin/import-questions', {
                  method: 'POST',
                  body: formData,
                  headers: {
                    'Authorization': 'Bearer ' + authToken
                  }
                });

                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'Помилка: ' + response.status);
                }

                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка імпорту:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Завантажити';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
    logger.info('Відображено форму імпорту питань', { url: req.url });
  } catch (error) {
    logger.error('Помилка в /admin/import-questions (GET)', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Помилка завантаження форми</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-questions (GET) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Обробка імпорту питань
app.post('/admin/import-questions', ensureInitialized, checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    const token = req.headers['authorization']?.split(' ')[1] || req.cookies.token;
    logger.info('Отримано JWT для /admin/import-questions', { token: token ? '[присутній]' : '[відсутній]' });

    if (!req.file) {
      logger.error('Файл не надано', { url: req.url });
      return res.status(400).send('Файл не надано');
    }
    const testNumber = req.body.testNumber;
    if (!testNumber || !testNames[testNumber]) {
      logger.error('Невірний номер тесту', { testNumber, url: req.url });
      return res.status(400).send('Невірний номер тесту');
    }

    const count = await importQuestionsToMongoDB(req.file.buffer, testNumber);
    logger.info(`Імпортовано ${count} питань для тесту ${testNumber}`);

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання імпортовано</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Успішно імпортовано ${count} питань для тесту ${testNames[testNumber].name.replace(/"/g, '\\"')}</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка імпорту питань (POST)', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка імпорту</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Помилка імпорту питань</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-questions'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-questions (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для перегляду результатів тестів (адмін/інструктор) — оновлено формат часу поза вкладкою
app.get('/admin/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
      return res.status(403).send('Доступно тільки для адміністраторів та інструкторів');
    }

    const search = req.query.search || '';
    const query = search ? { user: { $regex: search, $options: 'i' } } : {};

    const results = await db.collection('test_results')
      .find(query)
      .sort({ endTime: -1 })
      .toArray();

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Результати тестів</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 30px 20px; background: #f5f5f5; margin: 0; }
            .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
            h1 { text-align: center; color: #333; margin-bottom: 25px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; font-weight: bold; }
            .nav-btn, .action-btn, .search-btn { padding: 12px 24px; margin: 10px 5px; cursor: pointer; border: none; border-radius: 8px; font-size: 16px; }
            .nav-btn { background: #007bff; color: white; }
            .nav-btn:hover { background: #0056b3; }
            .search-btn { background: #28a745; color: white; }
            .action-btn.view { background: #4CAF50; color: white; }
            .action-btn.delete { background: #ef5350; color: white; }
            .suspicious { color: #d32f2f; font-weight: bold; background: #ffebee; }
            .normal { color: #388e3c; }
            input[type="text"] { padding: 10px; width: 250px; font-size: 16px; }
            form { display: inline-block; margin: 10px 0; }
            .activity-details { font-size: 14px; line-height: 1.5; }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Результати тестів (${results.length})</h1>

            <button class="nav-btn" onclick="window.location.href='/select-test'">Повернутися до вибору тесту</button>

            <form id="search-form">
              <input type="text" name="search" placeholder="Пошук за логіном" value="${search}">
              <button type="submit" class="search-btn">Пошук</button>
            </form>

            <table>
              <tr>
                <th>Користувач</th>
                <th>Тест</th>
                <th>Варіант</th>
                <th>Очки/%</th>
                <th>Максимум</th>
                <th>Початок</th>
                <th>Кінець</th>
                <th>Тривалість</th>
                <th>Підозріла активність</th>
                <th>Дія</th>
              </tr>
    `;

    if (results.length === 0) {
      html += '<tr><td colspan="10">Немає результатів</td></tr>';
    } else {
      for (const result of results) {
        // Перерахунок балів
        let allQuestions = await db.collection('questions')
          .find({ testNumber: result.testNumber })
          .sort({ order: 1 })
          .toArray();

        const questions = allQuestions.filter(q => 
          !q.variant || q.variant === '' || q.variant === result.variant
        );

        const scoresPerQuestion = questions.map((q, idx) => {
          const userAnswer = result.answers ? result.answers[idx] : null;
          return calculateQuestionScore(q, userAnswer);
        });

        const exactScore = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
        const roundedScore = Math.round(exactScore * 10) / 10;
        const totalPoints = questions.reduce((sum, q) => sum + (q.points || 0), 0);
        const percentage = totalPoints > 0 ? (exactScore / totalPoints) * 100 : 0;
        const roundedPercentage = Math.round(percentage * 10) / 10;

        const startTimeStr = result.startTime ? new Date(result.startTime).toLocaleString('uk-UA') : '—';
        const endTimeStr = result.endTime ? new Date(result.endTime).toLocaleString('uk-UA') : '—';

        const durationSec = result.duration || 
          (result.endTime && result.startTime ? 
            Math.round((new Date(result.endTime) - new Date(result.startTime)) / 1000) : 0);

        const minutes = Math.floor(durationSec / 60).toString().padStart(2, '0');
        const seconds = (durationSec % 60).toString().padStart(2, '0');

        // Підозріла активність
        const susp = result.suspiciousActivity || {};
        const timeAwaySeconds = susp.timeAway || 0;
        const timeAwayPercent = durationSec > 0 
          ? Math.round((timeAwaySeconds / durationSec) * 100) 
          : 0;
        const switchCount = susp.switchCount || 0;

        // === НОВЕ ФОРМАТУВАННЯ ЧАСУ ПОЗА ВКЛАДКОЮ ===
        let timeAwayFormatted = '';
        if (timeAwaySeconds < 60) {
          // менше хвилини — показуємо з десятими
          timeAwayFormatted = timeAwaySeconds.toFixed(1) + ' сек';
        } else {
          // 1 хвилина і більше
          const awayMinutes = Math.floor(timeAwaySeconds / 60);
          const awaySeconds = timeAwaySeconds % 60;
          timeAwayFormatted = `${awayMinutes} хв ${awaySeconds} сек`;
        }

        const isSuspicious = timeAwayPercent > (config?.suspiciousActivity?.timeAwayThreshold || 15) || 
                            switchCount > (config?.suspiciousActivity?.switchCountThreshold || 5);

        html += `
          <tr class="${isSuspicious ? 'suspicious' : ''}">
            <td>${result.user || '—'}</td>
            <td>${testNames[result.testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}</td>
            <td>${result.variant || 'Немає'}</td>
            <td>${roundedScore.toFixed(1)} / ${roundedPercentage.toFixed(1)}%</td>
            <td>${totalPoints.toFixed(1)}</td>
            <td>${startTimeStr}</td>
            <td>${endTimeStr}</td>
            <td>${minutes} хв ${seconds} сек</td>
            <td>
              <div class="activity-details">
                <strong>${timeAwayPercent}%</strong> поза вкладкою<br>
                Перемикань: <strong>${switchCount}</strong><br>
                Час поза: <strong>${timeAwayFormatted}</strong>
              </div>
            </td>
            <td>
              <button class="action-btn view" onclick="viewResult('${result._id}')">Перегляд</button>
              ${req.userRole === 'admin' ? `<button class="action-btn delete" onclick="deleteResult('${result._id}')">🗑️ Видалити</button>` : ''}
            </td>
          </tr>
        `;
      }
    }

    html += `
            </table>
          </div>

          <script>
            async function viewResult(id) {
              window.location.href = '/admin/view-result?id=' + id;
            }

            async function deleteResult(id) {
              if (confirm('Видалити цей результат?')) {
                const formData = new URLSearchParams();
                formData.append('id', id);
                formData.append('_csrf', '${res.locals._csrf || ''}');
                try {
                  const response = await fetch('/admin/delete-result', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (response.ok) {
                    location.reload();
                  } else {
                    alert('Помилка видалення');
                  }
                } catch (error) {
                  alert('Не вдалося видалити результат');
                }
              }
            }

            document.getElementById('search-form')?.addEventListener('submit', e => {
              e.preventDefault();
              const searchValue = e.target.search.value;
              window.location.href = '/admin/results?search=' + encodeURIComponent(searchValue);
            });
          </script>
        </body>
      </html>
    `;

    res.send(html);

  } catch (error) {
    logger.error('Помилка в /admin/results', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні результатів');
  } finally {
    logger.info('Маршрут /admin/results виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для перегляду детального результату (адмін/інструктор)
app.get('/admin/view-result', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
      return res.status(403).send('Доступно тільки для адміністраторів та інструкторів');
    }

    const { id } = req.query;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).send('Невірний ідентифікатор результату');
    }

    const result = await db.collection('test_results').findOne({ _id: new ObjectId(id) });
    if (!result) {
      return res.status(404).send('Результат не знайдено');
    }

    let allQuestions = await db.collection('questions')
      .find({ testNumber: result.testNumber })
      .sort({ order: 1 })
      .toArray();

    const questions = allQuestions.filter(q => 
      !q.variant || q.variant === '' || q.variant === result.variant
    );

    // Використовуємо функцію для підрахунку
    const scoresPerQuestion = questions.map((q, displayIndex) => {
      const userAnswer = result.answers[displayIndex];
      return calculateQuestionScore(q, userAnswer);
    });

    const exactScore = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
    const roundedScore = Math.round(exactScore * 10) / 10;
    const totalPoints = questions.reduce((sum, q) => sum + q.points, 0);
    const percentage = totalPoints > 0 ? (exactScore / totalPoints) * 100 : 0;
    const roundedPercentage = Math.round(percentage * 10) / 10;

    const totalQuestions = questions.length;
    const correctClicks = scoresPerQuestion.filter(s => s > 0).length;

    const timeAwayPercent = result.suspiciousActivity?.timeAway && result.duration
      ? Math.round((result.suspiciousActivity.timeAway / result.duration) * 100)
      : 0;

    const switchCount = result.suspiciousActivity?.switchCount || 0;
    const avgResponseTime = result.suspiciousActivity?.responseTimes?.length
      ? (result.suspiciousActivity.responseTimes.reduce((sum, t) => sum + (t || 0), 0) / result.suspiciousActivity.responseTimes.length).toFixed(2)
      : 0;

    const totalActivityCount = result.suspiciousActivity?.activityCounts
      ? result.suspiciousActivity.activityCounts.reduce((sum, c) => sum + (c || 0), 0)
      : 0;

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Деталі результату</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 30px 20px; background: #f5f5f5; }
            .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
            h1 { text-align: center; color: #333; margin-bottom: 25px; }
            table { border-collapse: collapse; width: 100%; margin: 20px 0; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
            th { background: #f2f2f2; font-weight: bold; }
            .summary { font-size: 20px; margin: 20px 0 40px; padding: 20px; background: #f8f9fa; border-radius: 8px; }
            .nav-btn { padding: 12px 24px; margin: 10px 5px; cursor: pointer; border: none; border-radius: 8px; background: #007bff; color: white; }
            .nav-btn:hover { background: #0056b3; }
          </style>
          <!-- Підключення pdfMake -->
          <script src="/pdfmake/pdfmake.min.js"></script>
          <script src="/pdfmake/vfs_fonts.js"></script>
        </head>
        <body>
          <div class="container">
            <h1>Деталі результату для користувача ${result.user}</h1>
            <div style="text-align: center; margin-bottom: 20px;">
              <button class="nav-btn" onclick="window.location.href='/admin/results'">Назад до списку</button>
              <button id="exportPDF" style="margin-left: 15px;">Експортувати в PDF</button>
            </div>

            <script>
              // Передаємо всі потрібні дані з сервера в JavaScript
              const viewResultData = {
                user: "${result.user.replace(/"/g, '\\"')}",
                testName: "${testNames[result.testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}",
                variant: "${result.variant || 'Немає'}",
                roundedScore: ${roundedScore.toFixed(1)},
                totalPoints: ${totalPoints.toFixed(1)},
                roundedPercentage: ${roundedPercentage.toFixed(1)},
                totalQuestions: ${totalQuestions},
                correctClicks: ${correctClicks},
                endDateTime: "${new Date(result.endTime).toLocaleString('uk-UA')}",
                timeAwayPercent: ${timeAwayPercent},
                switchCount: ${switchCount},
                avgResponseTime: ${avgResponseTime},
                totalActivityCount: ${totalActivityCount},
                questionsTable: ${JSON.stringify(questions.map((q, idx) => {
                  const userAns = result.answers[idx] !== undefined ? result.answers[idx] : 'Не відповіли';
                  let userDisplay = userAns;
                  if (Array.isArray(userAns)) {
                    if (q.type === 'matching') {
                      userDisplay = userAns.map(p => `${p[0]} → ${p[1]}`).join('<br>');
                    } else if (q.type === 'fillblank') {
                      userDisplay = userAns.join('<br>');
                    } else {
                      userDisplay = userAns.join(', ');
                    }
                  }
                  return {
                    text: q.text.replace(/"/g, '\\"').replace(/\n/g, '<br>'),
                    userAnswer: userDisplay,
                    score: scoresPerQuestion[idx].toFixed(3),
                    maxPoints: q.points
                  };
                }))}
              };

              // Функція експорту PDF з перевіркою
              function exportToPDF() {
                if (typeof pdfMake === 'undefined' || typeof pdfMake.createPdf === 'undefined') {
                  console.error('pdfMake не завантажився. Перевірте підключення скриптів:');
                  console.log('pdfmake.min.js:', document.querySelector('script[src="/pdfmake/pdfmake.min.js"]')?.src || 'не знайдено');
                  console.log('vfs_fonts.js:', document.querySelector('script[src="/pdfmake/vfs_fonts.js"]')?.src || 'не знайдено');
                  alert('PDF-генератор не завантажився. Спробуйте оновити сторінку (Ctrl+F5) або зверніться до адміністратора.');
                  return;
                }

                const docDefinition = {
                  pageSize: 'A4',
                  pageMargins: [40, 60, 40, 60],
                  content: [
                    { text: 'Деталі результату для користувача ' + viewResultData.user, style: 'mainHeader' },
                    { text: 'Тест: ' + viewResultData.testName, margin: [0, 10, 0, 5], style: 'subHeader' },
                    { text: 'Варіант: ' + viewResultData.variant, margin: [0, 0, 0, 15] },
                    {
                      table: {
                        widths: ['*', 'auto'],
                        body: [
                          [{ text: 'Бали:', bold: true }, viewResultData.roundedScore + ' з ' + viewResultData.totalPoints],
                          [{ text: 'Відсоток:', bold: true }, viewResultData.roundedPercentage + '%'],
                          [{ text: 'Питань:', bold: true }, viewResultData.totalQuestions],
                          [{ text: 'Правильних:', bold: true }, viewResultData.correctClicks],
                          [{ text: 'Дата завершення:', bold: true }, viewResultData.endDateTime]
                        ]
                      },
                      layout: 'lightHorizontalLines',
                      margin: [0, 0, 0, 20]
                    },
                    {
                      text: 'Підозріла активність:',
                      style: 'subHeader',
                      margin: [0, 0, 0, 5]
                    },
                    {
                      ul: [
                        'Час поза вкладкою: ' + viewResultData.timeAwayPercent + '%',
                        'Переключення вкладок: ' + viewResultData.switchCount,
                        'Середній час відповіді: ' + viewResultData.avgResponseTime + ' с',
                        'Загальна активність: ' + viewResultData.totalActivityCount
                      ],
                      margin: [0, 0, 0, 20]
                    },
                    { text: 'Деталі відповідей:', style: 'subHeader', margin: [0, 0, 0, 10] },
                    {
                      table: {
                        headerRows: 1,
                        widths: ['*', '*', 'auto'],
                        body: [
                          [
                            { text: 'Питання', bold: true, fillColor: '#f2f2f2' },
                            { text: 'Ваша відповідь', bold: true, fillColor: '#f2f2f2' },
                            { text: 'Бали', bold: true, fillColor: '#f2f2f2', alignment: 'center' }
                          ],
                          ...viewResultData.questionsTable.map(row => [
                            row.text,
                            row.userAnswer,
                            { text: row.score + ' / ' + row.maxPoints, alignment: 'center' }
                          ])
                        ]
                      },
                      layout: {
                        hLineWidth: () => 0.5,
                        vLineWidth: () => 0.5,
                        hLineColor: () => '#ddd',
                        vLineColor: () => '#ddd',
                        paddingLeft: () => 8,
                        paddingRight: () => 8,
                        paddingTop: () => 6,
                        paddingBottom: () => 6
                      }
                    }
                  ],
                  styles: {
                    mainHeader: { fontSize: 20, bold: true, alignment: 'center', margin: [0, 0, 0, 20] },
                    subHeader: { fontSize: 14, bold: true, margin: [0, 15, 0, 5] }
                  },
                  defaultStyle: {
                    fontSize: 11,
                    lineHeight: 1.4
                  }
                };

                pdfMake.createPdf(docDefinition).download(viewResultData.user + '_деталі_результату_' + '.pdf');
              }

              // Додаємо обробник після повного завантаження сторінки
              document.addEventListener('DOMContentLoaded', () => {
                const exportBtn = document.getElementById('exportPDF');
                if (exportBtn) {
                  exportBtn.addEventListener('click', exportToPDF);
                } else {
                  console.error('Кнопка #exportPDF не знайдена на сторінці');
                }
              });
            </script>

            <div class="summary">
              <strong>Тест:</strong> ${testNames[result.testNumber]?.name?.replace(/"/g, '\\"') || 'Невідомий тест'}<br>
              <strong>Варіант:</strong> ${result.variant || 'Немає'}<br>
              <strong>Бали:</strong> ${roundedScore.toFixed(1)} з ${totalPoints}<br>
              <strong>Відсоток:</strong> ${roundedPercentage.toFixed(1)}%<br>
              <strong>Питань:</strong> ${totalQuestions}<br>
              <strong>Правильних:</strong> ${correctClicks}<br>
              <strong>Дата завершення:</strong> ${new Date(result.endTime).toLocaleString('uk-UA')}<br><br>

              <strong>Підозріла активність:</strong><br>
              Час поза вкладкою: <span class="${timeAwayPercent > 50 ? 'suspicious' : ''}">${timeAwayPercent}%</span><br>
              Переключення вкладок: ${switchCount}<br>
              Середній час відповіді: ${avgResponseTime} с<br>
              Загальна активність: ${totalActivityCount}
            </div>

            <table>
              <tr>
                <th>Питання</th>
                <th>Ваша відповідь</th>
                <th>Бали</th>
              </tr>
    `;

    questions.forEach((question, index) => {
      const userAnswer = result.answers[index] !== undefined ? result.answers[index] : 'Не відповіли';
      const questionScore = scoresPerQuestion[index];

      let userAnswerDisplay = '—';
      if (Array.isArray(userAnswer)) {
        if (question.type === 'matching') {
          userAnswerDisplay = userAnswer.map(pair => `${pair[0]} → ${pair[1]}`).join('<br>');
        } else if (question.type === 'fillblank') {
          userAnswerDisplay = userAnswer.join('<br>');
        } else {
          userAnswerDisplay = userAnswer.join(', ');
        }
      } else {
        userAnswerDisplay = userAnswer;
      }

      html += `
        <tr>
          <td>${question.text.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>
          <td class="details">${userAnswerDisplay}</td>
          <td>${questionScore.toFixed(3)} / ${question.points}</td>
        </tr>
      `;
    });

    html += `
            </table>
            <br>
            <button class="nav-btn" onclick="window.location.href='/admin/results'">Назад до списку</button>
          </div>
        </body>
      </html>
    `;

    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/view-result', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка перегляду результату');
  } finally {
    logger.info('Маршрут /admin/view-result виконано');
  }
});

// Маршрут для редагування тестів
app.get('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Редагувати тести</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Редагувати тести</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <table>
            <tr>
              <th>Номер тесту</th>
              <th>Назва</th>
              <th>Ліміт часу (хв)</th>
              <th>Випадкові питання</th>
              <th>Випадкові відповіді</th>
              <th>Ліміт питань</th>
              <th>Ліміт спроб</th>
              <th>Швидкий тест</th>
              <th>Час на питання (с)</th>
              <th>Дії</th>
            </tr>
    `;
    if (!testNames || Object.keys(testNames).length === 0) {
      html += '<tr><td colspan="10">Немає тестів</td></tr>';
    } else {
      Object.entries(testNames).forEach(([num, data]) => {
        html += `
          <tr>
            <td>${num}</td>
            <td>${data.name.replace(/"/g, '\\"')}</td>
            <td>${data.timeLimit / 60}</td>
            <td>${data.randomQuestions ? 'Так' : 'Ні'}</td>
            <td>${data.randomAnswers ? 'Так' : 'Ні'}</td>
            <td>${data.questionLimit || 'Немає'}</td>
            <td>${data.attemptLimit || 1}</td>
            <td>${data.isQuickTest ? 'Так' : 'Ні'}</td>
            <td>${data.timePerQuestion || 'Немає'}</td>
            <td>
              <button class="action-btn edit" onclick="window.location.href='/admin/edit-test?testNumber=${num}'">Редагувати</button>
              <button class="action-btn delete" onclick="deleteTest('${num}')">Видалити</button>
            </td>
          </tr>
        `;
      });
    }
    html += `
          </table>
          <script>
            async function deleteTest(testNumber) {
              if (confirm('Ви впевнені, що хочете видалити тест ' + testNumber + '?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('testNumber', testNumber);
                  formData.append('_csrf', '${res.locals._csrf}');
                  const response = await fetch('/admin/delete-test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP-помилка! статус: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні тесту: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення тесту:', error);
                  alert('Не вдалося видалити тест. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-tests', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні тестів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-tests виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для редагування тесту
app.get('/admin/edit-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.query;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).send('Невірний номер тесту');
    }
    const test = testNames[testNumber];
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Редагувати тест</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select { padding: 5px; width: 300px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Редагувати тест ${testNumber}</h1>
          <form method="POST" action="/admin/edit-test" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="hidden" name="testNumber" value="${testNumber}">
            <label for="name">Назва тесту:</label>
            <input type="text" id="name" name="name" value="${test.name.replace(/"/g, '\\"')}" required>
            <label for="timeLimit">Ліміт часу (хвилини):</label>
            <input type="number" id="timeLimit" name="timeLimit" value="${test.timeLimit / 60}" min="1" required>
            <label for="randomQuestions">Випадкові питання:</label>
            <select id="randomQuestions" name="randomQuestions">
              <option value="true" ${test.randomQuestions ? 'selected' : ''}>Так</option>
              <option value="false" ${!test.randomQuestions ? 'selected' : ''}>Ні</option>
            </select>
            <label for="randomAnswers">Випадкові відповіді:</label>
            <select id="randomAnswers" name="randomAnswers">
              <option value="true" ${test.randomAnswers ? 'selected' : ''}>Так</option>
              <option value="false" ${!test.randomAnswers ? 'selected' : ''}>Ні</option>
            </select>
            <label for="questionLimit">Ліміт питань (опціонально):</label>
            <input type="number" id="questionLimit" name="questionLimit" value="${test.questionLimit || ''}" min="1">
            <label for="attemptLimit">Ліміт спроб:</label>
            <input type="number" id="attemptLimit" name="attemptLimit" value="${test.attemptLimit || 1}" min="1" required>
            <label for="isQuickTest">Швидкий тест:</label>
            <select id="isQuickTest" name="isQuickTest">
              <option value="true" ${test.isQuickTest ? 'selected' : ''}>Так</option>
              <option value="false" ${!test.isQuickTest ? 'selected' : ''}>Ні</option>
            </select>
            <label for="timePerQuestion">Час на питання (секунди, для швидкого тесту):</label>
            <input type="number" id="timePerQuestion" name="timePerQuestion" value="${test.timePerQuestion || ''}" min="1">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
          <script>
            function validateForm() {
              const name = document.getElementById('name').value;
              const timeLimit = document.getElementById('timeLimit').value;
              const questionLimit = document.getElementById('questionLimit').value;
              const attemptLimit = document.getElementById('attemptLimit').value;
              const timePerQuestion = document.getElementById('timePerQuestion').value;
              const isQuickTest = document.getElementById('isQuickTest').value;
              const errorMessage = document.getElementById('error-message');

              if (name.length < 1 || name.length > 100) {
                errorMessage.textContent = 'Назва тесту має бути від 1 до 100 символів';
                return false;
              }
              if (timeLimit < 1) {
                errorMessage.textContent = 'Ліміт часу має бути принаймні 1 хвилина';
                return false;
              }
              if (questionLimit && questionLimit < 1) {
                errorMessage.textContent = 'Ліміт питань має бути принаймні 1';
                return false;
              }
              if (attemptLimit < 1) {
                errorMessage.textContent = 'Ліміт спроб має бути принаймні 1';
                return false;
              }
              if (isQuickTest === 'true' && (!timePerQuestion || timePerQuestion < 1)) {
                errorMessage.textContent = 'Час на питання має бути принаймні 1 секунда для швидкого тесту';
                return false;
              }
              return true;
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки редагування тесту
app.post('/admin/edit-test', checkAuth, checkAdmin, [
  body('testNumber').notEmpty().withMessage('Номер тесту обов’язковий'),
  body('name')
    .isLength({ min: 1, max: 100 }).withMessage('Назва тесту має бути від 1 до 100 символів'),
  body('timeLimit')
    .isInt({ min: 1 }).withMessage('Ліміт часу має бути принаймні 1 хвилина'),
  body('questionLimit')
    .optional({ checkFalsy: true })
    .isInt({ min: 1 }).withMessage('Ліміт питань має бути принаймні 1'),
  body('attemptLimit')
    .isInt({ min: 1 }).withMessage('Ліміт спроб має бути принаймні 1'),
  body('timePerQuestion')
    .optional({ checkFalsy: true })
    .isInt({ min: 1 }).withMessage('Час на питання має бути принаймні 1 секунда')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).send(errors.array()[0].msg);
    }

    const { testNumber, name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    if (isQuickTest === 'true' && (!timePerQuestion || parseInt(timePerQuestion) < 1)) {
      return res.status(400).send('Час на питання має бути принаймні 1 секунда для швидкого тесту');
    }

    const testData = {
      name,
      timeLimit: parseInt(timeLimit) * 60,
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: isQuickTest === 'true' ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    testNames[testNumber] = testData;
    logger.info(`Тест ${testNumber} оновлено`, { testData });

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Тест оновлено</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Тест успішно оновлено</h1>
          <button onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка оновлення тесту', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-test (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для видалення тесту
app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.body;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).json({ success: false, message: 'Невірний номер тесту' });
    }
    await deleteTestFromMongoDB(testNumber);
    delete testNames[testNumber];
    await db.collection('questions').deleteMany({ testNumber });
    await db.collection('test_results').deleteMany({ testNumber });
    await CacheManager.invalidateCache('questions', testNumber);
    logger.info(`Тест ${testNumber} видалено разом із питаннями та результатами`);
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення тесту', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для створення нового тесту
app.get('/admin/create-test', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Створити тест</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select { padding: 5px; width: 300px; margin-bottom: 10px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Створити новий тест</h1>
          <form method="POST" action="/admin/create-test" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <label for="testNumber">Номер тесту:</label>
            <input type="text" id="testNumber" name="testNumber" required>
            <label for="name">Назва тесту:</label>
            <input type="text" id="name" name="name" required>
            <label for="timeLimit">Ліміт часу (хвилини):</label>
            <input type="number" id="timeLimit" name="timeLimit" value="60" min="1" required>
            <label for="randomQuestions">Випадкові питання:</label>
            <select id="randomQuestions" name="randomQuestions">
              <option value="true">Так</option>
              <option value="false" selected>Ні</option>
            </select>
            <label for="randomAnswers">Випадкові відповіді:</label>
            <select id="randomAnswers" name="randomAnswers">
              <option value="true">Так</option>
              <option value="false" selected>Ні</option>
            </select>
            <label for="questionLimit">Ліміт питань (опціонально):</label>
            <input type="number" id="questionLimit" name="questionLimit" min="1">
            <label for="attemptLimit">Ліміт спроб:</label>
            <input type="number" id="attemptLimit" name="attemptLimit" value="1" min="1" required>
            <label for="isQuickTest">Швидкий тест:</label>
            <select id="isQuickTest" name="isQuickTest">
              <option value="true">Так</option>
              <option value="false" selected>Ні</option>
            </select>
            <label for="timePerQuestion">Час на питання (секунди, для швидкого тесту):</label>
            <input type="number" id="timePerQuestion" name="timePerQuestion" min="1">
            <button type="submit" class="submit-btn">Створити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
          <script>
            function validateForm() {
              const testNumber = document.getElementById('testNumber').value;
              const name = document.getElementById('name').value;
              const timeLimit = document.getElementById('timeLimit').value;
              const questionLimit = document.getElementById('questionLimit').value;
              const attemptLimit = document.getElementById('attemptLimit').value;
              const timePerQuestion = document.getElementById('timePerQuestion').value;
              const isQuickTest = document.getElementById('isQuickTest').value;
              const errorMessage = document.getElementById('error-message');

              if (!/^[0-9]+$/.test(testNumber)) {
                errorMessage.textContent = 'Номер тесту має бути числом';
                return false;
              }
              if (name.length < 1 || name.length > 100) {
                errorMessage.textContent = 'Назва тесту має бути від 1 до 100 символів';
                return false;
              }
              if (timeLimit < 1) {
                errorMessage.textContent = 'Ліміт часу має бути принаймні 1 хвилина';
                return false;
              }
              if (questionLimit && questionLimit < 1) {
                errorMessage.textContent = 'Ліміт питань має бути принаймні 1';
                return false;
              }
              if (attemptLimit < 1) {
                errorMessage.textContent = 'Ліміт спроб має бути принаймні 1';
                return false;
              }
              if (isQuickTest === 'true' && (!timePerQuestion || timePerQuestion < 1)) {
                errorMessage.textContent = 'Час на питання має бути принаймні 1 секунда для швидкого тесту';
                return false;
              }
              return true;
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/create-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при створенні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/create-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки створення тесту
app.post('/admin/create-test', checkAuth, checkAdmin, [
  body('testNumber')
    .matches(/^[0-9]+$/).withMessage('Номер тесту має бути числом'),
  body('name')
    .isLength({ min: 1, max: 100 }).withMessage('Назва тесту має бути від 1 до 100 символів'),
  body('timeLimit')
    .isInt({ min: 1 }).withMessage('Ліміт часу має бути принаймні 1 хвилина'),
  body('questionLimit')
    .optional({ checkFalsy: true })
    .isInt({ min: 1 }).withMessage('Ліміт питань має бути принаймні 1'),
  body('attemptLimit')
    .isInt({ min: 1 }).withMessage('Ліміт спроб має бути принаймні 1'),
  body('timePerQuestion')
    .optional({ checkFalsy: true })
    .isInt({ min: 1 }).withMessage('Час на питання має бути принаймні 1 секунда')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).send(errors.array()[0].msg);
    }

    const { testNumber, name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    if (testNames[testNumber]) {
      return res.status(400).send('Тест із таким номером уже існує');
    }

    if (isQuickTest === 'true' && (!timePerQuestion || parseInt(timePerQuestion) < 1)) {
      return res.status(400).send('Час на питання має бути принаймні 1 секунда для швидкого тесту');
    }

    const testData = {
      name,
      timeLimit: parseInt(timeLimit) * 60,
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: isQuickTest === 'true' ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    testNames[testNumber] = testData;
    logger.info(`Створено новий тест ${testNumber}`, { testData });

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Тест створено</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Тест успішно створено</h1>
          <button onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка створення тесту', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при створенні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/create-test (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Видалення результатів
app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
  try {
    const { id } = req.body;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).json({ success: false, message: 'Невірний ID' });
    }

    const result = await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    if (result.deletedCount === 0) {
      return res.status(404).json({ success: false, message: 'Результат не знайдено' });
    }

    logger.info('Видалено результат', { id, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення результату', { error: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка сервера' });
  }
});

// Маршрут для очищення журналу активності
app.post('/admin/clear-activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    const result = await db.collection('activity_log').deleteMany({});
    logger.info('Журнал дій очищено', { deletedCount: result.deletedCount, user: req.user, ipAddress });
    await logActivity(req.user, 'очистив журнал дій', ipAddress, { deletedCount: result.deletedCount });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка очищення журналу дій', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при очищенні журналу' });
  } finally {
    logger.info('Маршрут /admin/clear-activity-log виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Маршрут для перегляду журналу активності
app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const search = req.query.search || '';

    const query = search ? { user: { $regex: search, $options: 'i' } } : {};
    const activities = await db.collection('activity_log')
      .find(query)
      .sort({ timestamp: -1 })
      .toArray();

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Журнал дій</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
            }
            .container {
              max-width: 800px;
              margin: 0 auto;
              background-color: white;
              padding: 20px;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            h1 {
              font-size: 24px;
              text-align: center;
              margin-bottom: 20px;
            }
            table {
              border-collapse: collapse;
              width: 100%;
              margin-top: 20px;
            }
            th, td {
              border: 1px solid #ddd;
              padding: 8px;
              text-align: left;
            }
            th {
              background-color: #f2f2f2;
            }
            .nav-btn, .search-btn, .clear-btn {
              padding: 10px 20px;
              margin: 10px 5px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
            }
            .nav-btn {
              background-color: #007bff;
              color: white;
            }
            .search-btn {
              background-color: #28a745;
              color: white;
            }
            .clear-btn {
              background-color: #ff4d4d;
              color: white;
            }
            .nav-btn:hover { background-color: #0056b3; }
            .search-btn:hover { background-color: #218838; }
            .clear-btn:hover { background-color: #d32f2f; }
            input[type="text"] {
              padding: 8px;
              margin: 5px;
              width: 200px;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 20px;
              }
              table {
                font-size: 14px;
              }
              .nav-btn, .search-btn, .clear-btn {
                width: 100%;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Журнал дій</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="clear-btn" onclick="clearActivityLog()">Очистити журнал дій</button>
            <div>
              <form id="search-form">
                <input type="text" id="search" name="search" placeholder="Пошук за логіном" value="${search}">
                <button type="submit" class="search-btn">Пошук</button>
              </form>
            </div>
            <table>
              <tr>
                <th>Користувач</th>
                <th>Дія</th>
                <th>IP-адреса</th>
                <th>Час</th>
              </tr>
              ${activities.length > 0 ? activities.map(a => `
                <tr>
                  <td>${a.user}</td>
                  <td>${a.action}</td>
                  <td>${a.ipAddress}</td>
                  <td>${new Date(a.timestamp).toLocaleString('uk-UA')}</td>
                </tr>
              `).join('') : '<tr><td colspan="4">Немає записів</td></tr>'}
            </table>
          </div>
          <script>
            document.getElementById('search-form').addEventListener('submit', (e) => {
              e.preventDefault();
              const search = document.getElementById('search').value;
              window.location.href = '/admin/activity-log?search=' + encodeURIComponent(search);
            });

            async function clearActivityLog() {
              if (confirm('Ви впевнені, що хочете очистити весь журнал дій? Цю дію неможливо скасувати.')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('_csrf', '${res.locals._csrf}');
                  const response = await fetch('/admin/clear-activity-log', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP-помилка! статус: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при очищенні журналу: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка очищення журналу:', error);
                  alert('Не вдалося очистити журнал. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/activity-log', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні журналу дій');
  } finally {
    logger.info('Маршрут /admin/activity-log виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Middleware для обробки помилок
app.use((err, req, res, next) => {
  logger.error('Неперехоплена помилка', { message: err.message, stack: err.stack });
  res.status(500).send('Щось пішло не так! Спробуйте ще раз або зверніться до адміністратора.');
});

// Запуск сервера
const port = process.env.PORT || 3000;
app.listen(port, () => {
  logger.info(`Сервер запущено на порту ${port}`);
});