// Імпорт необхідних модулів
require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const { MongoClient, ObjectId } = require('mongodb');
const bcrypt = require('bcrypt');
const multer = require('multer');
const fs = require('fs');
const nodemailer = require('nodemailer');
const { body, validationResult } = require('express-validator');
const jwt = require('jsonwebtoken');
const winston = require('winston');
const session = require('express-session');
const MongoStore = require('connect-mongo');

// Налаштування логування
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.File({ filename: 'error.log', level: 'error' }),
    new winston.transports.File({ filename: 'combined.log' })
  ]
});

if (process.env.NODE_ENV !== 'production') {
  logger.add(new winston.transports.Console({
    format: winston.format.simple()
  }));
}

// Налаштування multer для зображень і матеріалів
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['image/jpeg', 'image/png', 'image/gif', 'application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Непідтримуваний тип файлу'), false);
    }
  }
});

// Ініціалізація Express-додатку
const app = express();

// Увімкнення довіри до проксі
app.set('trust proxy', 1);

// Налаштування nodemailer
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
    pass: process.env.EMAIL_PASS || 'xfcd cvkl xiii qhtl'
  }
});

// Функція для відправки email про підозрілу активність
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
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/alpha?retryWrites=true&w=majority';
const client = new MongoClient(MONGODB_URI, {
  connectTimeoutMS: 5000,
  serverSelectionTimeoutMS: 5000
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
    logger.info(`Спроба підключення до MongoDB (${attempt}/${maxAttempts})`);
    const startTime = Date.now();
    await client.connect();
    db = client.db('alpha');
    logger.info(`Підключено до MongoDB за ${Date.now() - startTime} мс`, { databaseName: db.databaseName });
  } catch (error) {
    logger.error('Помилка підключення', { message: error.message, stack: error.stack });
    if (attempt < maxAttempts) {
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
  secret: process.env.SESSION_SECRET || 'your-secret-key',
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

// Генерація CSRF-токена
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

// CSRF-валідація
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

// Імпорт питань
const importQuestionsToMongoDB = async (buffer, testNumber) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) throw new Error('Лист "Questions" не знайдено');
    const MAX_ROWS = 1000;
    if (sheet.rowCount > MAX_ROWS + 1) throw new Error(`Занадто багато рядків (${sheet.rowCount - 1}). Макс: ${MAX_ROWS}`);
    const questions = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        try {
          const rowValues = row.values.slice(1);
          let questionText = rowValues[1];
          if (typeof questionText === 'object' && questionText) questionText = questionText.text || questionText.value || '[Невірний текст]';
          questionText = String(questionText || '').trim();
          if (!questionText) throw new Error('Текст питання відсутній');
          const picture = String(rowValues[0] || '').trim();
          let options = rowValues.slice(2, 14).filter(Boolean).map(val => String(val).trim());
          const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(val => String(val).trim());
          const type = String(rowValues[26] || 'multiple').toLowerCase();
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

          if (normalizedPicture) {
            const pictureMatch = normalizedPicture.match(/^Picture(\d+)$/i);
            if (pictureMatch) {
              const pictureNumber = parseInt(pictureMatch[1], 10);
              const targetFileNameBase = `Picture${pictureNumber}`;
              const extensions = ['.png', '.jpg', '.jpeg', '.gif'];
              const imageDir = path.join(__dirname, 'public', 'images');
              const filesInDir = fs.existsSync(imageDir) ? fs.readdirSync(imageDir) : [];
              for (const ext of extensions) {
                const expectedFileName = `${targetFileNameBase}${ext}`;
                const fileExists = filesInDir.some(file => file.toLowerCase() === expectedFileName.toLowerCase());
                if (fileExists) {
                  const matchedFile = filesInDir.find(file => file.toLowerCase() === expectedFileName.toLowerCase());
                  const imagePath = path.join(imageDir, matchedFile);
                  if (fs.existsSync(imagePath)) {
                    questionData.picture = `/images/Picture${pictureNumber}${ext.toLowerCase()}`;
                    break;
                  }
                }
              }
            }
          }

          if (type === 'matching') {
            questionData.pairs = options.map((opt, idx) => ({
              left: opt || '',
              right: questionData.correctAnswers[idx] || ''
            })).filter(pair => pair.left && pair.right);
            if (!questionData.pairs.length) throw new Error('Для Matching потрібні пари');
            questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
          }

          if (type === 'fillblank') {
            questionData.text = questionText.replace(/\s*___/g, '___');
            const blankCount = (questionData.text.match(/___/g) || []).length;
            if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
              throw new Error('Пропуски не відповідають відповідям');
            }
            questionData.blankCount = blankCount;
            questionData.correctAnswers.forEach((answer, idx) => {
              if (answer.includes('-')) {
                const [min, max] = answer.split('-').map(val => parseFloat(val.trim()));
                if (isNaN(min) || isNaN(max) || min > max) throw new Error(`Невірний діапазон для відповіді ${idx + 1}`);
              } else {
                const value = parseFloat(answer);
                if (isNaN(value)) throw new Error(`Відповідь ${idx + 1} має бути числом або діапазоном`);
              }
            });
          }

          if (type === 'singlechoice') {
            if (correctAnswers.length !== 1 || options.length < 2) throw new Error('Single Choice: потрібна 1 відповідь і ≥2 варіанти');
            questionData.correctAnswer = correctAnswers[0];
          }

          if (type === 'input') {
            if (correctAnswers.length !== 1) throw new Error('Input: потрібна 1 відповідь');
            const correctAnswer = correctAnswers[0];
            if (correctAnswer.includes('-')) {
              const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
              if (isNaN(min) || isNaN(max) || min > max) throw new Error('Невірний діапазон');
            } else {
              const value = parseFloat(correctAnswer);
              if (isNaN(value)) throw new Error('Відповідь має бути числом або діапазоном');
            }
          }

          questions.push(questionData);
        } catch (error) {
          throw new Error(`Помилка в рядку ${rowNumber}: ${error.message}`);
        }
      }
    });
    if (!questions.length) throw new Error('Не знайдено питань');
    await db.collection('questions').deleteMany({ testNumber });
    await db.collection('questions').insertMany(questions);
    await CacheManager.invalidateCache('questions', testNumber);
    return questions.length;
  } catch (error) {
    logger.error('Помилка імпорту питань', { message: error.message, stack: error.stack });
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
  const initStartTime = Date.now();
  try {
    logger.info('Початок ініціалізації сервера');

    // Підключення до MongoDB
    logger.info('Спроба підключення до MongoDB');
    await connectToMongoDB();
    logger.info(`MongoDB підключено за ${Date.now() - initStartTime} мс`);

    // Створення індексів (паралельно)
    logger.info('Створення індексів');
    const indexStartTime = Date.now();
    await Promise.all([
      db.collection('users').createIndex({ username: 1 }, { unique: true }).catch(err => logger.error('Помилка створення індексу users', { message: err.message })),
      db.collection('questions').createIndex({ testNumber: 1, variant: 1 }).catch(err => logger.error('Помилка створення індексу questions', { message: err.message })),
      db.collection('test_results').createIndex({ user: 1, testNumber: 1, endTime: -1 }).catch(err => logger.error('Помилка створення індексу test_results', { message: err.message })),
      db.collection('activity_log').createIndex({ user: 1, timestamp: -1 }).catch(err => logger.error('Помилка створення індексу activity_log', { message: err.message })),
      db.collection('test_attempts').createIndex({ user: 1, testNumber: 1, attemptDate: 1 }).catch(err => logger.error('Помилка створення індексу test_attempts', { message: err.message })),
      db.collection('login_attempts').createIndex({ ipAddress: 1, lastAttempt: 1 }).catch(err => logger.error('Помилка створення індексу login_attempts', { message: err.message })),
      db.collection('tests').createIndex({ testNumber: 1 }, { unique: true }).catch(err => logger.error('Помилка створення індексу tests', { message: err.message })),
      db.collection('active_tests').createIndex({ user: 1 }, { unique: true }).catch(err => logger.error('Помилка створення індексу active_tests', { message: err.message })),
      db.collection('sessions').createIndex({ expires: 1 }, { expireAfterSeconds: 0 }).catch(err => logger.error('Помилка створення індексу sessions', { message: err.message })),
      db.collection('sections').createIndex({ name: 1 }, { unique: true }).catch(err => logger.error('Помилка створення індексу sections', { message: err.message }))
    ]);
    logger.info(`Індекси створено за ${Date.now() - indexStartTime} мс`);

    // Створення папки для матеріалів
    const materialsDir = path.join(__dirname, 'public', 'materials');
    if (!fs.existsSync(materialsDir)) {
      fs.mkdirSync(materialsDir, { recursive: true });
      logger.info('Створено папку public/materials');
    }

    // Завантаження кешу користувачів
    const cacheLoadStartTime = Date.now();
    logger.info('Завантаження кешу користувачів');
    await loadUsersToCache().catch(err => logger.error('Помилка завантаження кешу користувачів', { message: err.message }));
    logger.info(`Кеш користувачів завантажено за ${Date.now() - cacheLoadStartTime} мс`, { userCacheLength: userCache.length });

    // Завантаження кешу тестів
    const testCacheLoadStartTime = Date.now();
    logger.info('Завантаження кешу тестів');
    await loadTestsFromMongoDB().catch(err => logger.error('Помилка завантаження кешу тестів', { message: err.message }));
    logger.info(`Кеш тестів завантажено за ${Date.now() - testCacheLoadStartTime} мс`, { testNames: Object.keys(testNames) });

    // Інвалідуємо кеш питань
    await CacheManager.invalidateCache('questions', null);
    logger.info('Кеш питань інвалідовано');

    isInitialized = true;
    initializationError = null;
    logger.info(`Сервер ініціалізовано успішно за ${Date.now() - initStartTime} мс`);
  } catch (error) {
    logger.error('Помилка ініціалізації', { message: error.message, stack: error.stack, duration: `${Date.now() - initStartTime} мс` });
    initializationError = error;
    isInitialized = false;
    throw error;
  }
};

// Пояснення: Додано створення папки `public/materials` для зберігання навчальних матеріалів і індекс для колекції `sections`.

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
  try {
    logger.info('Обробка запиту /favicon.ico');
    res.status(204).end();
  } catch (error) {
    logger.error('Помилка в /favicon.ico', { message: error.message, stack: error.stack });
    res.status(500).end();
  }
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
    logger.info('Початок обробки запиту /login', { body: req.body });

    // Перевірка сесії
    if (!req.session) {
      logger.error('Сесія відсутня в /login', { sessionID: req.sessionID || 'unknown' });
      return res.status(500).json({ success: false, message: 'Помилка сесії' });
    }
    logger.info('Сесія перевірена', { sessionID: req.sessionID });

    // Перевірка CSRF-токена
    const csrfToken = (req.body && req.body._csrf) || (req.headers && (req.headers['x-csrf-token'] || req.headers['xsrf-token'])) || '';
    logger.info('Отримано CSRF-токен', { csrfToken, body: req.body, headers: req.headers });
    if (!csrfToken || csrfToken !== res.locals._csrf) {
      logger.error('Невірний або відсутній CSRF-токен', { received: csrfToken, expected: res.locals._csrf });
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }

    // Очікування ініціалізації
    const maxWaitTime = 15000; // 15 секунд
    const waitStartTime = Date.now();
    while (!isInitialized && Date.now() - waitStartTime < maxWaitTime) {
      logger.info('Очікування ініціалізації сервера', { isInitialized, elapsed: Date.now() - waitStartTime });
      await new Promise(resolve => setTimeout(resolve, 500)); // Чекаємо 500 мс
    }
    if (!isInitialized) {
      logger.error('Сервер не ініціалізовано після очікування', { isInitialized, initializationError });
      return res.status(503).json({ success: false, message: 'Сервер ще ініціалізується' });
    }
    logger.info('Сервер ініціалізовано', { isInitialized });

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    logger.info('Перевірка лімітів входу', { ipAddress });
    await checkLoginAttempts(ipAddress);

    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації в /login', { errors: errors.array() });
      return res.status(400).json({ success: false, message: errors.array()[0].msg });
    }

    const { username, password } = req.body;
    logger.info('Отримано дані для входу', { username });

    if (!username || !password) {
      logger.warn('Логін або пароль не вказано');
      return res.status(400).json({ success: false, message: 'Логін або пароль не вказано' });
    }

    logger.info('Завантаження кешу користувачів', { userCacheLength: userCache.length });
    if (userCache.length === 0) {
      logger.warn('Кеш користувачів порожній, повторне завантаження');
      await loadUsersToCache();
      if (userCache.length === 0) {
        logger.error('Користувачів не знайдено після повторного завантаження');
        return res.status(500).json({ success: false, message: 'Користувачі недоступні в базі даних' });
      }
    }

    const foundUser = userCache.find(user => user.username === username);
    logger.info('Пошук користувача', { username, found: !!foundUser });
    if (!foundUser) {
      logger.warn('Користувача не знайдено', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    logger.info('Перевірка пароля', { username });
    if (!foundUser.password || typeof foundUser.password !== 'string') {
      logger.error('Некоректний пароль у базі', { username, password: foundUser.password });
      return res.status(500).json({ success: false, message: 'Помилка даних користувача' });
    }

    const passwordMatch = await bcrypt.compare(password, foundUser.password);
    logger.info('Результат перевірки пароля', { username, match: passwordMatch });
    if (!passwordMatch) {
      logger.warn('Невірний пароль', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    await checkLoginAttempts(ipAddress, true);
    logger.info('Ліміти входу скинуто', { ipAddress });

    const jwtSecret = process.env.JWT_SECRET || 'your-secret-key';
    logger.info('Генерація JWT-токена', { username, jwtSecret: jwtSecret ? 'defined' : 'undefined' });
    if (!jwtSecret) {
      logger.error('JWT_SECRET не визначено');
      return res.status(500).json({ success: false, message: 'Помилка конфігурації сервера' });
    }

    const jwtToken = jwt.sign(
      { username: foundUser.username, role: foundUser.role },
      jwtSecret,
      { expiresIn: '24h' }
    );

    res.cookie('token', jwtToken, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    res.cookie('auth_token', jwtToken, {
      httpOnly: false,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });
    logger.info('Cookies встановлено', { username });

    await logActivity(foundUser.username, 'увійшов на сайт', ipAddress);
    logger.info('Активність залогована', { username });

    res.json({ success: true, redirect: foundUser.role === 'admin' ? '/admin' : '/select-section' });
  } catch (error) {
    logger.error('Критична помилка в /login', {
      message: error.message,
      stack: error.stack,
      username: req.body ? req.body.username : 'unknown',
      sessionID: req.sessionID || 'unknown'
    });
    res.status(error.message.includes('Перевищено ліміт') ? 429 : 500).json({
      success: false,
      message: error.message || 'Помилка сервера під час авторизації'
    });
  } finally {
    logger.info('Маршрут /login завершився', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Змінено перенаправлення після входу на `/select-section` для всіх ролей, крім адміністратора.

// Middleware для перевірки авторизації через JWT
const checkAuth = (req, res, next) => {
  const token = req.headers['authorization']?.split(' ')[1] || req.cookies.token;
  if (!token) {
    return res.redirect('/');
  }

  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
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

// Middleware для перевірки ролі адміністратора або інструктора
const checkAdminOrInstructor = (req, res, next) => {
  if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
    return res.status(403).send('Доступно тільки для адміністраторів та інструкторів');
  }
  next();
};

// Сторінка вибору розділу
app.get('/select-section', ensureInitialized, checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    const sections = await db.collection('sections').find({}).toArray();

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Вибір розділу</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
              margin: 0;
            }
            .container {
              max-width: 1200px;
              margin: 0 auto;
              padding: 20px;
            }
            h1 {
              text-align: center;
              font-size: 28px;
              margin-bottom: 30px;
              color: #333;
            }
            .sections-grid {
              display: grid;
              grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
              gap: 20px;
            }
            .section-btn {
              position: relative;
              width: 100%;
              height: 200px;
              border: none;
              border-radius: 10px;
              overflow: hidden;
              cursor: pointer;
              transition: transform 0.3s ease;
              box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
              background-size: cover;
              background-position: center;
            }
            .section-btn:hover {
              transform: scale(1.05); /* Ефект зуму */
            }
            .section-label {
              position: absolute;
              bottom: 0;
              left: 0;
              right: 0;
              background-color: rgba(0, 0, 0, 0.7);
              color: white;
              font-size: 18px;
              font-weight: bold;
              text-align: center;
              padding: 10px;
              border-bottom-left-radius: 10px;
              border-bottom-right-radius: 10px;
            }
            .logout-btn {
              display: block;
              width: 200px;
              padding: 10px;
              margin: 30px auto;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #ff4d4d;
              color: white;
              font-size: 16px;
              text-align: center;
            }
            .logout-btn:hover {
              background-color: #d32f2f;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 24px;
              }
              .section-btn {
                height: 150px;
              }
              .section-label {
                font-size: 16px;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Виберіть розділ</h1>
            <div class="sections-grid">
    `;
    if (!sections.length) {
      html += `<p style="text-align: center; color: #666;">Немає доступних розділів</p>`;
    } else {
      sections.forEach(section => {
        html += `
          <button class="section-btn" style="background-image: url('${section.image || '/images/default-section.jpg'}');" onclick="window.location.href='/section/${section._id}'">
            <span class="section-label">${section.name.replace(/"/g, '\\"')}</span>
          </button>
        `;
      });
    }
    html += `
            </div>
            <button class="logout-btn" onclick="logout()">Вийти</button>
          </div>
          <script>
            async function logout() {
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) {
                  throw new Error('HTTP-помилка! статус: ' + response.status);
                }
                const result = await response.json();
                if (result.success) {
                  window.location.href = '/';
                } else {
                  alert('Помилка при виході: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка під час виходу:', error);
                alert('Не вдалося вийти.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /select-section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки вибору розділу');
  } finally {
    logger.info('Маршрут /select-section виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/select-section` відображає розділи як прямокутні кнопки з фоновим зображенням, підписами та ефектом зуму (`transform: scale(1.05)`). Використовує CSS Grid для адаптивного розташування.

// Сторінка вибору тесту (залишаємо для сумісності)
app.get('/select-test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
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
              padding-bottom: 80px; 
              margin: 0; 
              background-color: #f5f5f5;
            }
            h1 { 
              font-size: 24px; 
              margin-bottom: 20px; 
              color: #333;
            }
            .test-buttons { 
              display: flex; 
              flex-direction: column; 
              align-items: center; 
              gap: 10px; 
            }
            button, .instructions-btn, .feedback-btn { 
              padding: 10px; 
              font-size: 18px; 
              cursor: pointer; 
              width: 200px; 
              border: none; 
              border-radius: 5px; 
              color: white; 
              text-align: center;
              text-decoration: none;
            }
            button.test-btn { 
              background-color: #4CAF50; 
            }
            button.test-btn:hover { 
              background-color: #45a049; 
            }
            .instructions-btn { 
              background-color: #ffeb3b; 
              color: #333; 
            }
            .instructions-btn:hover { 
              background-color: #ffd700; 
            }
            .feedback-btn { 
              background-color: #ffeb3b; 
              color: #333; 
            }
            .feedback-btn:hover { 
              background-color: #ffd700; 
            }
            #logout { 
              background-color: #ef5350; 
              position: fixed; 
              bottom: 20px; 
              left: 50%; 
              transform: translateX(-50%); 
              width: 200px; 
            }
            #logout:hover { 
              background-color: #d32f2f; 
            }
            .no-tests { 
              color: red; 
              font-size: 18px; 
              margin-top: 20px; 
            }
            .results-btn { 
              background-color: #007bff; 
              margin-top: 20px; 
            }
            .results-btn:hover { 
              background-color: #0056b3; 
            }
            @media (max-width: 600px) {
              h1 { 
                font-size: 20px; 
              }
              button, .instructions-btn, .feedback-btn { 
                font-size: 16px; 
                width: 90%; 
                padding: 15px; 
              }
              #logout { 
                width: 90%; 
              }
            }
          </style>
        </head>
        <body>
          <h1>Виберіть тест</h1>
          <div class="test-buttons">
            ${Object.entries(testNames).length > 0
              ? Object.entries(testNames).map(([num, data]) => `
                  <button class="test-btn" onclick="window.location.href='/test?test=${num}'">${data.name.replace(/"/g, '\\"')}</button>
                `).join('')
              : '<p class="no-tests">Немає доступних тестів</p>'
            }
            ${req.userRole === 'instructor' ? `
              <button class="results-btn" onclick="window.location.href='/admin/results'">Переглянути результати</button>
            ` : ''}
            <a href="/instructions" class="instructions-btn">Інструкція до тестів</a>
            <a href="/feedback" class="feedback-btn">Зворотний зв’язок</a>
          </div>
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

// Сторінка розділу
app.get('/section/:sectionId', ensureInitialized, checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    const sectionId = req.params.sectionId;
    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });

    if (!section) {
      return res.status(404).send('Розділ не знайдено');
    }

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${section.name.replace(/"/g, '\\"')}</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
              margin: 0;
            }
            .container {
              max-width: 800px;
              margin: 0 auto;
              padding: 20px;
              background-color: white;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
              text-align: center;
            }
            h1 {
              font-size: 28px;
              margin-bottom: 20px;
              color: #333;
            }
            .section-options {
              display: flex;
              flex-direction: column;
              gap: 20px;
            }
            .option-btn {
              width: 100%;
              max-width: 400px;
              padding: 15px;
              margin: 0 auto;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #007bff;
              color: white;
              font-size: 18px;
              transition: background-color 0.3s;
            }
            .option-btn:hover {
              background-color: #0056b3;
            }
            .tests-list {
              margin-top: 20px;
              display: ${section.tests.length ? 'flex' : 'none'};
              flex-direction: column;
              gap: 10px;
            }
            .test-btn {
              width: 100%;
              max-width: 300px;
              padding: 10px;
              margin: 0 auto;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #28a745;
              color: white;
              font-size: 16px;
            }
            .test-btn:hover {
              background-color: #218838;
            }
            .back-btn {
              display: block;
              width: 200px;
              padding: 10px;
              margin: 30px auto;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #6c757d;
              color: white;
              font-size: 16px;
            }
            .back-btn:hover {
              background-color: #5a6268;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 24px;
              }
              .option-btn, .test-btn {
                font-size: 16px;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>${section.name.replace(/"/g, '\\"')}</h1>
            <div class="section-options">
              <button class="option-btn" onclick="window.location.href='/section/${sectionId}/materials'">Ознайомитись з навчальними матеріалами</button>
              <button class="option-btn" onclick="document.querySelector('.tests-list').style.display = document.querySelector('.tests-list').style.display === 'none' ? 'flex' : 'none';">Пройти тест</button>
              <div class="tests-list" style="display: none;">
                ${section.tests.length ? section.tests.map(testNumber => `
                  <button class="test-btn" onclick="window.location.href='/test?test=${testNumber}&sectionId=${sectionId}'">${testNames[testNumber]?.name.replace(/"/g, '\\"') || `Тест ${testNumber}`}</button>
                `).join('') : '<p>Немає тестів у цьому розділі</p>'}
              </div>
            </div>
            <button class="back-btn" onclick="window.location.href='/select-section'">Повернутися до розділів</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /section/:sectionId', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки розділу');
  } finally {
    logger.info('Маршрут /section/:sectionId виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/section/:sectionId` відображає сторінку розділу з двома кнопками: для перегляду матеріалів і вибору тестів (до 6). Тести ховаються/показуються через JavaScript.

// Перегляд навчальних матеріалів
app.get('/section/:sectionId/materials', ensureInitialized, checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const sectionId = req.params.sectionId;
    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });

    if (!section) {
      return res.status(404).send('Розділ не знайдено');
    }

    const isAdminOrInstructor = req.userRole === 'admin' || req.userRole === 'instructor';

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Навчальні матеріали - ${section.name.replace(/"/g, '\\"')}</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
              margin: 0;
            }
            .container {
              max-width: 800px;
              margin: 0 auto;
              padding: 20px;
              background-color: white;
              border-radius: 8px;
              box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            }
            h1 {
              font-size: 28px;
              text-align: center;
              margin-bottom: 20px;
              color: #333;
            }
            .materials-list {
              list-style: none;
              padding: 0;
            }
            .material-item {
              padding: 10px;
              border-bottom: 1px solid #ddd;
              display: flex;
              justify-content: space-between;
              align-items: center;
            }
            .material-item a {
              color: #007bff;
              text-decoration: none;
              font-size: 16px;
            }
            .material-item a:hover {
              text-decoration: underline;
            }
            .material-info {
              font-size: 14px;
              color: #666;
            }
            .upload-form {
              margin-top: 20px;
              padding: 15px;
              border: 1px solid #ddd;
              border-radius: 5px;
            }
            .upload-form label {
              display: block;
              margin-bottom: 10px;
              font-size: 16px;
            }
            .upload-form input[type="file"] {
              margin-bottom: 10px;
            }
            .upload-btn, .back-btn {
              padding: 10px 20px;
              margin: 10px 5px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              font-size: 16px;
            }
            .upload-btn {
              background-color: #28a745;
              color: white;
            }
            .upload-btn:hover {
              background-color: #218838;
            }
            .back-btn {
              background-color: #6c757d;
              color: white;
            }
            .back-btn:hover {
              background-color: #5a6268;
            }
            .error {
              color: red;
              text-align: center;
              margin-bottom: 10px;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 24px;
              }
              .material-item {
                flex-direction: column;
                align-items: flex-start;
                gap: 5px;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Навчальні матеріали - ${section.name.replace(/"/g, '\\"')}</h1>
            <ul class="materials-list">
              ${section.materials.length ? section.materials.map(material => `
                <li class="material-item">
                  <a href="${material.path}" target="_blank">${material.name}</a>
                  <span class="material-info">Завантажено: ${material.uploadedBy} (${new Date(material.uploadedAt).toLocaleString('uk-UA')})</span>
                </li>
              `).join('') : '<p>Немає навчальних матеріалів</p>'}
            </ul>
            ${isAdminOrInstructor ? `
              <div class="upload-form">
                <h2>Завантажити матеріал</h2>
                <form id="upload-form" enctype="multipart/form-data">
                  <label for="file">Виберіть файл:</label>
                  <input type="file" id="file" name="file" accept=".pdf,.doc,.docx" required>
                  <button type="submit" class="upload-btn">Завантажити</button>
                  <div id="error-message" class="error"></div>
                </form>
              </div>
            ` : ''}
            <button class="back-btn" onclick="window.location.href='/section/${sectionId}'">Повернутися до розділу</button>
          </div>
          ${isAdminOrInstructor ? `
            <script>
              document.getElementById('upload-form').addEventListener('submit', async (e) => {
                e.preventDefault();
                const fileInput = document.getElementById('file');
                const errorMessage = document.getElementById('error-message');
                const submitBtn = e.target.querySelector('.upload-btn');

                if (!fileInput.files[0]) {
                  errorMessage.textContent = 'Файл не вибрано.';
                  return;
                }

                submitBtn.disabled = true;
                submitBtn.textContent = 'Завантаження...';

                const formData = new FormData();
                formData.append('file', fileInput.files[0]);
                formData.append('_csrf', '${res.locals._csrf}');

                try {
                  const response = await fetch('/admin/section/${sectionId}/materials', {
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + document.cookie.split('; ').find(row => row.startsWith('auth_token=')).split('=')[1] },
                    body: formData
                  });

                  if (!response.ok) {
                    const result = await response.json();
                    throw new Error(result.message || 'Помилка: ' + response.status);
                  }

                  window.location.reload();
                } catch (error) {
                  console.error('Помилка:', error);
                  errorMessage.textContent = 'Помилка: ' + error.message;
                } finally {
                  submitBtn.disabled = false;
                  submitBtn.textContent = 'Завантажити';
                }
              });
            </script>
          ` : ''}
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /section/:sectionId/materials', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні матеріалів');
  } finally {
    logger.info('Маршрут /section/:sectionId/materials виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/section/:sectionId/materials` дозволяє переглядати матеріали (усі користувачі) і завантажувати нові (адмін/інструктор). Використовує `multer` для завантаження файлів.

// Завантаження навчальних матеріалів
app.post('/admin/section/:sectionId/materials', ensureInitialized, checkAuth, checkAdminOrInstructor, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    const sectionId = req.params.sectionId;
    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });

    if (!section) {
      return res.status(404).json({ success: false, message: 'Розділ не знайдено' });
    }

    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Файл не надано' });
    }

    const fileName = `${Date.now()}-${req.file.originalname}`;
    const filePath = `/materials/${fileName}`;
    fs.writeFileSync(path.join(__dirname, 'public', 'materials', fileName), req.file.buffer);

    await db.collection('sections').updateOne(
      { _id: new ObjectId(sectionId) },
      {
        $push: {
          materials: {
            name: req.file.originalname,
            path: filePath,
            uploadedBy: req.user,
            uploadedAt: new Date()
          }
        }
      }
    );

    logger.info('Навчальний матеріал завантажено', { sectionId, fileName, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка завантаження матеріалу', { message: error.message, stack: error.stack });
    res.status    (500).json({ success: false, message: 'Помилка при завантаженні матеріалу' });
  } finally {
    logger.info('Маршрут /admin/section/:sectionId/materials (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Завершено маршрут для завантаження матеріалів. Файли зберігаються в `public/materials`, а їхні метадані додаються до колекції `sections`.

// Обробка виходу користувача
app.post('/logout', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Отримано CSRF-токен у /logout', { token: req.body._csrf });
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    logActivity(req.user, 'покинув сайт', ipAddress);
    res.clearCookie('token');
    res.clearCookie('auth_token');
    req.session.destroy(err => {
      if (err) {
        logger.error('Помилка знищення сесії', { message: err.message, stack: err.stack });
        return res.status(500).json({ success: false, message: 'Помилка завершення сесії' });
      }
      logger.info('Сесію успішно знищено');
      res.json({ success: true });
    });
  } catch (error) {
    logger.error('Помилка в /logout', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при виході' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /logout виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Збереження результатів тесту
const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress, testSessionId) => {
  const startTimeLog = Date.now();
  const session = client.startSession();
  try {
    await session.withTransaction(async () => {
      const duration = Math.round((endTime - startTime) / 1000);
      const result = {
        user,
        testNumber,
        score,
        totalPoints,
        totalClicks,
        correctClicks,
        totalQuestions,
        percentage,
        startTime: new Date(startTime).toISOString(),
        endTime: new Date(endTime).toISOString(),
        duration,
        answers: Object.fromEntries(Object.entries(answers).sort((a, b) => parseInt(a[0]) - parseInt(b[0]))),
        scoresPerQuestion,
        suspiciousActivity,
        variant: `Variant ${variant}`,
        testSessionId
      };
      logger.info('Збереження результату в MongoDB із відповідями', {});
      if (!db) {
        throw new Error('Підключення до MongoDB не встановлено');
      }
      await db.collection('test_results').insertOne(result, { session });
      await logActivity(user, `завершив тест ${testNames[testNumber].name.replace(/"/g, '\\"')} з результатом ${Math.round(percentage)}%`, ipAddress, { percentage: Math.round(percentage) }, session);
    });
  } catch (error) {
    logger.error('Помилка збереження результату та логу активності', { message: error.message, stack: error.stack });
    throw error;
  } finally {
    await session.endSession();
    const endTimeLog = Date.now();
    logger.info('saveResult виконано', { duration: `${endTimeLog - startTimeLog}` });
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

    logger.info(`Користувач ${user} має ${attemptLimit - attempts} спроб для тесту ${testNumber}`);

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

// Форма зворотного зв’язку
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
            <button class="back-btn" onclick="window.location.href='/select-section'">Назад до вибору розділу</button>
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

// Пояснення: Оновлено маршрут `/feedback`, змінивши кнопку "Назад" для повернення на `/select-section` замість `/select-test`.

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

    await db.collection('feedback').insertOne({
      user,
      message,
      timestamp,
      ipAddress,
      read: false
    });

    logger.info('Зворотний зв’язок збережено', { user, message });

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
                  <td class="message">${f.message.replace(/</g, '<').replace(/>/g, '>')}</td>
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

// Інструкція до тестів
app.get('/instructions', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Інструкція до тестів</title>
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
              margin-top: 20px;
              margin-bottom: 10px;
            }
            p, li {
              font-size: 16px;
              margin-bottom: 10px;
            }
            ul {
              list-style-type: disc;
              padding-left: 20px;
            }
            img {
              max-width: 100%;
              height: auto;
              display: block;
              margin: 20px auto;
              border-radius: 5px;
            }
            .nav-btn {
              display: inline-block;
              padding: 10px 20px;
              margin-top: 20px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
              background-color: #4CAF50;
              color: white;
              text-decoration: none;
              font-size: 16px;
              text-align: center;
            }
            .nav-btn:hover {
              background-color: #45a049;
            }
            @media (max-width: 600px) {
              .container {
                padding: 15px;
              }
              h1 {
                font-size: 24px;
              }
              h2 {
                font-size: 18px;
              }
              p, li {
                font-size: 14px;
              }
              .nav-btn {
                width: 100%;
                box-sizing: border-box;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Інструкція для користувачів: Як правильно проходити тести</h1>
            <p>Вітаємо Вас в Центрі тестування! Щоб забезпечити найкращий досвід і отримати точні результати, будь ласка, дотримуйтесь цих рекомендацій:</p>
            
            <h2>1. Підготовка до тесту</h2>
            <ul>
              <li><strong>Перевірте з’єднання з Інтернетом:</strong> Переконайтеся, що ваше інтернет-з’єднання стабільне, щоб уникнути перерв під час тесту.</li>
              <li><strong>Використовуйте сумісний браузер:</strong> Рекомендуємо використовувати актуальні версії браузерів, таких як Google Chrome, Mozilla Firefox або Microsoft Edge.</li>
              <li><strong>Закрийте зайві вкладки та програми:</strong> Це допоможе уникнути відволікань і зменшить навантаження на пристрій.</li>
              <li><strong>Ознайомтеся з інструкціями:</strong> Перед початком тесту уважно прочитайте цю інструкцію.</li>
            </ul>

            <h2>2. Початок тесту</h2>
            <ul>
              <li><strong>Оберіть тест:</strong> На сторінці вибору розділу виберіть розділ, а потім тест із доступного списку.</li>
              <li><strong>Не залишайте сторінку без потреби:</strong> Якщо Ви плануєте перерву, завершіть тест перед тим, як закривати вкладку, щоб уникнути втрати прогресу.</li>
            </ul>

            <h2>3. Проведення тесту</h2>
            <ul>
              <li><strong>Відповідайте на питання послідовно:</strong> Пересувайтеся між питаннями за допомогою кнопок "Назад" і "Далі". Ви можете пропускати деякі питання і рухатись далі. Якщо Ви пропустили питання і не дали на нього відповідь, то в полосі прогресу кружечок з цим питанням буде червоного кольору і Ви зможете швидко знайти пропущене питання.</li>
              <li><strong>Перевіряйте відповіді:</strong> Перед завершенням тесту переконайтеся, що всі питання заповнені. Ви можете повертатися до попередніх питань, якщо це дозволено.</li>
              <li><strong>Дотримуйтесь таймера:</strong> Звертайте увагу на таймер у верхній частині екрана. Якщо час закінчиться, тест завершиться автоматично.</li>
              <li><strong>Увага до інструкцій під питаннями:</strong> Звертайте увагу на написи під текстом кожного питання, адже тести містять питання різних типів. Деякі питання мають лише одну правильну відповідь (питання типу "singlechoice"), напис під такими питаннями буде «Виберіть правильну відповідь». Питання мультивибору (типу "multiple") мають декілька правильних відповідей. Напис під цими питанням буде «Виберіть усі правильні відповіді». Вибір правильної кількості відповідей критично важливий для точного результату. Також є питання типу "input", в яких Вам необхідно у вікні відповіді ввести власноручно відповідь. У питаннях типу "fillblank" Вам необхідно буде заповнити пропуски у реченні. У питаннях типу "ordering" Вам будуть представлені варіанти відповідей (пункти), які необхідно буде розташувати у правильній послідовності переміщаючи (перетягуючи) їх. У питаннях типу "matching" Вам необхідно буде скласти пари, перетягуючи елементи і зіставляючи їх один навпроти іншого. Якщо Ви проходите тести з телефону, в яких зазвичай екрани мають невелике розширення, то на питаннях цього типу Вам необхідно буде розвернути телефон в альбомну розкладку, тоді Ви зможете коректно виконати такі пункти тесту.</li>
              <img src="/images/image1.jpg" alt="Інструкція для користувачів" onerror="this.style.display='none';">
            </ul>

            <h2>4. Завершення тесту</h2>
            <ul>
              <li><strong>Завершіть тест вручну:</strong> Натисніть кнопку "Завершити тест", коли закінчите, або дочекайтеся автоматичного завершення за таймером.</li>
              <li><strong>Перегляньте результати:</strong> Після завершення тесту програма перенаправить Вас на сторінку з результатами, де буде відображено ваш бал, відсоток правильних відповідей та іншу основну інформацію.</li>
              <li><strong>Експортуйте результати:</strong> Використовуйте кнопку "Експортувати в PDF", щоб зберегти результати тесту у зручному форматі.</li>
            </ul>

            <h2>5. На що звертати увагу</h2>
            <ul>
              <li><strong>Помилки сервера:</strong> Якщо з’являється повідомлення "Внутрішня помилка сервера", спробуйте перезавантажити сторінку. Якщо проблема повторюється, зверніться до адміністратора.</li>
              <li><strong>Збереження прогресу:</strong> Відповіді автоматично зберігаються під час переходу між питаннями, але при довгих перервах або збою з’єднання прогрес може бути втрачений. Завжди завершуйте тест у межах одного сеансу.</li>
              <li><strong>Сумнівна активність:</strong> Якщо ви багато перемикаєтеся між вкладками або проводите значний час поза тестом, це може бути зафіксовано системою для аналізу адміністратором.</li>
            </ul>

            <h2>6. Контактна інформація</h2>
            <p>Якщо у вас виникли труднощі або питання, зверніться до адміністратора через відповідний канал підтримки (наприклад, форму зворотного зв’язку).</p>

            <p style="text-align: center; font-size: 18px; margin-top: 20px;">Бажаємо успіхів у проходженні тестів! 😊</p>
            <a href="/select-section" class="nav-btn">Назад до вибору розділу</a>
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

// Пояснення: Оновлено маршрут `/instructions`, змінивши кнопку "Назад" для повернення на `/select-section`.

// Початок тесту
app.get('/test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');
    const testNumber = req.query.test;
    const sectionId = req.query.sectionId;
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
              <button onclick="window.location.href='/section/${sectionId}'">Повернутися до розділу</button>
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
    res.redirect(`/test/question?index=0${sectionId ? `&sectionId=${sectionId}` : ''}`);
  } catch (error) {
    logger.error('Помилка в /test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Оновлено маршрут `/test`, додавши параметр `sectionId` для повернення до розділу після тесту. Якщо спроби вичерпано, перенаправляє назад до `/section/:sectionId`.

// Відображення питання тесту
app.get('/test/question', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');

    let userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      return res.status(400).send('Тест не розпочато');
    }

    const { questions, testNumber, answers, currentQuestion, startTime: testStartTime, timeLimit, isQuickTest, timePerQuestion, suspiciousActivity, variant, testSessionId } = userTest;
    const sectionId = req.query.sectionId;

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
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'input' && userAnswer) {
          const normalizedUserAnswer = normalizeAnswer(userAnswer);
          const normalizedCorrectAnswer = normalizeAnswer(q.correctAnswers[0]);
          if (normalizedCorrectAnswer.includes('-')) {
            const [min, max] = normalizedCorrectAnswer.split('-').map(val => parseFloat(val.trim()));
            const userValue = parseFloat(normalizedUserAnswer);
            const isCorrect = !isNaN(userValue) && userValue >= min && userValue <= max;
            if (isCorrect) {
              questionScore = q.points;
            }
          } else {
            const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
            if (isCorrect) {
              questionScore = q.points;
            }
          }
        } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
          const userPairs = userAnswer.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const correctPairs = q.correctPairs.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const isCorrect = userPairs.length === correctPairs.length &&
            userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]));
          if (isCorrect) {
            questionScore = q.points;
          }
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
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswer = normalizeAnswer(q.correctAnswer);
          const isCorrect = userAnswers.length === 1 && userAnswers[0] === correctAnswer;
          if (isCorrect) {
            questionScore = q.points;
          }
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
            <button onclick="window.location.href='/section/${sectionId}'">Повернутися до розділу</button>
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
    const selectedOptionsString = JSON.stringify(selectedOptions).replace(/'/g, "\\'");

    const questionStartTime = userTest.questionStartTime || {};
    if (!questionStartTime[index]) {
      questionStartTime[index] = Date.now();
      await db.collection('active_tests').updateOne(
        { user: req.user },
        { $set: { questionStartTime } }
      );
    }

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${testNames[testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</title>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js"></script>
          <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 20px; padding-bottom: 80px; background-color: #f0f0f0; }
            h1 { font-size: 24px; text-align: center; }
            img { max-width: 100%; margin-bottom: 10px; display: block; margin-left: auto; margin-right: auto; }
            .progress-bar { 
              display: flex; 
              flex-wrap: wrap; 
              gap: 5px; 
              margin-bottom: 20px; 
              width: calc(100% - 40px); 
              margin-left: auto; 
              margin-right: auto; 
              box-sizing: border-box; 
              justify-content: center; 
              align-items: center; 
            }
            .progress-circle { 
              width: 40px; 
              height: 40px; 
              border-radius: 50%; 
              display: flex; 
              align-items: center; 
              justify-content: center; 
              font-size: 14px; 
              flex-shrink: 0; 
            }
            .progress-circle.unanswered { 
              background-color: red; 
              color: white; 
            }
            .progress-circle.answered { 
              background-color: green; 
              color: white; 
            }
            .progress-line { 
              width: 5px; 
              height: 2px; 
              background-color: #ccc; 
              margin: 0 2px; 
              align-self: center; 
              flex-shrink: 0; 
            }
            .progress-line.answered { 
              background-color: green; 
            }
            .option-box { border: 2px solid #ccc; padding: 10px; margin: 5px 0; border-radius: 5px; cursor: pointer; font-size: 16px; user-select: none; }
            .option-box.selected { background-color: #90ee90; }
            .button-container { position: fixed; bottom: 20px; left: 20px; right: 20px; display: flex; justify-content: space-between; }
            button { padding: 10px 20px; margin: 5px; border: none; cursor: pointer; border-radius: 5px; font-size: 16px; }
            .back-btn { background-color: red; color: white; }
            .next-btn { background-color: blue; color: white; }
            .finish-btn { background-color: green; color: white; }
            button:disabled { background-color: grey; cursor: not-allowed; }
            #timer { font-size: 24px; margin-bottom: 20px; text-align: center; }
            #question-timer { position: relative; width: 80px; height: 80px; margin: 0 auto 10px auto; }
            #question-timer svg { width: 100%; height: 100%; transform: rotate(-90deg); }
            #question-timer circle { fill: none; stroke-width: 8; }
            #question-timer .timer-circle-bg { stroke: #e0e0e0; }
            #question-timer .timer-circle { stroke: #ff4d4d; stroke-dasharray: 251; stroke-dashoffset: 0; transition: stroke-dashoffset 0.1s linear; }
            #question-timer .timer-text { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-size: 20px; font-weight: bold; color: #333; }
            #confirm-modal { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border: 2px solid black; z-index: 1000; }
            #confirm-modal button { margin: 0 10px; }
            .question-box { padding: 10px; margin: 5px 0; }
            .instruction { font-style: italic; color: #555; margin-bottom: 10px; font-size: 18px; }
            .option-box.draggable { cursor: move; }
            .option-box.dragging { opacity: 0.5; }
            #question-container { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); width: calc(100% - 40px); margin: 0 auto 20px auto; box-sizing: border-box; }
            #answers { margin-bottom: 20px; }
            .matching-container { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
            .matching-column { 
              width: 45%; 
              display: flex; 
              flex-direction: column; 
              gap: 5px; 
              box-sizing: border-box; 
            }
            .matching-item { 
              border: 2px solid #ccc; 
              padding: 10px; 
              margin: 0; 
              border-radius: 5px; 
              cursor: move; 
              font-family: Arial, sans-serif; 
              font-size: 16px; 
              line-height: 1.5; 
              min-height: 40px; 
              display: flex; 
              align-items: center; 
              justify-content: flex-start; 
              box-sizing: border-box; 
              white-space: normal; 
              overflow-wrap: break-word; 
            }
            .matching-item.matched { background-color: #90ee90; }
            .blank-input { width: 100px; margin: 0 5px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; display: inline-block; }
            .question-text { display: inline; }
            .image-error { color: red; font-style: italic; text-align: center; margin-bottom: 10px; }
            @media (max-width: 400px) {
              h1 { font-size: 24px; }
              .progress-bar { 
                gap: 2px; 
              }
              .progress-circle { 
                width: 20px; 
                height: 20px; 
                font-size: 8px; 
              }
              .progress-line { 
                width: 3px; 
              }
              button { font-size: 16px; padding: 10px; }
              #timer { font-size: 20px; }
              .question-box h2 { font-size: 18px; }
              .matching-container { flex-direction: column; }
              .matching-column { width: 100%; }
              .blank-input { width: 80px; }
              .option-box, .matching-item { 
                font-size: 14px; 
                padding: 8px; 
                min-height: 40px; 
                line-height: 1.5; 
              }
            }
            @media (min-width: 401px) and (max-width: 600px) {
              h1 { font-size: 28px; }
              .progress-bar { 
                gap: 3px; 
              }
              .progress-circle { 
                width: 25px; 
                height: 25px; 
                font-size: 10px; 
              }
              .progress-line { 
                width: 4px; 
              }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 20px; }
              .question-box h2 { font-size: 20px; }
              .matching-container { flex-direction: column; }
              .matching-column { width: 100%; }
              .blank-input { width: 80px; }
              .option-box, .matching-item { 
                font-size: 18px; 
                padding: 10px; 
                min-height: 50px; 
                line-height: 1.5; 
              }
            }
            @media (min-width: 601px) and (max-width: 900px) {
              h1 { font-size: 30px; }
              .progress-bar { 
                gap: 4px; 
              }
              .progress-circle { 
                width: 30px; 
                height: 30px; 
                font-size: 12px; 
              }
              .progress-line { 
                width: 5px; 
              }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 22px; }
              .question-box h2 { font-size: 22px; }
              .matching-column { width: 45%; }
              .blank-input { width: 100px; }
              .option-box, .matching-item { 
                font-size: 18px; 
                padding: 10px; 
                min-height: 50px; 
                line-height: 1.5; 
              }
            }
            @media (min-width: 901px) and (max-width: 1200px) {
              h1 { font-size: 32px; }
              .progress-bar { 
                gap: 5px; 
              }
              .progress-circle { 
                width: 35px; 
                height: 35px; 
                font-size: 14px; 
              }
              .progress-line { 
                width: 5px; 
              }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 24px; }
              .question-box h2 { font-size: 24px; }
              .matching-column { width: 45%; }
              .blank-input { width: 100px; }
              .option-box, .matching-item { 
                font-size: 18px; 
                padding: 10px; 
                min-height: 50px; 
                line-height: 1.5; 
              }
            }
            @media (min-width: 1201px) {
              h1 { font-size: 36px; }
              .progress-bar { 
                gap: 6px; 
              }
              .progress-circle { 
                width: 40px; 
                height: 40px; 
                font-size: 16px; 
              }
              .progress-line { 
                width: 6px; 
              }
              button { font-size: 20px; padding: 15px; }
              #timer { font-size: 26px; }
              .question-box h2 { font-size: 26px; }
              .matching-column { width: 45%; }
              .blank-input { width: 120px; }
              .option-box, .matching-item { 
                font-size: 20px; 
                padding: 12px; 
                min-height: 60px; 
                line-height: 1.5; 
              }
            }
          </style>
        </head>
        <body>
          <h1>${testNames[testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</h1>
          <div id="timer">Залишилось часу: ${minutes} хв ${seconds} с</div>
          <div class="progress-bar">
            ${progress.map((p, j) => `
              <div class="progress-circle ${p.answered ? 'answered' : 'unanswered'}" onclick="window.location.href='/test/question?index=${j}${sectionId ? `&sectionId=${sectionId}` : ''}'">${p.number}</div>
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
          <img src="${q.picture}" alt="Picture" onerror="this.style.display='none'; document.getElementById('image-error').style.display='block';">
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
      logger.info(`Частини питання fillblank для індексу ${index}`, { parts: q.text.split('___') });
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

    if (q.type === 'matching' && q.pairs) {
      const leftItems = shuffleArray([...q.pairs.map(p => p.left)]);
      const rightItems = shuffleArray([...q.pairs.map(p => p.right)]);
      const userPairs = Array.isArray(answers[index]) ? answers[index] : [];
      html += `
        <div class="matching-container">
          <div class="matching-column" id="left-column">
            ${leftItems.map((item, idx) => {
              const escapedItem = item.replace(/'/g, "\\'").replace(/"/g, '\\"');
              return `<div class="matching-item draggable" data-value="${escapedItem}">${item}</div>`;
            }).join('')}
          </div>
          <div class="matching-column" id="right-column">
            ${rightItems.map((item, idx) => {
              const escapedItem = item.replace(/'/g, "\\'").replace(/"/g, '\\"');
              const matchedLeft = userPairs.find(pair => pair[1] === item)?.[0] || '';
              return `
                <div class="matching-item droppable" data-value="${escapedItem}">
                  ${item}${matchedLeft ? `<span class="matched"> (Зіставлено: ${matchedLeft})</span>` : ''}
                </div>
              `;
            }).join('')}
          </div>
        </div>
        <button onclick="resetMatchingPairs()">Скинути зіставлення</button>
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
              <button class="back-btn" ${index === 0 ? 'disabled' : ''} onclick="window.location.href='/test/question?index=${index - 1}${sectionId ? `&sectionId=${sectionId}` : ''}'">Назад</button>
            ` : ''}
            <button id="submit-answer" class="next-btn" ${index === questions.length - 1 ? 'disabled' : ''} onclick="saveAndNext(${index})">Далі</button>
            <button class="finish-btn" onclick="showConfirm(${index})">Завершити тест</button>
          </div>
          <div id="confirm-modal">
            <h2>Ви дійсно бажаєте завершити тест?</h2>
            <button onclick="finishTest(${index})">Так</button>
            <button onclick="hideConfirm()">Ні</button>
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
            let questionStartTime = ${questionStartTime[index]};

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
                const safeAnswer = JSON.stringify(answers).replace(/'/g, "\\'").replace(/"/g, '\\"');
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');

                console.log('Автозбереження відповіді перед завершенням тесту:', { index, answers: safeAnswer, responseTime });

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  const errorText = await response.text();
                  throw new Error('HTTP-помилка! статус: ' + response.status + ' - ' + errorText);
                }

                const result = await response.json();
                if (!result.success) {
                  console.error('Помилка автозбереження відповіді:', result.error);
                }
              } catch (error) {
                console.error('Помилка в автозбереженні відповіді:', error);
              } finally {
                isSaving = false;
              }
            }

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
                const safeAnswer = JSON.stringify(answers).replace(/'/g, "\\'").replace(/"/g, '\\"');
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');

                console.log('Збереження даних у saveAndNext:', { timeAway, switchCount, responseTime, answer: safeAnswer });

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  throw new Error('HTTP-помилка! статус: ' + response.status);
                }

                const result = await response.json();
                if (result.success) {
                  const nextIndex = index + 1;
                  fetch('/set-question-start-time?index=' + nextIndex, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({ '_csrf': '${res.locals._csrf}' })
                  }).then(() => {
                    if (nextIndex < ${questions.length}) {
                      window.location.href = '/test/question?index=' + nextIndex + '${sectionId ? `&sectionId=${sectionId}` : ''}';
                    } else {
                      setTimeout(() => {
                        window.location.href = '/result${sectionId ? `?sectionId=${sectionId}` : ''}';
                      }, 300);
                    }
                  });
                } else {
                  console.error('Помилка збереження відповіді:', result.error);
                  alert('Помилка збереження відповіді: ' + result.error);
                }
              } catch (error) {
                console.error('Помилка в saveAndNext:', error);
                alert('Не вдалося зберегти відповідь: ' + error.message);
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

            async function finishTest(index) {
              if               isSaving) return;
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
                const safeAnswer = JSON.stringify(answers).replace(/'/g, "\\'").replace(/"/g, '\\"');
                formData.append('answer', safeAnswer);
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${res.locals._csrf}');

                console.log('Збереження даних у finishTest:', { timeAway, switchCount, responseTime, answer: safeAnswer });

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  throw new Error('HTTP-помилка! статус: ' + response.status);
                }

                const result = await response.json();
                if (result.success) {
                  setTimeout(() => {
                    window.location.href = '/result${sectionId ? `?sectionId=${sectionId}` : ''}';
                  }, 300);
                } else {
                  console.error('Помилка завершення тесту:', result.error);
                  alert('Помилка завершення тесту: ' + result.error);
                }
              } catch (error) {
                console.error('Помилка в finishTest:', error);
                alert('Не вдалося завершити тест: ' + error.message);
              } finally {
                isSaving = false;
              }
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
                console.log('Глобальний таймер закінчився, збереження відповіді та перенаправлення через 1.5с');
                saveCurrentAnswer(currentQuestionIndex).then(() => {
                  setTimeout(() => {
                    window.location.href = '/result${sectionId ? `?sectionId=${sectionId}` : ''}';
                  }, 1500);
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

              const questionTimerInterval = setInterval(() => {
                updateQuestionTimer();
                if (currentQuestionIndex >= totalQuestions - 1 && questionTimeRemaining <= 0 && !hasMovedToNext) {
                  console.log('Таймер швидкого тесту закінчився, збереження відповіді та перенаправлення через 1.5с');
                  clearInterval(questionTimerInterval);
                  saveCurrentAnswer(currentQuestionIndex).then(() => {
                    setTimeout(() => {
                      window.location.href = '/result${sectionId ? `?sectionId=${sectionId}` : ''}';
                    }, 1500);
                  });
                }
              }, 50);

              document.addEventListener('visibilitychange', () => {
                if (!document.hidden) {
                  updateGlobalTimer();
                  updateQuestionTimer();
                }
              });
            }

            window.addEventListener('blur', () => {
              if (!blurTimeout) {
                blurTimeout = setTimeout(() => {
                  if (lastBlurTime === 0) {
                    lastBlurTime = performance.now();
                    switchCount = Math.min(switchCount + 1, 1000);
                    console.log('Вкладка втратила фокус, початок підрахунку часу:', lastBlurTime, 'Кількість переключень:', switchCount);
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
                console.log('Вкладка отримала фокус, накопичено часу поза вкладкою:', awayDuration, 'Загальний timeAway:', timeAway);
                lastBlurTime = 0;
                saveCurrentAnswer(currentQuestionIndex);
              }
            });

            document.addEventListener('visibilitychange', () => {
              if (!document.hidden) {
                const now = Date.now();
                const timeSinceLastActivity = (now - lastActivityTime) / 1000;
                if (timeSinceLastActivity > 300) {
                  questionStartTime = now;
                  console.log('Виявлено тривалий простій, скидання questionStartTime:', questionStartTime);
                  fetch('/set-question-start-time?index=' + currentQuestionIndex, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({ '_csrf': '${res.locals._csrf}' })
                  });
                }
                updateGlobalTimer();
                if (isQuickTest) {
                  updateQuestionTimer();
                }
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
              box.addEventListener('click', () => {
                const questionType = '${q.type}';
                const option = box.getAttribute('data-value');
                if (questionType === 'truefalse' || questionType === 'multiple' || questionType === 'singlechoice') {
                  if (questionType === 'truefalse' || questionType === 'singlechoice') {
                    selectedOptions = [option];
                    document.querySelectorAll('.option-box').forEach(b => b.classList.remove('selected'));
                  } else {
                    const idx = selectedOptions.indexOf(option);
                    if (idx === -1) {
                      selectedOptions.push(option);
                    } else {
                      selectedOptions.splice(idx, 1);
                    }
                  }
                  box.classList.toggle('selected');
                }
              });
            });

            const sortable = document.getElementById('sortable-options');
            if (sortable) {
              new Sortable(sortable, { animation: 150 });
            }

            const leftColumn = document.getElementById('left-column');
            const rightColumn = document.getElementById('right-column');
            if (leftColumn && rightColumn && '${q.type}' === 'matching') {
              new Sortable(leftColumn, {
                group: 'matching',
                animation: 150,
                onStart: function(evt) {
                  evt.item.classList.add('dragging');
                },
                onEnd: function(evt) {
                  evt.item.classList.remove('dragging');
                  updateMatchingPairs();
                }
              });
              new Sortable(rightColumn, {
                group: 'matching',
                animation: 150,
                onStart: function(evt) {
                  evt.item.classList.add('dragging');
                },
                onEnd: function(evt) {
                  evt.item.classList.remove('dragging');
                  updateMatchingPairs();
                }
              });

              function updateMatchingPairs() {
                matchingPairs = [];
                const leftItems = Array.from(document.querySelectorAll('#left-column .draggable'));
                const rightItems = Array.from(document.querySelectorAll('#right-column .droppable'));
                rightItems.forEach((rightItem, idx) => {
                  const rightValue = rightItem.dataset.value || '';
                  const leftItem = leftItems[idx];
                  const leftValue = leftItem ? leftItem.dataset.value || '' : '';
                  if (leftValue && rightValue) {
                    matchingPairs.push([leftValue, rightValue]);
                  }
                });
              }

              function resetMatchingPairs() {
                matchingPairs = [];
                const rightItems = document.querySelectorAll('#right-column .droppable');
                rightItems.forEach(item => {
                  const rightValue = item.dataset.value || '';
                  item.innerHTML = rightValue;
                });
              }

              const droppableItems = document.querySelectorAll('.droppable');
              if (droppableItems.length > 0) {
                droppableItems.forEach(item => {
                  item.addEventListener('dragover', (e) => e.preventDefault());
                  item.addEventListener('drop', (e) => {
                    e.preventDefault();
                    const draggable = document.querySelector('.dragging');
                    if (draggable && draggable.classList.contains('draggable')) {
                      const leftValue = draggable.dataset.value || '';
                      const rightValue = item.dataset.value || '';
                      if (leftValue && rightValue) {
                        item.innerHTML = rightValue + ' <span class="matched"> (Зіставлено: ' + leftValue + ')</span>';
                        const leftColumn = document.getElementById('left-column');
                        const rightColumn = document.getElementById('right-column');
                        const leftItems = Array.from(leftColumn.children);
                        const rightItems = Array.from(rightColumn.children);
                        const rightIndex = rightItems.indexOf(item);
                        if (leftItems[rightIndex]) {
                          leftColumn.insertBefore(draggable, leftItems[rightIndex]);
                        } else {
                          leftColumn.appendChild(draggable);
                        }
                        updateMatchingPairs();
                      }
                    }
                  });
                });
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /test/question', { message: error.message, stack: error.stack, testNumber, testNames: Object.keys(testNames) });
    res.status(500).send('Внутрішня помилка сервера. Спробуйте ще раз або зверніться до адміністратора.');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test/question виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Завершено маршрут `/test/question`, додавши підтримку `sectionId` для перенаправлення назад до розділу після завершення тесту або переходу між питаннями. Додано клікабельність прогрес-кружечків для навігації.

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

    if (!index || !answer) {
      logger.error('Відсутні необхідні параметри в /answer', { index, answer });
      return res.status(400).json({ success: false, error: 'Необхідно надати index та answer' });
    }

    let parsedAnswer;
    try {
      if (typeof answer === 'string') {
        if (answer.trim() === '') {
          parsedAnswer = [];
        } else {
          logger.info('Парсинг відповіді в /answer', { answer });
          parsedAnswer = JSON.parse(answer);
        }
      } else {
        parsedAnswer = answer;
      }
    } catch (error) {
      logger.error('Помилка парсингу відповіді в /answer', { answer, message: error.message, stack: error.stack });
      return res.status(400).json({ success: false, error: 'Невірний формат відповіді' });
    }

    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      const recentResult = await db.collection('test_results').findOne(
        { user: req.user },
        { sort: { endTime: -1 } }
      );
      if (recentResult) {
        return res.json({ success: true });
      } else {
        logger.error('Тест не розпочато в /answer', { user: req.user });
        return res.status(400).json({ success: false, error: 'Тест не розпочато' });
      }
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

    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка в /answer', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /answer виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для відображення результатів тесту
app.get('/result', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');

    const sectionId = req.query.sectionId;
    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    let testData;

    if (!userTest) {
      const recentResult = await db.collection('test_results').findOne(
        { user: req.user },
        { sort: { endTime: -1 } }
      );
      if (!recentResult) {
        return res.status(400).send('Тест не розпочато або перерваний. Розпочніть новий тест.');
      }
      testData = recentResult;
    } else {
      testData = userTest;
    }

    const { questions, testNumber, answers, startTime: testStartTime, suspiciousActivity, variant, testSessionId, timeLimit } = userTest || testData;
    let score = testData.score || 0;
    const totalPoints = testData.totalPoints || (questions ? questions.reduce((sum, q) => sum + q.points, 0) : 0);
    let scoresPerQuestion = testData.scoresPerQuestion || [];

    if (!scoresPerQuestion.length && questions) {
      scoresPerQuestion = questions.map((q, index) => {
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
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'input' && userAnswer) {
          const normalizedUserAnswer = normalizeAnswer(userAnswer);
          const normalizedCorrectAnswer = normalizeAnswer(q.correctAnswers[0]);
          logger.info(`Порівняння відповіді input для питання ${index + 1}`, {
            userAnswer: normalizedUserAnswer,
            correctAnswer: normalizedCorrectAnswer
          });

          if (normalizedCorrectAnswer.includes('-')) {
            const [min, max] = normalizedCorrectAnswer.split('-').map(val => parseFloat(val.trim()));
            const userValue = parseFloat(normalizedUserAnswer);
            const isCorrect = !isNaN(userValue) && userValue >= min && userValue <= max;
            if (isCorrect) {
              questionScore = q.points;
            }
          } else {
            const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
            if (isCorrect) {
              questionScore = q.points;
            }
          }
        } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
          const userPairs = userAnswer.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const correctPairs = q.correctPairs.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const isCorrect = userPairs.length === correctPairs.length &&
            userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]));
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          logger.info(`Питання fillblank ${index + 1}`, { userAnswers, correctAnswers });

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

          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswer = normalizeAnswer(q.correctAnswer);
          logger.info(`Питання single choice ${index + 1}`, { userAnswers, correctAnswer });
          const isCorrect = userAnswers.length === 1 && userAnswers[0] === correctAnswer;
          if (isCorrect) {
            questionScore = q.points;
          }
        }
        return questionScore;
      });

      score = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
    }

    let endTime = testData.endTime ? new Date(testData.endTime).getTime() : Date.now();
    const maxEndTime = testStartTime + timeLimit;
    if (endTime > maxEndTime) {
      endTime = maxEndTime;
      logger.info(`Кориговано endTime до timeLimit для testSessionId: ${testSessionId}`);
    }

    const percentage = testData.percentage || (score / totalPoints) * 100;
    const totalClicks = testData.totalClicks || Object.keys(answers).length;
    const correctClicks = testData.correctClicks || scoresPerQuestion.filter(s => s > 0).length;
    const totalQuestions = testData.totalQuestions || (questions ? questions.length : 0);

    const duration = Math.round((endTime - testStartTime) / 1000);
    const timeAway = testData.timeAway || (suspiciousActivity ? suspiciousActivity.timeAway || 0 : 0);
    const correctedTimeAway = Math.min(timeAway, duration);
    const timeAwayPercent = Math.round((correctedTimeAway / duration) * 100);
    const switchCount = testData.switchCount || (suspiciousActivity ? suspiciousActivity.switchCount || 0 : 0);
    const avgResponseTime = testData.avgResponseTime || (suspiciousActivity && suspiciousActivity.responseTimes
      ? (suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / suspiciousActivity.responseTimes.length).toFixed(2)
      : 0);
    const totalActivityCount = testData.totalActivityCount || (suspiciousActivity && suspiciousActivity.activityCounts
      ? suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0)
      : 0);

    if (!testData.suspiciousActivity && (timeAwayPercent > config.suspiciousActivity.timeAwayThreshold || switchCount > config.suspiciousActivity.switchCountThreshold)) {
      const activityDetails = {
        timeAwayPercent,
        switchCount,
        avgResponseTime,
        totalActivityCount
      };
      await sendSuspiciousActivityEmail(req.user, activityDetails);
    }

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;

    if (userTest && !testData.isSaved) {
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
          { timeAway: correctedTimeAway, switchCount, responseTimes: suspiciousActivity?.responseTimes || [], activityCounts: suspiciousActivity?.activityCounts || [] },
          answers,
          scoresPerQuestion,
          variant,
          ipAddress,
          testSessionId
        );
        logger.info(`Результат збережено для testSessionId: ${testSessionId}`);
      }
    }

    if (userTest && testData.isSaved) {
      await db.collection('test_results').updateOne(
        { testSessionId },
        { $set: {
          score,
          totalPoints,
          endTime,
          totalClicks,
          correctClicks,
          totalQuestions,
          percentage,
          suspiciousActivity: { timeAway: correctedTimeAway, switchCount, responseTimes: suspiciousActivity?.responseTimes || [], activityCounts: suspiciousActivity?.activityCounts || [] },
          answers,
          scoresPerQuestion,
          variant,
          ipAddress
        } },
        { upsert: true }
      );
    }

    if (userTest) {
      await db.collection('active_tests').deleteOne({ user: req.user });
    }

    const endDateTime = new Date(endTime);
    const formattedTime = endDateTime.toLocaleTimeString('uk-UA', { hour12: false });
    const formattedDate = endDateTime.toLocaleDateString('uk-UA');
    const imagePath = path.join(__dirname, 'public', 'images', 'A.png');
    let imageBase64 = '';
    try {
      const imageBuffer = fs.readFileSync(imagePath);
      imageBase64 = imageBuffer.toString('base64');
    } catch (error) {
      logger.error('Помилка читання зображення A.png', { message: error.message, stack: error.stack });
    }

    const resultHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Результати ${testNames[testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f5f5f5; }
            .result-container { margin: 20px auto; width: 150px; height: 150px; position: relative; }
            .result-circle-bg { stroke: #e0e0e0; stroke-width: 10; fill: none; }
            .result-circle { stroke: #4CAF50; stroke-width: 10; fill: none; stroke-dasharray: 440; stroke-dashoffset: 440; animation: fillCircle 1.5s ease-in-out forwards; }
            .result-text { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-size: 24px; font-weight: bold; color: #333; }
            .buttons { margin-top: 20px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            #exportPDF { background-color: #ffeb3b; }
            #restart { background-color: #ef5350; }
            @keyframes fillCircle {
              to {
                stroke-dashoffset: ${(440 * (100 - percentage)) / 100};
              }
            }
          </style>
          <script src="/pdfmake/pdfmake.min.js"></script>
          <script src="/pdfmake/vfs_fonts.js"></script>
        </head>
        <body>
          <h1>Результат тесту</h1>
          <div class="result-container">
            <svg width="150" height="150">
              <circle class="result-circle-bg" cx="75" cy="75" r="70" />
              <circle class="result-circle" cx="75" cy="75" r="70" />
            </svg>
            <div class="result-text">${Math.round(percentage)}%</div>
          </div>
          <p>
            Кількість питань: ${totalQuestions}<br>
            Правильних відповідей: ${correctClicks}<br>
            Набрано балів: ${score}<br>
            Максимально можлива кількість балів: ${totalPoints}<br>
          </p>
          <div class="buttons">
            <button id="exportPDF">Експортувати в PDF</button>
            <button id="restart">Повернутися до розділу</button>
          </div>
          <script>
            const user = "${req.user.replace(/"/g, '\\"')}";
            const testName = "${testNames[testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}";
            const totalQuestions = ${totalQuestions};
            const correctClicks = ${correctClicks};
            const score = ${score};
            const totalPoints = ${totalPoints};
            const percentage = ${Math.round(percentage)};
            const time = "${formattedTime.replace(/"/g, '\\"')}";
            const date = "${formattedDate.replace(/"/g, '\\"')}";
            const imageBase64 = "${imageBase64.replace(/"/g, '\\"')}";

            console.log('Сторінка результатів завантажена з даними:', {
              user: user,
              testName: testName,
              totalQuestions: totalQuestions,
              correctClicks: correctClicks,
              score: score,
              totalPoints: totalPoints,
              percentage: percentage,
              time: time,
              date: date,
              imageBase64Length: imageBase64.length
            });

            const exportPDFButton = document.getElementById('exportPDF');
            const restartButton = document.getElementById('restart');

            if (!exportPDFButton) {
              console.error('Кнопка експорту PDF не знайдена!');
            } else {
              console.log('Кнопка експорту PDF знайдена, додаємо обробник події.');
              exportPDFButton.addEventListener('click', () => {
                try {
                  console.log('Натискання кнопки експорту PDF, генерація PDF...');
                  const docDefinition = {
                    content: [
                      imageBase64 ? {
                        image: 'data:image/png;base64,' + imageBase64,
                        width: 50,
                        alignment: 'center',
                        margin: [0, 0, 0, 20]
                      } : { text: 'Логотип відсутній', alignment: 'center', margin: [0, 0, 0, 20], lineHeight: 2 },
                      { text: 'Результат тесту користувача ' + user + ' з тесту ' + testName + ' складає ' + percentage + '%', style: 'header' },
                      { text: 'Кількість питань: ' + totalQuestions, lineHeight: 2 },
                      { text: 'Правильних відповідей: ' + correctClicks, lineHeight: 2 },
                      { text: 'Набрано балів: ' + score, lineHeight: 2 },
                      { text: 'Максимально можлива кількість балів: ' + totalPoints, lineHeight: 2 },
                      {
                        columns: [
                          { text: 'Час: ' + time, width: '50%', lineHeight: 2 },
                          { text: 'Дата: ' + date, width: '50%', alignment: 'right', lineHeight: 2 }
                        ],
                        margin: [0, 10, 0, 0]
                      }
                    ],
                    styles: {
                      header: { fontSize: 14, bold: true, margin: [0, 0, 0, 10], lineHeight: 2 }
                    }
                  };
                  pdfMake.createPdf(docDefinition).download('result.pdf');
                  console.log('PDF згенеровано успішно.');
                } catch (error) {
                  console.error('Помилка генерації PDF:', error);
                  alert('Не вдалося згенерувати PDF. Перевірте консоль браузера для деталей.');
                }
              });
            }

            if (!restartButton) {
              console.error('Кнопка повернення не знайдена!');
            } else {
              console.log('Кнопка повернення знайдена, додаємо обробник події.');
              restartButton.addEventListener('click', () => {
                console.log('Натискання кнопки повернення, перенаправлення на /section/${sectionId}');
                window.location.href = '${sectionId ? `/section/${sectionId}` : '/select-section'}';
              });
            }
          </script>
        </body>
      </html>
    `;
    res.send(resultHtml);
  } catch (error) {
    logger.error('Помилка в /result', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні результатів: ' + (error.message || 'Невідома помилка'));
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /result виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Оновлено маршрут `/result`, додавши підтримку `sectionId` для повернення до розділу після завершення тесту. Кнопка "Повернутися" перенаправляє на `/section/:sectionId` або `/select-section`, якщо `sectionId` відсутній.

// Маршрут для перегляду результатів користувача
app.get('/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    let resultsHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Результати</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .answers { white-space: pre-wrap; max-width: 300px; overflow-wrap: break-word; line-height: 1.8; }
            .buttons { margin-top: 20px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            #exportPDF { background-color: #ffeb3b; }
            #restart { background-color: #ef5350; }
          </style>
        </head>
        <body>
          <h1>Результати</h1>
    `;
    if (userTest) {
      const { questions, answers, testNumber, startTime, variant } = userTest;
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
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'input' && userAnswer) {
          const normalizedUserAnswer = normalizeAnswer(userAnswer);
          const normalizedCorrectAnswer = normalizeAnswer(q.correctAnswers[0]);
          if (normalizedCorrectAnswer.includes('-')) {
            const [min, max] = normalizedCorrectAnswer.split('-').map(val => parseFloat(val.trim()));
            const userValue = parseFloat(normalizedUserAnswer);
            const isCorrect = !isNaN(userValue) && userValue >= min && userValue <= max;
            if (isCorrect) {
              questionScore = q.points;
            }
          } else {
            const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
            if (isCorrect) {
              questionScore = q.points;
            }
          }
        } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
          const userPairs = userAnswer.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const correctPairs = q.correctPairs.map(pair => [normalizeAnswer(pair[0]), normalizeAnswer(pair[1])]);
          const isCorrect = userPairs.length === correctPairs.length &&
            userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]));
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswers = q.correctAnswers.map(val => normalizeAnswer(val));
          logger.info(`Питання fillblank ${index + 1} в /results`, { userAnswers, correctAnswers });
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
          if (isCorrect) {
            questionScore = q.points;
          }
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => normalizeAnswer(val));
          const correctAnswer = normalizeAnswer(q.correctAnswer);
          logger.info(`Питання single choice ${index + 1} в /results`, { userAnswers, correctAnswer });
          const isCorrect = userAnswers.length === 1 && userAnswers[0] === correctAnswer;
          if (isCorrect) {
            questionScore = q.points;
          }
        }
        return questionScore;
      });

      score = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
      const endTime = Date.now();
      const duration = Math.round((endTime - startTime) / 1000);
      const percentage = (score / totalPoints) * 100;
      const totalClicks = Object.keys(answers).length;
      const correctClicks = scoresPerQuestion.filter(s => s > 0).length;
      const totalQuestions = questions.length;
      const formattedTime = new Date(endTime).toLocaleTimeString('uk-UA', { hour12: false });
      const formattedDate = new Date(endTime).toLocaleDateString('uk-UA');
      const imagePath = path.join(__dirname, 'public', 'images', 'A.png');
      let imageBase64 = '';
      try {
        const imageBuffer = fs.readFileSync(imagePath);
        imageBase64 = imageBuffer.toString('base64');
      } catch (error) {
        logger.error('Помилка читання зображення A.png', { message: error.message, stack: error.stack });
      }

      resultsHtml += `
        <p>
          Результат тесту<br>
          ${Math.round(percentage)}%<br>
          Кількість питань: ${totalQuestions}<br>
          Правильних відповідей: ${correctClicks}<br>
          Набрано балів: ${score}<br>
          Максимально можлива кількість балів: ${totalPoints}<br>
        </p>
        <table>
          <tr>
            <th>Питання</th>
            <th>Ваш відповідь</th>
            <th>Правильна відповідь</th>
            <th>Бали</th>
          </tr>
      `;
      questions.forEach((q, index) => {
        const userAnswer = answers[index] || 'Не відповіли';
        let correctAnswer;
        if (q.type === 'matching') {
          correctAnswer = q.correctPairs.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ');
        } else if (q.type === 'fillblank') {
          correctAnswer = q.correctAnswers.join(', ');
        } else if (q.type === 'singlechoice') {
          correctAnswer = q.correctAnswer;
        } else {
          correctAnswer = q.correctAnswers.join(', ');
        }
        const questionScore = scoresPerQuestion[index];
        let userAnswerDisplay;
        if (q.type === 'matching' && Array.isArray(userAnswer)) {
          userAnswerDisplay = userAnswer.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ');
        } else if (q.type === 'fillblank' && Array.isArray(userAnswer)) {
          userAnswerDisplay = userAnswer.join(', ');
        } else if (Array.isArray(userAnswer)) {
          userAnswerDisplay = userAnswer.join(', ');
        } else {
          userAnswerDisplay = userAnswer;
        }
        resultsHtml += `
          <tr>
            <td>${q.text}</td>
            <td>${userAnswerDisplay}</td>
            <td>${correctAnswer}</td>
            <td>${questionScore} з ${q.points}</td>
          </tr>
        `;
      });
      resultsHtml += `
        </table>
        <div class="buttons">
          <button id="exportPDF">Експортувати в PDF</button>
          <button id="restart">Повернутися до розділу</button>
        </div>
        <script src="/pdfmake/pdfmake.min.js"></script>
        <script src="/pdfmake/vfs_fonts.js"></script>
        <script>
          const user = "${req.user.replace(/"/g, '\\"')}";
          const testName = "${testNames[testNumber].name.replace(/"/g, '\\"')}";
          const totalQuestions = ${totalQuestions};
          const correctClicks = ${correctClicks};
          const score = ${score};
          const totalPoints = ${totalPoints};
          const percentage = ${Math.round(percentage)};
          const time = "${formattedTime.replace(/"/g, '\\"')}";
          const date = "${formattedDate.replace(/"/g, '\\"')}";
          const imageBase64 = "${imageBase64.replace(/"/g, '\\"')}";

          document.getElementById('exportPDF').addEventListener('click', () => {
            const docDefinition = {
              content: [
                imageBase64 ? {
                  image: 'data:image/png;base64,' + imageBase64,
                  width: 150,
                  alignment: 'center',
                  margin: [0, 0, 0, 20]
                } : { text: 'Логотип відсутній', alignment: 'center', margin: [0, 0, 0, 20] },
                { text: 'Результат тесту користувача ' + user + ' з тесту ' + testName + ' складає ' + percentage + '%', style: 'header' },
                { text: 'Кількість питань: ' + totalQuestions },
                { text: 'Правильних відповідей: ' + correctClicks },
                { text: 'Набрано балів: ' + score },
                { text: 'Максимально можлива кількість балів: ' + totalPoints },
                {
                  columns: [
                    { text: 'Час: ' + time, width: '50%' },
                    { text: 'Дата: ' + date, width: '50%', alignment: 'right' }
                  ],
                  margin: [0, 10, 0, 0]
                }
              ],
              styles: {
                header: { fontSize: 14, bold: true, margin: [0, 0, 0, 10] }
              }
            };
            pdfMake.createPdf(docDefinition).download('results.pdf');
          });

          document.getElementById('restart').addEventListener('click', () => {
            window.location.href = '/select-section';
          });
        </script>
      `;

      await db.collection('active_tests').deleteOne({ user: req.user });
    } else {
      resultsHtml += '<p>Немає завершених тестів</p>';
    }
    resultsHtml += `
        </body>
      </html>
    `;
    res.send(resultsHtml);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /results виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Оновлено маршрут `/results`, змінивши кнопку "Повернутися" для повернення на `/select-section`.

// Маршрут для адмін-панелі
app.get('/admin', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const unreadFeedbackCount = await db.collection('feedback').countDocuments({ read: false });

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Адмін-панель</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; font-size: 24px; margin: 0; }
            h1 { font-size: 36px; margin-bottom: 20px; }
            button { padding: 15px 30px; margin: 10px; font-size: 24px; cursor: pointer; width: 300px; border: none; border-radius: 5px; background-color: #4CAF50; color: white; position: relative; }
            button:hover { background-color: #45a049; }
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
            }
            #logout { background-color: #ef5350; color: white; }
            @media (max-width: 600px) {
              body { padding: 20px; padding-bottom: 80px; }
              h1 { font-size: 32px; }
              button { font-size: 20px; width: 90%; padding: 15px; }
              #logout { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); width: 90%; }
              .notification-badge { font-size: 12px; padding: 3px 8px; }
            }
          </style>
        </head>
        <body>
          <h1>Адмін-панель</h1>
          <button onclick="window.location.href='/admin/users'">Керування користувачами</button><br>
          <button onclick="window.location.href='/admin/questions'">Керування питаннями</button><br>
          <button onclick="window.location.href='/admin/import-users'">Імпорт користувачів</button><br>
          <button onclick="window.location.href='/admin/import-questions'">Імпорт питань</button><br>
          <button onclick="window.location.href='/admin/results'">Перегляд результатів</button><br>
          <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button><br>
          <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button><br>
          <button onclick="window.location.href='/admin/sections'">Керування розділами</button><br>
          <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
          <button id="feedback-btn" onclick="window.location.href='/admin/feedback'">
            Зворотний зв’язок
            ${unreadFeedbackCount > 0 ? `<span class="notification-badge">${unreadFeedbackCount}</span>` : ''}
          </button><br>
          <button id="logout" onclick="logout()">Вийти</button>
          <script>
            async function logout() {
              console.log('Спроба виходу, CSRF-токен:', '${res.locals._csrf}');
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
                  throw new Error('HTTP-помилка! статус: ' + response.status);
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
    logger.error('Помилка в /admin', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні адмін-панелі');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Оновлено адмін-панель, додавши кнопку "Керування розділами" для доступу до `/admin/sections`.

// Маршрут для керування розділами
app.get('/admin/sections', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const sections = await db.collection('sections').find({}).toArray();

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Керування розділами</title>
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
            .nav-btn, .action-btn {
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
            .action-btn.edit {
              background-color: #4CAF50;
              color: white;
            }
            .action-btn.edit:hover {
              background-color: #45a049;
            }
            .action-btn.delete {
              background-color: #ef5350;
              color: white;
            }
            .action-btn.delete:hover {
              background-color: #d32f2f;
            }
            .tests-list {
              max-width: 300px;
              word-wrap: break-word;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 20px;
              }
              table {
                font-size: 14px;
              }
              .tests-list {
                max-width: 150px;
              }
              .nav-btn, .action-btn {
                width: 100%;
                box-sizing: border-box;
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Керування розділами</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="nav-btn" onclick="window.location.href='/admin/add-section'">Додати розділ</button>
            <table>
              <tr>
                <th>Назва</th>
                <th>Зображення</th>
                <th>Тести</th>
                <th>Дії</th>
              </tr>
              ${sections.length > 0 ? sections.map(section => `
                <tr>
                  <td>${section.name.replace(/"/g, '\\"')}</td>
                  <td>${section.image ? `<img src="${section.image}" alt="${section.name.replace(/"/g, '\\"')}" style="max-width: 100px;">` : 'Немає'}</td>
                  <td class="tests-list">${section.tests.map(t => testNames[t]?.name.replace(/"/g, '\\"') || `Тест ${t}`).join(', ')}</td>
                  <td>
                    <button class="action-btn edit" onclick="window.location.href='/admin/edit-section/${section._id}'">Редагувати</button>
                    <button class="action-btn delete" onclick="deleteSection('${section._id}')">Видалити</button>
                  </td>
                </tr>
              `).join('') : '<tr><td colspan="4">Немає розділів</td></tr>'}
            </table>
          </div>
          <script>
            async function deleteSection(id) {
              if (!confirm('Ви впевнені, що хочете видалити цей розділ?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/delete-section/' + id, {
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
                alert('Не вдалося видалити розділ.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/sections', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки керування розділами');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/sections виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/admin/sections` відображає список розділів із назвою, зображенням, тестами та діями (редагувати/видалити). Додано кнопки для повернення до адмін-панелі та створення нового розділу.

// Маршрут для додавання нового розділу
app.get('/admin/add-section', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const tests = await db.collection('tests').find({}).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Додати розділ</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
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
              text-align: center;
              margin-bottom: 20px;
            }
            label {
              display: block;
              font-size: 16px;
              margin-bottom: 5px;
            }
            input, select {
              width: 100%;
              padding: 10px;
              font-size: 16px;
              border: 1px solid #ccc;
              border-radius: 5px;
              margin-bottom: 10px;
              box-sizing: border-box;
            }
            .tests-container {
              margin-bottom: 20px;
            }
            .test-checkbox {
              margin-right: 10px;
            }
            button {
              padding: 10px 20px;
              font-size: 16px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
            }
            .submit-btn {
              background-color: #4CAF50;
              color: white;
            }
            .submit-btn:hover {
              background-color: #45a049;
            }
            .back-btn {
              background-color: #007bff;
              color: white;
            }
            .back-btn:hover {
              background-color: #0056b3;
            }
            .error {
              color: red;
              margin-top: 10px;
              font-size: 14px;
              text-align: center;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 20px;
              }
              input, select {
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
            <h1>Додати розділ</h1>
            <form id="add-section-form" enctype="multipart/form-data" action="/admin/add-section" method="post">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <input type="text" name="name" placeholder="Назва розділу" required>
              <input type="file" name="image" accept="image/*">
              <select name="tests" multiple>
                ${Object.keys(testNames).map(num => `
                  <option value="${num}">${testNames[num].name}</option>
                `).join('')}
              </select>
              <button type="submit">Створити</button>
            </form>
            <script>
              document.getElementById('add-section-form').addEventListener('submit', async (e) => {
                e.preventDefault();
                const form = e.target;
                const formData = new FormData(form);
                try {
                  const response = await fetch('/admin/add-section', {
                    method: 'POST',
                    body: formData
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.href = '/admin/sections';
                  } else {
                    alert('Помилка: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка:', error);
                  alert('Помилка створення розділу');
                }
              });
            </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/add-section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми додавання розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/admin/add-section` дозволяє додати розділ із назвою, зображенням і до 6 тестів. Використовує `multer` для завантаження зображення.

// Обробка додавання розділу
app.post('/admin/add-section', checkAuth, checkAdmin, upload.single('image'), [
  body('name')
    .isLength({ min: 1, max: 100 }).withMessage('Назва розділу має бути від 1 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Початок обробки запиту /admin/add-section', { body: req.body, file: req.file ? req.file.originalname : 'no file' });

    // Перевірка наявності req.body
    if (!req.body) {
      logger.error('req.body відсутній у /admin/add-section', { headers: req.headers });
      return res.status(400).json({ success: false, message: 'Некоректні дані форми' });
    }

    // Перевірка CSRF-токена
    const csrfToken = req.body._csrf || req.headers['x-csrf-token'] || req.headers['xsrf-token'] || '';
    logger.info('Отримано CSRF-токен', { csrfToken, body: req.body, headers: req.headers });
    if (!csrfToken || csrfToken !== res.locals._csrf) {
      logger.error('Невірний або відсутній CSRF-токен', { received: csrfToken, expected: res.locals._csrf });
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }

    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації', { errors: errors.array() });
      return res.status(400).json({ success: false, message: errors.array()[0].msg });
    }

    const { name, tests } = req.body;
    const testsArray = Array.isArray(tests) ? tests : tests ? [tests] : [];
    logger.info('Отримано дані для нового розділу', { name, tests: testsArray });

    if (testsArray.length > 6) {
      logger.warn('Занадто багато тестів', { testsCount: testsArray.length });
      return res.status(400).json({ success: false, message: 'Максимум 6 тестів на розділ' });
    }

    const existingSection = await db.collection('sections').findOne({ name });
    if (existingSection) {
      logger.warn('Розділ із такою назвою вже існує', { name });
      return res.status(400).json({ success: false, message: 'Розділ із такою назвою вже існує' });
    }

    let imagePath = '/images/default-section.png';
    if (req.file) {
      const fileName = `section_${Date.now()}-${req.file.originalname}`;
      imagePath = `/images/${fileName}`;
      fs.writeFileSync(path.join(__dirname, 'public', 'images', fileName), req.file.buffer);
      logger.info('Зображення збережено', { imagePath });
    }

    const newSection = {
      name,
      tests: testsArray,
      image: imagePath,
      materials: []
    };

    await db.collection('sections').insertOne(newSection);
    logger.info('Розділ створено', { name, tests: testsArray, user: req.user });

    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка створення розділу', { message: error.message, stack: error.stack, body: req.body });
    res.status(500).json({ success: false, message: 'Помилка при створенні розділу' });
  } finally {
    logger.info('Маршрут /admin/add-section виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Пояснення: Обробляє додавання розділу, перевіряючи унікальність назви та ліміт тестів (6). Зображення зберігається в `public/images`.

// Маршрут для редагування розділу
app.get('/admin/edit-section/:id', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const sectionId = req.params.id;
    if (!ObjectId.isValid(sectionId)) {
      return res.status(400).send('Невірний ідентифікатор розділу');
    }

    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });
    if (!section) {
      return res.status(404).send('Розділ не знайдено');
    }

    const tests = await db.collection('tests').find({}).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редагувати розділ</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              background-color: #f5f5f5;
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
              text-align: center;
              margin-bottom: 20px;
            }
            label {
              display: block;
              font-size: 16px;
              margin-bottom: 5px;
            }
            input, select {
              width: 100%;
              padding: 10px;
              font-size: 16px;
              border: 1px solid #ccc;
              border-radius: 5px;
              margin-bottom: 10px;
              box-sizing: border-box;
            }
            .tests-container {
              margin-bottom: 20px;
            }
            .test-checkbox {
              margin-right: 10px;
            }
            img#image-preview {
              max-width: 100px;
              margin-bottom: 10px;
            }
            button {
              padding: 10px 20px;
              font-size: 16px;
              cursor: pointer;
              border: none;
              border-radius: 5px;
            }
            .submit-btn {
              background-color: #4CAF50;
              color: white;
            }
            .submit-btn:hover {
              background-color: #45a049;
            }
            .back-btn {
              background-color: #007bff;
              color: white;
            }
            .back-btn:hover {
              background-color: #0056b3;
            }
            .error {
              color: red;
              margin-top: 10px;
              font-size: 14px;
              text-align: center;
            }
            @media (max-width: 600px) {
              h1 {
                font-size: 20px;
              }
              input, select {
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
            <h1>Редагувати розділ</h1>
            <form id="edit-section-form" enctype="multipart/form-data">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <input type="hidden" name="id" value="${sectionId}">
              <label for="name">Назва розділу:</label>
              <input type="text" id="name" name="name" value="${section.name.replace(/"/g, '\\"')}" required>
              <label for="image">Зображення (JPEG, PNG, GIF):</label>
              ${section.image ? `<img id="image-preview" src="${section.image}" alt="${section.name.replace(/"/g, '\\"')}">` : ''}
              <input type="file" id="image" name="image" accept="image/jpeg,image/png,image/gif">
              <label>Тести (до 6):</label>
              <div class="tests-container">
                ${tests.length > 0 ? tests.map(test => `
                  <label>
                    <input type="checkbox" class="test-checkbox" name="tests" value="${test.testNumber}" ${section.tests.includes(test.testNumber) ? 'checked' : ''}>
                    ${test.name.replace(/"/g, '\\"')}
                  </label><br>
                `).join('') : '<p>Немає доступних тестів</p>'}
              </div>
              <button type="submit" class="submit-btn">Зберегти</button>
            </form>
            <div id="error-message" class="error"></div>
            <button class="back-btn" onclick="window.location.href='/admin/sections'">Повернутися до розділів</button>
          </div>
          <script>
            document.getElementById('edit-section-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const name = document.getElementById('name').value;
              const imageInput = document.getElementById('image');
              const tests = Array.from(document.querySelectorAll('input[name="tests"]:checked')).map(cb => cb.value);
              const errorMessage = document.getElementById('error-message');
              const submitBtn = e.target.querySelector('.submit-btn');

              if (name.length <               1 || name.length > 100) {
                errorMessage.textContent = 'Назва розділу має бути від 1 до 100 символів';
                return;
              }
              if (tests.length > 6) {
                errorMessage.textContent = 'Максимум 6 тестів на розділ';
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Збереження...';

              const formData = new FormData();
              formData.append('id', '${sectionId}');
              formData.append('name', name);
              if (imageInput.files[0]) {
                formData.append('image', imageInput.files[0]);
              }
              tests.forEach(test => formData.append('tests', test));
              formData.append('_csrf', document.querySelector('input[name="_csrf"]').value);

              try {
                const response = await fetch('/admin/edit-section/${sectionId}', {
                  method: 'POST',
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  window.location.href = '/admin/sections';
                } else {
                  errorMessage.textContent = result.message || 'Помилка при редагуванні розділу';
                }
              } catch (error) {
                console.error('Помилка:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Зберегти';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-section/:id', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми редагування розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-section/:id виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Завершено маршрут `/admin/edit-section/:id`, додавши валідацію назви та кількості тестів. Форма дозволяє змінити назву, зображення та тести розділу.

// Обробка редагування розділу
app.post('/admin/edit-section/:id', checkAuth, checkAdmin, upload.single('image'), [
  body('name')
    .isLength({ min: 1, max: 100 }).withMessage('Назва розділу має бути від 1 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, message: errors.array()[0].msg });
    }

    const sectionId = req.params.id;
    if (!ObjectId.isValid(sectionId)) {
      return res.status(400).json({ success: false, message: 'Невірний ідентифікатор розділу' });
    }

    const { name, tests } = req.body;
    const testsArray = Array.isArray(tests) ? tests : tests ? [tests] : [];

    if (testsArray.length > 6) {
      return res.status(400).json({ success: false, message: 'Максимум 6 тестів на розділ' });
    }

    const existingSection = await db.collection('sections').findOne({ name, _id: { $ne: new ObjectId(sectionId) } });
    if (existingSection) {
      return res.status(400).json({ success: false, message: 'Розділ із такою назвою вже існує' });
    }

    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });
    if (!section) {
      return res.status(404).json({ success: false, message: 'Розділ не знайдено' });
    }

    let imagePath = section.image;
    if (req.file) {
      // Видаляємо старе зображення, якщо воно існує
      if (section.image && fs.existsSync(path.join(__dirname, 'public', section.image))) {
        fs.unlinkSync(path.join(__dirname, 'public', section.image));
      }
      const fileName = `section_${Date.now()}-${req.file.originalname}`;
      imagePath = `/images/${fileName}`;
      fs.writeFileSync(path.join(__dirname, 'public', 'images', fileName), req.file.buffer);
    }

    await db.collection('sections').updateOne(
      { _id: new ObjectId(sectionId) },
      { $set: { name, tests: testsArray, image: imagePath } }
    );

    logger.info('Розділ відредаговано', { sectionId, name, tests: testsArray, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка редагування розділу', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при редагуванні розділу' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-section/:id (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Обробляє редагування розділу, перевіряючи унікальність назви та ліміт тестів. Якщо нове зображення завантажено, старе видаляється.

// Видалення розділу
app.post('/admin/delete-section/:id', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const sectionId = req.params.id;
    if (!ObjectId.isValid(sectionId)) {
      return res.status(400).json({ success: false, message: 'Невірний ідентифікатор розділу' });
    }

    const section = await db.collection('sections').findOne({ _id: new ObjectId(sectionId) });
    if (!section) {
      return res.status(404).json({ success: false, message: 'Розділ не знайдено' });
    }

    // Видаляємо пов’язані матеріали
    if (section.materials && section.materials.length > 0) {
      section.materials.forEach(material => {
        if (material.path && fs.existsSync(path.join(__dirname, 'public', material.path))) {
          fs.unlinkSync(path.join(__dirname, 'public', material.path));
        }
      });
    }

    // Видаляємо зображення розділу
    if (section.image && fs.existsSync(path.join(__dirname, 'public', section.image))) {
      fs.unlinkSync(path.join(__dirname, 'public', section.image));
    }

    await db.collection('sections').deleteOne({ _id: new ObjectId(sectionId) });
    logger.info('Розділ видалено', { sectionId, name: section.name, user: req.user });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення розділу', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні розділу' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-section/:id виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Пояснення: Новий маршрут `/admin/delete-section/:id` видаляє розділ, його зображення та пов’язані матеріали з файлової системи та бази даних.

// Маршрут для керування користувачами
app.get('/admin/users', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const users = await db.collection('users').find({}).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Керування користувачами</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            table { border-collapse: collapse; width: 80%; margin: 20px auto; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .delete-btn { background-color: #ef5350; color: white; }
            .delete-btn:hover { background-color: #d32f2f; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            @media (max-width: 600px) {
              table { width: 100%; font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Керування користувачами</h1>
          <table>
            <tr>
              <th>Логін</th>
              <th>Роль</th>
              <th>Дії</th>
            </tr>
            ${users.map(user => `
              <tr>
                <td>${user.username}</td>
                <td>${user.role}</td>
                <td>
                  ${user.username !== 'admin' ? `
                    <button class="delete-btn" onclick="deleteUser('${user._id}')">Видалити</button>
                  ` : ''}
                </td>
              </tr>
            `).join('')}
          </table>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            async function deleteUser(userId) {
              if (!confirm('Ви впевнені, що хочете видалити цього користувача?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/users/delete/' + userId, {
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
                alert('Не вдалося видалити користувача.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/users', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки керування користувачами');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/users виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Видалення користувача
app.post('/admin/users/delete/:id', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const userId = req.params.id;
    if (!ObjectId.isValid(userId)) {
      return res.status(400).json({ success: false, message: 'Невірний ідентифікатор користувача' });
    }

    const user = await db.collection('users').findOne({ _id: new ObjectId(userId) });
    if (!user) {
      return res.status(404).json({ success: false, message: 'Користувача не знайдено' });
    }
    if (user.username === 'admin') {
      return res.status(403).json({ success: false, message: 'Неможливо видалити головного адміністратора' });
    }

    await db.collection('users').deleteOne({ _id: new ObjectId(userId) });
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Користувача видалено', { userId, username: user.username });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення користувача', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні користувача' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/users/delete/:id виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для імпорту користувачів
app.get('/admin/import-users', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Імпорт користувачів</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            input { margin: 10px; padding: 10px; font-size: 16px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .submit-btn:hover { background-color: #45a049; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .error { color: red; }
            @media (max-width: 600px) {
              input { width: 90%; font-size: 14px; }
              button { width: 90%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Імпорт користувачів</h1>
          <form id="import-users-form" enctype="multipart/form-data">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn">Імпортувати</button>
          </form>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <div id="error-message" class="error"></div>
          <script>
            document.getElementById('import-users-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const form = e.target;
              const formData = new FormData(form);
              const errorMessage = document.getElementById('error-message');
              const submitBtn = form.querySelector('.submit-btn');
              submitBtn.disabled = true;
              submitBtn.textContent = 'Імпортування...';
              try {
                const response = await fetch('/admin/import-users', {
                  method: 'POST',
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  alert('Імпортовано ' + result.count + ' користувачів');
                  window.location.href = '/admin/users';
                } else {
                  errorMessage.textContent = result.message || 'Помилка імпорту';
                }
              } catch (error) {
                console.error('Помилка:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Імпортувати';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/import-users', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми імпорту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-users виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Файл не надано' });
    }
    const count = await importUsersToMongoDB(req.file.buffer);
    res.json({ success: true, count });
  } catch (error) {
    logger.error('Помилка імпорту користувачів', { message: error.message, stack: error.stack });
    res.status(400).json({ success: false, message: error.message || 'Помилка імпорту користувачів' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-users (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для імпорту питань
app.get('/admin/import-questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Імпорт питань</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            input, select { margin: 10px; padding: 10px; font-size: 16px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .submit-btn:hover { background-color: #45a049; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .error { color: red; }
            @media (max-width: 600px) {
              input, select { width: 90%; font-size: 14px; }
              button { width: 90%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Імпорт питань</h1>
          <form id="import-questions-form" enctype="multipart/form-data">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <select name="testNumber" required>
              ${Object.keys(testNames).map(num => `
                <option value="${num}">${testNames[num].name.replace(/"/g, '\\"')}</option>
              `).join('')}
            </select>
            <input type="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn">Імпортувати</button>
          </form>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <div id="error-message" class="error"></div>
          <script>
            document.getElementById('import-questions-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const form = e.target;
              const formData = new FormData(form);
              const errorMessage = document.getElementById('error-message');
              const submitBtn = form.querySelector('.submit-btn');
              submitBtn.disabled = true;
              submitBtn.textContent = 'Імпортування...';
              try {
                const response = await fetch('/admin/import-questions', {
                  method: 'POST',
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  alert('Імпортовано ' + result.count + ' питань');
                  window.location.href = '/admin/questions';
                } else {
                  errorMessage.textContent = result.message || 'Помилка імпорту';
                }
              } catch (error) {
                console.error('Помилка:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Імпортувати';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/import-questions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми імпорту питань');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-questions виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    const testNumber = req.body.testNumber;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).json({ success: false, message: 'Номер тесту не вказано або тест не існує' });
    }
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Файл не надано' });
    }
    const count = await importQuestionsToMongoDB(req.file.buffer, testNumber);
    await CacheManager.invalidateCache('questions', testNumber);
    res.json({ success: true, count });
  } catch (error) {
    logger.error('Помилка імпорту питань', { message: error.message, stack: error.stack });
    res.status(400).json({ success: false, message: error.message || 'Помилка імпорту питань' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-questions (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для керування питаннями
app.get('/admin/questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;
    const questions = await CacheManager.getAllQuestions();
    const paginatedQuestions = questions.slice(skip, skip + limit);
    const totalPages = Math.ceil(questions.length / limit);

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Керування питаннями</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            table { border-collapse: collapse; width: 90%; margin: 20px auto; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            img { max-width: 100px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .delete-btn { background-color: #ef5350; color: white; }
            .delete-btn:hover { background-color: #d32f2f; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
            .options { max-width: 200px; overflow-wrap: break-word; }
            .correct-answers { max-width: 200px; overflow-wrap: break-word; }
            @media (max-width: 600px) {
              table { width: 100%; font-size: 14px; }
              button { width: 100%; font-size: 14px; }
              img { max-width: 80px; }
            }
          </style>
        </head>
        <body>
          <h1>Керування питаннями</h1>
          <table>
            <tr>
              <th>Тест</th>
              <th>Зображення</th>
              <th>Питання</th>
              <th>Тип</th>
              <th>Варіанти</th>
              <th>Правильні відповіді</th>
              <th>Бали</th>
              <th>Варіант</th>
              <th>Дії</th>
            </tr>
            ${paginatedQuestions.map(q => `
              <tr>
                <td>${testNames[q.testNumber]?.name || q.testNumber}</td>
                <td>${q.picture ? `<img src="${q.picture}" alt="Picture" onerror="this.style.display='none'">` : '-'}</td>
                <td>${q.text.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</td>
                <td>${q.type}</td>
                <td class="options">${q.options ? q.options.join(', ') : q.pairs ? q.pairs.map(p => p.left + ' -> ' + p.right).join(', ') : '-'}</td>
                <td class="correct-answers">${q.correctAnswers ? q.correctAnswers.join(', ') : q.correctPairs ? q.correctPairs.map(p => p[0] + ' -> ' + p[1]).join(', ') : q.correctAnswer || '-'}</td>
                <td>${q.points}</td>
                <td>${q.variant || '-'}</td>
                <td>
                  <button class="delete-btn" onclick="deleteQuestion('${q._id}')">Видалити</button>
                </td>
              </tr>
            `).join('')}
          </table>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/questions?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/questions?page=${page + 1}">Наступна</a>` : ''}
          </div>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            async function deleteQuestion(questionId) {
              if (!confirm('Ви впевнені, що хочете видалити це питання?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/questions/delete/' + questionId, {
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
                alert('Не вдалося видалити питання.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/questions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки керування питаннями');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/questions виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/questions/delete/:id', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const questionId = req.params.id;
    if (!ObjectId.isValid(questionId)) {
      return res.status(400).json({ success: false, message: 'Невірний ідентифікатор питання' });
    }

    const question = await db.collection('questions').findOne({ _id: new ObjectId(questionId) });
    if (!question) {
      return res.status(404).json({ success: false, message: 'Питання не знайдено' });
    }

    await db.collection('questions').deleteOne({ _id: new ObjectId(questionId) });
    await CacheManager.invalidateCache('questions', question.testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    logger.info('Питання видалено', { questionId, testNumber: question.testNumber });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення питання', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні питання' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/questions/delete/:id виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для перегляду результатів в адмін-панелі
app.get('/admin/results', checkAuth, checkAdminOrInstructor, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;

    let query = {};
    if (req.userRole === 'instructor') {
      query.user = { $ne: 'admin' };
    }

    const results = await db.collection('test_results')
      .find(query)
      .sort({ endTime: -1 })
      .skip(skip)
      .limit(limit)
      .toArray();

    const totalResults = await db.collection('test_results').countDocuments(query);
    const totalPages = Math.ceil(totalResults / limit);

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Результати тестів</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            table { border-collapse: collapse; width: 90%; margin: 20px auto; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
            .answers { max-width: 300px; overflow-wrap: break-word; }
            @media (max-width: 600px) {
              table { width: 100%; font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Результати тестів</h1>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Тест</th>
              <th>Результат</th>
              <th>Бали</th>
              <th>Час завершення</th>
              <th>Відповіді</th>
            </tr>
            ${results.map(r => `
              <tr>
                <td>${r.user}</td>
                <td>${testNames[r.testNumber]?.name || r.testNumber}</td>
                <td>${Math.round(r.percentage)}%</td>
                <td>${r.score}/${r.totalPoints}</td>
                <td>${new Date(r.endTime).toLocaleString('uk-UA')}</td>
                <td class="answers">${Object.entries(r.answers).map(([q, a]) => `Питання ${parseInt(q) + 1}: ${Array.isArray(a) ? a.join(', ') : a}`).join('<br>')}</td>
              </tr>
            `).join('')}
          </table>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/results?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/results?page=${page + 1}">Наступна</a>` : ''}
          </div>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/results', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні результатів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/results виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для редагування назв тестів
app.get('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const tests = await db.collection('tests').find({}).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редагувати тести</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            table { border-collapse: collapse; width: 80%; margin: 20px auto; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            input { padding: 5px; font-size: 16px; width: 100%; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .save-btn { background-color: #4CAF50; color: white; }
            .save-btn:hover { background-color: #45a049; }
            .delete-btn { background-color: #ef5350; color: white; }
            .delete-btn:hover { background-color: #d32f2f; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .error { color: red; }
            @media (max-width: 600px) {
              table { width: 100%; font-size: 14px; }
              input { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Редагувати тести</h1>
          <table>
            <tr>
              <th>Номер тесту</th>
              <th>Назва</th>
              <th>Ліміт часу (сек)</th>
              <th>Випадкові питання</th>
              <th>Випадкові відповіді</th>
              <th>Ліміт питань</th>
              <th>Ліміт спроб</th>
              <th>Швидкий тест</th>
              <th>Час на питання (сек)</th>
              <th>Дії</th>
            </tr>
            ${tests.map(test => `
              <tr>
                <td>${test.testNumber}</td>
                <td><input type="text" id="name_${test.testNumber}" value="${test.name.replace(/"/g, '\\"')}" data-test-number="${test.testNumber}"></td>
                <td><input type="number" id="timeLimit_${test.testNumber}" value="${test.timeLimit}" data-test-number="${test.testNumber}"></td>
                <td><input type="checkbox" id="randomQuestions_${test.testNumber}" ${test.randomQuestions ? 'checked' : ''} data-test-number="${test.testNumber}"></td>
                <td><input type="checkbox" id="randomAnswers_${test.testNumber}" ${test.randomAnswers ? 'checked' : ''} data-test-number="${test.testNumber}"></td>
                <td><input type="number" id="questionLimit_${test.testNumber}" value="${test.questionLimit || ''}" data-test-number="${test.testNumber}"></td>
                <td><input type="number" id="attemptLimit_${test.testNumber}" value="${test.attemptLimit}" data-test-number="${test.testNumber}"></td>
                <td><input type="checkbox" id="isQuickTest_${test.testNumber}" ${test.isQuickTest ? 'checked' : ''} data-test-number="${test.testNumber}"></td>
                <td><input type="number" id="timePerQuestion_${test.testNumber}" value="${test.timePerQuestion || ''}" data-test-number="${test.testNumber}"></td>
                <td>
                  <button class="save-btn" onclick="saveTest('${test.testNumber}')">Зберегти</button>
                  <button class="delete-btn" onclick="deleteTest('${test.testNumber}')">Видалити</button>
                </td>
              </tr>
            `).join('')}
          </table>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <div id="error-message" class="error"></div>
          <script>
            async function saveTest(testNumber) {
              const name = document.getElementById('name_' + testNumber).value;
              const timeLimit = parseInt(document.getElementById('timeLimit_' + testNumber).value);
              const randomQuestions = document.getElementById('randomQuestions_' + testNumber).checked;
              const randomAnswers = document.getElementById('randomAnswers_' + testNumber).checked;
              const questionLimit = parseInt(document.getElementById('questionLimit_' + testNumber).value) || null;
              const attemptLimit = parseInt(document.getElementById('attemptLimit_' + testNumber).value);
              const isQuickTest = document.getElementById('isQuickTest_' + testNumber).checked;
              const timePerQuestion = parseInt(document.getElementById('timePerQuestion_' + testNumber).value) || null;
              const errorMessage = document.getElementById('error-message');

              if (!name || name.length > 100) {
                errorMessage.textContent = 'Назва тесту має бути від 1 до 100 символів';
                return;
              }
              if (timeLimit < 60 || isNaN(timeLimit)) {
                errorMessage.textContent = 'Ліміт часу має бути не менше 60 секунд';
                return;
              }
              if (attemptLimit < 1 || isNaN(attemptLimit)) {
                errorMessage.textContent = 'Ліміт спроб має бути не менше 1';
                return;
              }
              if (isQuickTest && (!timePerQuestion || timePerQuestion < 5)) {
                errorMessage.textContent = 'Час на питання має бути не менше 5 секунд для швидкого тесту';
                return;
              }

              const formData = new URLSearchParams();
              formData.append('testNumber', testNumber);
              formData.append('name', name);
              formData.append('timeLimit', timeLimit);
              formData.append('randomQuestions', randomQuestions);
              formData.append('randomAnswers', randomAnswers);
              formData.append('questionLimit', questionLimit || '');
              formData.append('attemptLimit', attemptLimit);
              formData.append('isQuickTest', isQuickTest);
              formData.append('timePerQuestion', timePerQuestion || '');
              formData.append('_csrf', '${res.locals._csrf}');

              try {
                const response = await fetch('/admin/edit-tests', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  window.location.reload();
                } else {
                  errorMessage.textContent = result.message || 'Помилка збереження тесту';
                }
              } catch (error) {
                console.error('Помилка:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              }
            }

            async function deleteTest(testNumber) {
              if (!confirm('Ви впевнені, що хочете видалити цей тест?')) return;
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/admin/delete-test/' + testNumber, {
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
                alert('Не вдалося видалити тест.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-tests', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки редагування тестів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-tests виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber, name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    if (!testNumber || !name || name.length > 100) {
      return res.status(400).json({ success: false, message: 'Невірна назва тесту або номер' });
    }
    if (parseInt(timeLimit) < 60) {
      return res.status(400).json({ success: false, message: 'Ліміт часу має бути не менше 60 секунд' });
    }
    if (parseInt(attemptLimit) < 1) {
      return res.status(400).json({ success: false, message: 'Ліміт спроб має бути не менше 1' });
    }
    if (isQuickTest === 'true' && (!timePerQuestion || parseInt(timePerQuestion) < 5)) {
      return res.status(400).json({ success: false, message: 'Час на питання має бути не менше 5 секунд' });
    }

    const testData = {
      name,
      timeLimit: parseInt(timeLimit),
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: timePerQuestion ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    await loadTestsFromMongoDB();
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка збереження тесту', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при збереженні тесту' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-tests (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/delete-test/:testNumber', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const testNumber = req.params.testNumber;
    if (!testNames[testNumber]) {
      return res.status(404).json({ success: false, message: 'Тест не знайдено' });
    }

    // Видаляємо тест із усіх розділів
    await db.collection('sections').updateMany(
      { tests: testNumber },
      { $pull: { tests: testNumber } }
    );

    await deleteTestFromMongoDB(testNumber);
    await db.collection('questions').deleteMany({ testNumber });
    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    await loadTestsFromMongoDB();
    logger.info('Тест видалено', { testNumber });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення тесту', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-test/:testNumber виконано', { duration: `${endTime - startTime} мс` });
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
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Створити новий тест</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            input, select { margin: 10px; padding: 10px; font-size: 16px; width: 300px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .submit-btn:hover { background-color: #45a049; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .error { color: red; }
            @media (max-width: 600px) {
              input, select { width: 90%; font-size: 14px; }
              button { width: 90%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Створити новий тест</h1>
          <form id="create-test-form">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="text" name="name" placeholder="Назва тесту" required><br>
            <input type="number" name="timeLimit" placeholder="Ліміт часу (сек)" required><br>
            <label><input type="checkbox" name="randomQuestions"> Випадкові питання</label><br>
            <label><input type="checkbox" name="randomAnswers"> Випадкові відповіді</label><br>
            <input type="number" name="questionLimit" placeholder="Ліміт питань (необов’язково)"><br>
            <input type="number" name="attemptLimit" placeholder="Ліміт спроб" required><br>
            <label><input type="checkbox" name="isQuickTest"> Швидкий тест</label><br>
            <input type="number" name="timePerQuestion" placeholder="Час на питання (сек, для швидкого тесту)"><br>
            <button type="submit" class="submit-btn">Створити</button>
          </form>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <div id="error-message" class="error"></div>
          <script>
            document.getElementById('create-test-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const form = e.target;
              const formData = new URLSearchParams(new FormData(form));
              const errorMessage = document.getElementById('error-message');
              const submitBtn = form.querySelector('.submit-btn');
              submitBtn.disabled = true;
              submitBtn.textContent = 'Створення...';
              try {
                const response = await fetch('/admin/create-test', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                const result = await response.json();
                if (result.success) {
                  window.location.href = '/admin/edit-tests';
                } else {
                  errorMessage.textContent = result.message || 'Помилка створення тесту';
                }
              } catch (error) {
                console.error('Помилка:', error);
                errorMessage.textContent = 'Помилка: ' + error.message;
              } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = 'Створити';
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/create-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні форми створення тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/create-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

app.post('/admin/create-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    if (!name || name.length > 100) {
      return res.status(400).json({ success: false, message: 'Назва тесту має бути від 1 до 100 символів' });
    }
    if (parseInt(timeLimit) < 60) {
      return res.status(400).json({ success: false, message: 'Ліміт часу має бути не менше 60 секунд' });
    }
    if (parseInt(attemptLimit) < 1) {
      return res.status(400).json({ success: false, message: 'Ліміт спроб має бути не менше 1' });
    }
    if (isQuickTest === 'true' && (!timePerQuestion || parseInt(timePerQuestion) < 5)) {
      return res.status(400).json({ success: false, message: 'Час на питання має бути не менше 5 секунд' });
    }

    const existingTests = await db.collection('tests').find({}).toArray();
    const newTestNumber = (Math.max(...existingTests.map(t => parseInt(t.testNumber)), 0) + 1).toString();

    const testData = {
      name,
      timeLimit: parseInt(timeLimit),
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: timePerQuestion ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(newTestNumber, testData);
    await loadTestsFromMongoDB();
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка створення тесту', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при створенні тесту' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/create-test (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для журналу дій
app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;

    const logs = await db.collection('activity_log')
      .find({})
      .sort({ timestamp: -1 })
      .skip(skip)
      .limit(limit)
      .toArray();

    const totalLogs = await db.collection('activity_log').countDocuments();
    const totalPages = Math.ceil(totalLogs / limit);

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Журнал дій</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            table { border-collapse: collapse; width: 90%; margin: 20px auto; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            .back-btn { background-color: #007bff; color: white; }
            .back-btn:hover { background-color: #0056b3; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
            .details { max-width: 300px; overflow-wrap: break-word; }
            @media (max-width: 600px) {
              table { width: 100%; font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <h1>Журнал дій</h1>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Дія</th>
              <th>Час</th>
              <th>IP-адреса</th>
              <th>Деталі</th>
            </tr>
            ${logs.map(log => `
              <tr>
                <td>${log.user}</td>
                <td>${log.action}</td>
                <td>${new Date(log.timestamp).toLocaleString('uk-UA')}</td>
                <td>${log.ipAddress}</td>
                <td class="details">${JSON.stringify(log.additionalInfo || {}, null, 2).replace(/"/g, '&quot;')}</td>
              </tr>
            `).join('')}
          </table>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/activity-log?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/activity-log?page=${page + 1}">Наступна</a>` : ''}
          </div>
          <button class="back-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/activity-log', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні журналу дій');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/activity-log виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Запуск сервера
const port = process.env.PORT || 3000;
app.listen(port, () => {
  logger.info(`Сервер запущено на порту ${port}`);
});

// Обробка помилок
app.use((err, req, res, next) => {
  logger.error('Неперехоплена помилка', { message: err.message, stack: err.stack, url: req.url });
  res.status(500).send('Внутрішня помилка сервера');
});
