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
const Tokens = require('csrf');
const tokens = new Tokens();
const MongoStore = require('connect-mongo');

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
    pass: process.env.EMAIL_PASS || ':bnnz<fnmrsdobysxtcnmysrjve'
  }
});

// Функція для відправки email про підозрілу активність
const sendSuspiciousActivityEmail = async (user, activityDetails) => {
  try {
    const mailOptions = {
      from: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
      to: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
      subject: 'Підозріла активність у системі тестування',
      text: `
        Користувач: ${user}
        Деталі активності:
        Час поза вкладкою: ${activityDetails.timeAwayPercent}%
        Переключення вкладок: ${activityDetails.switchCount}
        Середній час відповіді (сек): ${activityDetails.avgResponseTime}
        Загальна кількість дій: ${activityDetails.totalActivityCount}
      `
    };
    await transporter.sendMail(mailOptions);
    logger.info(`Email про підозрілу активність відправлено для користувача ${user}`);
  } catch (error) {
    logger.error('Помилка відправки email', { message: error.message, stack: error.stack });
  }
};

// Конфігурація параметрів підозрілої активності
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
      logger.info(`Трафік кешу для ${cacheKey}`);
      return this.cache[cacheKey];
    }
    logger.info(`Промах кешу для ${cacheKey}, завантаження з БД`);
    const startTime = Date.now();
    const data = await fetchFn();
    this.cache[cacheKey] = data;
    logger.info(`Оновлено кеш ${key} для тесту ${testNumber} за ${Date.now() - startTime} мс`);
    return data;
  }

  static async invalidateCache(key, testNumber) {
    const cacheKey = `${key}:${testNumber}`;
    delete this.cache[cacheKey];
    logger.info(`Інвалідовано кеш для ${cacheKey}`);
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

// Підключення до MongoDB
const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    logger.info(`Спроба підключення до MongoDB (${attempt}/${maxAttempts})`);
    const startTime = Date.now();
    await client.connect();
    logger.info(`Підключено до MongoDB за ${Date.now() - startTime} мс`);
    db = client.db('alpha');
    logger.info('Ініціалізовано базу даних', { databaseName: db.databaseName });
  } catch (error) {
    logger.error('Помилка підключення до MongoDB', { message: error.message, stack: error.stack });
    if (attempt < maxAttempts) {
      logger.info('Повторна спроба через 5 секунд...');
      await new Promise(resolve => setTimeout(resolve, 5000));
      return connectToMongoDB(attempt + 1, maxAttempts);
    }
    throw error;
  }
};

let isInitialized = false;
let initializationError = null;
let testNames = {};

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
    logger.info(`Завантажено ${tests.length} тестів із MongoDB`);
  } catch (error) {
    logger.error('Помилка завантаження тестів', { message: error.message, stack: error.stack });
    throw error;
  }
};

const saveTestToMongoDB = async (testNumber, testData) => {
  try {
    await db.collection('tests').updateOne(
      { testNumber },
      {
        $set: {
          testNumber,
          name: testData.name,
          timeLimit: testData.timeLimit,
          randomQuestions: testData.randomQuestions,
          randomAnswers: testData.randomAnswers,
          questionLimit: testData.questionLimit,
          attemptLimit: testData.attemptLimit,
          isQuickTest: testData.isQuickTest || false,
          timePerQuestion: testData.timePerQuestion || null
        }
      },
      { upsert: true }
    );
    logger.info('Тест збережено', { testNumber });
  } catch (error) {
    logger.error('Помилка збереження тесту', { message: error.message, stack: error.stack });
    throw error;
  }
};

const deleteTestFromMongoDB = async (testNumber) => {
  try {
    await db.collection('tests').deleteOne({ testNumber });
    logger.info('Тест видалено', { testNumber });
  } catch (error) {
    logger.error('Помилка видалення тесту', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

app.use(session({
  secret: process.env.SESSION_SECRET || 'your-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    client: client,
    dbName: 'alpha',
    collectionName: 'sessions',
    ttl: 24 * 60 * 60
  }),
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
    maxAge: 24 * 60 * 60 * 1000
  }
}));

app.use((req, res, next) => {
  logger.info('Запит отримано', { url: req.url, method: req.method, userRole: req.userRole, timestamp: new Date().toISOString() });
  next();
});

app.use((req, res, next) => {
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = tokens.secretSync();
    logger.info('Згенеровано CSRF-секрет', { secret: req.session.csrfSecret });
  }
  const token = tokens.create(req.session.csrfSecret);
  res.locals._csrf = token;
  res.cookie('XSRF-TOKEN', token, { httpOnly: false });
  logger.info('Згенеровано CSRF-токен', { token });
  next();
});

app.use((req, res, next) => {
  if (['POST', 'PUT', 'DELETE'].includes(req.method)) {
    const token = req.body._csrf || req.headers['x-csrf-token'];
    if (!token) {
      logger.error('Відсутній CSRF-токен', { method: req.method, url: req.url });
      return res.status(403).json({ success: false, message: 'CSRF-токен відсутній' });
    }
    if (!req.session.csrfSecret) {
      logger.error('Відсутній CSRF-секрет', { sessionId: req.sessionID });
      return res.status(403).json({ success: false, message: 'Помилка сесії' });
    }
    if (!tokens.verify(req.session.csrfSecret, token)) {
      logger.error('Помилка валідації CSRF', { expectedSecret: req.session.csrfSecret, receivedToken: token });
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }
    logger.info('CSRF-токен валідовано', { token });
  }
  next();
});

app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

app.use((err, req, res, next) => {
  if (err.name === 'MongoNetworkError' || err.name === 'MongoServerError') {
    logger.error('Помилка MongoDB', { message: err.message, stack: err.stack });
    res.status(503).json({ success: false, message: 'Помилка з’єднання з БД' });
  } else {
    next(err);
  }
});

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
              (e.ctrlKey && (e.key === 'p' || e.key === 'P' || e.key === 's' || e.key === 'S')) ||
              (e.metaKey && (e.key === 'p' || e.key === 'P' || e.key === 's' || e.key === 'S')) ||
              (e.altKey && e.key === 'PrintScreen') ||
              (e.metaKey && e.shiftKey && (e.key === '3' || e.key === '4'))
            ) {
              e.preventDefault();
            }
          });
          document.addEventListener('contextmenu', (e) => e.preventDefault());
          document.addEventListener('selectstart', (e) => e.preventDefault());
          document.addEventListener('copy', (e) => e.preventDefault());
        </script>
      `;
      body = body.replace('</body>', `${watermarkScript}</body>`);
    }
    return originalSend.call(this, body);
  };
  next();
});

const importUsersToMongoDB = async (buffer) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    let sheet = workbook.getWorksheet('Users') || workbook.getWorksheet('Sheet1');
    if (!sheet) {
      throw new Error('Лист "Users" або "Sheet1" не знайдено');
    }
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
    if (users.length === 0) {
      throw new Error('Не знайдено користувачів');
    }
    await db.collection('users').deleteMany({});
    logger.info('Очищено всіх користувачів');
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

const importQuestionsToMongoDB = async (buffer, testNumber) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      throw new Error('Лист "Questions" не знайдено');
    }
    const MAX_ROWS = 1000;
    if (sheet.rowCount > MAX_ROWS + 1) {
      throw new Error(`Занадто багато рядків (${sheet.rowCount - 1}). Максимум: ${MAX_ROWS}`);
    }
    const questions = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        try {
          const rowValues = row.values.slice(1);
          let questionText = rowValues[1];
          if (typeof questionText === 'object' && questionText !== null) {
            questionText = questionText.text || questionText.value || '[Невірний текст]';
          }
          questionText = String(questionText || '').trim();
          if (questionText === '') throw new Error('Текст питання відсутній');
          const picture = String(rowValues[0] || '').trim();
          let options = rowValues.slice(2, 14).filter(Boolean).map(val => String(val).trim());
          const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(val => String(val).trim());
          const type = String(rowValues[26] || 'multiple').toLowerCase();
          const points = Number(rowValues[27]) || 1;
          const variant = String(rowValues[28] || '').trim();

          if (type === 'truefalse') {
            options = ["Правда", "Неправда"];
          }

          const normalizedPicture = picture
            ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '')
            : null;

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
              const pictureNumber = pictureMatch[1];
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
                    logger.info(`Зображення знайдено: ${questionData.picture}`);
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
            if (questionData.pairs.length === 0) throw new Error('Для Matching потрібні пари');
            questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
          }

          if (type === 'fillblank') {
            questionText = questionText.replace(/\s*___\s*/g, '___');
            const blankCount = (questionText.match(/___/g) || []).length;
            if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
              throw new Error('Пропуски не відповідають правильним відповідям');
            }
            questionData.text = questionText;
            questionData.blankCount = blankCount;

            questionData.correctAnswers.forEach((correctAnswer, idx) => {
              if (correctAnswer.includes('-')) {
                const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
                if (isNaN(min) || isNaN(max) || min > max) {
                  throw new Error(`Невірний формат діапазону для відповіді ${idx + 1}`);
                }
              } else {
                const value = parseFloat(correctAnswer);
                if (isNaN(value)) {
                  throw new Error(`Відповідь ${idx + 1} має бути числом або діапазоном`);
                }
              }
            });
          }

          if (type === 'singlechoice') {
            if (correctAnswers.length !== 1 || options.length < 2) {
              throw new Error('Single Choice: одна правильна відповідь і мінімум 2 варіанти');
            }
            questionData.correctAnswer = correctAnswers[0];
          }

          if (type === 'input') {
            if (questionData.correctAnswers.length !== 1) {
              throw new Error(`Input: потрібна одна правильна відповідь`);
            }
            const correctAnswer = questionData.correctAnswers[0];
            if (correctAnswer.includes('-')) {
              const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
              if (isNaN(min) || isNaN(max) || min > max) {
                throw new Error(`Невірний формат діапазону`);
              }
            } else {
              const value = parseFloat(correctAnswer);
              if (isNaN(value)) {
                throw new Error(`Відповідь для Input має бути числом або діапазоном`);
              }
            }
          }

          questions.push(questionData);
        } catch (error) {
          throw new Error(`Помилка в рядку ${rowNumber}: ${error.message}`);
        }
      }
    });
    if (questions.length === 0) {
      throw new Error('Не знайдено питань');
    }
    await db.collection('questions').deleteMany({ testNumber });
    await db.collection('questions').insertMany(questions);
    logger.info(`Імпортовано ${questions.length} питань для тесту ${testNumber}`);
    await CacheManager.invalidateCache('questions', testNumber);
    return questions.length;
  } catch (error) {
    logger.error('Помилка імпорту питань', { message: error.message, stack: error.stack });
    throw error;
  }
};

const shuffleArray = (array) => {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
};

const loadUsersToCache = async () => {
  try {
    const startTime = Date.now();
    userCache = await db.collection('users').find({}).toArray();
    logger.info(`Оновлено кеш користувачів: ${userCache.length} за ${Date.now() - startTime} мс`);
  } catch (error) {
    logger.error('Помилка оновлення кешу користувачів', { message: error.message, stack: error.stack });
    throw error;
  }
};

const loadQuestions = async (testNumber) => {
  try {
    const startTime = Date.now();
    if (questionsCache[testNumber]) {
      logger.info(`Завантажено ${questionsCache[testNumber].length} питань з кешу за ${Date.now() - startTime} мс`);
      return questionsCache[testNumber];
    }
    const questions = await db.collection('questions').find({ testNumber: testNumber.toString() }).sort({ order: 1 }).toArray();
    if (questions.length === 0) {
      throw new Error(`Не знайдено питань для тесту ${testNumber}`);
    }
    questionsCache[testNumber] = questions;
    logger.info(`Завантажено ${questions.length} питань із MongoDB за ${Date.now() - startTime} мс`);
    return questions;
  } catch (error) {
    logger.error(`Помилка в loadQuestions`, { message: error.message, stack: error.stack });
    throw error;
  }
};

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
  logger.info('Оновлено паролі', { duration: `${Date.now() - startTime} мс` });
  await CacheManager.invalidateCache('users', null);
};

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
    logger.info('Індекси MongoDB створено');

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
    if (testCount === 0) {
      const defaultTests = {
        "1": { name: "Тест 1", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "2": { name: "Тест 2", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "3": { name: "Тест 3", timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 }
      };
      for (const [testNumber, testData] of Object.entries(defaultTests)) {
        await saveTestToMongoDB(testNumber, testData);
      }
      logger.info('Міграція стандартних тестів', { count: Object.keys(defaultTests).length });
    }

    await updateUserPasswords();
    await loadUsersToCache();
    await loadTestsFromMongoDB();
    await CacheManager.invalidateCache('questions', null);
    isInitialized = true;
    initializationError = null;
  } catch (error) {
    logger.error('Помилка ініціалізації сервера', { message: error.message, stack: error.stack });
    initializationError = error;
    throw error;
  }
};

const cleanupActivityLog = async () => {
  try {
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const result = await db.collection('activity_log').deleteMany({
      timestamp: { $lt: thirtyDaysAgo.toISOString() }
    });
    logger.info('Очищено старі записи', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення журналу', { message: error.message, stack: error.stack });
  }
};

const cleanupActiveTests = async () => {
  try {
    const twentyFourHoursAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const result = await db.collection('active_tests').deleteMany({
      startTime: { $lt: twentyFourHoursAgo }
    });
    logger.info('Очищено старі тести', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення тестів', { message: error.message, stack: error.stack });
  }
};

setInterval(cleanupActivityLog, 24 * 60 * 60 * 1000);
setInterval(cleanupActiveTests, 24 * 60 * 60 * 1000);

(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
    await cleanupActivityLog();
    await cleanupActiveTests();
  } catch (error) {
    logger.error('Помилка запуску сервера', { message: error.message, stack: error.stack });
    process.exit(1);
  }
})();

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

const logActivity = async (user, action, ipAddress, additionalInfo = {}, session = null) => {
  try {
    const startTime = Date.now();
    const timestamp = new Date();
    const timeOffset = 3 * 60 * 60 * 1000;
    const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
    await db.collection('activity_log').insertOne({
      user,
      action,
      ipAddress,
      timestamp: adjustedTimestamp.toISOString(),
      additionalInfo
    }, { session });
    logger.info(`Залогована активність: ${user} - ${action}`, { duration: `${Date.now() - startTime} мс` });
  } catch (error) {
    logger.error('Помилка логування', { message: error.message, stack: error.stack });
    throw error;
  }
};

app.get('/test-mongo', async (req, res) => {
  try {
    if (!db) {
      throw new Error('Підключення до MongoDB не встановлено');
    }
    await db.collection('users').findOne();
    res.json({ success: true, message: 'Підключення до MongoDB успішне' });
  } catch (error) {
    logger.error('Тест MongoDB провалився', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка підключення', error: error.message });
  }
});

app.get('/api/test', (req, res) => {
  logger.info('Обробка /api/test');
  res.json({ success: true, message: 'Сервер Express працює' });
});

app.get('/favicon.ico', (req, res) => {
  res.status(204).end();
});

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
                  errorMessage.textContent = result.message || 'Перевищено ліміт спроб';
                } else if (response.status === 400) {
                  errorMessage.textContent = result.message || 'Некоректні дані';
                } else if (response.status === 401) {
                  errorMessage.textContent = result.message || 'Невірний логін або пароль';
                } else {
                  errorMessage.textContent = result.message || 'Помилка входу';
                }
              }
            } catch (error) {
              console.error('Помилка входу:', error);
              errorMessage.textContent = 'Не вдалося підключитися';
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

app.post('/login', [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Логін від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Логін: лише літери та цифри'),
  body('password')
    .isLength({ min: 6, max: 100 }).withMessage('Пароль від 6 до 100 символів')
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
      logger.warn('Кеш користувачів порожній');
      await loadUsersToCache();
      if (userCache.length === 0) {
        logger.error('Користувачів не знайдено');
        throw new Error('Не вдалося завантажити користувачів');
      }
    }

    const foundUser = userCache.find(user => user.username === username);
    logger.info('Користувача знайдено', { username });

    if (!foundUser) {
      logger.warn('Користувача не знайдено', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    const passwordMatch = await bcrypt.compare(password, foundUser.password);
    if (!passwordMatch) {
      logger.warn('Невірний пароль', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    await checkLoginAttempts(ipAddress, true);

    const token = jwt.sign(
      { username: foundUser.username, role: foundUser.role },
      process.env.JWT_SECRET || 'your-secret-key',
      { expiresIn: '24h' }
    );

    res.cookie('token', token, {
      httpOnly: true,
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
    res.status(error.message.includes('Перевищено ліміт') ? 429 : 500).json({ success: false, message: error.message });
  } finally {
    logger.info('Маршрут /login виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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

const checkAdmin = (req, res, next) => {
  if (req.userRole !== 'admin') {
    return res.status(403).send('Доступно тільки для адміністратора');
  }
  next();
};

app.get('/select-test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    if (Object.keys(testNames).length === 0) {
      logger.warn('testNames порожній');
      await loadTestsFromMongoDB();
      if (Object.keys(testNames).length === 0) {
        logger.error('Тести не знайдено');
        throw new Error('Не вдалося завантажити тести');
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
            body { font-family: Arial, sans-serif; text-align: center; padding: 20px; padding-bottom: 80px; margin: 0; }
            h1 { font-size: 24px; margin-bottom: 20px; }
            .test-buttons { display: flex; flex-direction: column; align-items: center; gap: 10px; }
            button { padding: 10px; font-size: 18px; cursor: pointer; width: 200px; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
            #logout { background-color: #ef5350; color: white; position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); width: 200px; }
            .no-tests { color: red; font-size: 18px; margin-top: 20px; }
            .results-btn { background-color: #007bff; color: white; margin-top: 20px; }
            @media (max-width: 600px) {
              h1 { font-size: 28px; }
              button { font-size: 20px; width: 90%; padding: 15px; }
              #logout { width: 90%; }
            }
          </style>
        </head>
        <body>
          <h1>Виберіть тест</h1>
          <div class="test-buttons">
            ${Object.entries(testNames).length > 0
              ? Object.entries(testNames).map(([num, data]) => `
                  <button onclick="window.location.href='/test?test=${num}'">${data.name.replace(/"/g, '\\"')}</button>
                `).join('')
              : '<p class="no-tests">Немає доступних тестів</p>'
            }
            ${req.userRole === 'instructor' ? `
              <button class="results-btn" onclick="window.location.href='/admin/results'">Переглянути результати</button>
            ` : ''}
          </div>
          <button id="logout" onclick="logout()">Вихийти</button>
          <script>
            async function logout() {
              console.log('POST /logout, CSRF-токен:', '${res.locals._csrf}');
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
                  throw new Error('HTTP error: ' + response.status);
                }
                const result = await response.json();
                console.log('Відповідь:', result);
                if (result.success) {
                  window.location.href = '/';
                } else {
                  throw new Error('Вихід не вдався: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка виходу:', error);
                alert('Не вдалося вийти');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /select-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка вибору тесту');
  } finally {
    logger.info('Маршрут /select-test виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/logout', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Отримано CSRF-токен', { token: req.body._csrf });
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    logActivity(req.user, 'покинув сайт', ipAddress);
    res.clearCookie('token');
    req.session.destroy(err => {
      if (err) {
        logger.error('Помилка знищення сесії', { message: err.message, stack: err.stack });
        return res.status(500).json({ success: false, message: 'Помилка завершення сесії' });
      }
      logger.info('Сесію знищено');
      res.json({ success: true });
    });
  } catch (error) {
    logger.error('Помилка в /logout', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при виході' });
  } finally {
    logger.info('Маршрут /logout виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress, testSessionId, attempt = 1, maxAttempts = 3) => {
  const startTimeLog = Date.now();
  const session = client.startSession();
  try {
    await session.withTransaction(async () => {
      const duration = Math.round((endTime - startTime) / 1000);
      const timeOffset = 3 * 60 * 60 * 1000;
      const adjustedStartTime = new Date(startTime + timeOffset);
      const adjustedEndTime = new Date(endTime + timeOffset);

      const result = {
        user,
        testNumber,
        score,
        totalPoints,
        totalClicks,
        correctClicks,
        totalQuestions,
        percentage,
        startTime: adjustedStartTime.toISOString(),
        endTime: adjustedEndTime.toISOString(),
        duration,
        answers: Object.fromEntries(Object.entries(answers).sort((a, b) => parseInt(a[0]) - parseInt(b[0]))),
        scoresPerQuestion,
        suspiciousActivity,
        variant: `Variant ${variant}`,
        testSessionId
      };
      logger.info('Збереження результату', { testSessionId });
      await db.collection('test_results').insertOne(result, { session });
      await logActivity(user, `завершив тест ${testNames[testNumber].name.replace(/"/g, '\\"')} з результатом ${Math.round(percentage)}%`, ipAddress, { percentage: Math.round(percentage) }, session);
    });
  } catch (error) {
    logger.error(`Помилка збереження (спроба ${attempt})`, { message: error.message, stack: error.stack });
    if (attempt < maxAttempts) {
      logger.info(`Повторна спроба (${attempt + 1}/${maxAttempts})`);
      await new Promise(resolve => setTimeout(resolve, 1000));
      return saveResult(user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress, testSessionId, attempt + 1, maxAttempts);
    }
    throw error;
  } finally {
    await session.endSession();
    logger.info('saveResult виконано', { duration: `${Date.now() - startTimeLog} мс` });
  }
};

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

    logger.info(`Користувач ${user} має ${attemptLimit - attempts} спроб`);

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
    logger.error('Помилка перевірки спроб', { message: error.message, stack: error.stack });
    throw error;
  }
};

app.get('/test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');
    const testNumber = req.query.test;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).send('Номер тесту не вказано');
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
              button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
              button:hover { background-color: #45a049; }
              .overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 999; }
              h2 { margin-bottom: 20px; font-size: 24px; color: #333; }
            </style>
          </head>
          <body>
            <div class="overlay"></div>
            <div id="modal">
              <h2>Ви вже проходили цей тест</h2>
              <button onclick="window.location.href='/select-test'">Повернутися до вибору</button>
            </div>
          </body>
        </html>
      `);
    }

    let questions = await loadQuestions(testNumber);
    const userVariant = Math.floor(Math.random() * 3) + 1;
    logger.info(`Призначено варіант ${userVariant} для ${req.user}`);

    questions = questions.filter(q => !q.variant || q.variant === '' || q.variant === `Variant ${userVariant}`);
    logger.info(`Фільтровано ${questions.length} питань`);

    if (questions.length === 0) {
      return res.status(400).send(`Немає питань для варіанту ${userVariant}`);
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
      testSessionId,
      isSavingResult: false,
      isSaved: false,
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
    res.status(500).send('Помилка завантаження тесту: ' + error.message);
  } finally {
    logger.info('Маршрут /test виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/test/question', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');

    let userTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!userTest) {
      return res.status(400).send('Тест не розпочато');
    }

    const { questions, testNumber, answers, currentQuestion, startTime: testStartTime, timeLimit, isQuickTest, timePerQuestion, suspiciousActivity, variant, testSessionId } = userTest;

    if (!testNames[testNumber]) {
      let score = 0;
      const totalPoints = questions.reduce((sum, q) => sum + q, 0);
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
        logger.info(`Результат збережено для testSessionId: ${testSessionId}`);
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
            <h2>Тест недоступний. Ваші відповіді збережено.</h2>
            <button onclick="window.location.href='/select-test'">Повернутися</button>
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
              .progress-bar { gap: 2px; }
              .progress-circle { width: 20px; height: 20px; font-size: 8px; }
              .progress-line { width: 3px; }
              button { font-size: 16px; padding: 10px; }
              #timer { font-size: 20px; }
              .question-box h2 { font-size: 18px; }
              .matching-container { flex-direction: column; }
              .matching-column { width: 100%; }
              .blank-input { width: 80px; }
              .option-box, .matching-item { font-size: 14px; padding: 8px; min-height: 40px; line-height: 1.5; }
            }
            @media (min-width: 401px) and (max-width: 600px) {
              h1 { font-size: 28px; }
              .progress-bar { gap: 3px; }
              .progress-circle { width: 25px; height: 25px; font-size: 10px; }
              .progress-line { width: 4px; }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 20px; }
              .question-box h2 { font-size: 20px; }
              .matching-container { flex-direction: column; }
              .matching-column { width: 100%; }
              .blank-input { width: 80px; }
              .option-box, .matching-item { font-size: 18px; padding: 10px; min-height: 50px; line-height: 1.5; }
            }
            @media (min-width: 601px) and (max-width: 900px) {
              h1 { font-size: 30px; }
              .progress-bar { gap: 4px; }
              .progress-circle { width: 30px; height: 30px; font-size: 12px; }
              .progress-line { width: 5px; }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 22px; }
              .question-box h2 { font-size: 22px; }
              .matching-column { width: 45%; }
              .blank-input { width: 100px; }
              .option-box, .matching-item { font-size: 18px; padding: 10px; min-height: 50px; line-height: 1.5; }
            }
            @media (min-width: 901px) and (max-width: 1200px) {
              h1 { font-size: 32px; }
              .progress-bar { gap: 5px; }
              .progress-circle { width: 35px; height: 35px; font-size: 14px; }
              .progress-line { width: 5px; }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 24px; }
              .question-box h2 { font-size: 24px; }
              .matching-column { width: 45%; }
              .blank-input { width: 100px; }
              .option-box, .matching-item { font-size: 18px; padding: 10px; min-height: 50px; line-height: 1.5; }
            }
            @media (min-width: 1201px) {
              h1 { font-size: 36px; }
              .progress-bar { gap: 6px; }
              .progress-circle { width: 40px; height: 40px; font-size: 16px; }
              .progress-line { width: 6px; }
              button { font-size: 20px; padding: 15px; }
              #timer { font-size: 26px; }
              .question-box h2 { font-size: 26px; }
              .matching-column { width: 45%; }
              .blank-input { width: 120px; }
              .option-box, .matching-item { font-size: 20px; padding: 12px; min-height: 60px; line-height: 1.5; }
            }
          </style>
        </head>
        <body>
          <h1>${testNames[testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</h1>
          <div id="timer">Залишилось часу: ${minutes} хв ${seconds} с</div>
          <div class="progress-bar">
            ${progress.map((p, j) => `
              <div class="progress-circle ${p.answered ? 'answered' : 'unanswered'}">${p.number}</div>
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
              <button class="back-btn" ${index === 0 ? 'disabled' : ''} onclick="window.location.href='/test/question?index=${index - 1}'">Назад</button>
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
            let switchCount = ${userTest.suspiciousActivity?.switchCount || 0};
            let lastHiddenTime = 0;
            let lastActivityTime = Date.now();
            let activityCount = ${userTest.suspiciousActivity?.activityCounts?.[index] || 0};
            let lastMouseMoveTime = 0;
            let lastKeydownTime = 0;
            const debounceDelay = 100;
            let selectedOptions = ${selectedOptionsString};
            let matchingPairs = ${JSON.stringify(answers[index] || [])};
            let questionTimeRemaining = timePerQuestion;
            let currentQuestionIndex = ${index};
            let lastGlobalUpdateTime = Date.now();
            let isSaving = false;
            let hasMovedToNext = false;
            let questionStartTime = ${questionStartTime[index]};

            async function saveCurrentAnswer(index, attempt = 1, maxAttempts = 3) {
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

                console.log('Автозбереження:', { index, answers: safeAnswer, timeAway, switchCount, responseTime });

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  if (attempt < maxAttempts) {
                    console.warn('Спроба ${attempt} невдала, повтор через 1с');
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    return saveCurrentAnswer(index, attempt + 1, maxAttempts);
                  }
                  throw new Error('HTTP-помилка: ' + response.status);
                }

                const result = await response.json();
                if (!result.success) {
                  throw new Error('Помилка збереження: ' + result.error);
                }
                return result;
              } catch (error) {
                console.error('Помилка автозбереження:', error);
                throw error;
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

                console.log('Збереження в saveAndNext:', { timeAway, switchCount, responseTime, answer: safeAnswer });

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  throw new Error('HTTP-помилка: ' + response.status);
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
                      window.location.href = '/test/question?index=' + nextIndex;
                    } else {
                      window.location.href = '/result';
                    }
                  });
                } else {
                  console.error('Помилка збереження:', result.error);
                  alert('Помилка збереження: ' + result.error);
                }
              } catch (error) {
                console.error('Помилка в saveAndNext:', error);
                alert('Не вдалося зберегти: ' + error.message);
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
              if (isSaving) return;
              isSaving = true;
              try {
                await saveCurrentAnswer(index);
                window.location.href = '/result';
              } catch (error) {
                console.error('Помилка в finishTest:', error);
                alert('Не вдалося завершити: ' + error.message);
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
                saveCurrentAnswer(currentQuestionIndex).then(() => {
                  window.location.href = '/result';
                }).catch(error => {
                  console.error('Помилка збереження при таймері:', error);
                  alert('Помилка збереження результатів');
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
                  clearInterval(questionTimerInterval);
                  saveCurrentAnswer(currentQuestionIndex).then(() => {
                    window.location.href = '/result';
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

            document.addEventListener('visibilitychange', () => {
              console.log(\`visibilitychange: document.hidden = \${document.hidden}, timeAway = \${timeAway}, switchCount = \${switchCount}\`);
              if (document.hidden) {
                if (!lastHiddenTime) {
                  lastHiddenTime = performance.now();
                  switchCount = Math.min(switchCount + 1, 1000);
                  console.log('Вкладка прихована:', lastHiddenTime, 'Переключень:', switchCount);
                }
              } else {
                if (lastHiddenTime > 0) {
                  const now = performance.now();
                  const awayDuration = Math.min((now - lastHiddenTime) / 1000, 60);
                  timeAway += awayDuration;
                  console.log('Вкладка видима, час поза вкладкою:', awayDuration, 'Загальний:', timeAway);
                  lastHiddenTime = 0;
                  saveCurrentAnswer(currentQuestionIndex);
                }
              }
            });

            setInterval(() => {
              if (timeAway > 0 || switchCount > 0) {
                saveCurrentAnswer(currentQuestionIndex);
              }
            }, 10000);

            document.addEventListener('visibilitychange', () => {
              if (!document.hidden) {
                const now = Date.now();
                const timeSinceLastActivity = (now - lastActivityTime) / 1000;
                if (timeSinceLastActivity > 300) {
                  questionStartTime = now;
                  console.log('Тривалий простій, скидання questionStartTime:', questionStartTime);
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
    logger.error('Помилка в /test/question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка сервера');
  } finally {
    logger.info('Маршрут /test/question виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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
    logger.info('Маршрут /set-question-start-time виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/answer', checkAuth, express.urlencoded({ extended: true }), async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;
    logger.info('Отримано дані в /answer', { index, timeAway, switchCount, responseTime, activityCount });

    if (!index || !answer) {
      logger.error('Відсутні параметри', { index, answer });
      return res.status(400).json({ success: false, error: 'Необхідно надати index та answer' });
    }

    let parsedAnswer;
    try {
      if (typeof answer === 'string') {
        if (answer.trim() === '') {
          parsedAnswer = [];
        } else {
          logger.info('Парсинг відповіді', { answer });
          parsedAnswer = JSON.parse(answer);
        }
      } else {
        parsedAnswer = answer;
      }
    } catch (error) {
      logger.error('Помилка парсингу', { answer, message: error.message });
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
        logger.error('Тест не розпочато', { user: req.user });
        return res.status(400).json({ success: false, error: 'Тест не розпочато' });
      }
    }

    userTest.answers[index] = parsedAnswer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = Math.max(0, parseFloat(timeAway) || userTest.suspiciousActivity.timeAway);
    userTest.suspiciousActivity.switchCount = Math.min(Math.max(0, parseInt(switchCount) || userTest.suspiciousActivity.switchCount), 1000);
    userTest.suspiciousActivity.responseTimes[index] = Math.max(0, parseFloat(responseTime) || 0);
    userTest.suspiciousActivity.activityCounts[index] = parseInt(activityCount) || 0;

    await db.collection('active_tests').updateOne(
      { user: req.user },
      { $set: { answers: userTest.answers, suspiciousActivity: userTest.suspiciousActivity } }
    );

    logger.info('Підозріла активність оновлена', { timeAway: userTest.suspiciousActivity.timeAway, switchCount: userTest.suspiciousActivity.switchCount });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка в /answer', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    logger.info('Маршрут /answer виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/result', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');

    const userTest = await db.collection('active_tests').findOne({ user: req.user });
    let testData;

    if (!userTest) {
      const recentResult = await db.collection('test_results').findOne(
        { user: req.user },
        { sort: { endTime: -1 } }
      );
      if (!recentResult) {
        return res.status(400).send('Тест не розпочато або перерваний');
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
          logger.info(`Порівняння input ${index + 1}`, {
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
          logger.info(`Fillblank ${index + 1}`, { userAnswers, correctAnswers });
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
          logger.info(`Single choice ${index + 1}`, { userAnswers, correctAnswer });
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
      logger.info(`Кориговано endTime для ${testSessionId}`);
    }

    const percentage = testData.percentage || (score / totalPoints) * 100;
    const totalClicks = testData.totalClicks || Object.keys(answers).length;
    const correctClicks = testData.correctClicks || scoresPerQuestion.filter(s => s > 0).length;
    const totalQuestions = testData.totalQuestions || (questions ? questions.length : 0);

    let duration = Math.round((endTime - testStartTime) / 1000);
    if (duration <= 0) {
      logger.warn('Некоректна тривалість', { endTime, testStartTime, duration });
      duration = 1;
    }

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

    logger.info('Підозріла активність в /result', {
      timeAway,
      correctedTimeAway,
      duration,
      timeAwayPercent,
      switchCount,
      suspiciousActivity
    });

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
        await db.collection('active_tests').updateOne(
          { user: req.user },
          { $set: { isSaved: true } }
        );
        logger.info(`Результат збережено: ${testSessionId}`);
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
      logger.error('Помилка читання A.png', { message: error.message, stack: error.stack });
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
            Час поза вкладкою: ${timeAwayPercent}%<br>
            Переключення вкладок: ${switchCount}
          </p>
          <div class="buttons">
            <button id="exportPDF">Експортувати в PDF</button>
            <button id="restart">Вихід</button>
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

            console.log('Результати:', {
              user, testName, totalQuestions, correctClicks, score, totalPoints, percentage, time, date
            });

            const exportPDFButton = document.getElementById('exportPDF');
            const restartButton = document.getElementById('restart');

            if (exportPDFButton) {
              exportPDFButton.addEventListener('click', () => {
                try {
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
                } catch (error) {
                  console.error('Помилка PDF:', error);
                  alert('Не вдалося згенерувати PDF');
                }
              });
            }

            if (restartButton) {
              restartButton.addEventListener('click', () => {
                window.location.href = '/select-test';
              });
            }
          </script>
        </body>
      </html>
    `;
    res.send(resultHtml);
  } catch (error) {
    logger.error('Помилка в /result', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка результатів: ' + error.message);
  } finally {
    logger.info('Маршрут /result виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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
          logger.info(`Fillblank ${index + 1} в /results`, { userAnswers, correctAnswers });
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
          logger.info(`Single choice ${index + 1} в /results`, { userAnswers, correctAnswer });
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
        logger.error('Помилка читання A.png', { message: error.message, stack: error.stack });
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
          <button id="restart">Повернутися на головну</button>
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
            window.location.href = '/';
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
    logger.info('Маршрут /results виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
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
            button { padding: 15px 30px; margin: 10px; font-size: 24px; cursor: pointer; width: 300px; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
            #logout { background-color: #ef5350; color: white; }
            @media (max-width: 600px) {
              body { padding: 20px; padding-bottom: 80px; }
              h1 { font-size: 32px; }
              button { font-size: 20px; width: 90%; padding: 15px; }
              #logout { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); width: 90%; }
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
          <button onclick="window.location.href='/admin/edit-tests'">Редагувати тести</button><br>
          <button onclick="window.location.href='/admin/create-test'">Створити тест</button><br>
          <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
          <button onclick="window.location.href='/admin/check-active-tests'">Перевірка активних тестів</button><br>
          <button id="logout" onclick="logout()">Вийти</button>
          <script>
            async function logout() {
              console.log('POST /logout, CSRF-токен:', '${res.locals._csrf}');
              const formData = new URLSearchParams();
              formData.append('_csrf', '${res.locals._csrf}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                console.log('Статус:', response.status);
                if (!response.ok) {
                  throw new Error('HTTP error: ' + response.status);
                }
                const result = await response.json();
                console.log('Відповідь:', result);
                if (result.success) {
                  window.location.href = '/';
                } else {
                  throw new Error('Вихід не вдався: ' + result.message);
                }
              } catch (error) {
                console.error('Помилка виходу:', error);
                alert('Не вдалося вийти');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    logger.info('Маршрут /admin виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/check-active-tests', checkAuth, checkAdmin, async (req, res) => {
  try {
    const activeTests = await db.collection('active_tests').find({}).toArray();
    res.json({
      success: true,
      count: activeTests.length,
      tests: activeTests.map(test => ({
        user: test.user,
        testNumber: test.testNumber,
        startTime: test.startTime,
        suspiciousActivity: test.suspiciousActivity
      }))
    });
  } catch (error) {
    logger.error('Помилка в /admin/check-active-tests', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  }
});

app.get('/admin/users', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let users = [];
    let errorMessage = '';
    try {
      users = await db.collection('users').find({}).toArray();
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
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Керування користувачами</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <button class="nav-btn" onclick="window.location.href='/admin/add-user'">Додати користувача</button>
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
                    alert('Помилка при видаленні: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення:', error);
                  alert('Не вдалося видалити користувача');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/users виконано', { duration: `${endTime - startTime} мс` });
  }
});

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
          <button class="nav-btn" onclick="window.location.href='/admin/users'">Повернутися до списку</button>
          <script>
            function validateForm() {
              const username = document.getElementById('username').value;
              const password = document.getElementById('password').value;
              const errorMessage = document.getElementById('error-message');
              if (username.length < 3 || username.length > 50) {
                errorMessage.textContent = 'Ім’я від 3 до 50 символів';
                return false;
              }
              if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
                errorMessage.textContent = 'Ім’я: лише літери та цифри';
                return false;
              }
              if (password.length < 6 || password.length > 100) {
                errorMessage.textContent = 'Пароль від 6 до 100 символів';
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
    logger.info('Маршрут /admin/add-user виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/add-user', checkAuth, checkAdmin, [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Ім’я від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Ім’я: лише літери та цифри'),
  body('password')
    .isLength({ min: 6, max: 100 }).withMessage('Пароль від 6 до 100 символів')
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
      return res.status(400).send('Користувач уже існує');
    }
    const saltRounds = 10;
    const hashedPassword = await bcrypt.hash(password, saltRounds);
    const newUser = { username, password: hashedPassword, role: username === 'Instructor' ? 'instructor' : 'user' };
    await db.collection('users').insertOne(newUser);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш користувачів оновлено після додавання');
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувача додано</title>
        </head>
        <body>
          <h1>Користувача ${username} додано</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка додавання', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні');
  } finally {
    logger.info('Маршрут /admin/add-user (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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
          <h1>Редагувати: ${username}</h1>
          <form method="POST" action="/admin/edit-user" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="hidden" name="oldUsername" value="${username}">
            <label for="username">Нове ім'я:</label>
            <input type="text" id="username" name="username" value="${username}" required>
            <label for="password">Новий пароль (порожнє — без змін):</label>
            <input type="text" id="password" name="password" placeholder="Новий пароль">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/users'">Повернутися</button>
          <script>
            function validateForm() {
              const username = document.getElementById('username').value;
              const password = document.getElementById('password').value;
              const errorMessage = document.getElementById('error-message');
              if (username.length < 3 || username.length > 50) {
                errorMessage.textContent = 'Ім’я від 3 до 50 символів';
                return false;
              }
              if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
                errorMessage.textContent = 'Ім’я: лише літери та цифри';
                return false;
              }
              if (password && (password.length < 6 || password.length > 100)) {
                errorMessage.textContent = 'Пароль від 6 до 100 символів';
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
    logger.info('Маршрут /admin/edit-user виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/edit-user', checkAuth, checkAdmin, [
  body('username')
    .isLength({ min: 3, max: 50 }).withMessage('Ім’я від 3 до 50 символів')
    .matches(/^[a-zA-Z0-9а-яА-Я]+$/).withMessage('Ім’я: лише літери та цифри'),
  body('password')
    .optional({ checkFalsy: true })
    .isLength({ min: 6, max: 100 }).withMessage('Пароль від 6 до 100 символів')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації', { errors: errors.array() });
      return res.status(400).send(errors.array()[0].msg);
    }

    const { oldUsername, username, password } = req.body;
    logger.info('Оновлення користувача', { oldUsername, username, passwordProvided: !!password });

    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser && username !== oldUsername) {
      logger.warn('Ім’я уже існує', { username });
      return res.status(400).send('Користувач із таким ім’ям існує');
    }

    const updateData = { username };
    if (password) {
      const saltRounds = 10;
      const hashedPassword = await bcrypt.hash(password, saltRounds);
      updateData.password = hashedPassword;
      logger.info('Пароль оновлено', { username });
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
    logger.info('Оновлення', { matchedCount: updateResult.matchedCount, modifiedCount: updateResult.modifiedCount });

    if (updateResult.matchedCount === 0) {
      logger.error('Користувача не знайдено', { oldUsername });
      return res.status(404).send('Користувача не знайдено');
    }

    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш оновлено після редагування');

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувача оновлено</title>
        </head>
        <body>
          <h1>Користувача ${username} оновлено</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка редагування', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні');
  } finally {
    logger.info('Маршрут /admin/edit-user (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/delete-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username } = req.body;
    await db.collection('users').deleteOne({ username });
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('Кеш оновлено після видалення');
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні' });
  } finally {
    logger.info('Маршрут /admin/delete-user виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const sortBy = req.query.sortBy || 'order';
    const limit = 10;
    const skip = (page - 1) * limit;

    let questions = [];
    let errorMessage = '';
    let totalQuestions = 0;
    let totalPages = 0;

    try {
      totalQuestions = await db.collection('questions').countDocuments();
      totalPages = Math.ceil(totalQuestions / limit);

      if (sortBy === 'testName') {
        questions = await db.collection('questions')
          .find({})
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
          .find({})
          .sort({ order: 1 })
          .skip(skip)
          .limit(limit)
          .toArray();
      }
      await CacheManager.invalidateCache('questions', null);
    } catch (error) {
      logger.error('Помилка отримання питань', { message: error.message, stack: error.stack });
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
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .sort-btn { padding: 5px 10px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
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
            <button class="sort-btn" onclick="window.location.href='/admin/questions?page=${page}&sortBy=order'">Сортувати за порядком</button>
            <button class="sort-btn" onclick="window.location.href='/admin/questions?page=${page}&sortBy=testName'">Сортувати за тестом</button>
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
            ${page > 1 ? `<a href="/admin/questions?page=${page - 1}&sortBy=${sortBy}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/questions?page=${page + 1}&sortBy=${sortBy}">Наступна</a>` : ''}
          </div>
          <script>
            async function deleteQuestion(id) {
              if (confirm('Ви впевнені, що хочете видалити питання?')) {
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
                    throw new Error('HTTP-помилка: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка видалення: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення:', error);
                  alert('Не вдалося видалити питання');
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
            <label for="picture">Файл зображення (напр., Picture1.png):</label>
            <p class="note">Файл має бути в public/images.</p>
            <input type="text" id="picture" name="picture" placeholder="Picture1.png">
            <label for="text">Текст питання:</label>
            <p class="note">Для Fillblank: ___ для пропусків.</p>
            <textarea id="text" name="text" required placeholder="Текст питання"></textarea>
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
              <label for="options">Варіанти (через ;):</label>
              <textarea id="options" name="options" placeholder="Варіанти через крапку з комою"></textarea>
            </div>
            <label for="correctAnswers">Правильні відповіді (через ;):</label>
            <p id="correctAnswersNote" class="note">Для Input/Fillblank: діапазон "число1-число2", напр., "12-14".</p>
            <textarea id="correctAnswers" name="correctAnswers" required placeholder="Правильні відповіді"></textarea>
            <label for="points">Бали:</label>
            <input type="number" id="points" name="points" value="1" min="1" required>
            <label for="variant">Варіант (опціонально):</label>
            <input type="text" id="variant" name="variant" placeholder="Variant 1">
            <button type="submit" class="submit-btn">Додати</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/questions'">Повернутися</button>
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
                errorMessage.textContent = 'Текст від 5 до 1000 символів';
                return false;
              }
              if (points < 1 || points > 100) {
                errorMessage.textContent = 'Бали від 1 до 100';
                return false;
              }
              if (variant && (variant.length < 1 || variant.length > 50)) {
                errorMessage.textContent = 'Варіант від 1 до 50 символів';
                return false;
              }
              if (picture && !/\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Файл: .jpeg, .jpg, .png або .gif';
                return false;
              }
              if (type === 'input' || type === 'fillblank') {
                const answersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (type === 'input' && answersArray.length !== 1) {
                  errorMessage.textContent = 'Input: одна відповідь';
                  return false;
                }
                if (type === 'fillblank') {
                  const blankCount = (text.match(/___/g) || []).length;
                  if (blankCount === 0 || blankCount !== answersArray.length) {
                    errorMessage.textContent = 'Пропуски не відповідають відповідям';
                    return false;
                  }
                }
                for (let i = 0; i < answersArray.length; i++) {
                  const answer = answersArray[i];
                  if (answer.includes('-')) {
                    const [min, max] = answer.split('-').map(val => parseFloat(val.trim()));
                    if (isNaN(min) || isNaN(max) || min > max) {
                      errorMessage.textContent = \`Відповідь \${i + 1}: невірний діапазон\`;
                      return false;
                    }
                  } else {
                    const value = parseFloat(answer);
                    if (isNaN(value)) {
                      errorMessage.textContent = \`Відповідь \${i + 1} для \${type}: число або діапазон\`;
                      return false;
                    }
                  }
                }
              }
              if (type === 'singlechoice') {
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (correctAnswersArray.length !== 1) {
                  errorMessage.textContent = 'Single Choice: одна відповідь';
                  return false;
                }
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                if (options.length < 2) {
                  errorMessage.textContent = 'Single Choice: мінімум 2 варіанти';
                  return false;
                }
              }
              if (type === 'matching') {
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean);
                if (options.length === 0 || options.length !== correctAnswersArray.length) {
                  errorMessage.textContent = 'Matching: варіанти = відповіді';
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
    res.status(500).send('Помилка при додаванні');
  } finally {
    logger.info('Маршрут /admin/add-question виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/add-question', checkAuth, checkAdmin, [
  body('testNumber').notEmpty().withMessage('Номер тесту обов’язковий'),
  body('text').isLength({ min: 5, max: 1000 }).withMessage('Текст від 5 до 1000 символів'),
  body('type').isIn(['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank']).withMessage('Невірний тип'),
  body('correctAnswers').notEmpty().withMessage('Правильні відповіді обов’язкові'),
  body('points').isInt({ min: 1, max: 100 }).withMessage('Бали від 1 до 100'),
  body('variant').optional({ checkFalsy: true }).isLength({ min: 1, max: 50 }).withMessage('Варіант від 1 до 50 символів'),
  body('picture').optional({ checkFalsy: true }).matches(/\.(jpeg|jpg|png|gif)$/i).withMessage('Файл: .jpeg, .jpg, .png або .gif')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      logger.warn('Помилки валідації', { errors: errors.array() });
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
        logger.warn(`Зображення не знайдено: ${imagePath}`);
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
        logger.warn('Matching: потрібні пари', { testNumber, text });
        return res.status(400).send('Matching: потрібні пари');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        logger.warn('Fillblank: пропуски ≠ відповіді', { blankCount, correctAnswersLength: questionData.correctAnswers.length });
        return res.status(400).send('Пропуски не відповідають відповідям');
      }
      questionData.blankCount = blankCount;

      questionData.correctAnswers.forEach((correctAnswer, idx) => {
        if (correctAnswer.includes('-')) {
          const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
          if (isNaN(min) || isNaN(max) || min > max) {
            return res.status(400).send(`Відповідь ${idx + 1}: невірний діапазон`);
          }
        } else {
          const value = parseFloat(correctAnswer);
          if (isNaN(value)) {
            return res.status(400).send(`Відповідь ${idx + 1} для Fillblank: число або діапазон`);
          }
        }
      });
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        logger.warn('Single Choice: одна відповідь, ≥2 варіанти', {
          correctAnswersLength: questionData.correctAnswers.length,
          optionsLength: questionData.options.length
        });
        return res.status(400).send('Single Choice: одна відповідь і ≥2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    if (type === 'input') {
      if (questionData.correctAnswers.length !== 1) {
        return res.status(400).send('Input: одна відповідь');
      }
      const correctAnswer = questionData.correctAnswers[0];
      if (correctAnswer.includes('-')) {
        const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
        if (isNaN(min) || isNaN(max) || min > max) {
          return res.status(400).send('Input: невірний діапазон');
        }
      } else {
        const value = parseFloat(correctAnswer);
        if (isNaN(value)) {
          return res.status(400).send('Input: число або діапазон');
        }
      }
    }

    await db.collection('questions').insertOne(questionData);
    logger.info('Питання додано', { testNumber, text, type });

    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    logger.info('Кеш очищено', { testNumber });

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
          <h1>Питання додано</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка додавання', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні: ' + error.message);
  } finally {
    logger.info('Маршрут /admin/add-question (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/edit-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.query;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).send('Невірний ідентифікатор');
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
      ? `Попередження: зображення "${normalizedOriginalPicture}" не знайдено в public/images`
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
            <label for="picture">Файл зображення (напр., Picture1.png):</label>
            <p class="note">Файл має бути в public/images.</p>
            <input type="text" id="picture" name="picture" value="${pictureName}" placeholder="Picture1.png">
            ${warningMessage ? `<p class="warning">${warningMessage}</p>` : ''}
            ${pictureName ? `<img id="image-preview" src="/images/${pictureName}" alt="Зображення" onerror="this.onerror=null;this.src='';this.alt='Зображення недоступне';">` : ''}
            <label for="text">Текст питання:</label>
            <p class="note">Для Fillblank: ___ для пропусків.</p>
            <textarea id="text" name="text" required placeholder="Текст питання">${question.text}</textarea>
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
              <label for="options">Варіанти (через ;):</label>
              <textarea id="options" name="options" placeholder="Варіанти через крапку з комою">${question.options.join('; ')}</textarea>
            </div>
            <label for="correctAnswers">Правильні відповіді (через ;):</label>
            <p id="correctAnswersNote" class="note">Для Input/Fillblank: діапазон "число1-число2", напр., "12-14".</p>
            <textarea id="correctAnswers" name="correctAnswers" required placeholder="Правильні відповіді">${question.correctAnswers.join('; ')}</textarea>
            <label for="points">Бали:</label>
            <input type="number" id="points" name="points" value="${question.points}" min="1" required>
            <label for="variant">Варіант:</label>
            <input type="text" id="variant" name="variant" value="${question.variant}" placeholder="Variant 1">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/questions'">Повернутися</button>
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
                errorMessage.textContent = 'Текст від 5 до 1000 символів';
                return false;
              }
              if (points < 1 || points > 100) {
                errorMessage.textContent = 'Бали від 1 до 100';
                return false;
              }
              if (variant && (variant.length < 1 || variant.length > 50)) {
                errorMessage.textContent = 'Варіант від 1 до 50 символів';
                return false;
              }
              if (picture && !/\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Файл: .jpeg, .jpg, .png або .gif';
                return false;
              }
              if (type === 'input' || type === 'fillblank') {
                const answersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (type === 'input' && answersArray.length !== 1) {
                  errorMessage.textContent = 'Input: одна відповідь';
                  return false;
                }
                if (type === 'fillblank') {
                  const blankCount = (text.match(/___/g) || []).length;
                  if (blankCount === 0 || blankCount !== answersArray.length) {
                    errorMessage.textContent = 'Пропуски не відповідають відповідям';
                    return false;
                  }
                }
                for (let i = 0; i < answersArray.length; i++) {
                  const answer = answersArray[i];
                  if (answer.includes('-')) {
                    const [min, max] = answer.split('-').map(val => parseFloat(val.trim()));
                    if (isNaN(min) || isNaN(max) || min > max) {
                      errorMessage.textContent = \`Відповідь \${i + 1}: невірний діапазон\`;
                      return false;
                    }
                  } else {
                    const value = parseFloat(answer);
                    if (isNaN(value)) {
                      errorMessage.textContent = \`Відповідь \${i + 1} для \${type}: число або діапазон\`;
                      return false;
                    }
                  }
                }
              }
              if (type === 'singlechoice') {
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim());
                if (correctAnswersArray.length !== 1) {
                  errorMessage.textContent = 'Single Choice: одна відповідь';
                  return false;
                }
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                if (options.length < 2) {
                  errorMessage.textContent = 'Single Choice: ≥2 варіанти';
                  return false;
                }
              }
              if (type === 'matching') {
                const options = document.getElementById('options').value.split(';').map(opt => opt.trim()).filter(Boolean);
                const correctAnswersArray = correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean);
                if (options.length === 0 || options.length !== correctAnswersArray.length) {
                  errorMessage.textContent = 'Matching: варіанти = відповіді';
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
    res.status(500).send('Помилка при редагуванні');
  } finally {
    logger.info('Маршрут /admin/edit-question виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/edit-question', checkAuth, checkAdmin, [
  body('testNumber').notEmpty().withMessage('Номер тесту обов’язковий'),
  body('text').isLength({ min: 5, max: 1000 }).withMessage('Текст від 5 до 1000 символів'),
  body('type').isIn(['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank']).withMessage('Невірний тип'),
  body('correctAnswers').notEmpty().withMessage('Правильні відповіді обов’язкові'),
  body('points').isInt({ min: 1, max: 100 }).withMessage('Бали від 1 до 100'),
  body('variant').optional({ checkFalsy: true }).isLength({ min: 1, max: 50 }).withMessage('Варіант від 1 до 50 символів'),
  body('picture').optional({ checkFalsy: true }).matches(/\.(jpeg|jpg|png|gif)$/i).withMessage('Файл: .jpeg, .jpg, .png або .gif')
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

      logger.info(`Перевірка зображення ${normalizedPicture} у ${imageDir}`);

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
        logger.warn(`Зображення ${normalizedPicture} не знайдено: ${filesInDir.join(', ')}`);
        questionData.picture = null;
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
        logger.warn('Matching: потрібні пари', { testNumber, text });
        return res.status(400).send('Matching: потрібні пари');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        logger.warn('Fillblank: пропуски ≠ відповіді', { blankCount, correctAnswersLength: questionData.correctAnswers.length });
        return res.status(400).send('Пропуски не відповідають відповідям');
      }
      questionData.blankCount = blankCount;

      questionData.correctAnswers.forEach((correctAnswer, idx) => {
        if (correctAnswer.includes('-')) {
          const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
          if (isNaN(min) || isNaN(max) || min > max) {
            return res.status(400).send(`Відповідь ${idx + 1}: невірний діапазон`);
          }
        } else {
          const value = parseFloat(correctAnswer);
          if (isNaN(value)) {
            return res.status(400).send(`Відповідь ${idx + 1} для Fillblank: число або діапазон`);
          }
        }
      });
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        logger.warn('Single Choice: одна відповідь, ≥2 варіанти', {
          correctAnswersLength: questionData.correctAnswers.length,
          optionsLength: questionData.options.length
        });
        return res.status(400).send('Single Choice: одна відповідь і ≥2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    if (type === 'input') {
      if (questionData.correctAnswers.length !== 1) {
        return res.status(400).send('Input: одна відповідь');
      }
      const correctAnswer = questionData.correctAnswers[0];
      if (correctAnswer.includes('-')) {
        const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
        if (isNaN(min) || isNaN(max) || min > max) {
          return res.status(400).send('Input: невірний діапазон');
        }
      } else {
        const value = parseFloat(correctAnswer);
        if (isNaN(value)) {
          return res.status(400).send('Input: число або діапазон');
        }
      }
    }

    await db.collection('questions').updateOne(
      { _id: new ObjectId(id) },
      { $set: questionData }
    );
    logger.info('Питання оновлено', { id, testNumber, text, type });

    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    logger.info('Кеш очищено', { testNumber });

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
          <h1>Питання оновлено</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка оновлення', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні: ' + error.message);
  } finally {
    logger.info('Маршрут /admin/edit-question (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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
    logger.error('Помилка видалення', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні' });
  } finally {
    logger.info('Маршрут /admin/delete-question виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/import-users', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
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
          <form id="import-form">
            <input type="hidden" name="_csrf" id="_csrf" value="${res.locals._csrf}">
            <label for="file">Виберіть users.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn" id="submit-btn">Завантажити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися</button>
          <script>
            document.getElementById('import-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const fileInput = document.getElementById('file');
              const errorMessage = document.getElementById('error-message');
              const submitBtn = document.getElementById('submit-btn');
              const csrfToken = document.getElementById('_csrf').value;

              if (!csrfToken) {
                errorMessage.textContent = 'CSRF-токен відсутній';
                return;
              }

              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано';
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Завантаження...';

              const formData = new FormData();
              formData.append('file', fileInput.files[0]);

              try {
                const response = await fetch('/admin/import-users', {
                  method: 'POST',
                  body: formData,
                  headers: { 'X-CSRF-Token': csrfToken }
                });
                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'HTTP-помилка: ' + response.status);
                }
                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка завантаження:', error);
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
  } finally {
    logger.info('Маршрут /admin/import-users виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      logger.error('Файл не надано');
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
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Імпортовано ${count} користувачів</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка імпорту', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка імпорту</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Помилка імпорту</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-users'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-users (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/import-questions', checkAuth, checkAdmin, (req, res) => {
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
          <form id="import-form">
            <input type="hidden" name="_csrf" id="_csrf" value="${res.locals._csrf}">
            <label for="testNumber">Тест:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}">${testNames[num].name.replace(/"/g, '\\"')}</option>`).join('')}
            </select>
            <label for="file">Виберіть questions.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn" id="submit-btn">Завантажити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися</button>
          <script>
            document.getElementById('import-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              const fileInput = document.getElementById('file');
              const testNumber = document.getElementById('testNumber').value;
              const errorMessage = document.getElementById('error-message');
              const submitBtn = document.getElementById('submit-btn');
              const csrfToken = document.getElementById('_csrf').value;

              if (!csrfToken) {
                errorMessage.textContent = 'CSRF-токен відсутній';
                return;
              }

              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано';
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Завантаження...';

              const formData = new FormData();
              formData.append('file', fileInput.files[0]);
              formData.append('testNumber', testNumber);

              try {
                const response = await fetch('/admin/import-questions', {
                  method: 'POST',
                  body: formData,
                  headers: { 'X-CSRF-Token': csrfToken }
                });
                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'HTTP-помилка: ' + response.status);
                }
                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка завантаження:', error);
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
  } catch (error) {
    logger.error('Помилка в /admin/import-questions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при імпорті: ' + error.message);
  } finally {
    logger.info('Маршрут /admin/import-questions виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file || !req.body.testNumber) {
      logger.error('Файл або номер тесту не надано', { testNumber: req.body.testNumber });
      return res.status(400).send('Файл або номер тесту не надано');
    }

    const testNumber = req.body.testNumber;
    if (!testNames[testNumber]) {
      logger.error('Невірний номер тесту', { testNumber });
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
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Імпортовано ${count} питань для тесту ${testNames[testNumber].name.replace(/"/g, '\\"')}</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка імпорту питань', { message: error.message, stack: error.stack });
    res.status(500).send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Помилка імпорту</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            .error { color: red; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Помилка імпорту</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-questions'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    logger.info('Маршрут /admin/import-questions (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    let results;
    if (req.userRole === 'instructor' || req.userRole === 'admin') {
      results = await db.collection('test_results').find({}).sort({ endTime: -1 }).limit(100).toArray();
    } else {
      results = await db.collection('test_results').find({ user: req.user }).sort({ endTime: -1 }).limit(20).toArray();
    }

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Результати тестів</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .details-btn { padding: 5px 10px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
            .nav-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Результати тестів</h1>
          <button class="nav-btn" onclick="window.location.href='${req.userRole === 'admin' ? '/admin' : '/select-test'}'">Повернутися</button>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Тест</th>
              <th>Відсоток</th>
              <th>Дата</th>
              <th>Деталі</th>
            </tr>
    `;
    if (!results || results.length === 0) {
      html += '<tr><td colspan="5">Немає результатів</td></tr>';
    } else {
      results.forEach(result => {
        const endDateTime = new Date(result.endTime);
        const formattedDate = endDateTime.toLocaleString('uk-UA');
        html += `
          <tr>
            <td>${result.user}</td>
            <td>${testNames[result.testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</td>
            <td>${Math.round(result.percentage)}%</td>
            <td>${formattedDate}</td>
            <td>
              <button class="details-btn" onclick="window.location.href='/admin/result-details?id=${result._id}'">Переглянути</button>
            </td>
          </tr>
        `;
      });
    }
    html += `
          </table>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/results', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка перегляду результатів');
  } finally {
    logger.info('Маршрут /admin/results виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/result-details', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
      return res.status(403).send('Доступ заборонено');
    }

    const { id } = req.query;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).send('Невірний ідентифікатор');
    }

    const result = await db.collection('test_results').findOne({ _id: new ObjectId(id) });
    if (!result) {
      return res.status(404).send('Результат не знайдено');
    }

    const questions = await db.collection('questions').find({ testNumber: result.testNumber }).toArray();
    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Деталі результату</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .answers { white-space: pre-wrap; max-width: 300px; overflow-wrap: break-word; }
            .nav-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Деталі результату</h1>
          <p>
            Користувач: ${result.user}<br>
            Тест: ${testNames[result.testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}<br>
            Відсоток: ${Math.round(result.percentage)}%<br>
            Бали: ${result.score} з ${result.totalPoints}<br>
            Час поза вкладкою: ${result.suspiciousActivity?.timeAway ? Math.round((result.suspiciousActivity.timeAway / result.duration) * 100) : 0}%<br>
            Переключення вкладок: ${result.suspiciousActivity?.switchCount || 0}<br>
            Дата завершення: ${new Date(result.endTime).toLocaleString('uk-UA')}
          </p>
          <button class="nav-btn" onclick="window.location.href='/admin/results'">Повернутися</button>
          <table>
            <tr>
              <th>Питання</th>
              <th>Відповідь користувача</th>
              <th>Правильна відповідь</th>
              <th>Бали</th>
            </tr>
    `;
    Object.keys(result.answers).forEach(index => {
      const question = questions.find(q => q.order === parseInt(index)) || {};
      const userAnswer = result.answers[index] || 'Не відповіли';
      let correctAnswer;
      if (question.type === 'matching') {
        correctAnswer = question.correctPairs ? question.correctPairs.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ') : question.correctAnswers.join(', ');
      } else if (question.type === 'fillblank') {
        correctAnswer = question.correctAnswers.join(', ');
      } else if (question.type === 'singlechoice') {
        correctAnswer = question.correctAnswer;
      } else {
        correctAnswer = question.correctAnswers ? question.correctAnswers.join(', ') : 'Невідомо';
      }
      const questionScore = result.scoresPerQuestion[index] || 0;
      let userAnswerDisplay;
      if (question.type === 'matching' && Array.isArray(userAnswer)) {
        userAnswerDisplay = userAnswer.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ');
      } else if (question.type === 'fillblank' && Array.isArray(userAnswer)) {
        userAnswerDisplay = userAnswer.join(', ');
      } else if (Array.isArray(userAnswer)) {
        userAnswerDisplay = userAnswer.join(', ');
      } else {
        userAnswerDisplay = userAnswer;
      }
      html += `
        <tr>
          <td>${question.text || 'Питання не знайдено'}</td>
          <td class="answers">${userAnswerDisplay}</td>
          <td class="answers">${correctAnswer}</td>
          <td>${questionScore} з ${question.points || 0}</td>
        </tr>
      `;
    });
    html += `
          </table>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/result-details', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка перегляду деталей');
  } finally {
    logger.info('Маршрут /admin/result-details виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const tests = await db.collection('tests').find({}).toArray();
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
            .action-btn { padding: 5px 10px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
          </style>
        </head>
        <body>
          <h1>Редагувати тести</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися</button>
          <table>
            <tr>
              <th>Номер тесту</th>
              <th>Назва</th>
              <th>Ліміт часу (с)</th>
              <th>Дії</th>
            </tr>
    `;
    if (!tests || tests.length === 0) {
      html += '<tr><td colspan="4">Немає тестів</td></tr>';
    } else {
      tests.forEach(test => {
        html += `
          <tr>
            <td>${test.testNumber}</td>
            <td>${test.name.replace(/"/g, '\\"')}</td>
            <td>${test.timeLimit}</td>
            <td>
              <button class="action-btn edit" onclick="window.location.href='/admin/edit-test?testNumber=${test.testNumber}'">Редагувати</button>
              <button class="action-btn delete" onclick="deleteTest('${test.testNumber}')">Видалити</button>
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
                    throw new Error('HTTP-помилка: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка видалення: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення:', error);
                  alert('Не вдалося видалити тест');
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
    res.status(500).send('Помилка перегляду тестів');
  } finally {
    logger.info('Маршрут /admin/edit-tests виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/edit-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.query;
    const test = await db.collection('tests').findOne({ testNumber });
    if (!test) {
      return res.status(404).send('Тест не знайдено');
    }

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Редагувати тест ${test.testNumber}</title>
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
          <h1>Редагувати тест ${test.testNumber}</h1>
          <form method="POST" action="/admin/edit-test" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <input type="hidden" name="testNumber" value="${test.testNumber}">
            <label for="name">Назва тесту:</label>
            <input type="text" id="name" name="name" value="${test.name.replace(/"/g, '\\"')}" required>
            <label for="timeLimit">Ліміт часу (секунди):</label>
            <input type="number" id="timeLimit" name="timeLimit" value="${test.timeLimit}" min="1" required>
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
            <label for="questionLimit">Ліміт питань (порожнє для всіх):</label>
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
          <button class="nav-btn" onclick="window.location.href='/admin/edit-tests'">Повернутися</button>
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
                errorMessage.textContent = 'Назва від 1 до 100 символів';
                return false;
              }
              if (timeLimit < 1) {
                errorMessage.textContent = 'Ліміт часу ≥ 1 секунда';
                return false;
              }
              if (questionLimit && questionLimit < 1) {
                errorMessage.textContent = 'Ліміт питань ≥ 1';
                return false;
              }
              if (attemptLimit < 1) {
                errorMessage.textContent = 'Ліміт спроб ≥ 1';
                return false;
              }
              if (isQuickTest === 'true' && (!timePerQuestion || timePerQuestion < 1)) {
                errorMessage.textContent = 'Час на питання ≥ 1 секунда для швидкого тесту';
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
    res.status(500).send('Помилка редагування тесту');
  } finally {
    logger.info('Маршрут /admin/edit-test виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/edit-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber, name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    const testData = {
      name: name.trim(),
      timeLimit: parseInt(timeLimit),
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: isQuickTest === 'true' && timePerQuestion ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    await loadTestsFromMongoDB();

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
          <h1>Тест ${testNumber} оновлено</h1>
          <button onclick="window.location.href='/admin/edit-tests'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка оновлення тесту', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при оновленні тесту');
  } finally {
    logger.info('Маршрут /admin/edit-test (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.body;
    await deleteTestFromMongoDB(testNumber);
    await db.collection('questions').deleteMany({ testNumber });
    await CacheManager.invalidateCache('questions', testNumber);
    await CacheManager.invalidateCache('allQuestions', 'all');
    await loadTestsFromMongoDB();
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення тесту', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні' });
  } finally {
    logger.info('Маршрут /admin/delete-test виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

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
          <h1>Створити тест</h1>
          <form method="POST" action="/admin/create-test" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${res.locals._csrf}">
            <label for="testNumber">Номер тесту:</label>
            <input type="text" id="testNumber" name="testNumber" required>
            <label for="name">Назва тесту:</label>
            <input type="text" id="name" name="name" required>
            <label for="timeLimit">Ліміт часу (секунди):</label>
            <input type="number" id="timeLimit" name="timeLimit" min="1" required>
            <label for="randomQuestions">Випадкові питання:</label>
            <select id="randomQuestions" name="randomQuestions">
              <option value="true">Так</option>
              <option value="false">Ні</option>
            </select>
            <label for="randomAnswers">Випадкові відповіді:</label>
            <select id="randomAnswers" name="randomAnswers">
              <option value="true">Так</option>
              <option value="false">Ні</option>
            </select>
            <label for="questionLimit">Ліміт питань (порожнє для всіх):</label>
            <input type="number" id="questionLimit" name="questionLimit" min="1">
            <label for="attemptLimit">Ліміт спроб:</label>
            <input type="number" id="attemptLimit" name="attemptLimit" value="1" min="1" required>
            <label for="isQuickTest">Швидкий тест:</label>
            <select id="isQuickTest" name="isQuickTest">
              <option value="true">Так</option>
              <option value="false">Ні</option>
            </select>
            <label for="timePerQuestion">Час на питання (секунди, для швидкого тесту):</label>
            <input type="number" id="timePerQuestion" name="timePerQuestion" min="1">
            <button type="submit" class="submit-btn">Створити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися</button>
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
                errorMessage.textContent = 'Номер тесту: лише цифри';
                return false;
              }
              if (name.length < 1 || name.length > 100) {
                errorMessage.textContent = 'Назва від 1 до 100 символів';
                return false;
              }
              if (timeLimit < 1) {
                errorMessage.textContent = 'Ліміт часу ≥ 1 секунда';
                return false;
              }
              if (questionLimit && questionLimit < 1) {
                errorMessage.textContent = 'Ліміт питань ≥ 1';
                return false;
              }
              if (attemptLimit < 1) {
                errorMessage.textContent = 'Ліміт спроб ≥ 1';
                return false;
              }
              if (isQuickTest === 'true' && (!timePerQuestion || timePerQuestion < 1)) {
                errorMessage.textContent = 'Час на питання ≥ 1 секунда для швидкого тесту';
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
    res.status(500).send('Помилка створення тесту');
  } finally {
    logger.info('Маршрут /admin/create-test виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.post('/admin/create-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber, name, timeLimit, randomQuestions, randomAnswers, questionLimit, attemptLimit, isQuickTest, timePerQuestion } = req.body;

    if (!/^[0-9]+$/.test(testNumber)) {
      return res.status(400).send('Номер тесту: лише цифри');
    }

    const existingTest = await db.collection('tests').findOne({ testNumber });
    if (existingTest) {
      return res.status(400).send('Тест із таким номером уже існує');
    }

    const testData = {
      name: name.trim(),
      timeLimit: parseInt(timeLimit),
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: isQuickTest === 'true' && timePerQuestion ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    await loadTestsFromMongoDB();

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
          <h1>Тест ${testNumber} створено</h1>
          <button onclick="window.location.href='/admin/edit-tests'">Повернутися</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка створення тесту', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при створенні тесту');
  } finally {
    logger.info('Маршрут /admin/create-test (POST) виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;

    const totalLogs = await db.collection('activity_log').countDocuments();
    const totalPages = Math.ceil(totalLogs / limit);

    const logs = await db.collection('activity_log')
      .find({})
      .sort({ timestamp: -1 })
      .skip(skip)
      .limit(limit)
      .toArray();

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Журнал дій</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .nav-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
          </style>
        </head>
        <body>
          <h1>Журнал дій</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися</button>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Дія</th>
              <th>IP-адреса</th>
              <th>Час</th>
              <th>Додаткова інформація</th>
            </tr>
    `;
    if (!logs || logs.length === 0) {
      html += '<tr><td colspan="5">Немає записів</td></tr>';
    } else {
      logs.forEach(log => {
        const timestamp = new Date(log.timestamp).toLocaleString('uk-UA');
        const additionalInfo = log.additionalInfo ? JSON.stringify(log.additionalInfo, null, 2) : '';
        html += `
          <tr>
            <td>${log.user}</td>
            <td>${log.action}</td>
            <td>${log.ipAddress}</td>
            <td>${timestamp}</td>
            <td><pre>${additionalInfo}</pre></td>
          </tr>
        `;
      });
    }
    html += `
          </table>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/activity-log?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/activity-log?page=${page + 1}">Наступна</a>` : ''}
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/activity-log', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка перегляду журналу');
  } finally {
    logger.info('Маршрут /admin/activity-log виконано', { duration: `${Date.now() - startTime} мс` });
  }
});

// Запуск сервера
const port = process.env.PORT || 3000;
app.listen(port, () => {
  logger.info(`Сервер запущено на порту ${port}`);
});

// Обробка завершення роботи
process.on('SIGTERM', async () => {
  logger.info('Отримано SIGTERM. Завершення роботи...');
  try {
    await client.close();
    logger.info('MongoDB з’єднання закрито');
  } catch (error) {
    logger.error('Помилка закриття MongoDB', { message: error.message, stack: error.stack });
  }
  process.exit(0);
});

process.on('SIGINT', async () => {
  logger.info('Отримано SIGINT. Завершення роботи...');
  try {
    await client.close();
    logger.info('MongoDB з’єднання закрито');
  } catch (error) {
    logger.error('Помилка закриття MongoDB', { message: error.message, stack: error.stack });
  }
  process.exit(0);
});

process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled Rejection', { promise, reason });
});

process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception', { message: error.message, stack: error.stack });
  process.exit(1);
});
