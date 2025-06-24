// Імпорт необхідних модулів
require('dotenv').config(); // Завантаження змінних середовища з .env файлу
const express = require('express'); // Фреймворк для створення веб-додатку
const cookieParser = require('cookie-parser'); // Парсинг cookies
const path = require('path'); // Робота з шляхами до файлів
const ExcelJS = require('exceljs'); // Робота з Excel-файлами
const { MongoClient, ObjectId } = require('mongodb'); // Клієнт MongoDB та ObjectId для роботи з ID
const bcrypt = require('bcrypt'); // Хешування паролів
const fs = require('fs'); // Робота з файловою системою
const multer = require('multer'); // Завантаження файлів
const nodemailer = require('nodemailer'); // Відправка email
const { body, validationResult } = require('express-validator'); // Валідація вхідних даних
const jwt = require('jsonwebtoken'); // Робота з JWT-токенами
const winston = require('winston'); // Логування
const session = require('express-session'); // Управління сесіями
const MongoStore = require('connect-mongo'); // Зберігання сесій у MongoDB
const { createClient } = require('@vercel/blob'); // Імпорт createClient з @vercel/blob

// Ініціалізація Express-додатку
const app = express();

// Увімкнення довіри до проксі для коректної роботи за проксі-серверами
app.set('trust proxy', 1);

// Налаштування логування
const logger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.File({ filename: 'error.log', level: 'error' }), // Логи помилок
    new winston.transports.File({ filename: 'combined.log' }), // Всі логи
    new winston.transports.Console() // Вивід у консоль
  ]
});

// Налаштування multer для імпорту файлів (Excel)
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 4 * 1024 * 1024 } // Ліміт 4MB
});

// Ініціалізація клієнта Vercel Blob
const blob = createClient({
  token: process.env.BLOB_READ_WRITE_TOKEN, // Токен з Vercel Dashboard
});

// Налаштування multer для матеріалів з використанням memoryStorage (для завантаження в пам’ять перед Blob)
const materialStorage = multer.memoryStorage(); // Використовуємо memoryStorage для тимчасового зберігання в пам’яті

// Ініціалізація uploadMaterial з фільтрами та лімітом
const uploadMaterial = multer({
  storage: materialStorage,
  limits: { fileSize: 10 * 1024 * 1024 }, // Ліміт 10MB
  fileFilter: (req, file, cb) => {
    const allowedTypes = ['application/pdf', 'image/jpeg', 'image/png', 'text/plain'];
    if (allowedTypes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Дозволені формати: PDF, JPEG, PNG, TXT'));
    }
  }
});

// Налаштування nodemailer для відправки email
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

// Конфігурація для підозрілої активності
const config = {
  suspiciousActivity: {
    timeAwayThreshold: 50,
    switchCountThreshold: 5
  }
};

// Налаштування підключення до MongoDB
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/alpha?retryWrites=true&w=majority';
const client = new MongoClient(MONGODB_URI, {
  connectTimeoutMS: 5000,
  serverSelectionTimeoutMS: 5000
});
let db;

// Клас для кешування даних
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

// Ініціалізація кешів та змінних
let userCache = [];
const questionsCache = {};
let isInitialized = false;
let initializationError = null;
let testNames = {};
let sectionNames = {};

// Підключення до MongoDB із повторними спробами
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

// Завантаження тестів із MongoDB
const loadTestsFromMongoDB = async () => {
  try {
    const tests = await db.collection('tests').find({}).toArray();
    testNames = {};
    tests.forEach(test => {
      testNames[test.testNumber] = {
        name: test.name,
        sectionId: test.sectionId || null, // Додано поле sectionId
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

// Завантаження розділів із MongoDB
const loadSectionsFromMongoDB = async () => {
  try {
    const sections = await db.collection('sections').find({}).toArray();
    sectionNames = {};
    sections.forEach(section => {
      sectionNames[section.sectionId] = {
        name: section.name,
        imageUrl: section.imageUrl || '/images/default-section.png'
      };
    });
    logger.info(`Завантажено ${sections.length} розділів`);
  } catch (error) {
    logger.error('Помилка завантаження розділів', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Збереження тесту в MongoDB
const saveTestToMongoDB = async (testNumber, testData) => {
  try {
    await db.collection('tests').updateOne(
      { testNumber },
      { $set: { 
        testNumber,
        sectionId: testData.sectionId || null,
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

// Видалення тесту з MongoDB
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

// Логування отриманих cookies
app.use((req, res, next) => {
  logger.info('Отримано cookie', { cookies: req.cookies, sessionID: req.sessionID || 'unknown' });
  next();
});

// Логування запитів
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

// Валідація CSRF-токена (виключено для маршрутів завантаження файлів)
app.use((req, res, next) => {
  if (['POST', 'PUT', 'DELETE'].includes(req.method) && 
      !req.url.startsWith('/admin/import-users') && 
      !req.url.startsWith('/admin/import-questions') &&
      !req.url.startsWith('/admin/upload-material')) {
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

// Запобігання кешуванню сторінок
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

// Додавання водяного знаку та блокування копіювання/скріншотів
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

// Імпорт користувачів із Excel
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

// Імпорт питань із Excel
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

// Завантаження користувачів у кеш
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

// Перевірка ініціалізації сервера
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

// Оновлення паролів користувачів
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
    // Створення індексів для всіх колекцій, включаючи нові sections та materials
    await db.collection('users').createIndex({ username: 1 }, { unique: true });
    await db.collection('questions').createIndex({ testNumber: 1, variant: 1 });
    await db.collection('test_results').createIndex({ user: 1, testNumber: 1, endTime: -1 });
    await db.collection('activity_log').createIndex({ user: 1, timestamp: -1 });
    await db.collection('test_attempts').createIndex({ user: 1, testNumber: 1, attemptDate: 1 });
    await db.collection('login_attempts').createIndex({ ipAddress: 1, lastAttempt: 1 });
    await db.collection('tests').createIndex({ testNumber: 1 }, { unique: true });
    await db.collection('active_tests').createIndex({ user: 1 }, { unique: true });
    await db.collection('sections').createIndex({ sectionId: 1 }, { unique: true });
    await db.collection('materials').createIndex({ sectionId: 1 });
    logger.info('Індекси створено');

    // Міграція ролей користувачів
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

    // Ініціалізація дефолтних тестів
    const testCount = await db.collection('tests').countDocuments();
    if (!testCount) {
      const defaultTests = {
        "1": { name: "Тест 1", sectionId: null, timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "2": { name: "Тест 2", sectionId: null, timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 },
        "3": { name: "Тест 3", sectionId: null, timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null, attemptLimit: 1 }
      };
      for (const [testNumber, testData] of Object.entries(defaultTests)) {
        await saveTestToMongoDB(testNumber, testData);
      }
      logger.info('Міграція тестів завершена', { count: Object.keys(defaultTests).length });
    }

    await updateUserPasswords();
    await loadUsersToCache();
    await loadTestsFromMongoDB();
    await loadSectionsFromMongoDB();
    await CacheManager.invalidateCache('questions', null);
    isInitialized = true;
    initializationError = null;
  } catch (error) {
    logger.error('Помилка ініціалізації', { message: error.message, stack: error.stack });
    initializationError = error;
    throw error;
  }
};

// Очищення старих записів активності
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

// Очищення активних тестів
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

// Запуск ініціалізації сервера
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

// Тест підключення до MongoDB
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

// Обробка favicon
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
      process.env.JWT_SECRET || 'your-secret-key',
      { expiresIn: '24h' }
    );

    res.cookie('token', token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    res.cookie('auth_token', token, {
      httpOnly: false,
      secure: process.env.NODE_ENV === 'production',
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    await logActivity(foundUser.username, 'увійшов на сайт', ipAddress);

    // Перенаправлення на сторінку вибору розділів для звичайних користувачів
    if (foundUser.role === 'admin') {
      res.json({ success: true, redirect: '/admin' });
    } else {
      res.json({ success: true, redirect: '/select-section' });
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

// Сторінка вибору розділів
app.get('/select-section', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    // Перевірка кешу розділів
    if (Object.keys(sectionNames).length === 0) {
      logger.warn('sectionNames порожній, повторне завантаження з MongoDB');
      await loadSectionsFromMongoDB();
      if (Object.keys(sectionNames).length === 0) {
        throw new Error('Не вдалося завантажити розділи з бази даних');
      }
    }
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Вибір розділу</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              display: flex; 
              justify-content: center; 
              align-items: center; 
              min-height: 100vh; 
              margin: 0; 
              background-color: #f0f0f0; 
            }
            .container { 
              display: grid; 
              grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); 
              gap: 20px; 
              padding: 20px; 
              max-width: 1200px; 
            }
            .section-card { 
              position: relative; 
              width: 100%; 
              height: 200px; 
              overflow: hidden; 
              border-radius: 10px; 
              cursor: pointer; 
              transition: transform 0.3s ease; 
            }
            .section-card:hover { 
              transform: scale(1.05); 
            }
            .section-card img { 
              width: 100%; 
              height: 100%; 
              object-fit: cover; 
            }
            .section-title { 
              position: absolute; 
              bottom: 0; 
              left: 0; 
              right: 0; 
              background: rgba(0, 0, 0, 0.6); 
              color: white; 
              padding: 10px; 
              text-align: center; 
              font-size: 1.2em; 
            }
            #logout { 
              position: fixed; 
              bottom: 20px; 
              left: 50%; 
              transform: translateX(-50%); 
              padding: 10px 20px; 
              font-size: 18px; 
              cursor: pointer; 
              border: none; 
              border-radius: 5px; 
              background-color: #ef5350; 
              color: white; 
            }
            #logout:hover { 
              background-color: #d32f2f; 
            }
            .no-sections { 
              color: red; 
              font-size: 18px; 
              margin-top: 20px; 
            }
            @media (max-width: 600px) {
              .container { 
                grid-template-columns: 1fr; 
              }
              .section-card { 
                height: 150px; 
              }
              .section-title { 
                font-size: 1em; 
              }
              #logout { 
                width: 90%; 
              }
            }
          </style>
        </head>
        <body>
          <div class="container">
            ${Object.entries(sectionNames).length > 0
              ? Object.entries(sectionNames).map(([id, data]) => `
                  <div class="section-card" onclick="window.location.href='/section?section=${id}'">
                    <img src="${data.imageUrl}" alt="${data.name.replace(/"/g, '\\"')}">
                    <div class="section-title">${data.name.replace(/"/g, '\\"')}</div>
                  </div>
                `).join('')
              : '<p class="no-sections">Немає доступних розділів</p>'
            }
          </div>
          <button id="logout" onclick="logout()">Вийти</button>
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
                if (response.ok) {
                  window.location.href = '/';
                } else {
                  alert('Помилка при виході');
                }
              } catch (error) {
                alert('Не вдалося вийти');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /select-section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки вибору розділів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /select-section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Сторінка розділу
app.get('/section', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
    }
    const sectionId = req.query.section;
    if (!sectionId || !sectionNames[sectionId]) {
      return res.status(400).send('Розділ не знайдено');
    }
    const tests = await db.collection('tests').find({ sectionId }).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${sectionNames[sectionId].name.replace(/"/g, '\\"')}</title>
          <style>
            body { 
              font-family: Arial, sans-serif; 
              display: flex; 
              justify-content: center; 
              align-items: center; 
              min-height: 100vh; 
              margin: 0; 
              background-color: #f0f0f0; 
            }
            .container { 
              text-align: center; 
              max-width: 600px; 
              padding: 20px; 
            }
            h1 { 
              font-size: 24px; 
              margin-bottom: 20px; 
            }
            .button { 
              display: inline-block; 
              padding: 15px 30px; 
              margin: 10px; 
              background-color: #007bff; 
              color: white; 
              text-decoration: none; 
              border-radius: 5px; 
              font-size: 1.2em; 
              transition: background-color 0.3s; 
            }
            .button:hover { 
              background-color: #0056b3; 
            }
            .test-buttons { 
              display: flex; 
              flex-direction: column; 
              align-items: center; 
              gap: 10px; 
              margin-top: 20px; 
            }
            .test-button { 
              padding: 10px; 
              font-size: 18px; 
              cursor: pointer; 
              width: 200px; 
              border: none; 
              border-radius: 5px; 
              background-color: #4CAF50; 
              color: white; 
            }
            .test-button:hover { 
              background-color: #45a049; 
            }
            #logout { 
              position: fixed; 
              bottom: 20px; 
              left: 50%; 
              transform: translateX(-50%); 
              padding: 10px 20px; 
              font-size: 18px; 
              cursor: pointer; 
              border: none; 
              border-radius: 5px; 
              background-color: #ef5350; 
              color: white; 
            }
            #logout:hover { 
              background-color: #d32f2f; 
            }
            .no-tests { 
              color: red; 
              font-size: 18px; 
              margin-top: 20px; 
            }
            @media (max-width: 600px) {
              h1 { font-size: 20px; }
              .button, .test-button { font-size: 16px; width: 90%; padding: 15px; }
              #logout { width: 90%; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>${sectionNames[sectionId].name.replace(/"/g, '\\"')}</h1>
            <a href="/materials?section=${sectionId}" class="button">Ознайомитись з навчальними матеріалами</a>
            <div class="test-buttons">
              ${tests.length > 0
                ? tests.map(test => `
                    <button class="test-button" onclick="window.location.href='/test?test=${test.testNumber}'">${test.name.replace(/"/g, '\\"')}</button>
                  `).join('')
                : '<p class="no-tests">Немає доступних тестів у цьому розділі</p>'
              }
            </div>
            <button id="logout" onclick="logout()">Вийти</button>
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
                if (response.ok) {
                  window.location.href = '/';
                } else {
                  alert('Помилка при виході');
                }
              } catch (error) {
                alert('Не вдалося вийти');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні сторінки розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка завантаження матеріалів
app.post('/admin/upload-material', checkAuth, async (req, res) => {
  if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
    return res.status(403).send('Доступно тільки для адміністраторів та інструкторів');
  }
  uploadMaterial.single('file')(req, res, async (err) => {
    const startTime = Date.now();
    try {
      if (err) {
        logger.error('Помилка завантаження файлу', { message: err.message, stack: err.stack });
        return res.status(400).send(err.message);
      }
      const { sectionId } = req.body;
      if (!sectionId || !sectionNames[sectionId]) {
        return res.status(400).send('Невірний ID розділу');
      }
      if (!req.file) {
        return res.status(400).send('Файл не надано');
      }

      // Завантаження файлу в Vercel Blob
      const blobResult = await blob.upload(req.file.originalname, req.file.buffer, {
        access: 'public', // Файл буде доступний публічно
      });

      const fileUrl = blobResult.url; // URL завантаженого файлу
      await db.collection('materials').insertOne({
        sectionId,
        fileUrl,
        uploadedBy: req.user,
        uploadDate: new Date().toISOString()
      });
      logger.info('Матеріал завантажено в Vercel Blob', { sectionId, fileUrl, uploadedBy: req.user });
      res.send(`
        <!DOCTYPE html>
        <html lang="uk">
          <head>
            <meta charset="UTF-8">
            <title>Матеріал завантажено</title>
            <style>
              body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
              button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            </style>
          </head>
          <body>
            <h1>Матеріал успішно завантажено</h1>
            <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          </body>
        </html>
      `);
    } catch (error) {
      logger.error('Помилка в /admin/upload-material', { message: error.message, stack: error.stack });
      res.status(500).send('Помилка при завантаженні матеріалу');
    } finally {
      const endTime = Date.now();
      logger.info('Маршрут /admin/upload-material виконано', { duration: `${endTime - startTime} мс` });
    }
  });
});

// Перегляд навчальних матеріалів
app.get('/materials', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const sectionId = req.query.section;
    if (!sectionId || !sectionNames[sectionId]) {
      return res.status(400).send('Розділ не знайдено');
    }
    const materials = await db.collection('materials').find({ sectionId }).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Навчальні матеріали</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f0f0f0; }
            .container { max-width: 800px; margin: 0 auto; }
            h1 { font-size: 24px; text-align: center; }
            .material { margin: 10px 0; }
            .material a { text-decoration: none; color: #007bff; }
            .material a:hover { text-decoration: underline; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
            button:hover { background-color: #0056b3; }
            @media (max-width: 600px) {
              h1 { font-size: 20px; }
              button { width: 100%; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Навчальні матеріали для ${sectionNames[sectionId].name.replace(/"/g, '\\"')}</h1>
            ${materials.length > 0
              ? materials.map(m => `
                  <div class="material">
                    <a href="${m.fileUrl}" target="_blank">${m.fileUrl.split('/').pop()}</a>
                    <p>Завантажено: ${new Date(m.uploadDate).toLocaleString('uk-UA')} | Автор: ${m.uploadedBy}</p>
                  </div>
                `).join('')
              : '<p>Немає доступних матеріалів</p>'
            }
            <button onclick="window.location.href='/section?section=${sectionId}'">Повернутися до розділу</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /materials', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні матеріалів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /materials виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка виходу користувача
app.post('/logout', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Отримано CSRF-токен у /logout', { token: req.body._csrf });
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    logActivity(req.user, 'покинув сайт', ipAddress);
    res.clearCookie('token');
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
      logger.info('Збереження результату в MongoDB із відповідями', { answers: result.answers });
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
    logger.info('saveResult виконано', { duration: `${endTimeLog - startTimeLog} мс` });
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
              <button onclick="window.location.href='/select-section'">Повернутися до вибору розділу</button>
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
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    logger.error('Помилка в /test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Сторінка інструкцій до тестів
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
              text-align: center;
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
              margin-bottom: 20px;
              color: #333;
            }
            p {
              font-size: 16px;
              line-height: 1.5;
              margin-bottom: 15px;
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
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              p { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Інструкція до тестів</h1>
            <p>Ласкаво просимо до системи тестування. Дотримуйтесь наступних інструкцій:</p>
            <p>1. Оберіть розділ на головній сторінці.</p>
            <p>2. У розділі ви можете ознайомитись із навчальними матеріалами або пройти тест.</p>
            <p>3. Під час тесту уважно читайте питання та обирайте правильні відповіді.</p>
            <p>4. Дотримуйтесь часових обмежень, якщо вони встановлені.</p>
            <p>5. Після завершення тесту ви побачите свої результати.</p>
            <button onclick="window.location.href='/select-section'">Повернутися до вибору розділу</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /instructions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні інструкцій');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /instructions виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Відображення питання тесту
app.get('/test/question', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const index = parseInt(req.query.index);
    if (isNaN(index)) {
      return res.status(400).send('Невірний індекс питання');
    }

    const activeTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!activeTest || !activeTest.questions || index < 0 || index >= activeTest.questions.length) {
      return res.redirect('/select-section');
    }

    const question = activeTest.questions[index];
    const testNumber = activeTest.testNumber;
    const timeLimit = activeTest.timeLimit;
    const isQuickTest = activeTest.isQuickTest;
    const timePerQuestion = activeTest.timePerQuestion ? activeTest.timePerQuestion * 1000 : null;

    let optionsHtml = '';
    if (question.type === 'multiple' || question.type === 'singlechoice') {
      optionsHtml = question.options.map((option, i) => `
        <label><input type="${question.type === 'singlechoice' ? 'radio' : 'checkbox'}" name="answer" value="${i}"> ${option.replace(/"/g, '\\"')}</label><br>
      `).join('');
    } else if (question.type === 'truefalse') {
      optionsHtml = question.options.map((option, i) => `
        <label><input type="radio" name="answer" value="${i}"> ${option.replace(/"/g, '\\"')}</label><br>
      `).join('');
    } else if (question.type === 'fillblank') {
      const blankCount = question.blankCount || 1;
      optionsHtml = Array.from({ length: blankCount }, (_, i) => `
        <input type="text" name="answer_${i}" placeholder="Введіть відповідь ${i + 1}" style="margin: 5px;"><br>
      `).join('');
    } else if (question.type === 'matching') {
      const leftItems = shuffleArray([...question.pairs.map(p => p.left)]);
      const rightItems = shuffleArray([...question.pairs.map(p => p.right)]);
      optionsHtml = `
        <div id="matching-container">
          <div style="float: left; width: 45%;">
            ${leftItems.map((left, i) => `
              <div class="matching-left" id="left-${i}" draggable="true" ondragstart="drag(event)" style="border: 1px solid #ccc; padding: 5px; margin: 5px;">${left}</div>
            `).join('')}
          </div>
          <div style="float: right; width: 45%;">
            ${rightItems.map((right, i) => `
              <div class="matching-right" id="right-${i}" ondrop="drop(event)" ondragover="allowDrop(event)" style="border: 1px solid #ccc; padding: 5px; margin: 5px;">${right}</div>
            `).join('')}
          </div>
        </div>
      `;
    } else if (question.type === 'input') {
      optionsHtml = `<input type="text" name="answer" placeholder="Введіть відповідь"><br>`;
    }

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Питання тесту ${testNames[testNumber].name.replace(/"/g, '\\"')}</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f5f5f5; margin: 0; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            p { font-size: 16px; }
            img { max-width: 100%; height: auto; margin: 10px 0; }
            label { display: block; margin: 10px 0; font-size: 14px; }
            input[type="text"] { padding: 5px; font-size: 14px; width: 200px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
            button:disabled { background-color: #cccccc; cursor: not-allowed; }
            #timer { font-size: 18px; color: red; margin: 10px 0; }
            #matching-container { overflow: hidden; }
            .matching-left, .matching-right { cursor: move; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              p, label { font-size: 14px; }
              input[type="text"] { width: 90%; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Питання ${index + 1} із ${activeTest.questions.length}</h1>
            <div id="timer">${isQuickTest && timePerQuestion ? `Залишилось: ${timePerQuestion / 1000} сек` : `Залишилось: ${Math.floor(timeLimit / 1000)} сек`}</div>
            <p>${question.text.replace(/"/g, '\\"')}</p>
            ${question.picture ? `<img src="${question.picture}" alt="Зображення">` : ''}
            <form id="answer-form" method="POST" action="/test/submit-answer">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <input type="hidden" name="index" value="${index}">
              ${optionsHtml}
              <button type="submit" id="submit-btn">${index === activeTest.questions.length - 1 ? 'Завершити тест' : 'Наступне питання'}</button>
            </form>
            <button onclick="window.location.href='/select-section'" id="cancel-btn">Скасувати тест</button>
          </div>
          <script>
            let timeLeft = ${isQuickTest && timePerQuestion ? timePerQuestion : timeLimit};
            const timerElement = document.getElementById('timer');
            const submitBtn = document.getElementById('submit-btn');
            const startTime = Date.now();

            function updateTimer() {
              timeLeft -= 1000;
              if (timeLeft <= 0) {
                submitBtn.disabled = true;
                document.getElementById('answer-form').submit();
                return;
              }
              timerElement.textContent = \`${isQuickTest && timePerQuestion ? 'Залишилось: ' : 'Залишилось: '}\${Math.floor(timeLeft / 1000)} сек\`;
              setTimeout(updateTimer, 1000);
            }
            updateTimer();

            let switchCount = 0;
            let timeAway = 0;
            let lastFocusTime = Date.now();
            let activityCount = 0;

            document.addEventListener('visibilitychange', () => {
              if (document.hidden) {
                lastFocusTime = Date.now();
              } else {
                timeAway += Date.now() - lastFocusTime;
                switchCount++;
              }
            });

            ['click', 'keydown', 'mousemove'].forEach(event => {
              document.addEventListener(event, () => activityCount++);
            });

            document.getElementById('answer-form').addEventListener('submit', async (e) => {
              e.preventDefault();
              submitBtn.disabled = true;
              const responseTime = Date.now() - startTime;
              const formData = new FormData(e.target);
              formData.append('timeAway', timeAway);
              formData.append('switchCount', switchCount);
              formData.append('responseTime', responseTime);
              formData.append('activityCount', activityCount);

              try {
                const response = await fetch('/test/submit-answer', {
                  method: 'POST',
                  body: formData
                });
                if (response.ok) {
                  const result = await response.json();
                  window.location.href = result.redirect;
                } else {
                  alert('Помилка при надсиланні відповіді');
                  submitBtn.disabled = false;
                }
              } catch (error) {
                alert('Не вдалося надіслати відповідь');
                submitBtn.disabled = false;
              }
            });

            function allowDrop(ev) { ev.preventDefault(); }
            function drag(ev) { ev.dataTransfer.setData('text', ev.target.id); }
            function drop(ev) {
              ev.preventDefault();
              const data = ev.dataTransfer.getData('text');
              const draggedElement = document.getElementById(data);
              const dropTarget = ev.target;
              if (dropTarget.classList.contains('matching-right')) {
                const temp = dropTarget.innerHTML;
                dropTarget.innerHTML = draggedElement.innerHTML;
                draggedElement.innerHTML = temp;
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /test/question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test/question виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка відповіді на питання
app.post('/test/submit-answer', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const index = parseInt(req.body.index);
    if (isNaN(index)) {
      return res.status(400).json({ success: false, message: 'Невірний індекс питання' });
    }

    const activeTest = await db.collection('active_tests').findOne({ user: req.user });
    if (!activeTest || !activeTest.questions || index < 0 || index >= activeTest.questions.length) {
      return res.status(400).json({ success: false, message: 'Тест не знайдено або завершено' });
    }

    const question = activeTest.questions[index];
    let userAnswer;

    if (question.type === 'multiple') {
      userAnswer = Array.isArray(req.body.answer) ? req.body.answer.map(i => parseInt(i)) : [parseInt(req.body.answer)];
      userAnswer = userAnswer.filter(i => !isNaN(i)).map(i => question.options[i]);
    } else if (question.type === 'singlechoice' || question.type === 'truefalse') {
      const answerIndex = parseInt(req.body.answer);
      userAnswer = !isNaN(answerIndex) ? [question.options[answerIndex]] : [];
    } else if (question.type === 'fillblank') {
      userAnswer = [];
      for (let i = 0; i < (question.blankCount || 1); i++) {
        const answer = req.body[`answer_${i}`]?.trim();
        if (answer) userAnswer.push(answer);
      }
    } else if (question.type === 'matching') {
      userAnswer = activeTest.questions[index].pairs.map((_, i) => {
        const left = document.querySelector(`#left-${i}`)?.innerHTML;
        const right = document.querySelector(`#right-${i}`)?.innerHTML;
        return [left, right];
      }).filter(pair => pair[0] && pair[1]);
    } else if (question.type === 'input') {
      userAnswer = [req.body.answer?.trim()];
    }

    activeTest.answers[index] = userAnswer;
    activeTest.answerTimestamps[index] = Date.now();
    activeTest.suspiciousActivity.timeAway += parseInt(req.body.timeAway) || 0;
    activeTest.suspiciousActivity.switchCount += parseInt(req.body.switchCount) || 0;
    activeTest.suspiciousActivity.responseTimes.push(parseInt(req.body.responseTime) || 0);
    activeTest.suspiciousActivity.activityCounts.push(parseInt(req.body.activityCount) || 0);

    await db.collection('active_tests').updateOne(
      { user: req.user },
      { $set: { answers: activeTest.answers, suspiciousActivity: activeTest.suspiciousActivity, answerTimestamps: activeTest.answerTimestamps } }
    );

    if (index < activeTest.questions.length - 1) {
      return res.json({ success: true, redirect: `/test/question?index=${index + 1}` });
    }

    let score = 0;
    let totalPoints = 0;
    let correctClicks = 0;
    let totalClicks = 0;
    const scoresPerQuestion = {};

    activeTest.questions.forEach((q, i) => {
      const userAns = activeTest.answers[i] || [];
      let questionScore = 0;
      totalPoints += q.points;

      if (q.type === 'multiple') {
        const correct = q.correctAnswers;
        const userCorrect = userAns.every(ans => correct.includes(ans)) && userAns.length === correct.length;
        if (userCorrect) {
          questionScore = q.points;
          correctClicks += userAns.length;
        }
        totalClicks += userAns.length;
      } else if (q.type === 'singlechoice' || q.type === 'truefalse') {
        if (userAns[0] === q.correctAnswers[0]) {
          questionScore = q.points;
          correctClicks += 1;
        }
        totalClicks += userAns.length;
      } else if (q.type === 'fillblank') {
        let allCorrect = true;
        (userAns || []).forEach((ans, idx) => {
          const correct = q.correctAnswers[idx];
          if (correct.includes('-')) {
            const [min, max] = correct.split('-').map(v => parseFloat(v.trim()));
            const value = parseFloat(ans);
            if (isNaN(value) || value < min || value > max) allCorrect = false;
          } else {
            const value = parseFloat(ans);
            const correctValue = parseFloat(correct);
            if (isNaN(value) || value !== correctValue) allCorrect = false;
          }
        });
        if (allCorrect && userAns.length === q.correctAnswers.length) {
          questionScore = q.points;
          correctClicks += userAns.length;
        }
        totalClicks += userAns.length;
      } else if (q.type === 'matching') {
        const userPairs = userAns || [];
        const correctPairs = q.correctPairs;
        const isCorrect = userPairs.every(([left, right], idx) => left === correctPairs[idx][0] && right === correctPairs[idx][1]);
        if (isCorrect && userPairs.length === correctPairs.length) {
          questionScore = q.points;
          correctClicks += userPairs.length;
        }
        totalClicks += userPairs.length;
      } else if (q.type === 'input') {
        const userValue = parseFloat(userAns[0]);
        const correct = q.correctAnswers[0];
        if (correct.includes('-')) {
          const [min, max] = correct.split('-').map(v => parseFloat(v.trim()));
          if (!isNaN(userValue) && userValue >= min && userValue <= max) {
            questionScore = q.points;
            correctClicks += 1;
          }
        } else {
          const correctValue = parseFloat(correct);
          if (!isNaN(userValue) && userValue === correctValue) {
            questionScore = q.points;
            correctClicks += 1;
          }
        }
        totalClicks += userAns.length;
      }

      score += questionScore;
      scoresPerQuestion[i] = questionScore;
    });

    const percentage = totalPoints > 0 ? (score / totalPoints) * 100 : 0;
    const timeAwayPercent = ((activeTest.suspiciousActivity.timeAway / (Date.now() - activeTest.startTime)) * 100).toFixed(2);
    const avgResponseTime = activeTest.suspiciousActivity.responseTimes.length
      ? (activeTest.suspiciousActivity.responseTimes.reduce((a, b) => a + b, 0) / activeTest.suspiciousActivity.responseTimes.length / 1000).toFixed(2)
      : 0;

    const suspiciousActivityDetails = {
      timeAwayPercent: parseFloat(timeAwayPercent),
      switchCount: activeTest.suspiciousActivity.switchCount,
      avgResponseTime: parseFloat(avgResponseTime),
      totalActivityCount: activeTest.suspiciousActivity.activityCounts.reduce((a, b) => a + b, 0, 0)
    };

    if (
      suspiciousActivityDetails.timeAwayPercent > config.suspiciousActivity.timeAwayThreshold ||
      suspiciousActivityDetails.switchCount > config.suspiciousActivity.switchCountThreshold
    ) {
      await sendSuspiciousActivityEmail(req.user, suspiciousActivityDetails);
    }

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await saveResult(
      req.user,
      activeTest.testNumber,
      score,
      totalPoints,
      activeTest.startTime,
      Date.now(),
      totalClicks,
      correctClicks,
      activeTest.questions.length,
      percentage,
      suspiciousActivityDetails,
      activeTest.answers,
      scoresPerQuestion,
      activeTest.variant,
      ipAddress,
      activeTest.testSessionId
    );

    await db.collection('active_tests').deleteOne({ user: req.user });

    res.json({ success: true, redirect: `/results?test=${activeTest.testNumber}&session=${activeTest.testSessionId}` });
  } catch (error) {
    logger.error('Помилка в /test/submit-answer', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при обробці відповіді' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /test/submit-answer виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Відображення результатів тесту
app.get('/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    const { test, session } = req.query;
    if (!test || !session) {
      return res.status(400).send('Невірні параметри');
    }

    const result = await db.collection('test_results').findOne({
      user: req.user,
      testNumber: test,
      testSessionId: session
    });

    if (!result) {
      return res.status(404).send('Результати не знайдено');
    }

    const suspiciousActivityHtml = `
      <p>Час поза вкладкою: ${result.suspiciousActivity.timeAwayPercent}%</p>
      <p>Переключення вкладок: ${result.suspiciousActivity.switchCount}</p>
      <p>Середній час відповіді: ${result.suspiciousActivity.avgResponseTime} сек</p>
      <p>Загальна кількість дій: ${result.suspiciousActivity.totalActivityCount}</p>
    `;

    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Результати тесту</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f5f5f5; margin: 0; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            p { font-size: 16px; margin: 5px 0; }
            button { padding: 10px 20px; margin: 10px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
            .details { margin-top: 20px; text-align: left; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              p { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Результати тесту "${testNames[test].name.replace(/"/g, '\\"')}"</h1>
            <p>Ваш результат: ${result.score} з ${result.totalPoints} балів</p>
            <p>Відсоток: ${Math.round(result.percentage)}%</p>
            <p>Час виконання: ${result.duration} секунд</p>
            <p>Варіант: ${result.variant}</p>
            <div class="details">
              <h2>Деталі активності:</h2>
              ${suspiciousActivityHtml}
            </div>
            <button onclick="window.location.href='/select-section'">Повернутися до вибору розділу</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /results', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні результатів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /results виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Адмін-панель
app.get('/admin', checkAuth, checkAdmin, async (req, res) => {
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
            body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f5f5f5; margin: 0; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
            input[type="file"] { margin: 10px 0; }
            .error { color: red; margin-top: 10px; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Адмін-панель</h1>
            <form id="import-users-form" method="POST" action="/admin/import-users" enctype="multipart/form-data">
              <input type="file" name="file" accept=".xlsx" required>
              <button type="submit">Імпортувати користувачів</button>
            </form>
            <form id="import-questions-form" method="POST" action="/admin/import-questions" enctype="multipart/form-data">
              <input type="text" name="testNumber" placeholder="Номер тесту" required>
              <input type="file" name="file" accept=".xlsx" required>
              <button type="submit">Імпортувати питання</button>
            </form>
            <form id="upload-material-form" method="POST" action="/admin/upload-material" enctype="multipart/form-data">
              <input type="text" name="sectionId" placeholder="ID розділу" required>
              <input type="file" name="file" accept=".pdf,.jpg,.jpeg,.png,.txt" required>
              <button type="submit">Завантажити матеріал</button>
            </form>
            <button onclick="window.location.href='/admin/edit-tests'">Редагувати тести</button>
            <button onclick="window.location.href='/admin/edit-sections'">Редагувати розділи</button>
            <button onclick="window.location.href='/admin/feedback'">Переглянути зворотний зв’язок</button>
            <button onclick="logout()">Вийти</button>
            <div id="error-message" class="error"></div>
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
                if (response.ok) {
                  window.location.href = '/';
                } else {
                  alert('Помилка при виході');
                }
              } catch (error) {
                alert('Не вдалося вийти');
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

// Імпорт користувачів
app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      return res.status(400).send('Файл не надано');
    }
    const count = await importUsersToMongoDB(req.file.buffer);
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
          <h1>Імпортовано ${count} користувачів</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка в /admin/import-users', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при імпорті користувачів: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-users виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Імпорт питань
app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.body;
    if (!req.file || !testNumber) {
      return res.status(400).send('Файл або номер тесту не надано');
    }
    const count = await importQuestionsToMongoDB(req.file.buffer, testNumber);
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
          <h1>Імпортовано ${count} питань для тесту ${testNumber}</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка в /admin/import-questions', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при імпорті питань: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-questions виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Список тестів для редагування
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
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .nav-btn:hover { background-color: #0056b3; }
            .action-btn:hover { opacity: 0.9; }
            @media (max-width: 600px) {
              table { font-size: 14px; }
              th, td { padding: 5px; }
              .nav-btn, .action-btn { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Редагувати тести</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="nav-btn" onclick="window.location.href='/admin/add-test'">Додати тест</button>
            <table>
              <tr>
                <th>Номер тесту</th>
                <th>Назва</th>
                <th>Розділ</th>
                <th>Дії</th>
              </tr>
              ${tests.length > 0
                ? tests.map(test => `
                    <tr>
                      <td>${test.testNumber}</td>
                      <td>${test.name.replace(/"/g, '\\"')}</td>
                      <td>${test.sectionId ? sectionNames[test.sectionId]?.name.replace(/"/g, '\\"') || 'Немає' : 'Немає'}</td>
                      <td>
                        <button class="action-btn edit" onclick="window.location.href='/admin/edit-test?testNumber=${test.testNumber}'">Редагувати</button>
                        <button class="action-btn delete" onclick="deleteTest('${test.testNumber}')">Видалити</button>
                      </td>
                    </tr>
                  `).join('')
                : '<tr><td colspan="4">Немає тестів</td></tr>'
              }
            </table>
          </div>
          <script>
            async function deleteTest(testNumber) {
              if (confirm('Ви впевнені, що хочете видалити тест ' + testNumber + '?')) {
                const formData = new URLSearchParams();
                formData.append('testNumber', testNumber);
                formData.append('_csrf', '${res.locals._csrf}');
                try {
                  const response = await fetch('/admin/delete-test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (response.ok) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні тесту');
                  }
                } catch (error) {
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
    res.status(500).send('Помилка при завантаженні тестів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-tests виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Додавання нового тесту
app.get('/admin/add-test', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Додати тест</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 600px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select { padding: 5px; width: 100%; max-width: 300px; margin-bottom: 10px; font-size: 14px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              input, select { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Додати тест</h1>
            <form method="POST" action="/admin/add-test" onsubmit="return validateForm()">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <label for="testNumber">Номер тесту:</label>
              <input type="text" id="testNumber" name="testNumber" required>
              <label for="sectionId">Розділ:</label>
              <select id="sectionId" name="sectionId">
                <option value="">Без розділу</option>
                ${Object.entries(sectionNames).map(([id, data]) => `<option value="${id}">${data.name.replace(/"/g, '\\"')}</option>`).join('')}
              </select>
              <label for="name">Назва тесту:</label>
              <input type="text" id="name" name="name" required>
              <label for="timeLimit">Ліміт часу (хвилини):</label>
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
              <label for="questionLimit">Ліміт питань (опціонально):</label>
              <input type="number" id="questionLimit" name="questionLimit" min="1">
              <label for="attemptLimit">Ліміт спроб:</label>
              <input type="number" id="attemptLimit" name="attemptLimit" min="1" required>
              <label for="isQuickTest">Швидкий тест:</label>
              <select id="isQuickTest" name="isQuickTest">
                <option value="true">Так</option>
                <option value="false">Ні</option>
              </select>
              <label for="timePerQuestion">Час на питання (секунди, для швидкого тесту):</label>
              <input type="number" id="timePerQuestion" name="timePerQuestion" min="1">
              <button type="submit" class="submit-btn">Додати</button>
            </form>
            <div id="error-message" class="error"></div>
            <button class="nav-btn" onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
          </div>
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

              if (!testNumber.match(/^[a-zA-Z0-9]+$/)) {
                errorMessage.textContent = 'Номер тесту може містити лише літери та цифри';
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
                errorMessage.textContent = 'Час на питання має бути принаймні 1 секунда';
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
    logger.error('Помилка в /admin/add-test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка додавання тесту
app.post('/admin/add-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const {
      testNumber,
      sectionId,
      name,
      timeLimit,
      randomQuestions,
      randomAnswers,
      questionLimit,
      attemptLimit,
      isQuickTest,
      timePerQuestion
    } = req.body;

    if (!testNumber || !name || !timeLimit || !attemptLimit) {
      return res.status(400).send('Обов’язкові поля не заповнені');
    }

    if (testNames[testNumber]) {
      return res.status(400).send('Тест із таким номером уже існує');
    }

    const testData = {
      sectionId: sectionId || null,
      name,
      timeLimit: parseInt(timeLimit) * 60,
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: timePerQuestion ? parseInt(timePerQuestion) : null
    };

    await saveTestToMongoDB(testNumber, testData);
    await loadTestsFromMongoDB();

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Тест додано</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Тест успішно додано</h1>
          <button onclick="window.location.href='/admin/edit-tests'">Повернутися до списку тестів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка в /admin/add-test (POST)', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-test (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Редагування тесту
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
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редагувати тест</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 600px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input, select { padding: 5px; width: 100%; max-width: 300px; margin-bottom: 10px; font-size: 14px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              input, select { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Редагувати тест ${testNumber}</h1>
            <form method="POST" action="/admin/edit-test" onsubmit="return validateForm()">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <input type="hidden" name="testNumber" value="${testNumber}">
              <label for="sectionId">Розділ:</label>
              <select id="sectionId" name="sectionId">
                <option value="">Без розділу</option>
                ${Object.entries(sectionNames).map(([id, data]) => `<option value="${id}" ${test.sectionId === id ? 'selected' : ''}>${data.name.replace(/"/g, '')}</option>`).join('')}
              </select>
              <label for="name">Назва тесту:</label>
              <input type="text" id="name" name="name" value="${test.name.replace(/"/g, '')}" required>
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
          </div>
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

// Обробка редагування тесту
app.post('/admin/edit-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const {
      testNumber,
      sectionId,
      name,
      timeLimit,
      randomQuestions,
      randomAnswers,
      questionLimit,
      attemptLimit,
      isQuickTest,
      timePerQuestion
    } = req.body;

    if (!testNumber || !name || !timeLimit || !attemptLimit) {
      return res.status(400).send('Обов’язкові поля не заповнені');
    }

    const testData = {
      sectionId: sectionId || null,
      name,
      timeLimit: parseInt(timeLimit) * 60,
      randomQuestions: randomQuestions === 'true',
      randomAnswers: randomAnswers === 'true',
      questionLimit: questionLimit ? parseInt(questionLimit) : null,
      attemptLimit: parseInt(attemptLimit),
      isQuickTest: isQuickTest === 'true',
      timePerQuestion: timePerQuestion ? parseInt(timePerQuestion) : null
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
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; }
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
    logger.error('Помилка в /admin/edit-test (POST)', { message: error.message, stack: error.stack});
    res.status(500).send('Помилка при оновлення тесту');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-test (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Видалити тест
app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.body; // Очікуємо testNumber як рядок
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).json({ success: false, message: 'Невірний номер тесту' });
    }
    await deleteTestFromMongoDB(testNumber);
    await loadTestsFromMongoDB(); // Оновлюємо кеш тестів

    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка в /admin/delete-test', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-test виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Список розділов для редактирования
app.get('/admin/edit-sections', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const sections = await db.collection('sections').find({}).toArray();
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редактировать разделы</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.edit { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            img { max-width: 100px; height: auto; }
            .nav-btn:hover { background-color: #0056b3; }
            .action-btn:hover { opacity: 0.9; }
            @media (max-width: 600px) {
              .container { padding: 10px; }
              table { font-size: 14px; }
              th, td { padding: 5px; }
              .nav-btn, .action-btn { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Редагувати розділи</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="nav-btn" onclick="window.location.href='/admin/add-section'">Додати розділ</button>
            <table>
              <tr>
                <th>ID розділу</th>
                <th>Назва</th>
                <th>Зображення</th>
                <th>Дії</th>
              </tr>
              ${sections.length > 0
                ? sections.map(section => `
                    <tr>
                      <td>${section.sectionId}</td>
                      <td>${section.name.replace(/"/g, '\\"')}</td>
                      <td><img src="${section.imageUrl}" alt="${section.name.replace(/"/g, '\\"')}" style="max-width: 100px;"></td>
                      <td>
                        <button class="action-btn edit" onclick="window.location.href='/admin/edit-section?sectionId=${section.sectionId}'">Редагувати</button>
                        <button class="action-btn delete" onclick="deleteSection('${section.sectionId}')">Видалити</button>
                      </td>
                    </tr>
                  `).join('')
                : '<tr><td colspan="4">Немає розділів</td></tr>'
              }
            </table>
          </div>
          <script>
            async function deleteSection(sectionId) {
              if (confirm('Ви впевнені, що хочете видалити розділ ' + sectionId + '?')) {
                const formData = new URLSearchParams();
                formData.append('sectionId', sectionId);
                formData.append('_csrf', '${res.locals._csrf}');
                try {
                  const response = await fetch('/admin/delete-section', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (response.ok) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні розділу');
                  }
                } catch (error) {
                  alert('Не вдалося видалити розділ');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-sections', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні розділів');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-sections виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Додавання нового розділу
app.get('/admin/add-section', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Додати розділ</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 600px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { padding: 5px; width: 100%; max-width: 300px; margin-bottom: 10px; font-size: 14px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              input { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Додати розділ</h1>
            <form method="POST" action="/admin/add-section" enctype="multipart/form-data">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <label for="sectionId">ID розділу:</label>
              <input type="text" id="sectionId" name="sectionId" required>
              <label for="name">Назва розділу:</label>
              <input type="text" id="name" name="name" required>
              <label for="image">Зображення:</label>
              <input type="file" id="image" name="image" accept="image/*">
              <button type="submit" class="submit-btn">Додати</button>
            </form>
            <div id="error-message" class="error"></div>
            <button class="nav-btn" onclick="window.location.href='/admin/edit-sections'">Повернутися до списку розділів</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/add-section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка додавання розділу
app.post('/admin/add-section', checkAuth, checkAdmin, uploadMaterial.single('image'), async (req, res) => {
  const startTime = Date.now();
  try {
    const { sectionId, name } = req.body;
    if (!sectionId || !name) {
      return res.status(400).send('ID та назва розділу обов’язкові');
    }
    if (sectionNames[sectionId]) {
      return res.status(400).send('Розділ із таким ID уже існує');
    }
    const imageUrl = req.file ? `/materials/${req.file.filename}` : '/images/default-section.png';
    await db.collection('sections').insertOne({
      sectionId,
      name,
      imageUrl
    });
    await loadSectionsFromMongoDB();
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Розділ додано</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        <body>
          <h1>Розділ успішно додано</h1>
          <button onclick="window.location.href='/admin/edit-sections'">Повернутися до списку розділів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка в /admin/add-section (POST)', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/add-section (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Редагування розділу
app.get('/admin/edit-section', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { sectionId } = req.query;
    if (!sectionId || !sectionNames[sectionId]) {
      return res.status(400).send('Невірний ID розділу');
    }
    const section = await db.collection('sections').findOne({ sectionId });
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Редагувати розділ</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 600px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { padding: 5px; width: 100%; max-width: 300px; margin-bottom: 10px; font-size: 14px; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; }
            .nav-btn { background-color: #007bff; color: white; }
            .submit-btn { background-color: #4CAF50; color: white; }
            .error { color: red; }
            img { max-width: 100px; margin-bottom: 10px; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              input { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Редагувати розділ ${sectionId}</h1>
            <form method="POST" action="/admin/edit-section" enctype="multipart/form-data">
              <input type="hidden" name="_csrf" value="${res.locals._csrf}">
              <input type="hidden" name="sectionId" value="${sectionId}">
              <label for="name">Назва:</label>
              <input type="text" id="name" name="name" value="${section.name.replace(/"/g, '\\"')}" required>
              <label>Поточне зображення:</label>
              <img src="${section.imageUrl}" alt="${section.name.replace(/"/g, '\\"')}">
              <label for="image">Нове зображення (опціонально):</label>
              <input type="file" id="image" name="image" accept="image/*">
              <button type="submit" class="submit-btn">Зберегти</button>
            </form>
            <div id="error-message" class="error"></div>
            <button class="nav-btn" onclick="window.location.href='/admin/edit-sections'">Повернутися до списку розділів</button>
          </div>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/edit-section', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Обробка редагування розділу
app.post('/admin/edit-section', checkAuth, checkAdmin, uploadMaterial.single('image'), async (req, res) => {
  const startTime = Date.now();
  try {
    const { sectionId, name } = req.body;
    if (!sectionId || !name) {
      return res.status(400).send('ID та назва розділу обов’язкові');
    }
    if (!sectionNames[sectionId]) {
      return res.status(400).send('Розділ не знайдено');
    }
    const imageUrl = req.file ? `/materials/${req.file.filename}` : undefined; // Зберігаємо нове зображення, якщо надано
    const updateData = { name };
    if (imageUrl) updateData.imageUrl = imageUrl;
    
    await db.collection('sections').updateOne(
      { sectionId },
      { $set: updateData }
    );
    await loadSectionsFromMongoDB(); // Оновлення кешу розділів
    logger.info('Розділ оновлено', { sectionId, name });
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Розділ оновлено</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
            button { padding: 10px 20px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
            button:hover { background-color: #45a049; }
          </style>
        </head>
        <body>
          <h1>Розділ успішно оновлено</h1>
          <button onclick="window.location.href='/admin/edit-sections'">Повернутися до списку розділів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка в /admin/edit-section (POST)', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при оновленні розділу');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/edit-section (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Видалення розділу
app.post('/admin/delete-section', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { sectionId } = req.body;
    if (!sectionId || !sectionNames[sectionId]) {
      return res.status(400).json({ success: false, message: 'Невірний ID розділу' });
    }
    await db.collection('sections').deleteOne({ sectionId });
    await db.collection('tests').updateMany({ sectionId }, { $set: { sectionId: null } }); // Видаляємо прив’язку тестів до розділу
    await db.collection('materials').deleteMany({ sectionId }); // Видаляємо всі матеріали розділу
    await loadSectionsFromMongoDB();
    logger.info('Розділ видалено', { sectionId });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка в /admin/delete-section', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні розділу' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-section виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Перегляд зворотного зв’язку
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
            body { font-family: Arial, sans-serif; padding: 20px; background-color: #f5f5f5; }
            .container { max-width: 800px; margin: auto; background-color: white; padding: 20px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
            h1 { font-size: 24px; margin-bottom: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .nav-btn:hover { background-color: #0056b3; }
            .action-btn:hover { opacity: 0.9; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; text-decoration: none; color: #007bff; }
            .pagination a:hover { text-decoration: underline; }
            @media (max-width: 600px) {
              .container { padding: 15px; }
              table { font-size: 14px; }
              th, td { padding: 5px; }
              .nav-btn, .action-btn { width: 100%; font-size: 14px; }
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>Зворотний зв’язок</h1>
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
            <button class="action-btn" onclick="deleteAllFeedback()">Видалити всі повідомлення</button>
            <table>
              <tr>
                <th>Користувач</th>
                <th>Повідомлення</th>
                <th>Час</th>
                <th>Дії</th>
              </tr>
              ${feedback.length > 0
                ? feedback.map(f => `
                    <tr>
                      <td>${f.user}</td>
                      <td>${f.message.replace(/"/g, '\\"')}</td>
                      <td>${new Date(f.timestamp).toLocaleString('uk-UA')}</td>
                      <td>
                        <button class="action-btn" onclick="deleteFeedback('${f._id}')">Видалити</button>
                      </td>
                    </tr>
                  `).join('')
                : '<tr><td colspan="4">Немає повідомлень</td></tr>'
              }
            </table>
            <div class="pagination">
              ${Array.from({ length: totalPages }, (_, i) => `
                <a href="/admin/feedback?page=${i + 1}" ${page === i + 1 ? 'style="font-weight: bold;"' : ''}>${i + 1}</a>
              `).join('')}
            </div>
          </div>
          <script>
            async function deleteFeedback(feedbackId) {
              if (confirm('Ви впевнені, що хочете видалити це повідомлення?')) {
                const formData = new URLSearchParams();
                formData.append('_csrf', '${res.locals._csrf}');
                try {
                  const response = await fetch('/admin/feedback/delete/' + feedbackId, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (response.ok) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні повідомлення');
                  }
                } catch (error) {
                  alert('Не вдалося видалити повідомлення');
                }
              }
            }

            async function deleteAllFeedback() {
              if (confirm('Ви впевнені, що хочете видалити всі повідомлення?')) {
                const formData = new URLSearchParams();
                formData.append('_csrf', '${res.locals._csrf}');
                try {
                  const response = await fetch('/admin/feedback/delete-all', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (response.ok) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні всіх повідомлень');
                  }
                } catch (error) {
                  alert('Не вдалося видалити всі повідомлення');
                }
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

// Видалення одного повідомлення зворотного зв’язку
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

// Видалення всіх повідомлень зворотного зв’язку
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

// Форма зворотного зв’язку для користувачів
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
              .container { padding: 15px; }
              h1 { font-size: 20px; }
              textarea { font-size: 14px; }
              button { width: 100%; font-size: 14px; }
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

// Обробка надсилання зворотного зв’язку
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

// Запуск сервера
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  logger.info(`Сервер запущено на порту ${PORT}`);
});
