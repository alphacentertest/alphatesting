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

// Увімкнення довіри до проксі для коректної роботи за розгортання
app.set('trust proxy', 1);

// Налаштування логування за допомогою winston
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

// Налаштування multer для зберігання файлів у пам’яті (сумісність із Vercel)
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 4 * 1024 * 1024 } // Обмеження розміру файлу до 4 МБ
});

// Налаштування nodemailer для відправки email
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
    logger.error('Помилка відправки email про підозрілу активність', { message: error.message, stack: error.stack });
  }
};

// Конфігурація параметрів підозрілої активності
const config = {
  suspiciousActivity: {
    timeAwayThreshold: 50, // Поріг часу поза вкладкою у відсотках
    switchCountThreshold: 5 // Поріг кількості переключень вкладок
  }
};

// Налаштування підключення до MongoDB
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/alpha?retryWrites=true&w=majority';
const client = new MongoClient(MONGODB_URI, {
  connectTimeoutMS: 5000,
  serverSelectionTimeoutMS: 5000
});
let db; // Змінна для збереження посилання на базу даних

// Клас CacheManager для кешування даних із сортуванням за полем order
class CacheManager {
  static cache = {};

  // Отримання даних із кешу або бази даних
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
    logger.info(`Оновлено кеш ${key} для тесту ${testNumber} з ${data.length} елементами за ${Date.now() - startTime} мс`);
    return data;
  }

  // Інвалідувати кеш для певного ключа
  static async invalidateCache(key, testNumber) {
    const cacheKey = `${key}:${testNumber}`;
    delete this.cache[cacheKey];
    logger.info(`Інвалідовано кеш для ${cacheKey}`);
  }

  // Отримання питань для конкретного тесту
  static async getQuestions(testNumber) {
    return await this.getOrFetch('questions', testNumber, async () => {
      const questions = await db.collection('questions').find({ testNumber }).sort({ order: 1 }).toArray();
      return questions;
    });
  }

  // Отримання всіх питань
  static async getAllQuestions() {
    return await this.getOrFetch('allQuestions', 'all', async () => {
      const questions = await db.collection('questions').find({}).sort({ order: 1 }).toArray();
      return questions;
    });
  }
}

// Кеш для зберігання даних користувачів і питань
let userCache = [];
const questionsCache = {};

// Функція для підключення до MongoDB із повторними спробами
const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    logger.info(`Спроба підключення до MongoDB (Спроба ${attempt} з ${maxAttempts}) з URI: ${MONGODB_URI}`);
    const startTime = Date.now();
    await client.connect();
    const endTime = Date.now();
    logger.info(`Підключено до MongoDB успішно за ${endTime - startTime} мс`);
    db = client.db('alpha');
    logger.info('Ініціалізовано базу даних', { databaseName: db.databaseName });
  } catch (error) {
    logger.error('Помилка підключення до MongoDB', { message: error.message, stack: error.stack });
    if (attempt < maxAttempts) {
      logger.info('Повторна спроба підключення до MongoDB через 5 секунд...');
      await new Promise(resolve => setTimeout(resolve, 5000));
      return connectToMongoDB(attempt + 1, maxAttempts);
    }
    throw error;
  }
};

let isInitialized = false; // Статус ініціалізації сервера
let initializationError = null; // Зберігання помилки ініціалізації
let testNames = {}; // Об’єкт для зберігання назв тестів

// Завантаження тестів із MongoDB
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
    logger.error('Помилка завантаження тестів із MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Збереження тесту в MongoDB
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
    logger.info('Тест збережено в MongoDB', { testNumber });
  } catch (error) {
    logger.error('Помилка збереження тесту в MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Видалення тесту з MongoDB
const deleteTestFromMongoDB = async (testNumber) => {
  try {
    await db.collection('tests').deleteOne({ testNumber });
    logger.info('Тест видалено з MongoDB', { testNumber });
  } catch (error) {
    logger.error('Помилка видалення тесту з MongoDB', { testNumber, message: error.message, stack: error.stack });
    throw error;
  }
};

// Налаштування middleware для Express
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Налаштування express-session із MongoStore
app.use(session({
  secret: process.env.SESSION_SECRET || 'your-secret-key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    client: client,
    dbName: 'alpha',
    collectionName: 'sessions',
    ttl: 24 * 60 * 60 // 24 години
  }),
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
    maxAge: 24 * 60 * 60 * 1000 // 24 години
  }
}));

// Middleware для генерації CSRF-токенів
app.use((req, res, next) => {
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = tokens.secretSync();
    logger.info('Згенеровано новий CSRF-секрет', { secret: req.session.csrfSecret });
  }
  const token = tokens.create(req.session.csrfSecret);
  res.locals._csrf = token;
  res.cookie('XSRF-TOKEN', token, { httpOnly: false });
  logger.info('Згенеровано CSRF-токен', { token });
  next();
});

// Валідація CSRF-токенів для POST-запитів
app.use((req, res, next) => {
  if (['POST', 'PUT', 'DELETE'].includes(req.method)) {
    const token = req.body._csrf || req.headers['x-csrf-token'];
    if (!token) {
      logger.error('Відсутній CSRF-токен у запиті', { method: req.method, url: req.url });
      return res.status(403).json({ success: false, message: 'CSRF-токен відсутній' });
    }
    if (!req.session.csrfSecret) {
      logger.error('Відсутній CSRF-секрет у сесії', { sessionId: req.sessionID });
      return res.status(403).json({ success: false, message: 'Помилка сесії: CSRF секрет відсутній' });
    }
    if (!tokens.verify(req.session.csrfSecret, token)) {
      logger.error('Помилка валідації CSRF-токена', {
        expectedSecret: req.session.csrfSecret,
        receivedToken: token
      });
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }
    logger.info('CSRF-токен успішно валідовано', { token });
  }
  next();
});

// Middleware для запобігання кешуванню
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Middleware для обробки помилок MongoDB
app.use((err, req, res, next) => {
  if (err.name === 'MongoNetworkError' || err.name === 'MongoServerError') {
    logger.error('Помилка MongoDB', { message: err.message, stack: err.stack });
    res.status(503).json({ success: false, message: 'Помилка з’єднання з базою даних. Спробуйте пізніше.' });
  } else {
    next(err);
  }
});

// Middleware для додавання водяного знака та захисту від скріншотів
app.use((req, res, next) => {
  const originalSend = res.send;
  res.send = function (body) {
    if (typeof body === 'string' && body.includes('</body>') && req.user) {
      // Додаємо водяний знак із ім’ям користувача та захист від скріншотів
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
          // Блокуємо комбінації клавіш для скріншотів
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

          // Блокуємо контекстне меню
          document.addEventListener('contextmenu', (e) => {
            e.preventDefault();
          });

          // Блокуємо виділення тексту та копіювання
          document.addEventListener('selectstart', (e) => {
            e.preventDefault();
          });
          document.addEventListener('copy', (e) => {
            e.preventDefault();
          });

          // Логуємо зміну видимості вкладки
          document.addEventListener('visibilitychange', () => {
            if (document.hidden) {
              console.log('Вкладка стала невидимою — можлива спроба зробити скріншот');
            }
          });
        </script>
      `;
      body = body.replace('</body>', `${watermarkScript}</body>`);
    }
    return originalSend.call(this, body);
  };
  next();
});

// Імпорт користувачів із Excel до MongoDB
const importUsersToMongoDB = async (buffer) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    let sheet = workbook.getWorksheet('Users') || workbook.getWorksheet('Sheet1');
    if (!sheet) {
      throw new Error('Лист "Users" або "Sheet1" не знайдено у файлі');
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
      throw new Error('Не знайдено користувачів у файлі');
    }
    await db.collection('users').deleteMany({});
    logger.info('Очищено всіх користувачів перед імпортом');
    await db.collection('users').insertMany(users);
    logger.info(`Імпортовано ${users.length} користувачів до MongoDB із хешованими паролями`);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    return users.length;
  } catch (error) {
    logger.error('Помилка імпорту користувачів до MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Імпорт питань із Excel до MongoDB
const importQuestionsToMongoDB = async (buffer, testNumber) => {
  try {
    logger.info('Відкриття workbook');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    logger.info('Workbook успішно відкрито');

    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      throw new Error('Лист "Questions" не знайдено у файлі');
    }
    logger.info('Знайдено лист "Questions"', { rowCount: sheet.rowCount });

    const MAX_ROWS = 1000;
    if (sheet.rowCount > MAX_ROWS + 1) {
      throw new Error(`Занадто багато рядків (${sheet.rowCount - 1} питань). Максимальна кількість питань: ${MAX_ROWS}.`);
    }

    const questions = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        try {
          logger.info(`Обробка рядка ${rowNumber}`);
          const rowValues = row.values.slice(1);
          let questionText = rowValues[1];
          if (typeof questionText === 'object' && questionText !== null) {
            questionText = questionText.text || questionText.value || '[Невірний текст питання]';
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

          // Нормалізація назви зображення
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

          // Перевірка наявності зображення
          if (normalizedPicture) {
            logger.info(`Обробка поля зображення: ${normalizedPicture}`, { testNumber, rowNumber });
            const pictureMatch = normalizedPicture.match(/^Picture(\d+)$/i);
            if (pictureMatch) {
              const pictureNumber = pictureMatch[1];
              const targetFileNameBase = `Picture${pictureNumber}`;
              const extensions = ['.png', '.jpg', '.jpeg', '.gif'];
              let found = false;
              const imageDir = path.join(__dirname, 'public', 'images');
              const filesInDir = fs.existsSync(imageDir) ? fs.readdirSync(imageDir) : [];
              logger.info(`Доступні файли в public/images: ${filesInDir.join(', ')}`, { testNumber, rowNumber });

              for (const ext of extensions) {
                const expectedFileName = `${targetFileNameBase}${ext}`;
                const fileExists = filesInDir.some(file => file.toLowerCase() === expectedFileName.toLowerCase());
                if (fileExists) {
                  const matchedFile = filesInDir.find(file => file.toLowerCase() === expectedFileName.toLowerCase());
                  const imagePath = path.join(imageDir, matchedFile);
                  if (fs.existsSync(imagePath)) {
                    questionData.picture = `/images/Picture${pictureNumber}${ext.toLowerCase()}`;
                    logger.info(`Зображення знайдено: ${questionData.picture}`, { testNumber, rowNumber });
                    found = true;
                    break;
                  } else {
                    logger.warn(`Файл ${matchedFile} знайдено в списку, але не існує на диску`, { testNumber, rowNumber });
                  }
                }
              }
              if (!found) {
                logger.warn(`Зображення не знайдено для ${normalizedPicture}. Доступні файли: ${filesInDir.join(', ')}`, { testNumber, rowNumber });
                questionData.picture = null;
              }
            } else {
              logger.warn(`Невірний формат зображення: ${picture}. Очікуваний формат: PictureX або PictureX.png`, { testNumber, rowNumber });
            }
          }

          if (type === 'matching') {
            questionData.pairs = options.map((opt, idx) => ({
              left: opt || '',
              right: questionData.correctAnswers[idx] || ''
            })).filter(pair => pair.left && pair.right);
            if (questionData.pairs.length === 0) throw new Error('Для типу Matching потрібні пари відповідей');
            questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
          }

          if (type === 'fillblank') {
            questionText = questionText.replace(/\s*___\s*/g, '___');
            const blankCount = (questionText.match(/___/g) || []).length;
            if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) throw new Error('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
            questionData.text = questionText;
            questionData.blankCount = blankCount;

            // Валідація правильних відповідей для fillblank
            questionData.correctAnswers.forEach((correctAnswer, idx) => {
              if (correctAnswer.includes('-')) {
                const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
                if (isNaN(min) || isNaN(max) || min > max) {
                  throw new Error(`Невірний формат діапазону для правильної відповіді ${idx + 1} у рядку ${rowNumber}. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.`);
                }
              } else {
                const value = parseFloat(correctAnswer);
                if (isNaN(value)) {
                  throw new Error(`Правильна відповідь ${idx + 1} у рядку ${rowNumber} для типу Fillblank має бути числом або діапазоном у форматі "число1-число2".`);
                }
              }
            });
          }

          if (type === 'singlechoice') {
            if (correctAnswers.length !== 1 || options.length < 2) throw new Error('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
            questionData.correctAnswer = correctAnswers[0];
          }

          if (type === 'input') {
            if (questionData.correctAnswers.length !== 1) {
              throw new Error(`Для типу Input у рядку ${rowNumber} потрібна одна правильна відповідь`);
            }
            const correctAnswer = questionData.correctAnswers[0];
            if (correctAnswer.includes('-')) {
              const [min, max] = correctAnswer.split('-').map(val => parseFloat(val.trim()));
              if (isNaN(min) || isNaN(max) || min > max) {
                throw new Error(`Невірний формат діапазону для правильної відповіді у рядку ${rowNumber}. Використовуйте формат "число1-число2", наприклад, "12-14", де число1 <= число2.`);
              }
            } else {
              const value = parseFloat(correctAnswer);
              if (isNaN(value)) {
                throw new Error(`Правильна відповідь у рядку ${rowNumber} для типу Input має бути числом або діапазоном у форматі "число1-число2".`);
              }
            }
          }

          questions.push(questionData);
          logger.info(`Рядок ${rowNumber} успішно оброблено`);
        } catch (error) {
          throw new Error(`Помилка в рядку ${rowNumber}: ${error.message}`);
        }
      }
    });
    if (questions.length === 0) {
      throw new Error('Не знайдено питань у файлі');
    }
    logger.info('Видалення існуючих питань для тесту', { testNumber });
    await db.collection('questions').deleteMany({ testNumber });
    logger.info('Вставка нових питань', { count: questions.length });
    await db.collection('questions').insertMany(questions);
    logger.info(`Імпортовано ${questions.length} питань для тесту ${testNumber} до MongoDB`);
    await CacheManager.invalidateCache('questions', testNumber);
    return questions.length;
  } catch (error) {
    logger.error('Помилка імпорту питань до MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Функція для випадкового перемішування масиву (Fisher-Yates)
const shuffleArray = (array) => {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
};

// Завантаження користувачів до кешу
const loadUsersToCache = async () => {
  try {
    const startTime = Date.now();
    userCache = await db.collection('users').find({}).toArray();
    const endTime = Date.now();
    logger.info(`Оновлено кеш користувачів із ${userCache.length} користувачами за ${endTime - startTime} мс`);
  } catch (error) {
    logger.error('Помилка оновлення кешу користувачів', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Завантаження питань для тесту
const loadQuestions = async (testNumber) => {
  try {
    const startTime = Date.now();
    if (questionsCache[testNumber]) {
      const endTime = Date.now();
      logger.info(`Завантажено ${questionsCache[testNumber].length} питань для тесту ${testNumber} з кешу за ${endTime - startTime} мс`);
      return questionsCache[testNumber];
    }

    const questions = await db.collection('questions').find({ testNumber: testNumber.toString() }).sort({ order: 1 }).toArray();
    const endTime = Date.now();
    if (questions.length === 0) {
      throw new Error(`Не знайдено питань у MongoDB для тесту ${testNumber}`);
    }
    questionsCache[testNumber] = questions;
    logger.info(`Завантажено ${questions.length} питань для тесту ${testNumber} із MongoDB за ${endTime - startTime} мс`);
    return questions;
  } catch (error) {
    logger.error(`Помилка в loadQuestions (тест ${testNumber})`, { message: error.message, stack: error.stack });
    throw error;
  }
};

// Middleware для перевірки ініціалізації сервера
const ensureInitialized = (req, res, next) => {
  if (!isInitialized) {
    if (initializationError) {
      logger.error('Сервер не ініціалізовано через помилку', { message: initializationError.message, stack: initializationError.stack });
      return res.status(500).json({ success: false, message: `Помилка ініціалізації сервера: ${initializationError.message}` });
    }
    logger.warn('Сервер ще ініціалізується...');
    return res.status(503).json({ success: false, message: 'Сервер ініціалізується, спробуйте знову пізніше' });
  }
  next();
};

// Оновлення паролів користувачів із хешуванням
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
  const endTime = Date.now();
  logger.info('Оновлено паролі користувачів із хешами', { duration: `${endTime - startTime} мс` });
  await CacheManager.invalidateCache('users', null);
};

// Ініціалізація сервера
const initializeServer = async () => {
  let attempt = 1;
  const maxAttempts = 5;
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
    logger.info('Індекси MongoDB успішно створено');

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
      logger.info('Міграція ролей для існуючих користувачів завершена');
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
      logger.info('Міграція стандартних тестів до MongoDB', { count: Object.keys(defaultTests).length });
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

// Очищення старих записів журналу активності (старше 30 днів)
const cleanupActivityLog = async () => {
  try {
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const result = await db.collection('activity_log').deleteMany({
      timestamp: { $lt: thirtyDaysAgo.toISOString() }
    });
    logger.info('Очищено старі записи журналу активності', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення журналу активності', { message: error.message, stack: error.stack });
  }
};

// Очищення старих записів active_tests (старше 24 годин)
const cleanupActiveTests = async () => {
  try {
    const twentyFourHoursAgo = new Date(Date.now() - 24 * 60 * 60 * 1000);
    const result = await db.collection('active_tests').deleteMany({
      startTime: { $lt: twentyFourHoursAgo.getTime() }
    });
    logger.info('Очищено старі активні тести', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Помилка очищення активних тестів', { message: error.message, stack: error.stack });
  }
};

// Запуск періодичних задач очищення
setInterval(cleanupActivityLog, 24 * 60 * 60 * 1000);
setInterval(cleanupActiveTests, 24 * 60 * 60 * 1000);

// Асинхронна ініціалізація сервера
(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
    await cleanupActivityLog();
    await cleanupActiveTests();
  } catch (error) {
    logger.error('Помилка запуску сервера через помилку ініціалізації', { message: error.message, stack: error.stack });
    process.exit(1);
  }
})();

// Налаштування обмеження спроб входу
const MAX_LOGIN_ATTEMPTS = 30;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;

// Перевірка кількості спроб входу
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
    throw new Error('Перевищено ліміт спроб входу (30 на день). Спробуйте знову завтра.');
  }

  await db.collection('login_attempts').updateOne(
    { ipAddress, lastAttempt: { $gte: startOfDay, $lt: endOfDay } },
    { $inc: { count: 1 }, $set: { lastAttempt: now } },
    { upsert: true }
  );
};

// Логування активності користувача
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
    const endTime = Date.now();
    logger.info(`Залогована активність: ${user} - ${action} о ${adjustedTimestamp}, IP: ${ipAddress}`, { duration: `${endTime - startTime} мс` });
  } catch (error) {
    logger.error('Помилка логування активності', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Тестовий маршрут для перевірки підключення до MongoDB
app.get('/test-mongo', async (req, res) => {
  try {
    if (!db) {
      throw new Error('Підключення до MongoDB не встановлено');
    }
    await db.collection('users').findOne();
    res.json({ success: true, message: 'Підключення до MongoDB успішне' });
  } catch (error) {
    logger.error('Тест MongoDB провалився', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка підключення до MongoDB', error: error.message });
  }
});

// Тестовий маршрут API
app.get('/api/test', (req, res) => {
  logger.info('Обробка запиту /api/test');
  res.json({ success: true, message: 'Сервер Express працює на /api/test' });
});

// Обробка запиту favicon.ico
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

    // Перевірка кешу користувачів
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
    const endTime = Date.now();
    logger.info('Маршрут /login виконано', { duration: `${endTime - startTime} мс` });
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
            let lastBlurTime = 0;
            let switchCount = ${userTest.suspiciousActivity?.switchCount || 0};
            let lastActivityTime = Date.now();
            let activityCount = ${userTest.suspiciousActivity?.activityCounts?.[index] || 0};
            let lastMouseMoveTime = 0;
            let lastKeydownTime = 0;
            const debounceDelay = 100;
            const blurDebounceDelay = 200; // Затримка для дебаунсингу подій blur
            let blurTimeout = null;
            let selectedOptions = ${selectedOptionsString};
            let matchingPairs = ${JSON.stringify(answers[index] || [])};
            let questionTimeRemaining = timePerQuestion;
            let currentQuestionIndex = ${index};
            let lastGlobalUpdateTime = Date.now();
            let isSaving = false;
            let hasMovedToNext = false;
            let questionStartTime = ${questionStartTime[index]};

            // Функція для автоматичного збереження відповіді
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

            // Функція для збереження відповіді та переходу до наступного питання
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
                      window.location.href = '/test/question?index=' + nextIndex;
                    } else {
                      setTimeout(() => {
                        window.location.href = '/result';
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

            // Показ модального вікна підтвердження завершення тесту
            function showConfirm(index) {
              document.getElementById('confirm-modal').style.display = 'block';
            }

            // Приховування модального вікна
            function hideConfirm() {
              document.getElementById('confirm-modal').style.display = 'none';
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
                    window.location.href = '/result';
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

            // Оновлення глобального таймера
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
                  setTimeout(() => {
                    window.location.href = '/result';
                  }, 300);
                });
              }
            }

            setInterval(updateGlobalTimer, 1000);

            // Таймер для швидкого тесту
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
                    setTimeout(() => {
                      window.location.href = '/result';
                    }, 300);
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

            // Відстеження втрати фокусу вкладки
            window.addEventListener('blur', () => {
              if (!blurTimeout) {
                blurTimeout = setTimeout(() => {
                  if (lastBlurTime === 0) {
                    lastBlurTime = performance.now();
                    switchCount = Math.min(switchCount + 1, 1000); // Обмеження switchCount
                    console.log('Вкладка втратила фокус, початок підрахунку часу:', lastBlurTime, 'Кількість переключень:', switchCount);
                  }
                  blurTimeout = null;
                }, blurDebounceDelay);
              }
            });

            // Відстеження повернення фокусу вкладки
            window.addEventListener('focus', () => {
              if (blurTimeout) {
                clearTimeout(blurTimeout);
                blurTimeout = null;
              }
              if (lastBlurTime > 0) {
                const now = performance.now();
                const awayDuration = Math.min((now - lastBlurTime) / 1000, 60); // Обмеження до 60 секунд
                timeAway += awayDuration;
                console.log('Вкладка отримала фокус, накопичено часу поза вкладкою:', awayDuration, 'Загальний timeAway:', timeAway);
                lastBlurTime = 0;
                saveCurrentAnswer(currentQuestionIndex);
              }
            });

            // Скидання questionStartTime після тривалого простою
            document.addEventListener('visibilitychange', () => {
              if (!document.hidden) {
                const now = Date.now();
                const timeSinceLastActivity = (now - lastActivityTime) / 1000;
                if (timeSinceLastActivity > 300) { // Якщо простій більше 5 хвилин
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

            // Дебаунсинг для руху миші
            function debounceMouseMove() {
              const now = Date.now();
              if (now - lastMouseMoveTime >= debounceDelay) {
                lastMouseMoveTime = now;
                lastActivityTime = now;
                activityCount++;
              }
            }

            // Дебаунсинг для натискання клавіш
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

            // Обробка кліків по варіантах відповідей
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

            // Ініціалізація sortable для ordering
            const sortable = document.getElementById('sortable-options');
            if (sortable) {
              new Sortable(sortable, { animation: 150 });
            }

            // Ініціалізація sortable для matching
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
                console.log('Натискання кнопки повернення, перенаправлення на /select-test');
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
    res.status(500).send('Помилка при завантаженні результатів: ' + (error.message || 'Невідома помилка'));
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /result виконано', { duration: `${endTime - startTime} мс` });
  }
});

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
    const endTime = Date.now();
    logger.info('Маршрут /results виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для адмін-панелі
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
          <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button><br>
          <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button><br>
          <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
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
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для керування користувачами
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
                    alert('Помилка при видаленні користувача: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення користувача:', error);
                  alert('Не вдалося видалити користувача. Перевірте ваше з’єднання з Інтернетом.');
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
            <button class="sort-btn" onclick="window.location.href='/admin/questions?page=${page}&sortBy=testName'">Сортувати за назвою тесту</button>
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
    const endTime = Date.now();
    logger.info('Маршрут /admin/questions виконано', { duration: `${endTime - startTime} мс` });
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
              const csrfToken = document.getElementById('_csrf').value;

              if (!csrfToken) {
                errorMessage.textContent = 'CSRF-токен відсутній. Оновіть сторінку та спробуйте знову.';
                return;
              }

              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано.';
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
                  headers: {
                    'X-CSRF-Token': csrfToken
                  }
                });

                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'HTTP-помилка! статус: ' + response.status);
                }

                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка під час завантаження файлу:', error);
                errorMessage.textContent = 'Помилка при завантаженні файлу: ' + error.message;
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
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-users виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки імпорту користувачів
app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      logger.error('Файл не надано для імпорту користувачів');
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
          <h1>Успішно імпортовано ${count} користувачів</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Помилка імпорту користувачів', { message: error.message, stack: error.stack });
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
          <h1>Помилка імпорту користувачів</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-users'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-users (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для імпорту питань
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
              const csrfToken = document.getElementById('_csrf').value;

              if (!csrfToken) {
                errorMessage.textContent = 'CSRF-токен відсутній. Оновіть сторінку та спробуйте знову.';
                return;
              }

              if (!fileInput.files[0]) {
                errorMessage.textContent = 'Файл не вибрано.';
                return;
              }

              submitBtn.disabled = true;
              submitBtn.textContent = 'Завантаження...';

              const formData = new FormData();
              formData.append('testNumber', testNumber);
              formData.append('file', fileInput.files[0]);

              try {
                const response = await fetch('/admin/import-questions', {
                  method: 'POST',
                  body: formData,
                  headers: {
                    'X-CSRF-Token': csrfToken
                  }
                });

                if (!response.ok) {
                  const result = await response.json();
                  throw new Error(result.message || 'HTTP-помилка! статус: ' + response.status);
                }

                const result = await response.text();
                document.body.innerHTML = result;
              } catch (error) {
                console.error('Помилка під час завантаження файлу:', error);
                errorMessage.textContent = 'Помилка при завантаженні файлу: ' + error.message;
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
    res.status(500).send('Помилка при імпорті питань');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-questions виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для обробки імпорту питань
app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      logger.error('Файл не надано для імпорту питань');
      return res.status(400).send('Файл не надано');
    }
    const testNumber = req.body.testNumber;
    if (!testNumber || !testNames[testNumber]) {
      logger.error('Невірний номер тесту для імпорту питань', { testNumber });
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
          <h1>Успішно імпортовано ${count} питань для тесту ${testNames[testNumber].name.replace(/"/g, '\\"')}</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
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
          <h1>Помилка імпорту питань</h1>
          <p class="error">${error.message}</p>
          <button onclick="window.location.href='/admin/import-questions'">Спробувати знову</button>
        </body>
      </html>
    `);
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/import-questions (POST) виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для перегляду результатів тестів
app.get('/admin/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
      return res.status(403).send('Доступно тільки для адміністраторів та інструкторів');
    }

    // Отримуємо всі результати та сортуємо за endTime у спадному порядку
    const results = await db.collection('test_results').find({}).sort({ endTime: -1 }).toArray();
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
            .nav-btn, .action-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; }
            .action-btn.view { background-color: #4CAF50; color: white; }
            .action-btn.delete { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .suspicious { color: red; }
            .details { white-space: pre-wrap; max-width: 300px; overflow-wrap: break-word; }
          </style>
        </head>
        <body>
          <h1>Результати тестів</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Тест</th>
              <th>Варіант</th>
              <th>Очки/%</th>
              <th>Максимум</th>
              <th>Початок</th>
              <th>Кінець</th>
              <th>Тривалість (хв:сек)</th>
              <th>Підозріла активність (%)</th>
              <th>Деталі активності</th>
              <th>Дія</th>
            </tr>
    `;
    if (!results || results.length === 0) {
      html += '<tr><td colspan="11">Немає результатів</td></tr>';
    } else {
      results.forEach(result => {
        const startTime = new Date(result.startTime).toLocaleTimeString('uk-UA', { hour12: false }) + ' ' + new Date(result.startTime).toLocaleDateString('uk-UA');
        const endTime = new Date(result.endTime).toLocaleTimeString('uk-UA', { hour12: false }) + ' ' + new Date(result.endTime).toLocaleDateString('uk-UA');
        const durationSec = result.duration || Math.round((new Date(result.endTime) - new Date(result.startTime)) / 1000);
        const minutes = Math.floor(durationSec / 60).toString().padStart(2, '0');
        const seconds = (durationSec % 60).toString().padStart(2, '0');
        const timeAwayPercent = result.suspiciousActivity?.timeAway
          ? Math.round((result.suspiciousActivity.timeAway / result.duration) * 100)
          : 0;
        const switchCount = result.suspiciousActivity?.switchCount || 0;
        const avgResponseTime = result.suspiciousActivity?.responseTimes
          ? (result.suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / result.suspiciousActivity.responseTimes.length).toFixed(2)
          : 0;
        const totalActivityCount = result.suspiciousActivity?.activityCounts
          ? result.suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0)
          : 0;

        const isSuspicious = timeAwayPercent > config.suspiciousActivity.timeAwayThreshold ||
                            switchCount > config.suspiciousActivity.switchCountThreshold;

        const activityDetails = `Час поза вкладкою: ${timeAwayPercent}%\n` +
                               `Переключення вкладок: ${switchCount}\n` +
                               `Середній час відповіді (сек): ${avgResponseTime}`;

        html += `
          <tr class="${isSuspicious ? 'suspicious' : ''}">
            <td>${result.user}</td>
            <td>${testNames[result.testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}</td>
            <td>${result.variant || 'Немає'}</td>
            <td>${result.score} / ${Math.round(result.percentage)}%</td>
            <td>${result.totalPoints}</td>
            <td>${startTime}</td>
            <td>${endTime}</td>
            <td>${minutes} хв ${seconds} сек</td>
            <td>${timeAwayPercent}%</td>
            <td class="details">${activityDetails}</td>
            <td>
              <button class="action-btn view" onclick="viewResult('${result._id}')">Перегляд</button>
              <button class="action-btn delete" onclick="deleteResult('${result._id}')">🗑️ Видалити</button>
            </td>
          </tr>
        `;
      });
    }
    html += `
          </table>
          <script>
            async function viewResult(id) {
              window.location.href = '/admin/view-result?id=' + id;
            }

            async function deleteResult(id) {
              if (confirm('Ви впевнені, що хочете видалити цей результат?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('id', id);
                  formData.append('_csrf', '${res.locals._csrf}');
                  const response = await fetch('/admin/delete-result', {
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
                    alert('Помилка при видаленні результату: ' + result.message);
                  }
                } catch (error) {
                  console.error('Помилка видалення результату:', error);
                  alert('Не вдалося видалити результат. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }
          </script>
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

// Маршрут для перегляду детального результату
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

    const questions = await db.collection('questions')
      .find({ testNumber: result.testNumber })
      .sort({ order: 1 })
      .toArray();

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
            .nav-btn { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; background-color: #007bff; color: white; }
            .answers { white-space: pre-wrap; max-width: 300px; overflow-wrap: break-word; line-height: 1.8; }
          </style>
        </head>
        <body>
          <h1>Деталі результату для користувача ${result.user}</h1>
          <p>
            Тест: ${testNames[result.testNumber]?.name.replace(/"/g, '\\"') || 'Невідомий тест'}<br>
            Результат: ${Math.round(result.percentage)}%<br>
            Кількість питань: ${result.totalQuestions}<br>
            Правильних відповідей: ${result.correctClicks}<br>
            Набрано балів: ${result.score}<br>
            Максимально можлива кількість балів: ${result.totalPoints}<br>
            Час поза вкладкою: ${result.suspiciousActivity?.timeAway ? Math.round((result.suspiciousActivity.timeAway / result.duration) * 100) : 0}%<br>
            Переключення вкладок: ${result.suspiciousActivity?.switchCount || 0}<br>
            Середній час відповіді: ${result.suspiciousActivity?.responseTimes
              ? (result.suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / result.suspiciousActivity.responseTimes.length).toFixed(2)
              : 0} с<br>
            Загальна активність: ${result.suspiciousActivity?.activityCounts
              ? result.suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0)
              : 0}<br>
            Дата завершення: ${new Date(result.endTime).toLocaleString('uk-UA')}<br>
            Варіант: ${result.variant || 'Немає'}
          </p>
          <table>
            <tr>
              <th>Питання</th>
              <th>Ваша відповідь</th>
              <th>Правильна відповідь</th>
              <th>Бали</th>
            </tr>
    `;

    // Перебираємо питання в порядку order і зіставляємо з відповідями за індексом
    questions.forEach((question, index) => {
      const userAnswer = result.answers[index] !== undefined ? result.answers[index] : 'Не відповіли';
      let correctAnswer;
      if (question.type === 'matching') {
        correctAnswer = question.correctPairs.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ');
      } else if (question.type === 'fillblank') {
        correctAnswer = question.correctAnswers.join(', ');
      } else if (question.type === 'singlechoice') {
        correctAnswer = question.correctAnswer;
      } else {
        correctAnswer = question.correctAnswers.join(', ');
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
          <td>${question.text}</td>
          <td class="answers">${userAnswerDisplay}</td>
          <td>${correctAnswer}</td>
          <td>${questionScore} з ${question.points}</td>
        </tr>
      `;
    });

    html += `
          </table>
          <button class="nav-btn" onclick="window.location.href='/admin/results'">Повернутися до результатів</button>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Помилка в /admin/view-result', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при перегляді результату');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/view-result виконано', { duration: `${endTime - startTime} мс` });
  }
});

// Маршрут для видалення результату
app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.body;
    if (!id || !ObjectId.isValid(id)) {
      return res.status(400).json({ success: false, message: 'Невірний ідентифікатор результату' });
    }
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    res.json({ success: true });
  } catch (error) {
    logger.error('Помилка видалення результату', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/delete-result виконано', { duration: `${endTime - startTime} мс` });
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

// Маршрут для перегляду журналу активності
app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 20;
    const skip = (page - 1) * limit;

    const activities = await db.collection('activity_log')
      .find({})
      .sort({ timestamp: -1 })
      .skip(skip)
      .limit(limit)
      .toArray();

    const totalActivities = await db.collection('activity_log').countDocuments();
    const totalPages = Math.ceil(totalActivities / limit);

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Журнал активності</title>
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
          <h1>Журнал активності</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <table>
            <tr>
              <th>Користувач</th>
              <th>Дія</th>
              <th>IP-адреса</th>
              <th>Час</th>
              <th>Додаткова інформація</th>
            </tr>
    `;
    if (!activities || activities.length === 0) {
      html += '<tr><td colspan="5">Немає записів активності</td></tr>';
    } else {
      activities.forEach(activity => {
        const timestamp = new Date(activity.timestamp).toLocaleString('uk-UA');
        const additionalInfo = JSON.stringify(activity.additionalInfo, null, 2);
        html += `
          <tr>
            <td>${activity.user}</td>
            <td>${activity.action}</td>
            <td>${activity.ipAddress}</td>
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
    res.status(500).send('Помилка при завантаженні журналу активності');
  } finally {
    const endTime = Date.now();
    logger.info('Маршрут /admin/activity-log виконано', { duration: `${endTime - startTime} мс` });
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
