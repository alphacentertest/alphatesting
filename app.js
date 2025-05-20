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
const csurf = require('csurf');
const jwt = require('jsonwebtoken');
const winston = require('winston');

const app = express();

app.set('trust proxy', 1);

// Логирование с использованием winston
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

// Налаштування multer з лімітом розміру файлу (5MB)
const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 5 * 1024 * 1024 } // 5MB ліміт
});

// Настройка nodemailer для отправки email
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER || 'alphacentertest@gmail.com',
    pass: process.env.EMAIL_PASS || ':bnnz<fnmrsdobysxtcnmysrjve'
  }
});

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
    logger.error('Error sending suspicious activity email', { message: error.message, stack: error.stack });
  }
};

// Конфигурация параметров подозрительной активности
const config = {
  suspiciousActivity: {
    timeAwayThreshold: 50,
    switchCountThreshold: 5
  }
};

// Подключение к MongoDB
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/alpha?retryWrites=true&w=majority';
const client = new MongoClient(MONGODB_URI, {
  connectTimeoutMS: 5000,
  serverSelectionTimeoutMS: 5000
});
let db;

// Новий клас CacheManager із сортуванням за полем order
class CacheManager {
  static cache = {};

  static async getOrFetch(key, testNumber, fetchFn) {
    const cacheKey = `${key}:${testNumber}`;
    if (this.cache[cacheKey]) {
      logger.info(`Cache hit for ${cacheKey}`);
      return this.cache[cacheKey];
    }

    logger.info(`Cache miss for ${cacheKey}, fetching from DB`);
    const startTime = Date.now();
    const data = await fetchFn();
    this.cache[cacheKey] = data;
    logger.info(`Refreshed ${key} cache for test ${testNumber} with ${data.length} items in ${Date.now() - startTime} ms`);
    return data;
  }

  static async invalidateCache(key, testNumber) {
    const cacheKey = `${key}:${testNumber}`;
    delete this.cache[cacheKey];
    logger.info(`Invalidated cache for ${cacheKey}`);
  }

  static async getQuestions(testNumber) {
    return await this.getOrFetch('questions', testNumber, async () => {
      const questions = await db.collection('questions').find({ testNumber }).sort({ order: 1 }).toArray();
      return questions;
    });
  }

  static async getAllQuestions() {
    return await this.getOrFetch('allQuestions', 'all', async () => {
      const questions = await db.collection('questions').find({}).sort({ order: 1 }).toArray();
      return questions;
    });
  }
}

// Кэш пользователей и вопросов
let userCache = [];
const questionsCache = {};

const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    logger.info(`Attempting to connect to MongoDB (Attempt ${attempt} of ${maxAttempts}) with URI: ${MONGODB_URI}`);
    const startTime = Date.now();
    await client.connect();
    const endTime = Date.now();
    logger.info(`Connected to MongoDB successfully in ${endTime - startTime} ms`);
    db = client.db('alpha');
    logger.info('Database initialized', { databaseName: db.databaseName });
  } catch (error) {
    logger.error('Failed to connect to MongoDB', { message: error.message, stack: error.stack });
    if (attempt < maxAttempts) {
      logger.info('Retrying MongoDB connection in 5 seconds...');
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
    logger.info(`Loaded ${tests.length} tests from MongoDB`);
  } catch (error) {
    logger.error('Error loading tests from MongoDB', { message: error.message, stack: error.stack });
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
          isQuickTest: testData.isQuickTest || false, // Додаємо isQuickTest
          timePerQuestion: testData.timePerQuestion || null // Додаємо timePerQuestion
        }
      },
      { upsert: true }
    );
  } catch (error) {
    logger.error('Error saving test to MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

const deleteTestFromMongoDB = async (testNumber) => {
  try {
    await db.collection('tests').deleteOne({ testNumber });
    logger.info('Deleted test from MongoDB', { testNumber });
  } catch (error) {
    logger.error('Error deleting test from MongoDB', { testNumber, message: error.message, stack: error.stack });
    throw error;
  }
};

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Middleware для предотвращения кэширования
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Настройка CSRF с использованием csurf
app.use(csurf({ cookie: true }));

// Middleware для обработки ошибок MongoDB
app.use((err, req, res, next) => {
  if (err.name === 'MongoNetworkError' || err.name === 'MongoServerError') {
    logger.error('MongoDB error', { message: err.message, stack: err.stack });
    res.status(503).json({ success: false, message: 'Помилка з’єднання з базою даних. Спробуйте пізніше.' });
  } else {
    next(err);
  }
});

// Middleware для обработки ошибок CSRF
app.use((err, req, res, next) => {
  if (err.code === 'EBADCSRFTOKEN') {
    logger.error('CSRF token validation failed', {
      expectedToken: req.csrfToken(),
      receivedToken: req.body._csrf || req.headers['x-csrf-token'] || 'not provided'
    });
    res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
  } else {
    next(err);
  }
});

const importUsersToMongoDB = async (filePath) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
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
      const role = String(row.getCell(3).value || '').trim().toLowerCase(); // Отримуємо роль із третього стовпця
      if (username && password) {
        const hashedPassword = await bcrypt.hash(password, saltRounds);
        const userRole = role === 'admin' ? 'admin' : role === 'instructor' ? 'instructor' : 'user'; // Якщо роль не вказана, за замовчуванням user
        users.push({ username, password: hashedPassword, role: userRole });
      }
    }
    if (users.length === 0) {
      throw new Error('Не знайдено користувачів у файлі');
    }
    await db.collection('users').deleteMany({});
    logger.info('Cleared all users before import');
    await db.collection('users').insertMany(users);
    logger.info(`Imported ${users.length} users to MongoDB with hashed passwords`);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    return users.length;
  } catch (error) {
    logger.error('Error importing users to MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

const importQuestionsToMongoDB = async (filePath, testNumber) => {
  try {
    logger.info('Opening workbook', { filePath });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    logger.info('Workbook opened successfully');

    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      throw new Error('Лист "Questions" не знайдено у файлі');
    }
    logger.info('Worksheet "Questions" found', { rowCount: sheet.rowCount });

    const MAX_ROWS = 1000;
    if (sheet.rowCount > MAX_ROWS + 1) {
      throw new Error(`Занадто багато рядків (${sheet.rowCount - 1} питань). Максимальна кількість питань: ${MAX_ROWS}.`);
    }

    const questions = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        try {
          logger.info(`Processing row ${rowNumber}`);
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

          // Нормалізація назви зображення для originalPicture
          const normalizedPicture = picture ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '') : null;

          let questionData = {
            testNumber,
            picture: null,
            originalPicture: normalizedPicture, // Зберігаємо нормалізовану назву
            text: questionText,
            options,
            correctAnswers,
            type,
            points,
            variant,
            order: rowNumber - 1 // Поле order для збереження порядку
          };

          if (picture) {
            logger.info(`Processing picture field: ${picture}`, { testNumber, rowNumber });
            const pictureMatch = normalizedPicture.match(/^Picture(\d+)$/i);
            if (pictureMatch) {
              const pictureNumber = pictureMatch[1];
              const pictureBaseName = `/images/Picture${pictureNumber}`;
              questionData.picture = pictureBaseName;

              const extensions = ['.png', '.jpg', '.jpeg', '.gif'];
              let found = false;
              const imageDir = path.join(__dirname, 'public', 'images');
              const filesInDir = fs.existsSync(imageDir) ? fs.readdirSync(imageDir) : [];
              const targetFileNameBase = `Picture${pictureNumber}`;
              for (const ext of extensions) {
                const expectedFileName = targetFileNameBase + ext;
                const fileExists = filesInDir.some(file => file.toLowerCase() === expectedFileName.toLowerCase());
                if (fileExists) {
                  const matchedFile = filesInDir.find(file => file.toLowerCase() === expectedFileName.toLowerCase());
                  questionData.picture = `/images/${matchedFile}`;
                  logger.info(`Image found: ${questionData.picture}`, { testNumber, rowNumber });
                  found = true;
                  break;
                }
              }
              if (!found) {
                logger.warn(`Image not found for Picture ${questionData.picture}. Available files: ${filesInDir.join(', ')}`, { testNumber, rowNumber });
                questionData.picture = null;
              }
            } else {
              logger.warn(`Invalid picture format: ${picture}. Expected format: PictureX or PictureX.png`, { testNumber, rowNumber });
            }
          }

          if (type === 'matching') {
            questionData.pairs = options.map((opt, idx) => ({
              left: opt || '',
              right: correctAnswers[idx] || ''
            })).filter(pair => pair.left && pair.right);
            if (questionData.pairs.length === 0) throw new Error('Для типу Matching потрібні пари відповідей');
            questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
          }

          if (type === 'fillblank') {
            questionText = questionText.replace(/\s*___\s*/g, '___');
            const blankCount = (questionText.match(/___/g) || []).length;
            if (blankCount === 0 || blankCount !== correctAnswers.length) throw new Error('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
            questionData.text = questionText;
            questionData.blankCount = blankCount;
          }

          if (type === 'singlechoice') {
            if (correctAnswers.length !== 1 || options.length < 2) throw new Error('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
            questionData.correctAnswer = correctAnswers[0];
          }

          questions.push(questionData);
          logger.info(`Row ${rowNumber} processed successfully`);
        } catch (error) {
          throw new Error(`Помилка в рядку ${rowNumber}: ${error.message}`);
        }
      }
    });
    if (questions.length === 0) {
      throw new Error('Не знайдено питань у файлі');
    }
    logger.info('Deleting existing questions for test', { testNumber });
    await db.collection('questions').deleteMany({ testNumber });
    logger.info('Inserting new questions', { count: questions.length });
    await db.collection('questions').insertMany(questions);
    logger.info(`Imported ${questions.length} questions for test ${testNumber} to MongoDB`);
    await CacheManager.invalidateCache('questions', testNumber);
    return questions.length;
  } catch (error) {
    logger.error('Error importing questions to MongoDB', { message: error.message, stack: error.stack });
    throw error;
  }
};

// Функция для случайного перемешивания массива (Fisher-Yates)
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
    const endTime = Date.now();
    logger.info(`Refreshed user cache with ${userCache.length} users in ${endTime - startTime} ms`);
  } catch (error) {
    logger.error('Error refreshing user cache', { message: error.message, stack: error.stack });
  }
};

const loadQuestions = async (testNumber) => {
  try {
    const startTime = Date.now();
    if (questionsCache[testNumber]) {
      const endTime = Date.now();
      logger.info(`Loaded ${questionsCache[testNumber].length} questions for test ${testNumber} from cache in ${endTime - startTime} ms`);
      return questionsCache[testNumber];
    }

    const questions = await db.collection('questions').find({ testNumber: testNumber.toString() }).sort({ order: 1 }).toArray();
    const endTime = Date.now();
    if (questions.length === 0) {
      throw new Error(`No questions found in MongoDB for test ${testNumber}`);
    }
    questionsCache[testNumber] = questions;
    logger.info(`Loaded ${questions.length} questions for test ${testNumber} from MongoDB in ${endTime - startTime} ms`);
    return questions;
  } catch (error) {
    logger.error(`Ошибка в loadQuestions (test ${testNumber})`, { message: error.message, stack: error.stack });
    throw error;
  }
};

const ensureInitialized = (req, res, next) => {
  if (!isInitialized) {
    if (initializationError) {
      logger.error('Server not initialized due to error', { message: initializationError.message, stack: initializationError.stack });
      return res.status(500).json({ success: false, message: `Server initialization failed: ${initializationError.message}` });
    }
    logger.warn('Server is still initializing...');
    return res.status(503).json({ success: false, message: 'Server is initializing, please try again later' });
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
  const endTime = Date.now();
  logger.info('User passwords updated with hashes', { duration: `${endTime - startTime} ms` });
  await CacheManager.invalidateCache('users', null);
};

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
    logger.info('MongoDB indexes created successfully');

    // Міграція ролей для існуючих користувачів
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
      logger.info('Migrated roles to existing users');
    }

    // Міграція тестов, якщо колекція пуста
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
      logger.info('Migrated default tests to MongoDB', { count: Object.keys(defaultTests).length });
    }

    await updateUserPasswords();
    await loadUsersToCache();
    await loadTestsFromMongoDB();
    await CacheManager.invalidateCache('questions', null);
    isInitialized = true;
    initializationError = null;
  } catch (error) {
    logger.error('Failed to initialize server', { message: error.message, stack: error.stack });
    initializationError = error;
    throw error;
  }
};

// Очистка старых записей журнала активности (старше 30 дней)
const cleanupActivityLog = async () => {
  try {
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    const result = await db.collection('activity_log').deleteMany({
      timestamp: { $lt: thirtyDaysAgo.toISOString() }
    });
    logger.info('Cleaned up old activity logs', { deletedCount: result.deletedCount });
  } catch (error) {
    logger.error('Error cleaning up activity logs', { message: error.message, stack: error.stack });
  }
};

// Запуск задачи очистки раз в день
setInterval(cleanupActivityLog, 24 * 60 * 60 * 1000);

(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
    // Выполняем первую очистку при старте
    await cleanupActivityLog();
  } catch (error) {
    logger.error('Failed to start server due to initialization error', { message: error.message, stack: error.stack });
    process.exit(1);
  }
})();

// Хранилище для отслеживания попыток входа (в MongoDB)
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
    throw new Error('Перевищено ліміт спроб входу (30 на день). Спробуйте знову завтра.');
  }

  await db.collection('login_attempts').updateOne(
    { ipAddress, lastAttempt: { $gte: startOfDay, $lt: endOfDay } },
    { $inc: { count: 1 }, $set: { lastAttempt: now } },
    { upsert: true }
  );
};

// Функция для логирования активности с поддержкой транзакций
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
    logger.info(`Logged activity: ${user} - ${action} at ${adjustedTimestamp}, IP: ${ipAddress}`, { duration: `${endTime - startTime} ms` });
  } catch (error) {
    logger.error('Error logging activity', { message: error.message, stack: error.stack });
    throw error;
  }
};

app.get('/test-mongo', async (req, res) => {
  try {
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    await db.collection('users').findOne();
    res.json({ success: true, message: 'MongoDB connection successful' });
  } catch (error) {
    logger.error('MongoDB test failed', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'MongoDB connection failed', error: error.message });
  }
});

app.get('/api/test', (req, res) => {
  logger.info('Handling /api/test request');
  res.json({ success: true, message: 'Express server is working on /api/test' });
});

app.get('/', (req, res) => {
  logger.info('Serving index.html');
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
          <input type="hidden" name="_csrf" value="${req.csrfToken()}">
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

            // Показуємо лоадер
            loginButton.disabled = true;
            loginButton.textContent = 'Завантаження...';

            const formData = new URLSearchParams();
            formData.append('username', username);
            formData.append('password', password);
            formData.append('_csrf', document.querySelector('input[name="_csrf"]').value);

            try {
              const response = await fetch('/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: formData
              });

              const result = await response.json();

              if (result.success) {
                window.location.href = result.redirect + '?nocache=' + Date.now();
              } else {
                // Деталізована обробка помилок за статус-кодом
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
              console.error('Error during login:', error);
              errorMessage.textContent = 'Не вдалося підключитися до сервера. Перевірте ваше з’єднання з Інтернетом.';
            } finally {
              // Прибираємо лоадер
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
    logger.info('Received login data', { username });

    if (!username || !password) {
      logger.warn('Username or password not provided');
      return res.status(400).json({ success: false, message: 'Логін або пароль не вказано' });
    }

    const foundUser = userCache.find(user => user.username === username);
    logger.info('User found in cache', { username, cachedPassword: foundUser?.password });

    if (!foundUser) {
      logger.warn('User not found', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    const passwordMatch = await bcrypt.compare(password, foundUser.password);
    if (!passwordMatch) {
      logger.warn('Invalid password for user', { username });
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    await checkLoginAttempts(ipAddress, true);

    const token = jwt.sign(
      { username: foundUser.username, role: foundUser.role }, // Додаємо роль у JWT
      process.env.JWT_SECRET || 'your-secret-key',
      { expiresIn: '24h' }
    );

    res.cookie('token', token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production' ? true : false,
      sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
      maxAge: 24 * 60 * 60 * 1000
    });

    await logActivity(foundUser.username, 'увійшов на сайт', ipAddress);

    if (foundUser.role === 'admin') {
      res.json({ success: true, redirect: '/admin' });
    } else {
      res.json({ success: true, redirect: '/select-test' }); // Для instructor і user перенаправляємо на вибір тестів
    }
  } catch (error) {
    logger.error('Ошибка в /login', { message: error.message, stack: error.stack });
    res.status(error.message.includes('Перевищено ліміт') ? 429 : 500).json({ success: false, message: error.message || 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /login executed', { duration: `${endTime - startTime} ms` });
  }
});

// Middleware для проверки JWT
const checkAuth = (req, res, next) => {
  const token = req.headers['authorization']?.split(' ')[1] || req.cookies.token;
  if (!token) {
    return res.redirect('/');
  }

  try {
    const decoded = jwt.verify(token, process.env.JWT_SECRET || 'your-secret-key');
    req.user = decoded.username;
    req.userRole = decoded.role; // Використовуємо роль із JWT
    next();
  } catch (error) {
    logger.error('JWT verification failed', { message: error.message, stack: error.stack });
    res.redirect('/');
  }
};

const checkAdmin = (req, res, next) => {
  if (req.user !== 'admin') {
    return res.status(403).send('Доступно тільки для адміністратора (403 Forbidden)');
  }
  next();
};

app.get('/select-test', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') {
      return res.redirect('/admin');
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
                  <button onclick="window.location.href='/test?test=${num}'">${data.name}</button>
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
              const formData = new URLSearchParams();
              formData.append('_csrf', '${req.csrfToken()}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) {
                  throw new Error('HTTP error! status: ' + response.status);
                }
                window.location.href = '/';
              } catch (error) {
                console.error('Error during logout:', error);
                alert('Не вдалося вийти. Перевірте ваше з’єднання з Інтернетом.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Route /select-test executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/logout', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    logActivity(req.user, 'покинув сайт', ipAddress);
    res.clearCookie('token');
    res.json({ success: true });
  } finally {
    const endTime = Date.now();
    logger.info('Route /logout executed', { duration: `${endTime - startTime} ms` });
  }
});

const userTests = new Map();

const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress) => {
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
        variant: `Variant ${variant}`
      };
      logger.info('Saving result to MongoDB with answers', { answers: result.answers });
      if (!db) {
        throw new Error('MongoDB connection not established');
      }
      await db.collection('test_results').insertOne(result, { session });
      await logActivity(user, `завершив тест ${testNames[testNumber].name} з результатом ${Math.round(percentage)}%`, ipAddress, { percentage: Math.round(percentage) }, session);
    });
  } catch (error) {
    logger.error('Ошибка сохранения результата и лога активности', { message: error.message, stack: error.stack });
    throw error;
  } finally {
    await session.endSession();
    const endTimeLog = Date.now();
    logger.info('saveResult executed', { duration: `${endTimeLog - startTimeLog} ms` });
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

    logger.info(`User ${user} has ${attemptLimit - attempts} attempts left for test ${testNumber} today`);

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
    logger.error('Error checking test attempts', { message: error.message, stack: error.stack });
    throw error;
  }
};

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
    logger.info(`Assigned variant to user ${req.user} for test ${testNumber}: Variant ${userVariant}`);

    questions = questions.filter(q => !q.variant || q.variant === '' || q.variant === `Variant ${userVariant}`);
    logger.info(`Filtered questions for test ${testNumber}, variant ${userVariant}: ${questions.length} questions found`);

    if (questions.length === 0) {
      return res.status(400).send(`Немає питань для варіанту ${userVariant} у тесті ${testNumber}`);
    }

    const questionLimit = testNames[testNumber].questionLimit;
    if (questionLimit && questions.length > questionLimit) {
      questions = shuffleArray([...questions]).slice(0, questionLimit);
    }

    // Якщо тест швидкий, увімкнемо випадковий вибір питань і відповідей
    const isQuickTest = testNames[testNumber].isQuickTest;
    if (isQuickTest || testNames[testNumber].randomQuestions) {
      questions = shuffleArray([...questions]);
    }
    if (isQuickTest || testNames[testNumber].randomAnswers) {
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

    userTests.set(req.user, {
      testNumber,
      questions,
      answers: {},
      currentQuestion: 0,
      startTime: Date.now(),
      timeLimit: testNames[testNumber].timeLimit * 1000,
      variant: userVariant,
      isQuickTest: isQuickTest,
      timePerQuestion: testNames[testNumber].timePerQuestion // Зберігаємо час на питання
    });

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await logActivity(req.user, `розпочав тест ${testNames[testNumber].name}`, ipAddress);
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    logger.error('Ошибка в /test', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Route /test executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/test/question', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole === 'admin') return res.redirect('/admin');
    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).send('Тест не розпочато');
    }
    const { questions, testNumber, answers, currentQuestion, startTime, timeLimit, isQuickTest, timePerQuestion } = userTest;
    const index = parseInt(req.query.index) || 0;
    if (index < 0 || index >= questions.length) {
      return res.status(400).send('Невірний номер питання');
    }
    userTest.currentQuestion = index;
    userTest.answerTimestamps = userTest.answerTimestamps || {};
    userTest.answerTimestamps[index] = userTest.answerTimestamps[index] || Date.now();
    const q = questions[index];
    const progress = Array.from({ length: questions.length }, (_, i) => ({
      number: i + 1,
      answered: answers[i] && (Array.isArray(answers[i]) ? answers[i].length > 0 : answers[i].trim() !== '')
    }));

    // Розрахунок загального часу для тесту
    let totalTestTime = timeLimit / 1000; // Загальний час у секундах для звичайного тесту
    if (isQuickTest) {
      totalTestTime = questions.length * timePerQuestion; // Для Quick Test: кількість питань * час на питання
    }
    const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
    const remainingTime = Math.max(0, totalTestTime - elapsedTime);
    const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
    const seconds = (remainingTime % 60).toString().padStart(2, '0');

    const selectedOptions = answers[index] || [];
    const selectedOptionsString = JSON.stringify(selectedOptions).replace(/'/g, "\\'");

    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${testNames[testNumber].name}</title>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js"></script>
          <style>
            body { font-family: Arial, sans-serif; margin: 0; padding: 20px; padding-bottom: 80px; background-color: #f0f0f0; }
            h1 { font-size: 24px; text-align: center; }
            img { max-width: 100%; margin-bottom: 10px; display: block; margin-left: auto; margin-right: auto; }
            .progress-bar { display: flex; flex-direction: column; gap: 5px; margin-bottom: 20px; width: calc(100% - 40px); margin-left: auto; margin-right: auto; box-sizing: border-box; }
            .progress-circle { width: 40px; height: 40px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px; flex-shrink: 0; }
            .progress-circle.unanswered { background-color: red; color: white; }
            .progress-circle.answered { background-color: green; color: white; }
            .progress-line { width: 5px; height: 2px; background-color: #ccc; margin: 0 2px; align-self: center; flex-shrink: 0; }
            .progress-line.answered { background-color: green; }
            .progress-row { display: flex; align-items: center; justify-content: space-around; gap: 2px; flex-wrap: wrap; }
            .option-box { border: 2px solid #ccc; padding: 10px; margin: 5px 0; border-radius: 5px; cursor: pointer; font-size: 16px; user-select: none; }
            .option-box.selected { background-color: #90ee90; }
            .button-container { position: fixed; bottom: 20px; left: 20px; right: 20px; display: flex; justify-content: space-between; }
            button { padding: 10px 20px; margin: 5px; border: none; cursor: pointer; border-radius: 5px; font-size: 16px; }
            .back-btn { background-color: red; color: white; }
            .next-btn { background-color: blue; color: white; }
            .finish-btn { background-color: green; color: white; }
            .results-btn { background-color: #007bff; color: white; }
            button:disabled { background-color: grey; cursor: not-allowed; }
            #timer { font-size: 24px; margin-bottom: 20px; text-align: center; }
            #question-timer { position: relative; width: 80px; height: 80px; margin: 0 auto 10px auto; }
            #question-timer svg { width: 100%; height: 100%; transform: rotate(-90deg); }
            #question-timer circle { fill: none; stroke-width: 8; }
            #question-timer .timer-circle-bg { stroke: #e0e0e0; }
            #question-timer .timer-circle { stroke: #ff4d4d; stroke-dasharray: 251; stroke-dashoffset: 0; transition: stroke-dashoffset 1s linear; }
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
            @media (max-width: 600px) {
              h1 { font-size: 28px; }
              .progress-bar { flex-direction: column; }
              .progress-circle { width: 20px; height: 20px; font-size: 10px; }
              .progress-line { width: 5px; }
              .progress-row { justify-content: center; gap: 2px; flex-wrap: wrap; }
              .option-box, .matching-item { 
                font-size: 18px; 
                padding: 10px; 
                min-height: 50px; 
                line-height: 1.5; 
              }
              button { font-size: 18px; padding: 15px; }
              #timer { font-size: 20px; }
              .question-box h2 { font-size: 20px; }
              .matching-container { flex-direction: column; }
              .matching-column { width: 100%; }
              .blank-input { width: 80px; }
            }
            @media (min-width: 601px) {
              .progress-bar { flex-direction: row; justify-content: center; }
              .progress-circle { width: 40px; height: 40px; font-size: 14px; }
              .progress-line { width: 5px; }
              .progress-row { justify-content: space-around; }
            }
          </style>
        </head>
        <body>
          <h1>${testNames[testNumber].name}</h1>
          <div id="timer">Залишилось часу: ${minutes} мм ${seconds} с</div>
          <div class="progress-bar">
    `;
    if (progress.length <= 10) {
      html += `
        <div class="progress-row">
          ${progress.map((p, j) => `
            <div class="progress-circle ${p.answered ? 'answered' : 'unanswered'}">${p.number}</div>
            ${j < progress.length - 1 ? '<div class="progress-line ' + (p.answered ? 'answered' : '') + '"></div>' : ''}
          `).join('')}
        </div>
      `;
    } else {
      for (let i = 0; i < progress.length; i += 10) {
        const rowCircles = progress.slice(i, i + 10);
        html += `
          <div class="progress-row">
            ${rowCircles.map((p, j) => `
              <div class="progress-circle ${p.answered ? 'answered' : 'unanswered'}">${p.number}</div>
              ${j < rowCircles.length - 1 ? '<div class="progress-line ' + (p.answered ? 'answered' : '') + '"></div>' : ''}
            `).join('')}
          </div>
        `;
      }
    }
    html += `
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
      logger.info(`Fillblank question parts for index ${index}`, { parts: q.text.split('___') });
      const parts = q.text.split('___');
      let inputHtml = '';
      parts.forEach((part, i) => {
        inputHtml += `<span class="question-text">${part}</span>`;
        if (i < parts.length - 1) {
          const userAnswer = userAnswers[i] || '';
          inputHtml += `<input type="text" class="blank-input" id="blank_${i}" value="${userAnswer.replace(/"/g, '"')}" placeholder="Введіть відповідь" autocomplete="off">`;
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
              <button id="submit-answer" class="next-btn" ${index === questions.length - 1 ? 'disabled' : ''} onclick="saveAndNext(${index})">Далі</button>
            ` : ''}
            <button class="finish-btn" onclick="showConfirm(${index})">Завершити тест</button>
            ${req.userRole === 'instructor' ? `
              <button class="results-btn" onclick="window.location.href='/admin/results'">Переглянути результати</button>
            ` : ''}
          </div>
          <div id="confirm-modal">
            <h2>Ви дійсно бажаєте завершити тест?</h2>
            <button onclick="finishTest(${index})">Так</button>
            <button onclick="hideConfirm()">Ні</button>
          </div>
          <script>
            const startTime = ${startTime};
            const timeLimit = ${totalTestTime * 1000};
            const timerElement = document.getElementById('timer');
            const isQuickTest = ${isQuickTest};
            const timePerQuestion = ${timePerQuestion || 0};
            let timeAway = 0;
            let lastBlurTime = 0;
            let switchCount = 0;
            let lastActivityTime = Date.now();
            let activityCount = 0;
            let lastMouseMoveTime = 0;
            let lastKeydownTime = 0;
            const debounceDelay = 100;
            const questionStartTime = ${userTest.answerTimestamps[index] || Date.now()};
            let selectedOptions = ${selectedOptionsString};
            let matchingPairs = ${JSON.stringify(answers[index] || [])};
            let questionTimeRemaining = timePerQuestion;

            function updateTimer() {
              try {
                const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
                const remainingTime = Math.max(0, Math.floor(timeLimit / 1000) - elapsedTime);
                const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
                const seconds = (remainingTime % 60).toString().padStart(2, '0');
                timerElement.textContent = 'Залишилось часу: ' + minutes + ' мм ' + seconds + ' с';
                if (remainingTime <= 0) {
                  window.location.href = '/result';
                } else {
                  requestAnimationFrame(updateTimer);
                }
              } catch (error) {
                console.error('Error in updateTimer:', error);
              }
            }
            requestAnimationFrame(updateTimer);

            if (isQuickTest) {
              function updateQuestionTimer() {
                questionTimeRemaining = Math.max(0, timePerQuestion - Math.floor((Date.now() - questionStartTime) / 1000));
                const timerText = document.getElementById('timer-text');
                const timerCircle = document.querySelector('#question-timer .timer-circle');
                if (timerText && timerCircle) {
                  timerText.textContent = questionTimeRemaining;
                  const circumference = 251; // Довжина кола (2 * π * r, де r = 36)
                  const offset = (1 - questionTimeRemaining / timePerQuestion) * circumference; // Зменшуємо коло
                  timerCircle.style.strokeDashoffset = offset;
                }
                if (questionTimeRemaining <= 0) {
                  saveAndNext(${index});
                } else {
                  requestAnimationFrame(updateQuestionTimer);
                }
              }
              requestAnimationFrame(updateQuestionTimer);
            }

            window.addEventListener('blur', () => {
              lastBlurTime = Date.now();
              switchCount++;
            });

            window.addEventListener('focus', () => {
              if (lastBlurTime) {
                timeAway += Date.now() - lastBlurTime;
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

            async function saveAndNext(index) {
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
                const responseTime = Date.now() - questionStartTime;

                const formData = new URLSearchParams();
                formData.append('index', index);
                formData.append('answer', JSON.stringify(answers));
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${req.csrfToken()}');

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  throw new Error('HTTP error! status: ' + response.status);
                }

                const result = await response.json();
                if (result.success) {
                  if (isQuickTest && index < ${questions.length - 1}) {
                    window.location.href = '/test/question?index=' + (index + 1);
                  } else {
                    window.location.href = '/result';
                  }
                } else {
                  console.error('Error saving answer:', result.error);
                  alert('Помилка збереження відповіді');
                }
              } catch (error) {
                console.error('Error in saveAndNext:', error);
                alert('Не вдалося зберегти відповідь. Перевірте ваше з’єднання з Інтернетом.');
              }
            }

            function showConfirm(index) {
              document.getElementById('confirm-modal').style.display = 'block';
            }

            function hideConfirm() {
              document.getElementById('confirm-modal').style.display = 'none';
            }

            async function finishTest(index) {
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
                const responseTime = Date.now() - questionStartTime;

                const formData = new URLSearchParams();
                formData.append('index', index);
                formData.append('answer', JSON.stringify(answers));
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);
                formData.append('_csrf', '${req.csrfToken()}');

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });

                if (!response.ok) {
                  throw new Error('HTTP error! status: ' + response.status);
                }

                const result = await response.json();
                if (result.success) {
                  window.location.href = '/result';
                } else {
                  console.error('Error finishing test:', result.error);
                  alert('Помилка завершення тесту');
                }
              } catch (error) {
                console.error('Error in finishTest:', error);
                alert('Не вдалося завершити тест. Перевірте ваше з’єднання з Інтернетом.');
              }
            }

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
  } finally {
    const endTime = Date.now();
    logger.info('Route /test/question executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/answer', checkAuth, express.urlencoded({ extended: true }), async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;

    if (!index || !answer) {
      return res.status(400).json({ success: false, error: 'Необхідно надати index та answer' });
    }

    let parsedAnswer;
    try {
      if (typeof answer === 'string') {
        if (answer.trim() === '') {
          parsedAnswer = [];
        } else {
          parsedAnswer = JSON.parse(answer);
        }
      } else {
        parsedAnswer = answer;
      }
    } catch (error) {
      logger.error('Ошибка парсинга ответа в /answer', { message: error.message, stack: error.stack });
      return res.status(400).json({ success: false, error: 'Невірний формат відповіді' });
    }

    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).json({ success: false, error: 'Тест не розпочато' });
    }

    userTest.answers[index] = parsedAnswer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = (userTest.suspiciousActivity.timeAway || 0) + (parseInt(timeAway) || 0);
    userTest.suspiciousActivity.switchCount = (userTest.suspiciousActivity.switchCount || 0) + (parseInt(switchCount) || 0);
    userTest.suspiciousActivity.responseTimes[index] = parseInt(responseTime) || 0;
    userTest.suspiciousActivity.activityCounts[index] = parseInt(activityCount) || 0;

    res.json({ success: true });
  } catch (error) {
    logger.error('Ошибка в /answer', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /answer executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/result', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).json({ error: 'Тест не розпочато' });
    }
    const { questions, answers, testNumber, startTime, suspiciousActivity, variant } = userTest;
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
        logger.info(`Comparing input answer for question ${index + 1}`, {
          userAnswer: normalizedUserAnswer,
          correctAnswer: normalizedCorrectAnswer
        });
        const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
        if (isCorrect) {
          questionScore = q.points;
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
        logger.info(`Fillblank question ${index + 1}`, { userAnswers, correctAnswers });
        const isCorrect = userAnswers.length === correctAnswers.length &&
          userAnswers.every((answer, idx) => answer === correctAnswers[idx]);
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => normalizeAnswer(val));
        const correctAnswer = normalizeAnswer(q.correctAnswer);
        logger.info(`Single choice question ${index + 1}`, { userAnswers, correctAnswer });
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

    const duration = Math.round((endTime - startTime) / 1000);
    const timeAwayPercent = suspiciousActivity && suspiciousActivity.timeAway
      ? Math.round((suspiciousActivity.timeAway / (duration * 1000)) * 100)
      : 0;
    const switchCount = suspiciousActivity ? suspiciousActivity.switchCount || 0 : 0;
    const avgResponseTime = suspiciousActivity && suspiciousActivity.responseTimes
      ? (suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / suspiciousActivity.responseTimes.length / 1000).toFixed(2)
      : 0;
    const totalActivityCount = suspiciousActivity && suspiciousActivity.activityCounts
      ? suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0).toFixed(0)
      : 0;

    if (timeAwayPercent > config.suspiciousActivity.timeAwayThreshold || switchCount > config.suspiciousActivity.switchCountThreshold) {
      const activityDetails = {
        timeAwayPercent,
        switchCount,
        avgResponseTime,
        totalActivityCount
      };
      await sendSuspiciousActivityEmail(req.user, activityDetails);
    }

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    await saveResult(req.user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant, ipAddress);

    const endDateTime = new Date(endTime);
    const formattedTime = endDateTime.toLocaleTimeString('uk-UA', { hour12: false });
    const formattedDate = endDateTime.toLocaleDateString('uk-UA');
    const imagePath = path.join(__dirname, 'public', 'images', 'A.png');
    let imageBase64 = '';
    try {
      const imageBuffer = fs.readFileSync(imagePath);
      imageBase64 = imageBuffer.toString('base64');
    } catch (error) {
      logger.error('Error reading image A.png', { message: error.message, stack: error.stack });
    }

    const resultHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Результати ${testNames[testNumber].name}</title>
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
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
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
            const user = "${req.user}";
            const testName = "${testNames[testNumber].name}";
            const totalQuestions = ${totalQuestions};
            const correctClicks = ${correctClicks};
            const score = ${score};
            const totalPoints = ${totalPoints};
            const percentage = ${Math.round(percentage)};
            const time = "${formattedTime}";
            const date = "${formattedDate}";
            const imageBase64 = "${imageBase64}";

            document.getElementById('exportPDF').addEventListener('click', () => {
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
            });

            document.getElementById('restart').addEventListener('click', () => {
              window.location.href = '/select-test';
            });
          </script>
        </body>
      </html>
    `;
    res.send(resultHtml);
  } finally {
    const endTime = Date.now();
    logger.info('Route /result executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const userTest = userTests.get(req.user);
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
        if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
          const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
          if (correctAnswers.length === userAnswers.length &&
              correctAnswers.every(val => userAnswers.includes(val)) &&
              userAnswers.every(val => correctAnswers.includes(val))) {
            questionScore = q.points;
          }
        } else if (q.type === 'input' && userAnswer) {
          if (String(userAnswer).trim().toLowerCase() === String(q.correctAnswers[0]).trim().toLowerCase()) {
            questionScore = q.points;
          }
        } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
          const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
          if (userAnswers.join(',') === correctAnswers.join(',')) {
            questionScore = q.points;
          }
        } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
          const userPairs = userAnswer.map(pair => [String(pair[0]).trim().toLowerCase(), String(pair[1]).trim().toLowerCase()]);
          const correctPairs = q.correctPairs.map(pair => [String(pair[0]).trim().toLowerCase(), String(pair[1]).trim().toLowerCase()]);
          if (userPairs.length === correctPairs.length &&
              userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]))) {
            questionScore = q.points;
          }
        } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.'));
          const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.'));
          logger.info(`Fillblank question ${index + 1} in /results`, { userAnswers, correctAnswers });
          if (userAnswers.length === correctAnswers.length &&
              userAnswers.every((answer, idx) => answer === correctAnswers[idx])) {
            questionScore = q.points;
          }
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
          const correctAnswer = String(q.correctAnswer).trim().toLowerCase();
          logger.info(`Single choice question ${index + 1} in /results`, { userAnswers, correctAnswer });
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
        logger.error('Error reading image A.png', { message: error.message, stack: error.stack });
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
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
        <script>
          const user = "${req.user}";
          const testName = "${testNames[testNumber].name}";
          const totalQuestions = ${totalQuestions};
          const correctClicks = ${correctClicks};
          const score = ${score};
          const totalPoints = ${totalPoints};
          const percentage = ${Math.round(percentage)};
          const time = "${formattedTime}";
          const date = "${formattedDate}";
          const imageBase64 = "${imageBase64}";

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
      userTests.delete(req.user);
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
    logger.info('Route /results executed', { duration: `${endTime - startTime} ms` });
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
          <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button><br>
          <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button><br>
          <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
          <button id="logout" onclick="logout()">Вийти</button>
          <script>
            async function logout() {
              const formData = new URLSearchParams();
              formData.append('_csrf', '${req.csrfToken()}');
              try {
                const response = await fetch('/logout', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  body: formData
                });
                if (!response.ok) {
                  throw new Error('HTTP error! status: ' + response.status);
                }
                window.location.href = '/';
              } catch (error) {
                console.error('Error during logout:', error);
                alert('Не вдалося вийти. Перевірте ваше з’єднання з Інтернетом.');
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin executed', { duration: `${endTime - startTime} ms` });
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
      logger.error('Error fetching users from MongoDB', { message: error.message, stack: error.stack });
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
                  formData.append('_csrf', '${req.csrfToken()}');
                  const response = await fetch('/admin/delete-user', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP error! status: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні користувача: ' + result.message);
                  }
                } catch (error) {
                  console.error('Error deleting user:', error);
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
    logger.info('Route /admin/users executed', { duration: `${endTime - startTime} ms` });
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
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
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
    logger.info('Route /admin/add-user executed', { duration: `${endTime - startTime} ms` });
  }
});

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
    const newUser = { username, password: hashedPassword, role: username === 'Instructor' ? 'instructor' : 'user' }; // Додаємо роль
    await db.collection('users').insertOne(newUser);
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('User cache reloaded after adding new user');
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
    logger.error('Error adding user', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні користувача');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/add-user (POST) executed', { duration: `${endTime - startTime} ms` });
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
          <h1>Редагувати користувача: ${username}</h1>
          <form method="POST" action="/admin/edit-user" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
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
    logger.info('Route /admin/edit-user executed', { duration: `${endTime - startTime} ms` });
  }
});

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
      logger.warn('Validation errors in /admin/edit-user', { errors: errors.array() });
      return res.status(400).send(errors.array()[0].msg);
    }

    const { oldUsername, username, password } = req.body;
    logger.info('Received data for user update', { oldUsername, username, passwordProvided: !!password });

    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser && username !== oldUsername) {
      logger.warn('Username already exists', { username });
      return res.status(400).send('Користувач із таким ім’ям уже існує');
    }

    const updateData = { username };
    if (password) {
      const saltRounds = 10;
      const hashedPassword = await bcrypt.hash(password, saltRounds);
      updateData.password = hashedPassword;
      logger.info('Password updated for user', { username });
    } else {
      logger.info('Password not provided, skipping password update', { username });
    }

    // Зберігаємо або оновлюємо роль
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
    logger.info('Update result', { matchedCount: updateResult.matchedCount, modifiedCount: updateResult.modifiedCount });

    if (updateResult.matchedCount === 0) {
      logger.error('No user found to update', { oldUsername });
      return res.status(404).send('Користувача не знайдено');
    }

    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache();
    logger.info('User cache reloaded after update');

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
    logger.error('Error editing user', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні користувача');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/edit-user (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/delete-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username } = req.body;
    await db.collection('users').deleteOne({ username });
    await CacheManager.invalidateCache('users', null);
    await loadUsersToCache(); // Оновлюємо userCache після видалення
    logger.info('User cache reloaded after deletion');
    res.json({ success: true });
  } catch (error) {
    logger.error('Error deleting user', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні користувача' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/delete-user executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const sortBy = req.query.sortBy || 'order'; // За замовчуванням сортуємо за порядком
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
        // Отримуємо всі питання без сортування з бази
        questions = await db.collection('questions')
          .find({})
          .skip(skip)
          .limit(limit)
          .toArray();

        // Сортуємо питання за назвою тесту (на стороні сервера)
        questions.sort((a, b) => {
          const testNameA = testNames[a.testNumber]?.name || '';
          const testNameB = testNames[b.testNumber]?.name || '';
          return testNameA.localeCompare(testNameB, 'uk'); // Сортування з урахуванням української локалізації
        });
      } else {
        // Сортування за порядком (order), як було раніше
        questions = await db.collection('questions')
          .find({})
          .sort({ order: 1 })
          .skip(skip)
          .limit(limit)
          .toArray();
      }

      await CacheManager.invalidateCache('questions', null);
    } catch (error) {
      logger.error('Error fetching questions from MongoDB', { message: error.message, stack: error.stack });
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
            <td>${testNames[question.testNumber]?.name || 'Невідомий тест'}</td>
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
                  formData.append('_csrf', '${req.csrfToken()}');
                  const response = await fetch('/admin/delete-question', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP error! status: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні питання: ' + result.message);
                  }
                } catch (error) {
                  console.error('Error deleting question:', error);
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
    logger.info('Route /admin/questions executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/add-question', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
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
          </style>
        </head>
        <body>
          <h1>Додати питання</h1>
          <form method="POST" action="/admin/add-question" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}">${testNames[num].name}</option>`).join('')}
            </select>
            <label for="picture">Назва файлу зображення (опціонально, наприклад, Picture1.png):</label>
            <input type="text" id="picture" name="picture" placeholder="Picture1.png">
            <label for="text">Текст питання:</label>
            <textarea id="text" name="text" required></textarea>
            <label for="type">Тип питання:</label>
            <select id="type" name="type" required>
              <option value="multiple">Multiple Choice</option>
              <option value="singlechoice">Single Choice</option>
              <option value="truefalse">True/False</option>
              <option value="input">Input</option>
              <option value="ordering">Ordering</option>
              <option value="matching">Matching</option>
              <option value="fillblank">Fill in the Blank</option>
            </select>
            <label for="options">Варіанти відповідей (через крапку з комою):</label>
            <textarea id="options" name="options" placeholder="Введіть варіанти через крапку з комою"></textarea>
            <label for="correctAnswers">Правильні відповіді (через крапку з комою):</label>
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
            function validateForm() {
              const text = document.getElementById('text').value;
              const points = document.getElementById('points').value;
              const variant = document.getElementById('variant').value;
              const picture = document.getElementById('picture').value;
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
              return true;
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/add-question executed', { duration: `${endTime - startTime} ms` });
  }
});

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
      return res.status(400).send(errors.array()[0].msg);
    }

    const { testNumber, text, type, options, correctAnswers, points, variant, picture } = req.body;

    // Нормалізація назви зображення для originalPicture
    const normalizedPicture = picture ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '') : null;

    let questionData = {
      testNumber,
      picture: picture ? `/images/${normalizedPicture}` : null,
      originalPicture: normalizedPicture, // Зберігаємо нормалізовану назву
      text,
      type: type.toLowerCase(),
      options: options ? options.split(';').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean),
      points: Number(points),
      variant: variant || '',
      order: await db.collection('questions').countDocuments({ testNumber }) // Порядок на основі кількості існуючих питань
    };

    // Перевірка наявності зображення
    if (questionData.picture) {
      const imagePath = path.join(__dirname, 'public', questionData.picture);
      if (!fs.existsSync(imagePath)) {
        questionData.picture = null; // Якщо зображення не знайдено, встановлюємо picture в null
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
        return res.status(400).send('Для типу Matching потрібні пари відповідей');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        return res.status(400).send('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
      }
      questionData.blankCount = blankCount;
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        return res.status(400).send('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    await db.collection('questions').insertOne(questionData);
    await CacheManager.invalidateCache('questions', testNumber);
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання додано</title>
        </head>
        <body>
          <h1>Питання успішно додано</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Error adding question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при додаванні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/add-question (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/edit-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.query;
    const question = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    if (!question) {
      return res.status(404).send('Питання не знайдено');
    }
    const pictureName = question.picture ? question.picture.replace('/images/', '') : '';
    const normalizedOriginalPicture = question.originalPicture ? question.originalPicture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '') : '';
    const warningMessage = question.picture === null && question.originalPicture && question.originalPicture.trim() !== ''
      ? `Попередження: зображення "${normalizedOriginalPicture}" не було знайдено під час імпорту. Перевірте, чи файл зображення є в папці public/images.`
      : '';
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
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
            img#image-preview { max-width: 200px; margin-top: 10px; }
          </style>
        </head>
        <body>
          <h1>Редагувати питання</h1>
          <form method="POST" action="/admin/edit-question" onsubmit="return validateForm()">
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
            <input type="hidden" name="id" value="${id}">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}" ${num === question.testNumber ? 'selected' : ''}>${testNames[num].name}</option>`).join('')}
            </select>
            <label for="picture">Назва файлу зображення (опціонально, наприклад, Picture1.png):</label>
            <input type="text" id="picture" name="picture" value="${pictureName}" placeholder="Picture1.png">
            ${warningMessage ? `<p class="warning">${warningMessage}</p>` : ''}
            ${pictureName ? `<img id="image-preview" src="/images/${pictureName}" alt="Зображення питання" onerror="this.onerror=null;this.src='';this.alt='Зображення недоступне';">` : ''}
            <label for="text">Текст питання:</label>
            <textarea id="text" name="text" required>${question.text}</textarea>
            <label for="type">Тип питання:</label>
            <select id="type" name="type" required>
              <option value="multiple" ${question.type === 'multiple' ? 'selected' : ''}>Multiple Choice</option>
              <option value="singlechoice" ${question.type === 'singlechoice' ? 'selected' : ''}>Single Choice</option>
              <option value="truefalse" ${question.type === 'truefalse' ? 'selected' : ''}>True/False</option>
              <option value="input" ${question.type === 'input' ? 'selected' : ''}>Input</option>
              <option value="ordering" ${question.type === 'ordering' ? 'selected' : ''}>Ordering</option>
              <option value="matching" ${question.type === 'matching' ? 'selected' : ''}>Matching</option>
              <option value="fillblank" ${question.type === 'fillblank' ? 'selected' : ''}>Fill in the Blank</option>
            </select>
            <label for="options">Варіанти відповідей (через крапку з комою):</label>
            <textarea id="options" name="options">${question.options.join('; ')}</textarea>
            <label for="correctAnswers">Правильні відповіді (через крапку з комою):</label>
            <textarea id="correctAnswers" name="correctAnswers" required>${question.correctAnswers.join('; ')}</textarea>
            <label for="points">Бали за питання:</label>
            <input type="number" id="points" name="points" value="${question.points}" min="1" required>
            <label for="variant">Варіант:</label>
            <input type="text" id="variant" name="variant" value="${question.variant}">
            <button type="submit" class="submit-btn">Зберегти</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
          <script>
            function validateForm() {
              const text = document.getElementById('text').value;
              const points = document.getElementById('points').value;
              const variant = document.getElementById('variant').value;
              const picture = document.getElementById('picture').value;
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
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } catch (error) {
    logger.error('Error in /admin/edit-question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/edit-question executed', { duration: `${endTime - startTime} ms` });
  }
});

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

    // Нормалізація назви зображення для originalPicture
    const normalizedPicture = picture ? picture.replace(/\.png$/i, '').replace(/^picture/i, 'Picture').replace(/\s+/g, '') : null;

    let questionData = {
      testNumber,
      picture: picture ? `/images/${normalizedPicture}` : null,
      originalPicture: normalizedPicture, // Зберігаємо нормалізовану назву
      text,
      type: type.toLowerCase(),
      options: options ? options.split(';').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(';').map(ans => ans.trim()).filter(Boolean),
      points: Number(points),
      variant: variant || ''
    };

    // Перевірка наявності зображення
    if (questionData.picture) {
      const imagePath = path.join(__dirname, 'public', questionData.picture);
      if (!fs.existsSync(imagePath)) {
        questionData.picture = null; // Якщо зображення не знайдено, встановлюємо picture в null
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
        return res.status(400).send('Для типу Matching потрібні пари відповідей');
      }
      questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
    }

    if (type === 'fillblank') {
      questionData.text = questionData.text.replace(/\s*___\s*/g, '___');
      const blankCount = (questionData.text.match(/___/g) || []).length;
      if (blankCount === 0 || blankCount !== questionData.correctAnswers.length) {
        return res.status(400).send('Кількість пропусків у тексті питання не відповідає кількості правильних відповідей');
      }
      questionData.blankCount = blankCount;
    }

    if (type === 'singlechoice') {
      if (questionData.correctAnswers.length !== 1 || questionData.options.length < 2) {
        return res.status(400).send('Для типу Single Choice потрібна одна правильна відповідь і мінімум 2 варіанти');
      }
      questionData.correctAnswer = questionData.correctAnswers[0];
    }

    const oldQuestion = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    const oldTestNumber = oldQuestion.testNumber;

    // Зберігаємо існуючий order, щоб не змінювати порядок питання
    questionData.order = oldQuestion.order;

    await db.collection('questions').updateOne(
      { _id: new ObjectId(id) },
      { $set: questionData }
    );

    await CacheManager.invalidateCache('questions', oldTestNumber);
    if (testNumber !== oldTestNumber) {
      await CacheManager.invalidateCache('questions', testNumber);
    }

    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання оновлено</title>
        </head>
        <body>
          <h1>Питання успішно оновлено</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Error editing question', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при редагуванні питання');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/edit-question (POST) executed', { duration: `${endTime - startTime} ms` });
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
    const testNumber = question.testNumber;
    const deletedOrder = question.order;

    // Видаляємо питання
    await db.collection('questions').deleteOne({ _id: new ObjectId(id) });

    // Оновлюємо order для всіх наступних питань у цьому тесті
    await db.collection('questions').updateMany(
      { testNumber, order: { $gt: deletedOrder } },
      { $inc: { order: -1 } }
    );

    await CacheManager.invalidateCache('questions', testNumber);
    res.json({ success: true });
  } catch (error) {
    logger.error('Error deleting question', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні питання' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/delete-question executed', { duration: `${endTime - startTime} ms` });
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
          <form id="import-users-form">
            <input type="hidden" name="_csrf" id="_csrf" value="${req.csrfToken()}">
            <label for="file">Виберіть файл users.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn" id="submit-btn">Завантажити</button>
          </form>
          <div id="error-message" class="error"></div>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            document.getElementById('import-users-form').addEventListener('submit', async (e) => {
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

                const result = await response.text();
                if (response.ok) {
                  document.body.innerHTML = result;
                } else {
                  errorMessage.textContent = 'Помилка при завантаженні файлу: ' + result;
                }
              } catch (error) {
                console.error('Error during file upload:', error);
                errorMessage.textContent = 'Не вдалося підключитися до сервера. Перевірте ваше з’єднання з Інтернетом.';
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
    logger.info('Route /admin/import-users executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Received POST request for /admin/import-users', {
      body: req.body,
      headers: { 'x-csrf-token': req.headers['x-csrf-token'] },
      file: req.file ? { originalname: req.file.originalname, size: req.file.size } : 'no file'
    });

    if (!req.file) {
      logger.warn('Файл не завантажено: req.file відсутній');
      return res.status(400).send('Файл не завантажено');
    }

    const filePath = req.file.path;
    const importedCount = await importUsersToMongoDB(filePath);
    fs.unlinkSync(filePath);
    logger.info(`Успішно імпортовано ${importedCount} користувачів`);
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Користувачів імпортовано</title>
        </head>
        <body>
          <h1>Імпортовано ${importedCount} користувачів</h1>
          <button onclick="window.location.href='/admin/users'">Повернутися до списку користувачів</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Error importing users', { message: error.message, stack: error.stack });
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).send('Помилка при імпорті користувачів: ' + error.message);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/import-users (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/import-questions', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Імпорт питань</title>
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
          <h1>Імпорт питань із Excel</h1>
          <form id="import-form">
            <input type="hidden" name="_csrf" id="_csrf" value="${req.csrfToken()}">
            <label for="file">Виберіть файл questions*.xlsx:</label>
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
                const response = await fetch('/admin/import-questions', {
                  method: 'POST',
                  body: formData,
                  headers: {
                    'X-CSRF-Token': csrfToken
                  }
                });

                const result = await response.json();
                if (result.success) {
                  window.location.href = result.redirect || '/admin/questions';
                } else {
                  errorMessage.textContent = result.message || 'Помилка при імпорті.';
                }
              } catch (error) {
                console.error('Error during file upload:', error);
                errorMessage.textContent = 'Не вдалося підключитися до сервера. Перевірте ваше з’єднання з Інтернетом.';
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
    logger.info('Route /admin/import-questions executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    logger.info('Received POST request for /admin/import-questions', {
      body: req.body,
      headers: { 'x-csrf-token': req.headers['x-csrf-token'] },
      file: req.file ? { originalname: req.file.originalname, size: req.file.size } : 'no file'
    });

    if (!req.file) {
      logger.warn('Файл не завантажено: req.file отсутствует');
      return res.status(400).json({ success: false, message: 'Файл не завантажено' });
    }

    logger.info('File uploaded successfully', { path: req.file.path, size: req.file.size });

    const testNumber = req.file.originalname.match(/^questions(\d+)\.xlsx$/)?.[1];
    if (!testNumber) {
      logger.warn(`Неверное имя файла: ${req.file.originalname}. Ожидается формат questionsX.xlsx`);
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ success: false, message: 'Файл повинен мати назву у форматі questionsX.xlsx, де X — номер тесту' });
    }

    logger.info(`Загружаем файл ${req.file.originalname} для теста ${testNumber}`);

    const timeoutPromise = new Promise((_, reject) => {
      setTimeout(() => reject(new Error('Перевищено час обробки файлу (20 секунд). Спробуйте завантажити менший файл.')), 20000);
    });

    logger.info('Starting questions import from file');
    const importPromise = importQuestionsToMongoDB(req.file.path, testNumber);
    const importedCount = await Promise.race([importPromise, timeoutPromise]);
    logger.info('Questions import completed', { importedCount });

    fs.unlinkSync(req.file.path);
    logger.info(`Успешно импортировано ${importedCount} вопросов для теста ${testNumber}`);
    res.json({ success: true, message: `Імпортовано ${importedCount} питань для тесту ${testNumber}`, redirect: '/admin/questions' });
  } catch (error) {
    logger.error('Error importing questions', { message: error.message, stack: error.stack });
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ success: false, message: `Помилка при імпорті питань: ${error.message}` });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/import-questions (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/results', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.userRole !== 'admin' && req.userRole !== 'instructor') {
      return res.status(403).send('Доступ заборонено (403 Forbidden)');
    }

    const page = parseInt(req.query.page) || 1;
    const limit = 10;
    const skip = (page - 1) * limit;

    let results = [];
    let errorMessage = '';
    let totalResults = 0;
    let totalPages = 0;

    try {
      totalResults = await db.collection('test_results').countDocuments();
      totalPages = Math.ceil(totalResults / limit);
      results = await db.collection('test_results')
        .find({})
        .sort({ endTime: -1 })
        .skip(skip)
        .limit(limit)
        .toArray();
    } catch (fetchError) {
      logger.error('Ошибка при получении данных из MongoDB в /admin/results', { message: fetchError.message, stack: fetchError.stack });
      errorMessage = `Ошибка MongoDB: ${fetchError.message}`;
    }

    let adminHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Результати всіх користувачів</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            table { border-collapse: collapse; width: 100%; margin-top: 20px; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            .error { color: red; }
            .view-btn { background-color: #4CAF50; color: white; padding: 5px 10px; border: none; cursor: pointer; border-radius: 5px; }
            .view-btn:hover { background-color: #45a049; }
            .delete-btn { background-color: #ff4d4d; color: white; padding: 5px 10px; border: none; cursor: pointer; border-radius: 5px; }
            .delete-all-btn { background-color: #ff4d4d; color: white; padding: 10px 20px; margin: 10px 0; border: none; cursor: pointer; border-radius: 5px; }
            .nav-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; background-color: #007bff; color: white; border: none; border-radius: 5px; }
            .details { white-space: pre-wrap; max-width: 300px; line-height: 1.8; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
            #answers-modal {
              display: none;
              position: fixed;
              top: 0;
              left: 0;
              width: 100%;
              height: 100%;
              background: rgba(0,0,0,0.5);
              z-index: 1000;
              opacity: 0;
              transition: opacity 0.3s ease-in-out;
            }
            #answers-modal.visible {
              display: block;
              opacity: 1;
            }
            #answers-modal .modal-content {
              background: white;
              margin: 5% auto;
              padding: 20px;
              width: 50%;
              max-height: 70vh;
              overflow-y: auto;
              border-radius: 5px;
              position: relative;
              box-shadow: 0 4px 8px rgba(0,0,0,0.2);
              transform: scale(0.8);
              transition: transform 0.3s ease-in-out;
            }
            #answers-modal.visible .modal-content {
              transform: scale(1);
            }
            #answers-modal .modal-content h2 {
              margin: 0 0 20px 0;
              font-size: 18px;
              color: #333;
              border-bottom: 1px solid #e0e0e0;
              padding-bottom: 10px;
            }
            #answers-modal .close-btn {
              position: absolute;
              top: 10px;
              right: 10px;
              cursor: pointer;
              background-color: #ff4d4d;
              color: white;
              border: none;
              padding: 5px 10px;
              border-radius: 5px;
            }
            #modal-content {
              white-space: pre-wrap;
              line-height: 1.8;
              max-height: 60vh;
              overflow-y: auto;
              padding-right: 10px;
            }
            #modal-content::-webkit-scrollbar {
              width: 8px;
            }
            #modal-content::-webkit-scrollbar-track {
              background: #f1f1f1;
              border-radius: 4px;
            }
            #modal-content::-webkit-scrollbar-thumb {
              background: #888;
              border-radius: 4px;
            }
            #modal-content::-webkit-scrollbar-thumb:hover {
              background: #555;
            }
          </style>
        </head>
        <body>
          <h1>Результати всіх користувачів</h1>
          ${req.userRole === 'admin' ? `
            <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          ` : `
            <button class="nav-btn" onclick="window.location.href='/select-test'">Повернутися до вибору тестів</button>
          `}
    `;
    if (errorMessage) {
      adminHtml += `<p class="error">${errorMessage}</p>`;
    }
    adminHtml += `
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
              <th>Відповіді та бали</th>
              ${req.userRole === 'admin' ? '<th>Дія</th>' : ''}
            </tr>
    `;
    if (!results || results.length === 0) {
      adminHtml += '<tr><td colspan="' + (req.userRole === 'admin' ? '12' : '11') + '">Немає результатів</td></tr>';
    } else {
      results.forEach((r, index) => {
        const answersArray = [];
        if (r.answers) {
          Object.keys(r.answers).sort((a, b) => parseInt(a) - parseInt(b)).forEach(key => {
            const idx = parseInt(key);
            answersArray[idx] = r.answers[key];
          });
        }
        logger.info(`User ${r.user} answers array`, { answersArray });
        const answersDisplay = answersArray.length > 0
          ? answersArray.map((a, i) => {
              if (a === undefined || a === null) return null;
              let userAnswer;
              if (Array.isArray(a) && a.length > 0 && Array.isArray(a[0])) {
                userAnswer = a.map(pair => {
                  if (Array.isArray(pair) && pair.length === 2) {
                    return `${pair[0]} -> ${pair[1]}`;
                  }
                  return 'Невірний формат пари';
                }).join(', ');
              } else if (Array.isArray(a)) {
                userAnswer = a.join(', ');
              } else {
                userAnswer = a;
              }
              const questionScore = r.scoresPerQuestion[i] || 0;
              return `Питання ${i + 1}: ${userAnswer ? userAnswer.replace(/\\'/g, "'") : 'Немає відповіді'} (${questionScore} балів)`;
            }).filter(line => line !== null).join('\n')
          : 'Немає відповідей';
        const formatDateTime = (isoString) => {
          if (!isoString) return 'N/A';
          const date = new Date(isoString);
          return `${date.toLocaleTimeString('uk-UA', { hour12: false })} ${date.toLocaleDateString('uk-UA')}`;
        };
        const suspiciousActivityPercent = r.suspiciousActivity && r.suspiciousActivity.suspiciousScore
          ? Math.round(r.suspiciousActivity.suspiciousScore)
          : 0;
        const timeAwayPercent = r.suspiciousActivity && r.suspiciousActivity.timeAway
          ? Math.round((r.suspiciousActivity.timeAway / (r.duration * 1000)) * 100)
          : 0;
        const switchCount = r.suspiciousActivity ? r.suspiciousActivity.switchCount || 0 : 0;
        const avgResponseTime = r.suspiciousActivity && r.suspiciousActivity.responseTimes
          ? (r.suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / r.suspiciousActivity.responseTimes.length / 1000).toFixed(2)
          : 0;
        const totalActivityCount = r.suspiciousActivity && r.suspiciousActivity.activityCounts
          ? r.suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0).toFixed(0)
          : 0;
        const activityDetails = `
Время вне вкладки: ${timeAwayPercent}%
Переключения вкладок: ${switchCount}
Среднее время ответа (сек): ${avgResponseTime}
Общее количество действий: ${totalActivityCount}
        `;
        const percentage = r.totalPoints > 0 ? Math.round((r.score / r.totalPoints) * 100) : 0;
        const durationMinutes = Math.floor(r.duration / 60);
        const durationSeconds = r.duration % 60;
        const formattedDuration = `${durationMinutes} хв ${durationSeconds} сек`;
        adminHtml += `
          <tr>
            <td>${r.user || 'N/A'}</td>
            <td>${testNames[r.testNumber]?.name || 'N/A'}</td>
            <td>${r.variant || 'N/A'}</td>
            <td>${r.score || '0'} / ${percentage}%</td>
            <td>${r.totalPoints || '0'}</td>
            <td>${formatDateTime(r.startTime)}</td>
            <td>${formatDateTime(r.endTime)}</td>
            <td>${formattedDuration}</td>
            <td>${suspiciousActivityPercent}%</td>
            <td class="details">${activityDetails}</td>
            <td>
              <button class="view-btn" onclick="showAnswersModal('answers-${index}', '${r.user || 'N/A'}', '${testNames[r.testNumber]?.name || 'N/A'}')">Перегляд</button>
              <input type="hidden" id="answers-${index}" value="${answersDisplay.replace(/"/g, '"').replace(/\n/g, '<br>')}">
            </td>
            ${req.userRole === 'admin' ? `
              <td><button class="delete-btn" onclick="deleteResult('${r._id}')">🗑️ Видалити</button></td>
            ` : ''}
          </tr>
        `;
      });
    }
    adminHtml += `
          </table>
          ${req.userRole === 'admin' ? `
            <button class="delete-all-btn" onclick="deleteAllResults()">Видалити всі результати</button>
          ` : ''}
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/results?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/results?page=${page + 1}">Наступна</a>` : ''}
          </div>
          <div id="answers-modal">
            <div class="modal-content">
              <h2 id="modal-title"></h2>
              <button class="close-btn" onclick="closeAnswersModal()">Закрити</button>
              <div id="modal-content"></div>
            </div>
          </div>
          <script>
            ${req.userRole === 'admin' ? `
              async function deleteResult(id) {
                if (confirm('Ви впевнені, що хочете видалити цей результат?')) {
                  try {
                    const formData = new URLSearchParams();
                    formData.append('id', id);
                    formData.append('_csrf', '${req.csrfToken()}');
                    const response = await fetch('/admin/delete-result', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                      body: formData
                    });
                    if (!response.ok) {
                      throw new Error('HTTP error! status: ' + response.status);
                    }
                    const result = await response.json();
                    if (result.success) {
                      window.location.reload();
                    } else {
                      alert('Помилка при видаленні результату: ' + result.message);
                    }
                  } catch (error) {
                    console.error('Error deleting result:', error);
                    alert('Не вдалося видалити результат. Перевірте ваше з’єднання з Інтернетом.');
                  }
                }
              }

              async function deleteAllResults() {
                if (confirm('Ви впевнені, що хочете видалити всі результати? Цю дію не можна скасувати!')) {
                  try {
                    const formData = new URLSearchParams();
                    formData.append('_csrf', '${req.csrfToken()}');
                    const response = await fetch('/admin/delete-all-results', {
                      method: 'POST',
                      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                      body: formData
                    });
                    if (!response.ok) {
                      throw new Error('HTTP error! status: ' + response.status);
                    }
                    const result = await response.json();
                    if (result.success) {
                      window.location.reload();
                    } else {
                      alert('Помилка при видаленні всіх результатів: ' + result.message);
                    }
                  } catch (error) {
                    console.error('Error deleting all results:', error);
                    alert('Не вдалося видалити всі результати. Перевірте ваше з’єднання з Інтернетом.');
                  }
                }
              }
            ` : ''}

            function showAnswersModal(id, user, testName) {
              const answers = document.getElementById(id).value;
              const modal = document.getElementById('answers-modal');
              document.getElementById('modal-title').textContent = 'Відповіді користувача ' + user + ' (Тест: ' + testName + ')';
              document.getElementById('modal-content').innerHTML = answers;
              modal.classList.add('visible');
            }

            function closeAnswersModal() {
              const modal = document.getElementById('answers-modal');
              modal.classList.remove('visible');
            }

            document.getElementById('answers-modal').addEventListener('click', (e) => {
              if (e.target === document.getElementById('answers-modal')) {
                closeAnswersModal();
              }
            });
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml.trim());
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/results executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/delete-all-results', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    const deleteResult = await db.collection('test_results').deleteMany({});
    logger.info(`Deleted ${deleteResult.deletedCount} results from test_results collection`);
    res.json({ success: true, message: `Успішно видалено ${deleteResult.deletedCount} результатів` });
  } catch (error) {
    logger.error('Ошибка при удалении всех результатов', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні всіх результатів' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/delete-all-results executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.body;
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    res.json({ success: true });
  } catch (error) {
    logger.error('Ошибка при удалении результата', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/delete-result executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/edit-tests', checkAuth, checkAdmin, (req, res) => {
  const startTime = Date.now();
  try {
    const html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Редагувати назви тестів</title>
          <style>
            body { font-size: 24px; margin: 20px; }
            input[type="text"], input[type="number"] { font-size: 24px; padding: 5px; margin: 5px; }
            input[type="checkbox"] { width: 20px; height: 20px; margin: 5px; }
            button { font-size: 24px; padding: 10px 20px; margin: 5px; }
            .delete-btn { background-color: #ff4d4d; color: white; }
            .test-row { display: flex; align-items: center; margin-bottom: 10px; flex-wrap: wrap; }
            label { margin-right: 10px; }
          </style>
        </head>
        <body>
          <h1>Редагувати назви та налаштування тестів</h1>
          <form method="POST" action="/admin/edit-tests">
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
            ${Object.entries(testNames).map(([num, data]) => `
              <div class="test-row">
                <label for="test${num}">Назва Тесту ${num}:</label>
                <input type="text" id="test${num}" name="test${num}" value="${data.name}" required>
                <label for="time${num}">Час (сек):</label>
                <input type="number" id="time${num}" name="time${num}" value="${data.timeLimit}" required min="1">
                <label for="quickTest${num}">Quick Test:</label>
                <input type="checkbox" id="quickTest${num}" name="quickTest${num}" ${data.isQuickTest ? 'checked' : ''}>
                <label for="timePerQuestion${num}">Час на питання (сек):</label>
                <input type="number" id="timePerQuestion${num}" name="timePerQuestion${num}" value="${data.timePerQuestion || ''}" placeholder="10" min="1">
                <label for="randomQuestions${num}">Випадковий вибір питань:</label>
                <input type="checkbox" id="randomQuestions${num}" name="randomQuestions${num}" ${data.randomQuestions ? 'checked' : ''}>
                <label for="randomAnswers${num}">Випадковий вибір відповідей:</label>
                <input type="checkbox" id="randomAnswers${num}" name="randomAnswers${num}" ${data.randomAnswers ? 'checked' : ''}>
                <label for="questionLimit${num}">Кількість питань:</label>
                <input type="number" id="questionLimit${num}" name="questionLimit${num}" value="${data.questionLimit || ''}" min="1" placeholder="Без обмеження">
                <label for="attemptLimit${num}">Ліміт спроб на день:</label>
                <input type="number" id="attemptLimit${num}" name="attemptLimit${num}" value="${data.attemptLimit || 1}" min="1" required>
                <button type="button" class="delete-btn" onclick="deleteTest('${num}')">Видалити</button>
              </div>
            `).join('')}
            <button type="submit">Зберегти</button>
          </form>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            async function deleteTest(testNumber) {
              if (confirm('Ви впевнені, що хочете видалити Тест ' + testNumber + '?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('testNumber', testNumber);
                  formData.append('_csrf', '${req.csrfToken()}');
                  const response = await fetch('/admin/delete-test', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP error! status: ' + response.status);
                  }
                  window.location.reload();
                } catch (error) {
                  console.error('Error deleting test:', error);
                  alert('Не вдалося видалити тест. Перевірте ваше з’єднання з Інтернетом.');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/edit-tests executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const validationErrors = [];
    Object.keys(testNames).forEach(num => {
      const testName = req.body[`test${num}`];
      const timeLimit = req.body[`time${num}`];
      const questionLimit = req.body[`questionLimit${num}`];
      const attemptLimit = req.body[`attemptLimit${num}`];
      const timePerQuestion = req.body[`timePerQuestion${num}`];

      if (!testName || testName.length < 1 || testName.length > 100) {
        validationErrors.push(`Назва тесту ${num} має бути від 1 до 100 символів`);
      }
      if (!timeLimit || isNaN(parseInt(timeLimit)) || parseInt(timeLimit) < 1) {
        validationErrors.push(`Час для тесту ${num} має бути числом більше 0`);
      }
      if (questionLimit && (isNaN(parseInt(questionLimit)) || parseInt(questionLimit) < 1)) {
        validationErrors.push(`Кількість питань для тесту ${num} має бути числом більше 0`);
      }
      if (!attemptLimit || isNaN(parseInt(attemptLimit)) || parseInt(attemptLimit) < 1) {
        validationErrors.push(`Ліміт спроб для тесту ${num} має бути числом більше 0`);
      }
      if (timePerQuestion && (isNaN(parseInt(timePerQuestion)) || parseInt(timePerQuestion) < 1)) {
        validationErrors.push(`Час на питання для тесту ${num} має бути числом більше 0`);
      }
    });

    if (validationErrors.length > 0) {
      return res.status(400).send(validationErrors.join('<br>'));
    }

    for (const num of Object.keys(testNames)) {
      const testName = req.body[`test${num}`];
      const timeLimit = req.body[`time${num}`];
      const randomQuestions = req.body[`randomQuestions${num}`] === 'on';
      const randomAnswers = req.body[`randomAnswers${num}`] === 'on';
      const questionLimit = req.body[`questionLimit${num}`] ? parseInt(req.body[`questionLimit${num}`]) : null;
      const attemptLimit = parseInt(req.body[`attemptLimit${num}`]);
      const isQuickTest = req.body[`quickTest${num}`] === 'on';
      const timePerQuestion = req.body[`timePerQuestion${num}`] ? parseInt(req.body[`timePerQuestion${num}`]) : null;

      if (testName && timeLimit && attemptLimit) {
        testNames[num] = {
          name: testName,
          timeLimit: parseInt(timeLimit) || testNames[num].timeLimit,
          randomQuestions,
          randomAnswers,
          questionLimit,
          attemptLimit,
          isQuickTest, // Зберігаємо Quick Test
          timePerQuestion // Зберігаємо час на питання
        };
        await saveTestToMongoDB(num, testNames[num]);
      }
    }
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Назви оновлено</title>
        </head>
        <body>
          <h1>Назви та налаштування тестів успішно оновлено</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Ошибка при редактировании названий тестов', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при оновленні назв тестів');
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/edit-tests (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  const session = client.startSession();
  try {
    const { testNumber } = req.body;
    if (!testNames[testNumber]) {
      return res.status(404).json({ success: false, message: 'Тест не знайдено' });
    }
    await session.withTransaction(async () => {
      delete testNames[testNumber];
      await deleteTestFromMongoDB(testNumber);
      await db.collection('questions').deleteMany({ testNumber }, { session });
      if (questionsCache[testNumber]) {
        delete questionsCache[testNumber];
      }
    });
    res.json({ success: true });
  } catch (error) {
    logger.error('Ошибка при удалении теста', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  } finally {
    await session.endSession();
    const endTime = Date.now();
    logger.info('Route /admin/delete-test executed', { duration: `${endTime - startTime} ms` });
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
          <title>Створити новий тест</title>
          <style>
            body { font-size: 24px; margin: 20px; }
            label { display: block; margin: 10px 0 5px; }
            input { font-size: 24px; padding: 5px; margin: 5px; }
            select { font-size: 24px; padding: 5px; margin: 5px; }
            button { font-size: 24px; padding: 10px 20px; margin: 5px; }
          </style>
        </head>
        <body>
          <h1>Створити новий тест</h1>
          <form method="POST" action="/admin/create-test">
            <input type="hidden" name="_csrf" value="${req.csrfToken()}">
            <div>
              <label for="testName">Назва нового тесту:</label>
              <input type="text" id="testName" name="testName" required>
            </div>
            <div>
              <label for="timeLimit">Час (сек):</label>
              <input type="number" id="timeLimit" name="timeLimit" value="3600" required min="1">
            </div>
            <button type="submit">Створити</button>
          </form>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/create-test executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/create-test', checkAuth, checkAdmin, [
  body('testName')
    .isLength({ min: 1, max: 100 }).withMessage('Назва тесту має бути від 1 до 100 символів'),
  body('timeLimit')
    .isInt({ min: 1 }).withMessage('Час має бути числом більше 0')
], async (req, res) => {
  const startTime = Date.now();
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).send(errors.array()[0].msg);
    }

    const { testName, timeLimit } = req.body;
    const testNumber = String(Object.keys(testNames).length + 1);
    if (testNames[testNumber]) {
      return res.status(400).send('Тест з таким номером вже існує');
    }
    testNames[testNumber] = {
      name: testName,
      timeLimit: parseInt(timeLimit) || 3600,
      randomQuestions: false,
      randomAnswers: false,
      questionLimit: null,
      attemptLimit: 1
    };
    await saveTestToMongoDB(testNumber, testNames[testNumber]);
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
            a { display: inline-block; padding: 10px 20px; margin: 5px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            a:hover { background-color: #0056b3; }
          </style>
        </head>
        <body>
          <h1>Новий тест "${testName}" створено (Тест ${testNumber})</h1>
          <p>Скачайте шаблон для додавання питань:</p>
          <a href="/download-template?testNumber=${testNumber}" download="questions${testNumber}.xlsx">Скачати questions${testNumber}.xlsx</a>
          <p>Після заповнення шаблону завантажте його:</p>
          <button onclick="window.location.href='/admin/import-questions'">Завантажити питання</button>
          <br>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    logger.error('Ошибка при создании нового теста', { message: error.message, stack: error.stack });
    res.status(500).send(`Помилка при створенні тесту: ${error.message}`);
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/create-test (POST) executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/download-template', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const testNumber = req.query.testNumber;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).send('Номер тесту не вказано або тест не існує');
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Questions');

    sheet.columns = [
      { header: 'Picture', key: 'picture', width: 15 },
      { header: 'Question Text', key: 'text', width: 30 },
      { header: 'Option 1', key: 'option1', width: 15 },
      { header: 'Option 2', key: 'option2', width: 15 },
      { header: 'Option 3', key: 'option3', width: 15 },
      { header: 'Option 4', key: 'option4', width: 15 },
      { header: 'Option 5', key: 'option5', width: 15 },
      { header: 'Option 6', key: 'option6', width: 15 },
      { header: 'Option 7', key: 'option7', width: 15 },
      { header: 'Option 8', key: 'option8', width: 15 },
      { header: 'Option 9', key: 'option9', width: 15 },
      { header: 'Option 10', key: 'option10', width: 15 },
      { header: 'Option 11', key: 'option11', width: 15 },
      { header: 'Option 12', key: 'option12', width: 15 },
      { header: 'Correct Answer 1', key: 'correct1', width: 15 },
      { header: 'Correct Answer 2', key: 'correct2', width: 15 },
      { header: 'Correct Answer 3', key: 'correct3', width: 15 },
      { header: 'Correct Answer 4', key: 'correct4', width: 15 },
      { header: 'Correct Answer 5', key: 'correct5', width: 15 },
      { header: 'Correct Answer 6', key: 'correct6', width: 15 },
      { header: 'Correct Answer 7', key: 'correct7', width: 15 },
      { header: 'Correct Answer 8', key: 'correct8', width: 15 },
      { header: 'Correct Answer 9', key: 'correct9', width: 15 },
      { header: 'Correct Answer 10', key: 'correct10', width: 15 },
      { header: 'Correct Answer 11', key: 'correct11', width: 15 },
      { header: 'Correct Answer 12', key: 'correct12', width: 15 },
      { header: 'Type', key: 'type', width: 15 },
      { header: 'Points', key: 'points', width: 10 },
      { header: 'Variant', key: 'variant', width: 15 }
    ];

    sheet.addRow({
      picture: 'Picture1 (наприклад, Picture1, Picture2 тощо)',
      text: 'Приклад питання',
      option1: 'Варіант 1',
      option2: 'Варіант 2',
      option3: 'Варіант 3',
      option4: 'Варіант 4',
      correct1: 'Варіант 1; Варіант 2',
      type: 'multiple',
      points: 1,
      variant: 'Variant 1'
    });

    res.setHeader('Content-Disposition', `attachment; filename=questions${testNumber}.xlsx`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    logger.error('Error generating template', { message: error.message, stack: error.stack });
    res.status(500).send('Помилка при генерації шаблону');
  } finally {
    const endTime = Date.now();
    logger.info('Route /download-template executed', { duration: `${endTime - startTime} ms` });
  }
});

app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const page = parseInt(req.query.page) || 1;
    const limit = 10;
    const skip = (page - 1) * limit;

    let activities = [];
    let errorMessage = '';
    let totalActivities = 0;
    let totalPages = 0;

    try {
      totalActivities = await db.collection('activity_log').countDocuments();
      totalPages = Math.ceil(totalActivities / limit);
      activities = await db.collection('activity_log')
        .find({})
        .sort({ timestamp: -1 })
        .skip(skip)
        .limit(limit)
        .toArray();
    } catch (fetchError) {
      logger.error('Ошибка при получении данных из MongoDB в /admin/activity-log', { message: fetchError.message, stack: fetchError.stack });
      errorMessage = `Ошибка MongoDB: ${fetchError.message}`;
    }
    let adminHtml = `
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
            .nav-btn, .clear-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; border: none; border-radius: 5px; }
            .clear-btn { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
            .pagination { margin-top: 20px; }
            .pagination a { margin: 0 5px; padding: 5px 10px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px; }
            .pagination a:hover { background-color: #0056b3; }
          </style>
        </head>
        <body>
          <h1>Журнал дій</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
    `;
    if (errorMessage) {
      adminHtml += `<p class="error">${errorMessage}</p>`;
    }
    adminHtml += `
          <table>
            <tr>
              <th>Дата-Час</th>
              <th>Користувач</th>
              <th>IP-адреса</th>
              <th>Дія</th>
            </tr>
    `;
    if (!activities || activities.length === 0) {
      adminHtml += '<tr><td colspan="4">Немає записів</td></tr>';
    } else {
      activities.forEach(activity => {
        const timestamp = new Date(activity.timestamp);
        const formattedDateTime = `${timestamp.toLocaleDateString('uk-UA')} ${timestamp.toLocaleTimeString('uk-UA', { hour12: false })}`;
        const actionWithInfo = activity.additionalInfo && activity.additionalInfo.percentage
          ? `${activity.action} (${activity.additionalInfo.percentage}%)`
          : activity.action;
        adminHtml += `
          <tr>
            <td>${formattedDateTime}</td>
            <td>${activity.user || 'N/A'}</td>
            <td>${activity.ipAddress || 'N/A'}</td>
            <td>${actionWithInfo}</td>
          </tr>
        `;
      });
    }
    adminHtml += `
          </table>
          <button class="clear-btn" onclick="clearActivityLog()">Видалити усі записи журналу</button>
          <div class="pagination">
            ${page > 1 ? `<a href="/admin/activity-log?page=${page - 1}">Попередня</a>` : ''}
            <span>Сторінка ${page} з ${totalPages}</span>
            ${page < totalPages ? `<a href="/admin/activity-log?page=${page + 1}">Наступна</a>` : ''}
          </div>
          <script>
            async function clearActivityLog() {
              if (confirm('Ви впевнені, що хочете видалити усі записи журналу дій?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('_csrf', '${req.csrfToken()}');
                  const response = await fetch('/admin/delete-activity-log', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: formData
                  });
                  if (!response.ok) {
                    throw new Error('HTTP error! status: ' + response.status);
                  }
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні записів журналу: ' + result.message);
                  }
                } catch (error) {
                  console.error('Error clearing activity log:', error);
                  alert('Не вдалося видалити записи журналу. Перевірте ваше з’єднання з Інтернетом.');
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
    logger.info('Route /admin/activity-log executed', { duration: `${endTime - startTime} ms` });
  }
});

app.post('/admin/delete-activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    await db.collection('activity_log').deleteMany({});
    res.json({ success: true });
  } catch (error) {
    logger.error('Ошибка при удалении записей журнала действий', { message: error.message, stack: error.stack });
    res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
  } finally {
    const endTime = Date.now();
    logger.info('Route /admin/delete-activity-log executed', { duration: `${endTime - startTime} ms` });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  logger.info('Server is running', { port: PORT });
});

module.exports = app;
