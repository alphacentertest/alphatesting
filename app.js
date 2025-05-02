require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { MongoClient, ObjectId } = require('mongodb');
const fs = require('fs');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const crypto = require('crypto');

const app = express();

// Налаштування для хешування паролів
const saltRounds = 10;

// Налаштування підключення до MongoDB
const MONGO_URL = process.env.MONGO_URL || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/testdb?retryWrites=true&w=majority';
const client = new MongoClient(MONGO_URL, { connectTimeoutMS: 5000, serverSelectionTimeoutMS: 5000 });
let db;

// Підключення до MongoDB із повторними спробами
const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    console.log(`Attempting to connect to MongoDB (Attempt ${attempt} of ${maxAttempts}) with URL:`, MONGO_URL);
    await client.connect();
    console.log('Connected to MongoDB successfully');
    db = client.db('testdb');
    console.log('Database initialized:', db.databaseName);
  } catch (error) {
    console.error('Failed to connect to MongoDB:', error.message, error.stack);
    if (attempt < maxAttempts) {
      console.log(`Retrying MongoDB connection in 5 seconds...`);
      await new Promise(resolve => setTimeout(resolve, 5000));
      return connectToMongoDB(attempt + 1, maxAttempts);
    }
    throw error;
  }
};

app.set('trust proxy', 1);

// Ініціалізація сервера
let isInitialized = false;
let initializationError = null;
let testNames = { 
  '1': { name: 'Тест 1', timeLimit: 3600 },
  '2': { name: 'Тест 2', timeLimit: 3600 },
  '3': { name: 'Тест 3', timeLimit: 3600 }
};

// Налаштування middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Обмеження кількості запитів: 100 запитів за 15 хвилин на одного користувача
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 хвилин
  max: 100, // Максимум 100 запитів
  message: 'Забагато запитів, спробуйте ще раз через 15 хвилин.'
});
app.use(limiter);

// Налаштування сесій із MongoStore
app.use(session({
  store: MongoStore.create({ 
    mongoUrl: MONGO_URL,
    collectionName: 'sessions',
    ttl: 24 * 60 * 60,
    clientPromise: client.connect().then(() => {
      console.log('MongoStore client connected successfully');
      return client;
    }).catch(err => {
      console.error('MongoStore client connection error:', err.message, err.stack);
      throw err;
    })
  }),
  secret: process.env.SESSION_SECRET || 'a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8s9t0',
  resave: false,
  saveUninitialized: false,
  cookie: { 
    secure: process.env.NODE_ENV === 'production', // true на Heroku
    httpOnly: true,
    sameSite: 'lax',
    maxAge: 24 * 60 * 60 * 1000
  }
}));

// Функція для генерації CSRF-токена
const generateCsrfToken = () => {
  return crypto.randomBytes(16).toString('hex');
};

// Middleware для додавання CSRF-токена до сесії та передачі його у відповідь
app.use((req, res, next) => {
  if (!req.session.csrfToken) {
    req.session.csrfToken = generateCsrfToken();
  }
  res.locals.csrfToken = req.session.csrfToken;
  console.log('CSRF Token in session:', req.session.csrfToken);
  next();
});

// Middleware для перевірки CSRF-токена
const verifyCsrfToken = (req, res, next) => {
  console.log('Request body:', req.body);
  console.log('Request headers:', req.headers);
  const csrfToken = req.body._csrf || req.headers['x-csrf-token'];
  console.log('CSRF Token in session:', req.session.csrfToken);
  console.log('CSRF Token received:', csrfToken);
  if (!req.session.csrfToken || csrfToken !== req.session.csrfToken) {
    console.warn('CSRF token validation failed');
    return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
  }
  // Оновлення CSRF-токена після успішного запиту
  req.session.csrfToken = generateCsrfToken();
  res.locals.csrfToken = req.session.csrfToken;
  next();
};

// Завантаження користувачів із файлу users.xlsx та хешування їх паролів
const loadUsers = async () => {
  try {
    const filePath = path.join(__dirname, 'users.xlsx');
    console.log('Attempting to load users from:', filePath);

    if (!fs.existsSync(filePath)) {
      throw new Error(`File users.xlsx not found at path: ${filePath}`);
    }
    console.log('File users.xlsx exists at:', filePath);

    const workbook = new ExcelJS.Workbook();
    console.log('Reading users.xlsx file...');
    await workbook.xlsx.readFile(filePath);
    console.log('File read successfully');

    let sheet = workbook.getWorksheet('Users');
    if (!sheet) {
      console.warn('Worksheet "Users" not found, trying "Sheet1"');
      sheet = workbook.getWorksheet('Sheet1');
      if (!sheet) {
        console.error('Worksheet "Sheet1" not found in users.xlsx');
        throw new Error('Ни один из листов ("Users" или "Sheet1") не найден');
      }
    }
    console.log('Worksheet found:', sheet.name);

    const users = {};
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        const username = String(row.getCell(1).value || '').trim();
        const password = String(row.getCell(2).value || '').trim();
        if (username && password) {
          // Хешуємо пароль перед збереженням
          const hashedPassword = bcrypt.hashSync(password, saltRounds);
          users[username] = hashedPassword;
        }
      }
    });
    if (Object.keys(users).length === 0) {
      console.error('No valid users found in users.xlsx');
      throw new Error('Не найдено пользователей в файле');
    }
    console.log('Loaded users from Excel with hashed passwords:', Object.keys(users));
    return users;
  } catch (error) {
    console.error('Error loading users from users.xlsx:', error.message, error.stack);
    throw error;
  }
};

// Завантаження питань із файлу questionsX.xlsx
const loadQuestions = async (testNumber) => {
  try {
    const filePath = path.join(__dirname, `questions${testNumber}.xlsx`);
    console.log(`Attempting to load questions from: ${filePath}`);
    if (!fs.existsSync(filePath)) {
      console.error(`File questions${testNumber}.xlsx not found at path: ${filePath}`);
      throw new Error(`File questions${testNumber}.xlsx not found at path: ${filePath}`);
    }
    console.log(`File questions${testNumber}.xlsx exists at: ${filePath}`);
    
    const workbook = new ExcelJS.Workbook();
    console.log(`Reading questions${testNumber}.xlsx file...`);
    await workbook.xlsx.readFile(filePath);
    console.log('File read successfully');

    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      console.error(`Worksheet "Questions" not found in questions${testNumber}.xlsx`);
      throw new Error(`Лист "Questions" не знайдено в questions${testNumber}.xlsx`);
    }
    console.log('Worksheet found:', sheet.name);

    const jsonData = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        const rowValues = row.values.slice(1);
        const picture = String(rowValues[0] || '').trim();
        const questionText = String(rowValues[1] || '').trim();
        const options = rowValues.slice(2, 14).filter(Boolean).map(String);
        const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(String);
        const type = String(rowValues[26] || 'multiple').toLowerCase();
        const points = Number(rowValues[27]) || 1;

        jsonData.push({
          picture: picture.match(/^Picture (\d+)/i) ? `/images/Picture ${picture.match(/^Picture (\d+)/i)[1]}.png` : null,
          text: questionText,
          options,
          correctAnswers,
          type,
          points
        });
      }
    });
    console.log(`Loaded questions for test ${testNumber}:`, jsonData);
    if (jsonData.length === 0) {
      console.error(`No questions loaded from questions${testNumber}.xlsx`);
      throw new Error(`No questions found in questions${testNumber}.xlsx`);
    }
    return jsonData;
  } catch (error) {
    console.error(`Ошибка в loadQuestions (test ${testNumber}):`, error.message, error.stack);
    throw error;
  }
};

// Middleware для перевірки ініціалізації сервера
const ensureInitialized = (req, res, next) => {
  if (!isInitialized) {
    if (initializationError) {
      console.error('Server not initialized due to error:', initializationError.message, initializationError.stack);
      return res.status(500).json({ success: false, message: `Server initialization failed: ${initializationError.message}` });
    }
    console.warn('Server is still initializing...');
    return res.status(503).json({ success: false, message: 'Server is initializing, please try again later' });
  }
  next();
};

let validPasswords = null;

// Ініціалізація сервера: підключення до MongoDB та завантаження користувачів
const initializeServer = async () => {
  let attempt = 1;
  const maxAttempts = 5;

  // Ініціалізація MongoDB
  try {
    await connectToMongoDB();
  } catch (error) {
    console.error('Failed to initialize server due to MongoDB connection error:', error.message, error.stack);
    throw error;
  }

  while (attempt <= maxAttempts) {
    try {
      console.log(`Starting server initialization (Attempt ${attempt} of ${maxAttempts})...`);
      validPasswords = await loadUsers();
      console.log('Users initialized successfully from Excel');
      isInitialized = true;
      initializationError = null;
      break;
    } catch (err) {
      console.error(`Failed to initialize server (Attempt ${attempt}):`, err.message, err.stack);
      initializationError = err;
      if (attempt < maxAttempts) {
        console.log(`Retrying initialization in 5 seconds...`);
        await new Promise(resolve => setTimeout(resolve, 5000));
      } else {
        console.error('Maximum initialization attempts reached. Server remains uninitialized.');
      }
      attempt++;
    }
  }
};

// Виконуємо ініціалізацію сервера
(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
  } catch (error) {
    console.error('Failed to start server due to initialization error:', error.message, error.stack);
    process.exit(1);
  }
})();

// Тестовий маршрут для перевірки підключення до MongoDB
app.get('/test-mongo', async (req, res) => {
  try {
    console.log('Testing MongoDB connection...');
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    await db.collection('users').findOne();
    console.log('MongoDB test successful');
    res.json({ success: true, message: 'MongoDB connection successful' });
  } catch (error) {
    console.error('MongoDB test failed:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'MongoDB connection failed', error: error.message });
  }
});

// Тестовий маршрут для перевірки роботи API
app.get('/api/test', (req, res) => {
  console.log('Handling /api/test request...');
  res.json({ success: true, message: 'Express server is working on /api/test' });
});

// Головна сторінка (форма входу)
app.get('/', (req, res) => {
  console.log('Rendering login page with CSRF token (Updated Version 2):', res.locals.csrfToken);
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Вхід</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 20px; }
          h1 { font-size: 24px; margin-bottom: 20px; }
          input { padding: 10px; font-size: 18px; width: 200px; margin-bottom: 10px; }
          button { padding: 10px 20px; font-size: 18px; cursor: pointer; background-color: #4CAF50; color: white; border: none; border-radius: 5px; }
          button:hover { background-color: #45a049; }
          .error { color: red; margin-top: 10px; }
          @media (max-width: 600px) {
            h1 { font-size: 20px; }
            input, button { font-size: 16px; width: 90%; padding: 8px; }
          }
        </style>
      </head>
      <body>
        <h1>Введіть пароль для входу (Updated Version 2)</h1>
        <input type="password" id="password" placeholder="Пароль">
        <br>
        <input type="hidden" id="csrfToken" value="${res.locals.csrfToken || 'undefined'}">
        <button onclick="login()">Увійти</button>
        <div id="error" class="error"></div>

        <script>
          console.log('Login page loaded (Updated Version 2) with CSRF token:', document.getElementById('csrfToken').value);
          async function login() {
            console.log('Login function called (Updated Version 2)');
            const password = document.getElementById('password').value;
            const csrfToken = document.getElementById('csrfToken').value;
            console.log('CSRF Token being sent (Updated Version 2):', csrfToken);
            if (csrfToken === 'undefined') {
              document.getElementById('error').textContent = 'CSRF-токен відсутній. Оновіть сторінку.';
              return;
            }
            const response = await fetch('/login', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ password, _csrf: csrfToken })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = result.redirect;
            } else {
              document.getElementById('error').textContent = 'Помилка: ' + result.message;
            }
          }

          document.getElementById('password').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') login();
          });
        </script>
      </body>
    </html>
  `);
});

// Функція для логування дій користувача
const logActivity = async (req, user, action) => {
  try {
    const timestamp = new Date();
    // Додаємо зсув +3 години (UTC+3)
    const timeOffset = 3 * 60 * 60 * 1000; // 3 години в мілісекундах
    const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
    // Отримуємо IP-адресу
    const ipAddress = req.headers['x-forwarded-for'] || req.ip || 'N/A';
    // Отримуємо ідентифікатор сесії
    const sessionId = req.sessionID || 'N/A';
    await db.collection('activity_log').insertOne({
      user,
      action,
      ipAddress,  // Додаємо IP-адресу
      sessionId,  // Додаємо ідентифікатор сесії
      timestamp: adjustedTimestamp.toISOString()
    });
    console.log(`Logged activity: ${user} - ${action} at ${adjustedTimestamp}, IP: ${ipAddress}, SessionID: ${sessionId}`);
  } catch (error) {
    console.error('Error logging activity:', error.message, error.stack);
  }
};

// Маршрут для авторизації користувача
app.post('/login', verifyCsrfToken, async (req, res) => {
  try {
    console.log('Handling /login request...');
    console.log('Request body:', req.body);
    const { password } = req.body;

    // Валідація введення пароля
    if (!password || typeof password !== 'string' || password.length < 6 || !/^[a-zA-Z0-9]+$/.test(password)) {
      console.warn('Invalid password provided in /login request');
      return res.status(400).json({ success: false, message: 'Пароль має бути довжиною не менше 6 символів і містити лише латинські літери та цифри' });
    }

    console.log('Checking password against hashed passwords...');
    const user = Object.keys(validPasswords).find(u => bcrypt.compareSync(password, validPasswords[u]));

    if (!user) {
      console.warn('Password not found in validPasswords');
      return res.status(401).json({ success: false, message: 'Невірний пароль' });
    }

    req.session.user = user;
    await logActivity(req, user, 'увійшов на сайт');
    console.log('Session after setting user:', req.session);
    console.log('Session ID after setting user:', req.sessionID);
    console.log('Cookies after setting session:', req.cookies);

    req.session.save(err => {
      if (err) {
        console.error('Error saving session in /login:', err.message, err.stack);
        return res.status(500).json({ success: false, message: 'Помилка сервера' });
      }
      console.log('Session saved successfully');
      if (user === 'admin') {
        console.log('Redirecting to /admin for user:', user);
        res.json({ success: true, redirect: '/admin' });
      } else {
        console.log('Redirecting to /select-test for user:', user);
        res.json({ success: true, redirect: '/select-test' });
      }
    });
  } catch (error) {
    console.error('Ошибка в /login:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка сервера' });
  }
});

// Middleware для перевірки авторизації
const checkAuth = (req, res, next) => {
  console.log('checkAuth: Session data:', req.session);
  console.log('checkAuth: Cookies:', req.cookies);
  console.log('checkAuth: Session ID:', req.sessionID);
  const user = req.session.user;
  console.log('checkAuth: user from session:', user);
  if (!user) {
    console.log('checkAuth: No valid auth, redirecting to /');
    return res.redirect('/');
  }
  req.user = user;
  next();
};

// Middleware для перевірки прав адміністратора
const checkAdmin = (req, res, next) => {
  const user = req.session.user;
  console.log('checkAdmin: user from session:', user);
  if (user !== 'admin') {
    console.log('checkAdmin: Not admin, returning 403');
    return res.status(403).send('Доступно тільки для адміністратора (403 Forbidden)');
  }
  next();
};

// Сторінка вибору тесту
app.get('/select-test', checkAuth, (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  console.log('Serving /select-test for user:', req.user);
  res.send(`
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
          }
          h1 { 
            font-size: 24px; 
            margin-bottom: 20px; 
          }
          .test-buttons { 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            gap: 10px; 
          }
          button { 
            padding: 10px; 
            font-size: 18px; 
            cursor: pointer; 
            width: 200px; 
            border: none; 
            border-radius: 5px; 
            background-color: #4CAF50; 
            color: white; 
          }
          button:hover { 
            background-color: #45a049; 
          }
          #logout { 
            background-color: #ef5350; 
            color: white; 
            position: fixed; 
            bottom: 20px; 
            left: 50%; 
            transform: translateX(-50%); 
            width: 200px; 
          }
          @media (max-width: 600px) {
            h1 { 
              font-size: 28px; 
            }
            button { 
              font-size: 20px; 
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
          ${Object.entries(testNames).map(([num, data]) => `
            <button onclick="window.location.href='/test?test=${num}'">${data.name}</button>
          `).join('')}
        </div>
        <button id="logout" onclick="logout()">Вийти</button>
        <script>
          async function logout() {
            const response = await fetch('/logout', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ _csrf: "${res.locals.csrfToken}" })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = '/';
            } else {
              alert('Помилка при виході');
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Маршрут для виходу з системи
app.post('/logout', verifyCsrfToken, (req, res) => {
  const user = req.session.user;
  const userTest = userTests.get(user);
  if (userTest) {
    logActivity(req, user, `завершив сесію не закінчивши тест`);
    userTests.delete(user);
  } else {
    logActivity(req, user, `завершив сесію`);
  }

  req.session.destroy(err => {
    if (err) {
      console.error('Error destroying session:', err);
      return res.status(500).json({ success: false, message: 'Помилка при виході' });
    }
    res.json({ success: true });
  });
});

const userTests = new Map();

// Збереження результатів тесту в MongoDB
const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage) => {
  try {
    console.log('Starting saveResult for user:', user, 'testNumber:', testNumber);
    const duration = Math.round((endTime - startTime) / 1000);
    const userTest = userTests.get(user);
    console.log('User test data:', userTest);
    const answers = userTest ? userTest.answers : {};
    const questions = userTest ? userTest.questions : [];
    const suspiciousActivity = userTest ? userTest.suspiciousActivity : { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    console.log('Answers:', answers, 'Questions:', questions);
    console.log('Suspicious activity:', suspiciousActivity);

    // Обчислення балів за кожне питання
    const scoresPerQuestion = questions.map((q, index) => {
      const userAnswer = answers[index];
      let questionScore = 0;
      if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\\'/g, "'"));
        const isCorrect = correctAnswers.length === userAnswers.length &&
          correctAnswers.every(val => userAnswers.includes(val)) &&
          userAnswers.every(val => correctAnswers.includes(val));
        console.log(`Question ${index + 1} (multiple): userAnswer=${userAnswers}, correctAnswer=${correctAnswers}, isCorrect=${isCorrect}`);
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'input' && userAnswer) {
        const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
        console.log(`Question ${index + 1} (input): userAnswer=${normalizedUserAnswer}, correctAnswer=${normalizedCorrectAnswer}, isCorrect=${isCorrect}`);
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
        const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
        console.log(`Question ${index + 1} (ordering): userAnswer=${userAnswers}, correctAnswer=${correctAnswers}, isCorrect=${isCorrect}`);
        if (isCorrect) {
          questionScore = q.points;
        }
      }
      return questionScore;
    });

    const calculatedScore = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
    console.log(`Calculated score: ${calculatedScore}, Provided score: ${score}`);

    let suspiciousScore = 0;
    const timeAwayPercent = suspiciousActivity.timeAway ? 
      Math.round((suspiciousActivity.timeAway / (duration * 1000)) * 100) : 0;
    suspiciousScore += timeAwayPercent;

    const switchCount = suspiciousActivity.switchCount || 0;
    if (switchCount > totalQuestions * 2) {
      suspiciousScore += 20;
    }

    const responseTimes = suspiciousActivity.responseTimes || [];
    const avgResponseTime = responseTimes.length > 0 ? 
      (responseTimes.reduce((sum, time) => sum + (time || 0), 0) / responseTimes.length / 1000).toFixed(2) : 0;
    responseTimes.forEach(time => {
      if (time < 5000) {
        suspiciousScore += 10;
      } else if (time > 5 * 60 * 1000) {
        suspiciousScore += 10;
      }
    });

    const activityCounts = suspiciousActivity.activityCounts || [];
    const avgActivityCount = activityCounts.length > 0 ? 
      (activityCounts.reduce((sum, count) => sum + (count || 0), 0) / activityCounts.length).toFixed(2) : 0;
    activityCounts.forEach((count, idx) => {
      if (count < 5 && responseTimes[idx] > 30 * 1000) {
        suspiciousScore += 10;
      }
    });

    let typicalResponseTime = 30 * 1000;
    let typicalSwitchCount = totalQuestions;
    const allResults = await db.collection('test_results').find({}).toArray();
    if (allResults.length > 0) {
      const allResponseTimes = allResults.flatMap(r => r.suspiciousActivity.responseTimes || []);
      typicalResponseTime = allResponseTimes.length > 0 ? 
        allResponseTimes.reduce((sum, time) => sum + (time || 0), 0) / allResponseTimes.length : typicalResponseTime;
      const allSwitchCounts = allResults.map(r => r.suspiciousActivity.switchCount || 0);
      typicalSwitchCount = allSwitchCounts.length > 0 ? 
        allSwitchCounts.reduce((sum, count) => sum + count, 0) / allSwitchCounts.length : typicalSwitchCount;
    }
    if (avgResponseTime < typicalResponseTime * 0.5 || avgResponseTime > typicalResponseTime * 1.5) {
      suspiciousScore += 15;
    }
    if (switchCount > typicalSwitchCount * 1.5) {
      suspiciousScore += 15;
    }

    suspiciousScore = Math.min(suspiciousScore, 100);

    const timeOffset = 3 * 60 * 60 * 1000; // 3 часа в миллисекундах
    const adjustedStartTime = new Date(startTime + timeOffset);
    const adjustedEndTime = new Date(endTime + timeOffset);

    const result = {
      user,
      testNumber,
      score: calculatedScore,
      totalPoints,
      totalClicks,
      correctClicks,
      totalQuestions,
      percentage,
      startTime: adjustedStartTime.toISOString(),
      endTime: adjustedEndTime.toISOString(),
      duration,
      answers: Object.fromEntries(Object.entries(answers).sort((a, b) => parseInt(a[0]) - parseInt(b[0]))),
      scoresPerQuestion: scoresPerQuestion.map((score, idx) => {
        console.log(`Saving score for question ${idx + 1}: ${score}`);
        return score;
      }),
      suspiciousActivity: {
        ...suspiciousActivity,
        suspiciousScore,
        responseTimes: responseTimes.map(time => Math.min(time, 5 * 60 * 1000))
      }
    };
    console.log('Saving result to MongoDB:', result);
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    const insertResult = await db.collection('test_results').insertOne(result);
    console.log(`Successfully saved result for ${user} in MongoDB with ID:`, insertResult.insertedId);
  } catch (error) {
    console.error('Ошибка сохранения в MongoDB:', error.message, error.stack);
    throw error;
  }
};

// Початок тесту
app.get('/test', checkAuth, async (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  const testNumber = req.query.test;
  console.log(`Processing /test request for testNumber: ${testNumber}, user: ${req.user}`);
  if (!testNumber) {
    console.warn('Test number not provided in query');
    return res.status(400).send('Номер тесту не вказано');
  }
  if (!testNames[testNumber]) {
    console.warn(`Test ${testNumber} not found`);
    return res.status(404).send('Тест не знайдено');
  }
  try {
    console.log(`Loading questions for test ${testNumber}...`);
    const questions = await loadQuestions(testNumber);
    userTests.set(req.user, {
      testNumber,
      questions,
      answers: {},
      currentQuestion: 0,
      startTime: Date.now(),
      timeLimit: testNames[testNumber].timeLimit * 1000
    });
    await logActivity(req, req.user, `розпочав тест ${testNames[testNumber].name}`);
    console.log(`Redirecting to first question for user ${req.user}`);
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    console.error('Ошибка в /test:', error.message, error.stack);
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  }
});

// Відображення питання тесту для користувача
app.get('/test/question', checkAuth, async (req, res) => {
  const { index } = req.query;
  const idx = parseInt(index, 10);
  const testNumber = req.session.currentTest;

  if (!testNumber || isNaN(idx)) {
    return res.redirect('/select-test');
  }

  const questions = await loadQuestions(testNumber);
  if (!questions || idx < 0 || idx >= questions.length) {
    return res.redirect('/select-test');
  }

  const question = { ...questions[idx], index: idx };
  console.log('Rendering question:', question);

  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Тест</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }
          h2 { font-size: 24px; margin-bottom: 20px; }
          img { max-width: 100%; height: auto; margin-bottom: 20px; }
          .options { display: flex; flex-direction: column; gap: 10px; }
          label { display: flex; align-items: center; gap: 10px; font-size: 18px; }
          input[type="checkbox"], input[type="text"] { margin: 0; }
          input[type="text"] { padding: 5px; font-size: 16px; width: 100%; max-width: 300px; }
          button { padding: 10px 20px; font-size: 18px; cursor: pointer; margin-top: 20px; border: none; border-radius: 5px; }
          #next { background-color: #4CAF50; color: white; }
          #next:hover { background-color: #45a049; }
          #finish { background-color: #2196F3; color: white; }
          #finish:hover { background-color: #1e88e5; }
          #logout { background-color: #ef5350; color: white; position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); width: 200px; }
          .sortable { list-style-type: none; padding: 0; margin: 0; }
          .sortable li { padding: 10px; background-color: #f0f0f0; margin-bottom: 5px; cursor: move; }
          @media (max-width: 600px) {
            body { padding: 10px; }
            h2 { font-size: 20px; }
            label { font-size: 16px; }
            button { font-size: 16px; padding: 8px 16px; }
            #logout { width: 90%; }
          }
        </style>
      </head>
      <body>
        <h2>${question.text}</h2>
        ${question.picture ? `<img src="${question.picture}" alt="Question Image">` : ''}
        ${question.type === 'multiple' ? `
          <div class="options">
            ${question.options.map((option, i) => `
              <label>
                <input type="checkbox" name="answer" value="${option}">
                ${option}
              </label>
            `).join('')}
          </div>
        ` : question.type === 'input' ? `
          <div class="options">
            <input type="text" id="answerInput" placeholder="Введіть відповідь">
          </div>
        ` : question.type === 'ordering' ? `
          <ul class="sortable">
            ${question.options.map((option, i) => `
              <li data-id="${i}">${option}</li>
            `).join('')}
          </ul>
        ` : ''}

        <input type="hidden" id="csrfToken" value="${res.locals.csrfToken || 'undefined'}">
        <button id="next" onclick="submitAnswer(${idx}, ${questions.length - 1})">Далі</button>
        ${idx === questions.length - 1 ? `<button id="finish" onclick="submitAnswer(${idx}, ${questions.length - 1}, true)">Завершити тест</button>` : ''}
        <button id="logout" onclick="logout()">Вийти</button>

        <script>
          let startTime = Date.now();
          let timeAway = 0;
          let lastFocus = Date.now();
          let switchCount = 0;
          let activityCount = 0;

          document.addEventListener('visibilitychange', () => {
            if (document.hidden) {
              lastFocus = Date.now();
            } else {
              timeAway += Date.now() - lastFocus;
              switchCount++;
            }
          });

          document.addEventListener('mousemove', () => activityCount++);
          document.addEventListener('keydown', () => activityCount++);

          async function submitAnswer(index, lastIndex, finish = false) {
            const responseTime = Date.now() - startTime - timeAway;
            let answer;
            const csrfToken = document.getElementById('csrfToken').value;

            if (${question.type === 'multiple'}) {
              answer = Array.from(document.querySelectorAll('input[name="answer"]:checked')).map(input => input.value);
            } else if (${question.type === 'input'}) {
              answer = [document.getElementById('answerInput').value.trim()];
            } else if (${question.type === 'ordering'}) {
              const items = Array.from(document.querySelectorAll('.sortable li'));
              answer = items.map(item => item.textContent.trim());
            }

            console.log('Submitting answer with CSRF token:', csrfToken);
            const response = await fetch('/answer', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                index,
                answer,
                timeAway,
                switchCount,
                responseTime,
                activityCount,
                _csrf: csrfToken
              })
            });

            const result = await response.json();
            console.log('Server response:', result);
            if (result.success) {
              if (finish) {
                window.location.href = '/result';
              } else {
                window.location.href = '/test/question?index=' + (index + 1);
              }
            } else {
              alert('Помилка: ' + result.message);
            }
          }

          async function logout() {
            const csrfToken = document.getElementById('csrfToken').value;
            const response = await fetch('/logout', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ _csrf: csrfToken })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = '/';
            } else {
              alert('Помилка при виході');
            }
          }

          // Drag-and-drop для ordering
          if (${question.type === 'ordering'}) {
            const sortableList = document.querySelector('.sortable');
            let draggedItem = null;

            sortableList.addEventListener('dragstart', (e) => {
              draggedItem = e.target;
              setTimeout(() => draggedItem.style.display = 'none', 0);
            });

            sortableList.addEventListener('dragend', (e) => {
              setTimeout(() => {
                draggedItem.style.display = 'block';
                draggedItem = null;
              }, 0);
            });

            sortableList.addEventListener('dragover', (e) => e.preventDefault());

            sortableList.addEventListener('dragenter', (e) => {
              e.preventDefault();
              if (e.target.classList.contains('sortable') || e.target.tagName === 'LI') {
                e.target.classList.add('drag-over');
              }
            });

            sortableList.addEventListener('dragleave', (e) => {
              if (e.target.classList.contains('sortable') || e.target.tagName === 'LI') {
                e.target.classList.remove('drag-over');
              }
            });

            sortableList.addEventListener('drop', (e) => {
              e.preventDefault();
              const target = e.target.classList.contains('sortable') ? e.target : e.target.closest('li');
              if (target && target !== draggedItem) {
                const allItems = Array.from(sortableList.children);
                const draggedIndex = allItems.indexOf(draggedItem);
                const targetIndex = allItems.indexOf(target);
                if (draggedIndex < targetIndex) {
                  target.after(draggedItem);
                } else {
                  target.before(draggedItem);
                }
              }
              sortableList.querySelectorAll('.drag-over').forEach(item => item.classList.remove('drag-over'));
            });
          }
        </script>
      </body>
    </html>
  `);
});

// Збереження відповіді на питання
app.post('/answer', checkAuth, verifyCsrfToken, (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  try {
    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;

    // Валідація вхідних даних
    if (typeof index !== 'number' || index < 0) {
      return res.status(400).json({ error: 'Невірний індекс питання' });
    }
    if (answer === undefined || (typeof answer !== 'string' && !Array.isArray(answer))) {
      return res.status(400).json({ error: 'Невірний формат відповіді' });
    }
    if (typeof timeAway !== 'number' || timeAway < 0) {
      return res.status(400).json({ error: 'Невірне значення timeAway' });
    }
    if (typeof switchCount !== 'number' || switchCount < 0) {
      return res.status(400).json({ error: 'Невірне значення switchCount' });
    }
    if (typeof responseTime !== 'number' || responseTime < 0) {
      return res.status(400).json({ error: 'Невірне значення responseTime' });
    }
    if (typeof activityCount !== 'number' || activityCount < 0) {
      return res.status(400).json({ error: 'Невірне значення activityCount' });
    }

    const userTest = userTests.get(req.user);
    if (!userTest) {
      console.warn(`Test not started for user ${req.user} in /answer`);
      return res.status(400).json({ error: 'Тест не розпочато' });
    }
    userTest.answers[index] = answer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = (userTest.suspiciousActivity.timeAway || 0) + (timeAway || 0);
    userTest.suspiciousActivity.switchCount = (userTest.suspiciousActivity.switchCount || 0) + (switchCount || 0);
    userTest.suspiciousActivity.responseTimes[index] = responseTime || 0;
    userTest.suspiciousActivity.activityCounts[index] = activityCount || 0;
    console.log(`Saved answer for user ${req.user}, question ${index}:`, answer);
    console.log(`Updated suspicious activity for user ${req.user}:`, userTest.suspiciousActivity);
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка в /answer:', error.message, error.stack);
    res.status(500).json({ error: 'Помилка сервера' });
  }
});

// Відображення результатів тесту
app.get('/result', checkAuth, async (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  const userTest = userTests.get(req.user);
  if (!userTest) {
    console.warn(`Test not started for user ${req.user} in /result`);
    return res.status(400).json({ error: 'Тест не розпочато' });
  }

  const { questions, answers, testNumber, startTime } = userTest;
  let score = 0;
  const totalPoints = questions.reduce((sum, q) => sum + q.points, 0);

  const scoresPerQuestion = questions.map((q, index) => {
    const userAnswer = answers[index];
    let questionScore = 0;
    if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
      const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
      const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\\'/g, "'"));
      const isCorrect = correctAnswers.length === userAnswers.length &&
        correctAnswers.every(val => userAnswers.includes(val)) &&
        userAnswers.every(val => correctAnswers.includes(val));
      console.log(`Question ${index + 1} (multiple): userAnswer=${userAnswers}, correctAnswer=${correctAnswers}, isCorrect=${isCorrect}`);
      if (isCorrect) {
        questionScore = q.points;
      }
    } else if (q.type === 'input' && userAnswer) {
      const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
      console.log(`Question ${index + 1} (input): userAnswer=${normalizedUserAnswer}, correctAnswer=${normalizedCorrectAnswer}, isCorrect=${isCorrect}`);
      if (isCorrect) {
        questionScore = q.points;
      }
    } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
      const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
      const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
      const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
      console.log(`Question ${index + 1} (ordering): userAnswer=${userAnswers}, correctAnswer=${correctAnswers}, isCorrect=${isCorrect}`);
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

  try {
    await saveResult(req.user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage);
    await logActivity(req, req.user, `завершив тест ${testNames[testNumber].name} з результатом ${score} балів`);
  } catch (error) {
    console.error('Error saving result in /result:', error.message, error.stack);
    return res.status(500).send('Помилка при збереженні результату');
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
    console.error('Error reading image A.png:', error.message, error.stack);
    imageBase64 = '';
  }

  const resultHtml = `
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Результати ${testNames[testNumber].name}</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 20px; background-color: #f5f5f5; }
          .result-circle { width: 100px; height: 100px; background-color: #ff4d4d; color: white; font-size: 24px; line-height: 100px; border-radius: 50%; margin: 0 auto; }
          .buttons { margin-top: 20px; }
          button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
          #exportPDF { background-color: #ffeb3b; }
          #restart { background-color: #ef5350; }
        </style>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js" onerror="console.error('Failed to load pdfmake.min.js')"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js" onerror="console.error('Failed to load vfs_fonts.js')"></script>
      </head>
      <body>
        <h1>Результат тесту</h1>
        <div class="result-circle">${Math.round(percentage)}%</div>
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
          if (typeof pdfMake === 'undefined') {
            console.error('pdfMake is not loaded');
            document.getElementById('exportPDF').disabled = true;
            document.getElementById('exportPDF').textContent = 'PDF не доступно';
          }

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
            console.log('Export PDF button clicked');
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
              console.log('PDF generated successfully');
            } catch (error) {
              console.error('Error generating PDF:', error);
            }
          });

          document.getElementById('restart').addEventListener('click', () => {
            console.log('Restart button clicked');
            window.location.href = '/select-test';
          });
        </script>
      </body>
    </html>
  `;
  res.send(resultHtml);
});

// Перегляд детальних результатів для користувача
app.get('/results', checkAuth, async (req, res) => {
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
    const { questions, answers, testNumber, startTime } = userTest;
    let score = 0;
    const totalPoints = questions.reduce((sum, q) => sum + q.points, 0);

    const scoresPerQuestion = questions.map((q, index) => {
      const userAnswer = answers[index];
      let questionScore = 0;
      if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
        const correctAnswers = q.correctAnswers.map(String).map(val => val.trim().toLowerCase());
        const userAnswers = userAnswer.map(String).map(val => val.trim().toLowerCase());
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
        const userAnswers = userAnswer.map(String).map(val => val.trim().toLowerCase());
        const correctAnswers = q.correctAnswers.map(String).map(val => val.trim().toLowerCase());
        if (userAnswers.join(',') === correctAnswers.join(',')) {
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

    const endDateTime = new Date(endTime);
    const formattedTime = endDateTime.toLocaleTimeString('uk-UA', { hour12: false });
    const formattedDate = endDateTime.toLocaleDateString('uk-UA');

    const imagePath = path.join(__dirname, 'public', 'images', 'A.png');
    let imageBase64 = '';
    try {
      const imageBuffer = fs.readFileSync(imagePath);
      imageBase64 = imageBuffer.toString('base64');
    } catch (error) {
      console.error('Error reading image A.png:', error.message, error.stack);
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
      const correctAnswer = q.correctAnswers.join(', ');
      const questionScore = scoresPerQuestion[index];
      resultsHtml += `
        <tr>
          <td>${q.text}</td>
          <td>${Array.isArray(userAnswer) ? userAnswer.join(', ') : userAnswer}</td>
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
              {
                image: 'data:image/png;base64,' + imageBase64,
                width: 150,
                alignment: 'center',
                margin: [0, 0, 0, 20]
              },
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
});

// Адміністративна панель
app.get('/admin', checkAuth, checkAdmin, (req, res) => {
  console.log('Serving /admin for user:', req.user);
  res.send(`
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
          }
          h1 { 
            font-size: 36px; 
            margin-bottom: 20px; 
          }
          button { 
            padding: 15px 30px; 
            margin: 10px; 
            font-size: 24px; 
            cursor: pointer; 
            width: 300px; 
            border: none; 
            border-radius: 5px; 
            background-color: #4CAF50; 
            color: white; 
          }
          button:hover { 
            background-color: #45a049; 
          }
          #logout { 
            background-color: #ef5350; 
            color: white; 
          }
          @media (max-width: 600px) {
            body { 
              padding: 20px; 
              padding-bottom: 80px; 
            }
            h1 { 
              font-size: 32px; 
            }
            button { 
              font-size: 20px; 
              width: 90%; 
              padding: 15px; 
            }
            #logout { 
              position: fixed; 
              bottom: 20px; 
              left: 50%; 
              transform: translateX(-50%); 
              width: 90%; 
            }
          }
        </style>
      </head>
      <body>
        <h1>Адмін-панель</h1>
        <button onclick="window.location.href='/admin/results'">Перегляд результатів</button><br>
        <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button><br>
        <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button><br>
        <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
        <button id="logout" onclick="logout()">Вийти</button>
        <script>
          async function logout() {
            const response = await fetch('/logout', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ _csrf: "${res.locals.csrfToken}" })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = '/';
            } else {
              alert('Помилка при виході');
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Перегляд результатів усіх користувачів (для адміна)
app.get('/admin/results', checkAuth, checkAdmin, async (req, res) => {
  let results = [];
  let errorMessage = '';
  try {
    console.log('Fetching test results from MongoDB...');
    results = await db.collection('test_results').find({}).sort({ endTime: -1 }).toArray();
    console.log('Fetched results from MongoDB:', results);
  } catch (fetchError) {
    console.error('Ошибка при получении данных из MongoDB:', fetchError.message, fetchError.stack);
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
          .answers { white-space: pre-wrap; max-width: 300px; overflow-wrap: break-word; line-height: 1.8; }
          .delete-btn { background-color: #ff4d4d; color: white; padding: 5px 10px; border: none; cursor: pointer; }
          .nav-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; }
          .details { white-space: pre-wrap; max-width: 300px; line-height: 1.8; }
        </style>
      </head>
      <body>
        <h1>Результати всіх користувачів</h1>
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
  `;
  if (errorMessage) {
    adminHtml += `<p class="error">${errorMessage}</p>`;
  }
  adminHtml += `
        <table>
          <tr>
            <th>Користувач</th>
            <th>Тест</th>
            <th>Очки</th>
            <th>Максимум</th>
            <th>Початок</th>
            <th>Кінець</th>
            <th>Тривалість (сек)</th>
            <th>Підозріла активність (%)</th>
            <th>Деталі активності</th>
            <th>Відповіді та бали</th>
            <th>Дія</th>
          </tr>
  `;
  if (!results || results.length === 0) {
    adminHtml += '<tr><td colspan="11">Немає результатів</td></tr>';
    console.log('No results found in test_results');
  } else {
    results.forEach((r, index) => {
      const answersArray = [];
      if (r.answers) {
        Object.keys(r.answers).sort((a, b) => parseInt(a) - parseInt(b)).forEach(key => {
          const idx = parseInt(key);
          answersArray[idx] = r.answers[key];
        });
      }

      // Форматування відповідей для відображення
      const answersDisplay = answersArray.length > 0 
        ? answersArray.map((a, i) => {
            if (!a) return null;
            const userAnswer = Array.isArray(a) ? a.join(', ') : a;
            const questionScore = r.scoresPerQuestion[i] || 0;
            console.log(`Result ${index + 1}, Question ${i + 1}: userAnswer=${userAnswer}, savedScore=${questionScore}`);
            return `Питання ${i + 1}: ${userAnswer.replace(/\\'/g, "'")} (${questionScore} балів)`;
          }).filter(line => line !== null).join('\n')
        : 'Немає відповідей';

      // Форматування дати та часу
      const formatDateTime = (isoString) => {
        if (!isoString) return 'N/A';
        const date = new Date(isoString);
        return `${date.toLocaleTimeString('uk-UA', { hour12: false })} ${date.toLocaleDateString('uk-UA')}`;
      };

      // Обчислення підозрілої активності
      const suspiciousActivityPercent = r.suspiciousActivity && r.suspiciousActivity.suspiciousScore ? 
        Math.round(r.suspiciousActivity.suspiciousScore) : 0;
      const timeAwayPercent = r.suspiciousActivity && r.suspiciousActivity.timeAway ? 
        Math.round((r.suspiciousActivity.timeAway / (r.duration * 1000)) * 100) : 0;
      const switchCount = r.suspiciousActivity ? r.suspiciousActivity.switchCount || 0 : 0;
      const avgResponseTime = r.suspiciousActivity && r.suspiciousActivity.responseTimes ? 
        (r.suspiciousActivity.responseTimes.reduce((sum, time) => sum + (time || 0), 0) / r.suspiciousActivity.responseTimes.length / 1000).toFixed(2) : 0;
      const totalActivityCount = r.suspiciousActivity && r.suspiciousActivity.activityCounts ? 
        r.suspiciousActivity.activityCounts.reduce((sum, count) => sum + (count || 0), 0).toFixed(0) : 0;
      const activityDetails = `
Время вне вкладки: ${timeAwayPercent}%
Переключения вкладок: ${switchCount}
Среднее время ответа (сек): ${avgResponseTime}
Общее количество действий: ${totalActivityCount}
      `;
      adminHtml += `
        <tr>
          <td>${r.user || 'N/A'}</td>
          <td>${testNames[r.testNumber]?.name || 'N/A'}</td>
          <td>${r.score || '0'}</td>
          <td>${r.totalPoints || '0'}</td>
          <td>${formatDateTime(r.startTime)}</td>
          <td>${formatDateTime(r.endTime)}</td>
          <td>${r.duration || 'N/A'}</td>
          <td>${suspiciousActivityPercent}%</td>
          <td class="details">${activityDetails}</td>
          <td class="answers">${answersDisplay}</td>
          <td><button class="delete-btn" onclick="deleteResult('${r._id}')">🗑️ Видалити</button></td>
        </tr>
      `;
    });
  }
  adminHtml += `
        </table>
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <script>
          async function deleteResult(id) {
            if (confirm('Ви впевнені, що хочете видалити цей результат?')) {
              try {
                const response = await fetch('/admin/delete-result', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ id, _csrf: "${res.locals.csrfToken}" })
                });
                const result = await response.json();
                if (result.success) {
                  console.log('Result deleted successfully');
                  window.location.reload();
                } else {
                  console.error('Failed to delete result:', result.message);
                  alert('Помилка при видаленні результату: ' + result.message);
                }
              } catch (error) {
                console.error('Error deleting result:', error);
                alert('Помилка при видаленні результату');
              }
            }
          }
        </script>
      </body>
    </html>
  `;
  res.send(adminHtml);
});

// Видалення результату тесту (для адміна)
app.post('/admin/delete-result', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    const { id } = req.body;
    // Валідація ID
    if (!id || typeof id !== 'string') {
      return res.status(400).json({ success: false, message: 'Невірний ID результату' });
    }
    console.log(`Deleting result with id ${id}...`);
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    console.log(`Result with id ${id} deleted from MongoDB`);
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении результата:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  }
});

// Видалення всіх результатів тестів (для адміна)
app.get('/admin/delete-results', checkAuth, checkAdmin, async (req, res) => {
  try {
    console.log('Deleting all test results...');
    await db.collection('test_results').deleteMany({});
    console.log('Test results deleted from MongoDB');
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Видалено результати</title>
        </head>
        <body>
          <h1>Результати успішно видалено</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Ошибка при удалении результатов:', error.message, error.stack);
    res.status(500).send('Помилка при видаленні результатів');
  }
});

// Редагування назв та часу тестів (для адміна)
app.get('/admin/edit-tests', checkAuth, checkAdmin, (req, res) => {
  console.log('Serving /admin/edit-tests for user:', req.user);
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Редагувати назви тестів</title>
        <style>
          body { font-size: 24px; margin: 20px; }
          input { font-size: 24px; padding: 5px; margin: 5px; }
          button { font-size: 24px; padding: 10px 20px; margin: 5px; }
          .delete-btn { background-color: #ff4d4d; color: white; }
          .test-row { display: flex; align-items: center; margin-bottom: 10px; }
        </style>
      </head>
      <body>
        <h1>Редагувати назви та час тестів</h1>
        <form method="POST" action="/admin/edit-tests">
          <input type="hidden" name="_csrf" value="${res.locals.csrfToken}">
          ${Object.entries(testNames).map(([num, data]) => `
            <div class="test-row">
              <label for="test${num}">Назва Тесту ${num}:</label>
              <input type="text" id="test${num}" name="test${num}" value="${data.name}" required>
              <label for="time${num}">Час (сек):</label>
              <input type="number" id="time${num}" name="time${num}" value="${data.timeLimit}" required min="1">
              <button type="button" class="delete-btn" onclick="deleteTest('${num}')">Видалити</button>
            </div>
          `).join('')}
          <button type="submit">Зберегти</button>
        </form>
        <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <script>
          async function deleteTest(testNumber) {
            if (confirm('Ви впевнені, що хочете видалити Тест ' + testNumber + '?')) {
              await fetch('/admin/delete-test', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ testNumber, _csrf: "${res.locals.csrfToken}" })
              });
              window.location.reload();
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Збереження змін назв та часу тестів (для адміна)
app.post('/admin/edit-tests', checkAuth, checkAdmin, verifyCsrfToken, (req, res) => {
  try {
    console.log('Updating test names and time limits...');
    Object.keys(testNames).forEach(num => {
      const testName = req.body[`test${num}`];
      const timeLimit = req.body[`time${num}`];
      // Валідація
      if (!testName || typeof testName !== 'string' || testName.length < 3) {
        throw new Error(`Невірна назва тесту ${num}: має бути не менше 3 символів`);
      }
      if (!timeLimit || isNaN(timeLimit) || parseInt(timeLimit) < 60) {
        throw new Error(`Невірний час для тесту ${num}: має бути не менше 60 секунд`);
      }
      testNames[num] = {
        name: testName,
        timeLimit: parseInt(timeLimit)
      };
    });
    console.log('Updated test names and time limits:', testNames);
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Назви оновлено</title>
        </head>
        <body>
          <h1>Назви та час тестів успішно оновлено</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Ошибка при редактировании названий тестов:', error.message, error.stack);
    res.status(400).send(`Помилка при оновленні назв тестів: ${error.message}`);
  }
});

// Видалення тесту (для адміна)
app.post('/admin/delete-test', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    const { testNumber } = req.body;
    // Валідація
    if (!testNumber || !testNames[testNumber]) {
      return res.status(404).json({ success: false, message: 'Тест не знайдено' });
    }
    delete testNames[testNumber];
    console.log(`Deleted test ${testNumber}, updated testNames:`, testNames);
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении теста:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  }
});

// Створення нового тесту (для адміна)
app.get('/admin/create-test', checkAuth, checkAdmin, (req, res) => {
  const excelFiles = fs.readdirSync(__dirname).filter(file => file.endsWith('.xlsx') && file.startsWith('questions'));
  console.log('Available Excel files:', excelFiles);
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Створити новий тест</title>
        <style>
          body { font-size: 24px; margin: 20px; }
          input { font-size: 24px; padding: 5px; margin: 5px; }
          select { font-size: 24px; padding: 5px; margin: 5px; }
          button { font-size: 24px; padding: 10px 20px; margin: 5px; }
        </style>
      </head>
      <body>
        <h1>Створити новий тест</h1>
        <form method="POST" action="/admin/create-test">
          <input type="hidden" name="_csrf" value="${res.locals.csrfToken}">
          <div>
            <label for="testName">Назва нового тесту:</label>
            <input type="text" id="testName" name="testName" required>
          </div>
          <div>
            <label for="timeLimit">Час (сек):</label>
            <input type="number" id="timeLimit" name="timeLimit" value="3600" required min="1">
          </div>
          <div>
            <label for="excelFile">Оберіть файл Excel з питаннями:</label>
            <select id="excelFile" name="excelFile" required>
              ${excelFiles.map(file => `<option value="${file}">${file}</option>`).join('')}
            </select>
          </div>
          <button type="submit">Створити</button>
        </form>
        <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
      </body>
    </html>
  `);
});

// Збереження нового тесту (для адміна)
app.post('/admin/create-test', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    const { testName, excelFile, timeLimit } = req.body;
    // Валідація
    if (!testName || typeof testName !== 'string' || testName.length < 3) {
      throw new Error('Назва тесту має бути не менше 3 символів');
    }
    if (!excelFile || !excelFile.match(/^questions(\d+)\.xlsx$/)) {
      throw new Error('Невірний формат файлу Excel');
    }
    if (!timeLimit || isNaN(timeLimit) || parseInt(timeLimit) < 60) {
      throw new Error('Час тесту має бути не менше 60 секунд');
    }

    const match = excelFile.match(/^questions(\d+)\.xlsx$/);
    const testNumber = match[1];
    if (testNames[testNumber]) throw new Error('Тест з таким номером вже існує');

    testNames[testNumber] = {
      name: testName,
      timeLimit: parseInt(timeLimit) || 3600
    };
    console.log('Created new test:', { testNumber, testName, timeLimit, excelFile });
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Тест створено</title>
        </head>
        <body>
          <h1>Новий тест "${testName}" створено</h1>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Ошибка при создании нового теста:', error.message, error.stack);
    res.status(400).send(`Помилка при створенні тесту: ${error.message}`);
  }
});

// Перегляд журналу дій (для адміна)
app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  let activities = [];
  let errorMessage = '';
  try {
    console.log('Fetching activity log from MongoDB...');
    activities = await db.collection('activity_log').find({}).sort({ timestamp: -1 }).toArray();
    console.log('Fetched activities from MongoDB:', activities);
  } catch (fetchError) {
    console.error('Ошибка при получении данных из MongoDB:', fetchError.message, fetchError.stack);
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
          .nav-btn, .clear-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; }
          .clear-btn { background-color: #ff4d4d; color: white; }
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
            <th>Користувач</th>
            <th>Дія</th>
            <th>IP-адреса</th>
            <th>Ідентифікатор сесії</th>
            <th>Час</th>
            <th>Дата</th>
          </tr>
  `;
  if (!activities || activities.length === 0) {
    adminHtml += '<tr><td colspan="6">Немає записів</td></tr>';
    console.log('No activities found in activity_log');
  } else {
    activities.forEach(activity => {
      const timestamp = new Date(activity.timestamp);
      const formattedTime = timestamp.toLocaleTimeString('uk-UA', { hour12: false });
      const formattedDate = timestamp.toLocaleDateString('uk-UA');
      adminHtml += `
        <tr>
          <td>${activity.user || 'N/A'}</td>
          <td>${activity.action || 'N/A'}</td>
          <td>${activity.ipAddress || 'N/A'}</td>
          <td>${activity.sessionId || 'N/A'}</td>
          <td>${formattedTime}</td>
          <td>${formattedDate}</td>
        </tr>
      `;
    });
  }
  adminHtml += `
        </table>
        <button class="clear-btn" onclick="clearActivityLog()">Видалення записів журналу</button>
        <script>
          async function clearActivityLog() {
            if (confirm('Ви впевнені, що хочете видалити усі записи журналу дій?')) {
              try {
                const response = await fetch('/admin/delete-activity-log', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ _csrf: "${res.locals.csrfToken}" })
                });
                const result = await response.json();
                if (result.success) {
                  console.log('Activity log cleared successfully');
                  window.location.reload();
                } else {
                  console.error('Failed to clear activity log:', result.message);
                  alert('Помилка при видаленні записів журналу: ' + result.message);
                }
              } catch (error) {
                console.error('Error clearing activity log:', error);
                alert('Помилка при видаленні записів журналу');
              }
            }
          }
        </script>
      </body>
    </html>
  `;
  res.send(adminHtml);
});

// Очищення журналу дій (для адміна)
app.post('/admin/delete-activity-log', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    console.log('Deleting all activity log entries...');
    await db.collection('activity_log').deleteMany({});
    console.log('Activity log cleared from MongoDB');
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении записей журнала действий:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
  }
});

// Запуск сервера
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

module.exports = app;
