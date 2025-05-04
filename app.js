require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { MongoClient, ObjectId } = require('mongodb');
const fs = require('fs');
const bcrypt = require('bcrypt');
const { v4: uuidv4 } = require('uuid');
const rateLimit = require('express-rate-limit');
const { body, validationResult } = require('express-validator');

// Ініціалізація додатка
const app = express();

// Налаштування trust proxy для Heroku
app.set('trust proxy', 1);

// Налаштування шаблонізатора
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'public'));

// Конфігурація MongoDB
const MONGO_URL = process.env.MONGO_URL || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/testdb?retryWrites=true&w=majority';
const client = new MongoClient(MONGO_URL, { connectTimeoutMS: 5000, serverSelectionTimeoutMS: 5000 });

// Підключення до MongoDB із повторними спробами
let db;
const connectToMongoDB = (attempt = 1, maxAttempts = 3) => {
  console.log(`Attempting to connect to MongoDB (Attempt ${attempt} of ${maxAttempts}) with URL:`, MONGO_URL);
  return client.connect()
    .then(() => {
      console.log('Connected to MongoDB successfully');
      db = client.db('testdb');
      console.log('Database initialized:', db.databaseName);
    })
    .catch(error => {
      console.error('Failed to connect to MongoDB:', error.message, error.stack);
      if (attempt < maxAttempts) {
        console.log(`Retrying MongoDB connection in 5 seconds...`);
        return new Promise(resolve => setTimeout(resolve, 5000))
          .then(() => connectToMongoDB(attempt + 1, maxAttempts));
      }
      throw error;
    });
};

// Глобальні змінні для ініціалізації
let isInitialized = false;
let initializationError = null;
let testNames = { 
  '1': { name: 'Тест 1', timeLimit: 3600 },
  '2': { name: 'Тест 2', timeLimit: 3600 },
  '3': { name: 'Тест 3', timeLimit: 3600 }
};
const userTests = new Map();

// Налаштування middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Обмеження кількості запитів (захист від DDoS)
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 хвилин
  max: 100, // Максимум 100 запитів на IP
  message: 'Занадто багато запитів з вашої IP-адреси. Спробуйте знову через 15 хвилин.'
});
app.use(limiter);

// Ініціалізація MongoStore для сесій
const sessionStore = MongoStore.create({
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
});

// Налаштування сесій
app.use(session({
  store: sessionStore,
  secret: process.env.SESSION_SECRET || 'a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8s9t0',
  resave: false,
  saveUninitialized: false,
  cookie: { 
    secure: true, // Для HTTPS на Heroku
    httpOnly: true,
    sameSite: 'none', // Змінено для крос-доменних запитів
    maxAge: 24 * 60 * 60 * 1000
  }
}));

// Middleware для генерації CSRF-токена
app.use((req, res, next) => {
  console.log('Session before CSRF middleware:', req.session);
  if (!req.session.csrfToken) {
    req.session.csrfToken = uuidv4();
    console.log('Generated new CSRF token:', req.session.csrfToken);
  } else {
    console.log('Using existing CSRF token:', req.session.csrfToken);
  }
  res.locals.csrfToken = req.session.csrfToken;
  next();
});

// Middleware для перевірки CSRF-токена
const checkCsrfToken = (req, res, next) => {
  if (['POST', 'PUT', 'DELETE'].includes(req.method)) {
    const token = req.body.csrfToken || req.headers['x-csrf-token'];
    if (!token || token !== req.session.csrfToken) {
      return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
    }
  }
  next();
};

// Middleware для дебаг-логування сесій
app.use((req, res, next) => {
  console.log(`Session ID: ${req.sessionID}, Session data before request:`, req.session);
  const originalSave = req.session.save;
  req.session.save = function (callback) {
    console.log(`Saving session ${req.sessionID} with data:`, req.session);
    originalSave.call(this, err => {
      if (err) {
        console.error(`Failed to save session ${req.sessionID}:`, err.message);
      } else {
        console.log(`Session ${req.sessionID} saved successfully`);
      }
      callback(err);
    });
  };
  next();
});

// Завантаження користувачів із хешуванням паролів
const loadUsers = () => {
  const filePath = path.join(__dirname, 'users.xlsx');
  console.log('Attempting to load users from:', filePath);

  if (!fs.existsSync(filePath)) {
    throw new Error(`File ${path.basename(filePath)} not found at path: ${filePath}`);
  }
  console.log(`File ${path.basename(filePath)} exists at:`, filePath);

  const workbook = new ExcelJS.Workbook();
  console.log(`Reading ${path.basename(filePath)} file...`);
  return workbook.xlsx.readFile(filePath)
    .then(() => {
      console.log('File read successfully');
      let sheet = workbook.getWorksheet('Users');
      if (!sheet) {
        console.warn('Worksheet "Users" not found, trying "Sheet1"');
        sheet = workbook.getWorksheet('Sheet1');
        if (!sheet) {
          console.error('Worksheet "Sheet1" not found in', path.basename(filePath));
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
            // Хешування пароля
            const hashedPassword = bcrypt.hashSync(password, 10);
            users[username] = hashedPassword;
          }
        }
      });
      if (Object.keys(users).length === 0) {
        console.error('No valid users found in', path.basename(filePath));
        throw new Error('Не найдено пользователей в файле');
      }
      console.log('Loaded users from Excel (passwords hashed):', Object.keys(users));
      return users;
    })
    .catch(error => {
      console.error('Error loading users from', path.basename(filePath), ':', error.message, error.stack);
      throw error;
    });
};

// Завантаження питань
const loadQuestions = (testNumber) => {
  const filePath = path.join(__dirname, `questions${testNumber}.xlsx`);
  console.log(`Attempting to load questions from: ${filePath}`);
  if (!fs.existsSync(filePath)) {
    console.error(`File ${path.basename(filePath)} not found at path: ${filePath}`);
    throw new Error(`File ${path.basename(filePath)} not found at path: ${filePath}`);
  }
  console.log(`File ${path.basename(filePath)} exists at: ${filePath}`);
  
  const workbook = new ExcelJS.Workbook();
  console.log(`Reading ${path.basename(filePath)} file...`);
  return workbook.xlsx.readFile(filePath)
    .then(() => {
      console.log('File read successfully');
      const sheet = workbook.getWorksheet('Questions');
      if (!sheet) {
        console.error(`Worksheet "Questions" not found in ${path.basename(filePath)}`);
        throw new Error(`Лист "Questions" не знайдено в ${path.basename(filePath)}`);
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
        console.error(`No questions loaded from ${path.basename(filePath)}`);
        throw new Error(`No questions found in ${path.basename(filePath)}`);
      }
      return jsonData;
    })
    .catch(error => {
      console.error(`Ошибка в loadQuestions (test ${testNumber}):`, error.message, error.stack);
      throw error;
    });
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

// Ініціалізація сервера
const initializeServer = () => {
  let attempt = 1;
  const maxAttempts = 5;

  return connectToMongoDB()
    .then(() => {
      while (attempt <= maxAttempts) {
        try {
          console.log(`Starting server initialization (Attempt ${attempt} of ${maxAttempts})...`);
          return loadUsers().then(users => {
            console.log('Users initialized successfully from Excel');
            isInitialized = true;
            initializationError = null;
            app.use(ensureInitialized);
            return users;
          });
        } catch (err) {
          console.error(`Failed to initialize server (Attempt ${attempt}):`, err.message, err.stack);
          initializationError = err;
          if (attempt < maxAttempts) {
            console.log(`Retrying initialization in 5 seconds...`);
            return new Promise(resolve => setTimeout(resolve, 5000))
              .then(() => {
                attempt++;
                return initializeServer();
              });
          }
          console.error('Maximum initialization attempts reached. Server remains uninitialized.');
          throw err;
        }
      }
    })
    .catch(error => {
      console.error('Failed to start server due to initialization error:', error.message, error.stack);
      process.exit(1);
    });
};

// Запуск ініціалізації сервера
initializeServer();

// Тестовий маршрут для перевірки MongoDB
app.get('/test-mongo', (req, res) => {
  console.log('Testing MongoDB connection...');
  if (!db) {
    return res.status(500).json({ success: false, message: 'MongoDB connection not established' });
  }
  return db.collection('users').findOne()
    .then(() => {
      console.log('MongoDB test successful');
      res.json({ success: true, message: 'MongoDB connection successful' });
    })
    .catch(error => {
      console.error('MongoDB test failed:', error.message, error.stack);
      res.status(500).json({ success: false, message: 'MongoDB connection failed', error: error.message });
    });
});

// Тестовий маршрут з префіксом /api
app.get('/api/test', (req, res) => {
  console.log('Handling /api/test request...');
  res.json({ success: true, message: 'Express server is working on /api/test' });
});

// Головна сторінка
app.get('/', (req, res) => {
  console.log('Serving index.html');
  console.log('CSRF token for rendering:', res.locals.csrfToken);
  res.render('index', { csrfToken: res.locals.csrfToken });
});

// Логування дій користувача
const logActivity = (req, user, action) => {
  const timestamp = new Date();
  const timeOffset = 3 * 60 * 60 * 1000; // 3 години в мілісекундах (UTC+3)
  const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
  const ipAddress = req.headers['x-forwarded-for'] || req.ip || 'N/A';
  const sessionId = req.sessionID || 'N/A';

  return db.collection('activity_log').insertOne({
    user,
    action,
    ipAddress,
    sessionId,
    timestamp: adjustedTimestamp.toISOString()
  })
    .then(() => {
      console.log(`Logged activity: ${user} - ${action} at ${adjustedTimestamp}, IP: ${ipAddress}, SessionID: ${sessionId}`);
    })
    .catch(error => {
      console.error('Error logging activity:', error.message, error.stack);
    });
};

// Валідація введення для логіну
const validateLogin = [
  body('password')
    .trim()
    .notEmpty().withMessage('Пароль не вказано')
    .isLength({ min: 6, max: 20 }).withMessage('Пароль має бути від 6 до 20 символів')
    .matches(/^[a-zA-Z0-9]+$/).withMessage('Пароль має містити лише латинські літери та цифри')
];

// Маршрут для логіну
app.post('/login', checkCsrfToken, validateLogin, (req, res) => {
  console.log('Handling /login request...');
  console.log('Request body:', req.body);
  console.log('Session CSRF token:', req.session.csrfToken);
  console.log('Session ID:', req.sessionID);
  console.log('Cookies:', req.cookies);
  
  const { password } = req.body;

  return loadUsers()
    .then(validUsers => {
      console.log('Checking password against valid users...');
      const user = Object.keys(validUsers).find(u => {
        const match = bcrypt.compareSync(password, validUsers[u]);
        console.log(`Comparing ${u} with provided password -> ${match}`);
        return match;
      });

      if (!user) {
        console.warn('Password not found in validUsers');
        return res.status(401).json({ success: false, message: 'Невірний пароль' });
      }

      req.session.user = user;
      return logActivity(req, user, 'увійшов на сайт')
        .then(() => {
          console.log('Session after setting user:', req.session);
          console.log('Session ID after setting user:', req.sessionID);
          console.log('Cookies after setting session:', req.cookies);

          return new Promise((resolve, reject) => {
            req.session.save(err => {
              if (err) {
                console.error('Error saving session in /login:', err.message, err.stack);
                return reject(err);
              }
              console.log('Session saved successfully');
              resolve();
            });
          });
        })
        .then(() => {
          if (user === 'admin') {
            console.log('Redirecting to /admin for user:', user);
            res.json({ success: true, redirect: '/admin' });
          } else {
            console.log('Redirecting to /select-test for user:', user);
            res.json({ success: true, redirect: '/select-test' });
          }
        });
    })
    .catch(error => {
      console.error('Ошибка в /login:', error.message, error.stack);
      res.status(500).json({ success: false, message: 'Помилка сервера' });
    });
});

// Middleware для перевірки авторизації
const checkAuth = (req, res, next) => {
  console.log('checkAuth: Session data:', req.session);
  console.log('checkAuth: Cookies:', req.cookies);
  console.log('checkAuth: Session ID:', req.sessionID);
  console.log('checkAuth: Cookie connect.sid:', req.cookies['connect.sid']);
  const user = req.session.user;
  console.log('checkAuth: user from session:', user);
  if (!user) {
    console.log('checkAuth: No valid auth, redirecting to /');
    return res.redirect('/');
  }
  req.user = user;
  next();
};

// Middleware для перевірки адміністратора
const checkAdmin = (req, res, next) => {
  const user = req.session.user;
  console.log('checkAdmin: user from session:', user);
  if (user !== 'admin') {
    console.log('checkAdmin: Not admin, returning 403');
    return res.status(403).send('Доступно тільки для адміністратора (403 Forbidden)');
  }
  next();
};

// Вибір тесту
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
        <form id="logout-form" action="/logout" method="POST">
          <input type="hidden" name="csrfToken" value="${res.locals.csrfToken}">
          <button type="submit" id="logout">Вийти</button>
        </form>
        <script>
          document.getElementById('logout-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            await fetch('/logout', { 
              method: 'POST',
              headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
              body: new URLSearchParams(new FormData(e.target)).toString(),
              credentials: 'same-origin'
            });
            window.location.href = '/';
          });
        </script>
      </body>
    </html>
  `);
});

// Маршрут для виходу з системи
app.post('/logout', checkCsrfToken, (req, res) => {
  const user = req.session.user;
  const userTest = userTests.get(user);
  if (userTest) {
    logActivity(req, user, `покинув сайт не завершивши тест ${testNames[userTest.testNumber].name}`);
    userTests.delete(user);
  } else {
    logActivity(req, user, `покинув сайт`);
  }

  req.session.destroy(err => {
    if (err) {
      console.error('Error destroying session:', err);
      return res.status(500).json({ success: false, message: 'Помилка при виході' });
    }
    res.json({ success: true });
  });
});

// Збереження результатів тесту
const saveResult = (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage) => {
  const duration = Math.round((endTime - startTime) / 1000);
  const userTest = userTests.get(user);
  const answers = userTest ? userTest.answers : {};
  const questions = userTest ? userTest.questions : [];
  const suspiciousActivity = userTest ? userTest.suspiciousActivity : { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };

  // Розрахунок балів за питання
  const scoresPerQuestion = questions.map((q, index) => {
    const userAnswer = answers[index];
    let questionScore = 0;
    if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
      const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
      const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\\'/g, "'"));
      const isCorrect = correctAnswers.length === userAnswers.length &&
        correctAnswers.every(val => userAnswers.includes(val)) &&
        userAnswers.every(val => correctAnswers.includes(val));
      if (isCorrect) questionScore = q.points;
    } else if (q.type === 'input' && userAnswer) {
      const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      if (normalizedUserAnswer === normalizedCorrectAnswer) questionScore = q.points;
    } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
      const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
      const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
      if (userAnswers.join(',') === correctAnswers.join(',')) questionScore = q.points;
    }
    return questionScore;
  });

  const calculatedScore = scoresPerQuestion.reduce((sum, s) => sum + s, 0);

  // Розрахунок підозрілої активності
  let suspiciousScore = 0;
  const timeAwayPercent = suspiciousActivity.timeAway ? 
    Math.round((suspiciousActivity.timeAway / (duration * 1000)) * 100) : 0;
  suspiciousScore += timeAwayPercent;

  const switchCount = suspiciousActivity.switchCount || 0;
  if (switchCount > totalQuestions * 2) suspiciousScore += 20;

  const responseTimes = suspiciousActivity.responseTimes || [];
  const avgResponseTime = responseTimes.length > 0 ? 
    (responseTimes.reduce((sum, time) => sum + (time || 0), 0) / responseTimes.length / 1000).toFixed(2) : 0;
  responseTimes.forEach(time => {
    if (time < 5000) suspiciousScore += 10;
    else if (time > 5 * 60 * 1000) suspiciousScore += 10;
  });

  const activityCounts = suspiciousActivity.activityCounts || [];
  const avgActivityCount = activityCounts.length > 0 ? 
    (activityCounts.reduce((sum, count) => sum + (count || 0), 0) / activityCounts.length).toFixed(2) : 0;
  activityCounts.forEach((count, idx) => {
    if (count < 5 && responseTimes[idx] > 30 * 1000) suspiciousScore += 10;
  });

  let typicalResponseTime = 30 * 1000;
  let typicalSwitchCount = totalQuestions;
  return db.collection('test_results').find({}).toArray()
    .then(allResults => {
      if (allResults.length > 0) {
        const allResponseTimes = allResults.flatMap(r => r.suspiciousActivity.responseTimes || []);
        typicalResponseTime = allResponseTimes.length > 0 ? 
          allResponseTimes.reduce((sum, time) => sum + (time || 0), 0) / allResponseTimes.length : typicalResponseTime;
        const allSwitchCounts = allResults.map(r => r.suspiciousActivity.switchCount || 0);
        typicalSwitchCount = allSwitchCounts.length > 0 ? 
          allSwitchCounts.reduce((sum, count) => sum + count, 0) / allSwitchCounts.length : typicalSwitchCount;
      }

      if (avgResponseTime < typicalResponseTime * 0.5 || avgResponseTime > typicalResponseTime * 1.5) suspiciousScore += 15;
      if (switchCount > typicalSwitchCount * 1.5) suspiciousScore += 15;

      suspiciousScore = Math.min(suspiciousScore, 100);

      const timeOffset = 3 * 60 * 60 * 1000;
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
        scoresPerQuestion,
        suspiciousActivity: {
          ...suspiciousActivity,
          suspiciousScore,
          responseTimes: responseTimes.map(time => Math.min(time, 5 * 60 * 1000))
        }
      };

      if (!db) throw new Error('MongoDB connection not established');
      return db.collection('test_results').insertOne(result)
        .then(insertResult => {
          console.log(`Successfully saved result for ${user} in MongoDB with ID:`, insertResult.insertedId);
        });
    })
    .catch(error => {
      console.error('Error calculating typical behavior or saving result:', error);
      throw error;
    });
};

// Початок тесту
app.get('/test', checkAuth, (req, res) => {
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

  return loadQuestions(testNumber)
    .then(questions => {
      userTests.set(req.user, {
        testNumber,
        questions,
        answers: {},
        currentQuestion: 0,
        startTime: Date.now(),
        timeLimit: testNames[testNumber].timeLimit * 1000
      });
      return logActivity(req, req.user, `розпочав тест ${testNames[testNumber].name}`)
        .then(() => {
          console.log(`Redirecting to first question for user ${req.user}`);
          res.redirect(`/test/question?index=0`);
        });
    })
    .catch(error => {
      console.error('Ошибка в /test:', error.message, error.stack);
      res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
    });
});

// Відображення питання тесту
app.get('/test/question', checkAuth, (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  const userTest = userTests.get(req.user);
  if (!userTest) {
    console.warn(`Test not started for user ${req.user}`);
    return res.status(400).send('Тест не розпочато');
  }

  const { questions, testNumber, answers, currentQuestion, startTime, timeLimit } = userTest;
  const index = parseInt(req.query.index) || 0;

  if (index < 0 || index >= questions.length) {
    console.warn(`Invalid question index ${index} for user ${req.user}`);
    return res.status(400).send('Невірний номер питання');
  }

  userTest.currentQuestion = index;
  userTest.answerTimestamps = userTest.answerTimestamps || {};
  userTest.answerTimestamps[index] = userTest.answerTimestamps[index] || Date.now();
  const q = questions[index];
  console.log('Rendering question:', { index, picture: q.picture, text: q.text, options: q.options });

  const progress = Array.from({ length: questions.length }, (_, i) => ({
    number: i + 1,
    answered: answers[i] && (Array.isArray(answers[i]) ? answers[i].length > 0 : answers[i].trim() !== '')
  }));

  const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
  const remainingTime = Math.max(0, Math.floor(timeLimit / 1000) - elapsedTime);
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
        <script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.0/Sortable.min.js" onerror="console.error('Failed to load Sortable.js')"></script>
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 20px; padding-bottom: 80px; background-color: #f0f0f0; }
          h1 { font-size: 24px; text-align: center; }
          img { max-width: 100%; margin-bottom: 10px; display: block; margin-left: auto; margin-right: auto; }
          .progress-bar { 
            display: flex; 
            flex-direction: column; 
            gap: 5px; 
            margin-bottom: 20px; 
            width: calc(100% - 40px); 
            margin-left: auto; 
            margin-right: auto; 
            box-sizing: border-box; 
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
          .progress-circle.unanswered { background-color: red; color: white; }
          .progress-circle.answered { background-color: green; color: white; }
          .progress-line { 
            width: 5px; 
            height: 2px; 
            background-color: #ccc; 
            margin: 0 2px; 
            align-self: center; 
            flex-shrink: 0; 
          }
          .progress-line.answered { background-color: green; }
          .progress-row { 
            display: flex; 
            align-items: center; 
            justify-content: space-around; 
            gap: 2px; 
            flex-wrap: wrap; 
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
          #confirm-modal { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border: 2px solid black; z-index: 1000; }
          #confirm-modal button { margin: 0 10px; }
          .question-box { padding: 10px; margin: 5px 0; }
          .instruction { font-style: italic; color: #555; margin-bottom: 10px; font-size: 18px; }
          .option-box.draggable { cursor: move; }
          .option-box.dragging { opacity: 0.5; }
          #question-container { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0, 0, 0, 0.1); width: calc(100% - 40px); margin: 0 auto 20px auto; box-sizing: border-box; }
          #answers { margin-bottom: 20px; }
          @media (max-width: 600px) {
            h1 { font-size: 28px; }
            .progress-bar { flex-direction: column; }
            .progress-circle { width: 20px; height: 20px; font-size: 10px; }
            .progress-line { width: 5px; }
            .progress-row { justify-content: center; gap: 2px; flex-wrap: wrap; }
            .option-box { font-size: 18px; padding: 15px; }
            button { font-size: 18px; padding: 15px; }
            #timer { font-size: 20px; }
            .question-box h2 { font-size: 20px; }
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
  if (q.picture) {
    html += `<img src="${q.picture}" alt="Picture" onerror="this.src='/images/placeholder.png'; console.log('Image failed to load: ${q.picture}')"><br>`;
  }

  const instructionText = q.type === 'multiple' ? 'Виберіть усі правильні відповіді' :
                         q.type === 'input' ? 'Введіть правильну відповідь' :
                         q.type === 'ordering' ? 'Розташуйте відповіді у правильній послідовності' : '';
  html += `
          <div class="question-box">
            <h2 id="question-text">${index + 1}. ${q.text}</h2>
          </div>
          <p id="instruction" class="instruction">${instructionText}</p>
          <div id="answers">
  `;

  if (!q.options || q.options.length === 0) {
    const userAnswer = answers[index] || '';
    html += `
      <input type="text" name="q${index}" id="q${index}_input" value="${userAnswer}" placeholder="Введіть відповідь" class="answer-option"><br>
    `;
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
          <button class="back-btn" ${index === 0 ? 'disabled' : ''} onclick="window.location.href='/test/question?index=${index - 1}'">Назад</button>
          <button id="submit-answer" class="next-btn" ${index === questions.length - 1 ? 'disabled' : ''} onclick="saveAndNext(${index})">Далі</button>
          <button class="finish-btn" onclick="showConfirm(${index})">Завершити тест</button>
        </div>
        <div id="confirm-modal">
          <h2>Ви дійсно бажаєте завершити тест?</h2>
          <button onclick="finishTest(${index})">Так</button>
          <button onclick="hideConfirm()">Ні</button>
        </div>
        <script>
          let startTime = ${startTime};
          let timeLimit = ${timeLimit};
          const timerElement = document.getElementById('timer');
          let timeAway = 0;
          let lastBlurTime = 0;
          let switchCount = 0;
          let lastActivityTime = Date.now();
          let activityCount = 0;
          let lastMouseMoveTime = 0;
          const debounceDelay = 100;
          const questionStartTime = ${userTest.answerTimestamps[index] || Date.now()};
          let selectedOptions = ${selectedOptionsString};
          const csrfToken = "${res.locals.csrfToken}";

          function updateTimer() {
            const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
            const remainingTime = Math.max(0, Math.floor(timeLimit / 1000) - elapsedTime);
            const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
            const seconds = (remainingTime % 60).toString().padStart(2, '0');
            timerElement.textContent = 'Залишилось часу: ' + minutes + ' мм ' + seconds + ' с';
            if (remainingTime <= 0) {
              window.location.href = '/result';
            }
          }
          updateTimer();
          setInterval(updateTimer, 1000);

          window.addEventListener('blur', () => {
            lastBlurTime = Date.now();
            switchCount++;
            console.log('Tab blurred, switch count:', switchCount);
          });

          window.addEventListener('focus', () => {
            if (lastBlurTime) {
              const timeSpentAway = Date.now() - lastBlurTime;
              timeAway += timeSpentAway;
              console.log('Tab focused, time away:', timeAway);
            }
          });

          function debounceMouseMove() {
            const now = Date.now();
            if (now - lastMouseMoveTime >= debounceDelay) {
              lastMouseMoveTime = now;
              lastActivityTime = now;
              activityCount++;
              console.log('Mouse activity detected, count:', activityCount);
            }
          }

          document.addEventListener('mousemove', debounceMouseMove);

          document.addEventListener('keydown', () => {
            lastActivityTime = Date.now();
            activityCount++;
            console.log('Keyboard activity detected, count:', activityCount);
          });

          document.querySelectorAll('.option-box:not(.draggable)').forEach(box => {
            box.addEventListener('click', () => {
              const option = box.getAttribute('data-value');
              const idx = selectedOptions.indexOf(option);
              if (idx === -1) {
                selectedOptions.push(option);
                box.classList.add('selected');
              } else {
                selectedOptions.splice(idx, 1);
                box.classList.remove('selected');
              }
              console.log('Selected options for question ${index}:', selectedOptions);
            });
          });

          async function saveAndNext(index) {
            console.log('Save and Next button clicked for index:', index);
            try {
              let answers = selectedOptions;
              console.log('Selected options:', selectedOptions);
              if (document.querySelector('input[name="q' + index + '"]')) {
                answers = document.getElementById('q' + index + '_input').value;
                console.log('Input answer:', answers);
              } else if (document.getElementById('sortable-options')) {
                answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
                console.log('Sortable options answer:', answers);
              }
              const responseTime = Date.now() - questionStartTime;
              console.log('Sending answer with data:', { index, answers, timeAway, switchCount, responseTime, activityCount });
              const response = await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount, csrfToken }),
                credentials: 'same-origin'
              });
              console.log('Response status:', response.status);
              if (response.status === 302 || response.redirected) {
                console.warn('Session expired, redirecting to login');
                window.location.href = '/';
                return;
              }
              const result = await response.json();
              console.log('Response result:', result);
              if (result.success) {
                console.log('Answer saved successfully, redirecting to next question');
                window.location.href = '/test/question?index=' + (index + 1);
              } else {
                console.error('Failed to save answer:', result);
                alert('Помилка при збереженні відповіді: ' + (result.error || 'Невідома помилка'));
              }
            } catch (error) {
              console.error('Error in saveAndNext:', error);
              alert('Помилка при збереженні відповіді: ' + error.message);
            }
          }

          function showConfirm(index) {
            document.getElementById('confirm-modal').style.display = 'block';
          }

          function hideConfirm() {
            document.getElementById('confirm-modal').style.display = 'none';
          }

          async function finishTest(index) {
            console.log('Finish Test button clicked for index:', index);
            try {
              let answers = selectedOptions;
              console.log('Selected options:', selectedOptions);
              if (document.querySelector('input[name="q' + index + '"]')) {
                answers = document.getElementById('q' + index + '_input').value;
                console.log('Input answer:', answers);
              } else if (document.getElementById('sortable-options')) {
                answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
                console.log('Sortable options answer:', answers);
              }
              const responseTime = Date.now() - questionStartTime;
              console.log('Finishing test with data:', { index, answers, timeAway, switchCount, responseTime, activityCount });
              const response = await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount, csrfToken }),
                credentials: 'same-origin'
              });
              console.log('Response status:', response.status);
              if (response.status === 302 || response.redirected) {
                console.warn('Session expired, redirecting to login');
                window.location.href = '/';
                return;
              }
              const result = await response.json();
              console.log('Response result:', result);
              if (result.success) {
                console.log('Answer saved successfully, redirecting to result');
                hideConfirm();
                window.location.href = '/result';
              } else {
                console.error('Failed to save answer:', result);
                alert('Помилка при завершенні тесту: ' + (result.error || 'Невідома помилка'));
              }
            } catch (error) {
              console.error('Error in finishTest:', error);
              alert('Помилка при завершенні тесту: ' + error.message);
            }
          }

          const sortable = document.getElementById('sortable-options');
          if (sortable) {
            if (typeof Sortable === 'undefined') {
              console.error('Sortable.js is not loaded');
            } else {
              new Sortable(sortable, {
                animation: 150,
                onStart: function(evt) {
                  console.log('Drag started on:', evt.item.dataset.value);
                },
                onEnd: function(evt) {
                  console.log('Drag ended:', evt.item.dataset.value, 'from index:', evt.oldIndex, 'to index:', evt.newIndex);
                }
              });
            }
          } else {
            console.log('Sortable options not found');
          }
        </script>
      </body>
    </html>
  `;
  res.send(html);
});

// Збереження відповіді на питання
app.post('/answer', checkCsrfToken, checkAuth, (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;
  console.log('Received answer data:', { index, answer, timeAway, switchCount, responseTime, activityCount });
  const userTest = userTests.get(req.user);
  if (!userTest) {
    console.warn(`Test not started for user ${req.user} in /answer`);
    return res.status(400).json({ success: false, error: 'Тест не розпочато' });
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
});

// Відображення результатів тесту
app.get('/result', checkAuth, (req, res) => {
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
      if (isCorrect) questionScore = q.points;
    } else if (q.type === 'input' && userAnswer) {
      const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
      if (normalizedUserAnswer === normalizedCorrectAnswer) questionScore = q.points;
    } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
      const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
      const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
      if (userAnswers.join(',') === correctAnswers.join(',')) questionScore = q.points;
    }
    return questionScore;
  });

  score = scoresPerQuestion.reduce((sum, s) => sum + s, 0);
  const endTime = Date.now();
  const percentage = (score / totalPoints) * 100;
  const totalClicks = Object.keys(answers).length;
  const correctClicks = scoresPerQuestion.filter(s => s > 0).length;
  const totalQuestions = questions.length;

  return saveResult(req.user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage)
    .then(() => logActivity(req, req.user, `закінчив тест ${testNames[testNumber].name} з результатом ${score} балів`))
    .then(() => {
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
    })
    .catch(error => {
      console.error('Error saving result in /result:', error.message, error.stack);
      res.status(500).send('Помилка при збереженні результату');
    });
});

// Детальні результати
app.get('/results', checkAuth, (req, res) => {
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
        const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        if (normalizedUserAnswer === normalizedCorrectAnswer) questionScore = q.points;
      } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(String).map(val => val.trim().toLowerCase());
        const correctAnswers = q.correctAnswers.map(String).map(val => val.trim().toLowerCase());
        if (userAnswers.join(',') === correctAnswers.join(',')) questionScore = q.points;
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

// Адмін-панель
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
        <form id="logout-form" action="/logout" method="POST">
          <input type="hidden" name="csrfToken" value="${res.locals.csrfToken}">
          <button type="submit" id="logout">Вийти</button>
        </form>
        <script>
          document.getElementById('logout-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            await fetch('/logout', { 
              method: 'POST',
              headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
              body: new URLSearchParams(new FormData(e.target)).toString(),
              credentials: 'same-origin'
            });
            window.location.href = '/';
          });
        </script>
      </body>
    </html>
  `);
});

// Перегляд результатів у адмін-панелі
app.get('/admin/results', checkAuth, checkAdmin, (req, res) => {
  let results = [];
  let errorMessage = '';
  return db.collection('test_results').find({}).sort({ endTime: -1 }).toArray()
    .then(fetchedResults => {
      results = fetchedResults;
      console.log('Fetched results from MongoDB:', results);
      
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

          const answersDisplay = answersArray.length > 0 
            ? answersArray.map((a, i) => {
                if (!a) return null;
                const userAnswer = Array.isArray(a) ? a.join(', ') : a;
                const questionScore = r.scoresPerQuestion[i] || 0;
                console.log(`Result ${index + 1}, Question ${i + 1}: userAnswer=${userAnswer}, savedScore=${questionScore}`);
                return `Питання ${i + 1}: ${userAnswer.replace(/\\'/g, "'")} (${questionScore} балів)`;
              }).filter(line => line !== null).join('\n')
            : 'Немає відповідей';
          const formatDateTime = (isoString) => {
            if (!isoString) return 'N/A';
            const date = new Date(isoString);
            return `${date.toLocaleTimeString('uk-UA', { hour12: false })} ${date.toLocaleDateString('uk-UA')}`;
          };
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
                      body: JSON.stringify({ id, csrfToken: "${res.locals.csrfToken}" }),
                      credentials: 'same-origin'
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
    })
    .catch(fetchError => {
      console.error('Ошибка при получении данных из MongoDB:', fetchError.message, fetchError.stack);
      errorMessage = `Ошибка MongoDB: ${fetchError.message}`;
      res.status(500).send(`Помилка при отриманні результатів: ${errorMessage}`);
    });
});

// Видалення одного результату
app.post('/admin/delete-result', checkCsrfToken, checkAuth, checkAdmin, (req, res) => {
  const { id } = req.body;
  console.log(`Deleting result with id ${id}...`);
  return db.collection('test_results').deleteOne({ _id: new ObjectId(id) })
    .then(() => {
      console.log(`Result with id ${id} deleted from MongoDB`);
      res.json({ success: true });
    })
    .catch(error => {
      console.error('Ошибка при удалении результата:', error.message, error.stack);
      res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
    });
});

// Видалення всіх результатів
app.get('/admin/delete-results', checkAuth, checkAdmin, (req, res) => {
  console.log('Deleting all test results...');
  return db.collection('test_results').deleteMany({})
    .then(() => {
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
    })
    .catch(error => {
      console.error('Ошибка при удалении результатов:', error.message, error.stack);
      res.status(500).send('Помилка при видаленні результатів');
    });
});

// Редагування назв тестів
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
          <input type="hidden" name="csrfToken" value="${res.locals.csrfToken}">
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
                body: JSON.stringify({ testNumber, csrfToken: "${res.locals.csrfToken}" }),
                credentials: 'same-origin'
              });
              window.location.reload();
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Збереження змін у назвах тестів
app.post('/admin/edit-tests', checkCsrfToken, checkAuth, checkAdmin, (req, res) => {
  console.log('Updating test names and time limits...');
  Object.keys(testNames).forEach(num => {
    const testName = req.body[`test${num}`];
    const timeLimit = req.body[`time${num}`];
    if (testName && timeLimit) {
      testNames[num] = {
        name: testName,
        timeLimit: parseInt(timeLimit) || testNames[num].timeLimit
      };
    }
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
});

// Видалення тесту
app.post('/admin/delete-test', checkCsrfToken, checkAuth, checkAdmin, (req, res) => {
  const { testNumber } = req.body;
  if (!testNames[testNumber]) {
    return res.status(404).json({ success: false, message: 'Тест не знайдено' });
  }
  delete testNames[testNumber];
  console.log(`Deleted test ${testNumber}, updated testNames:`, testNames);
  res.json({ success: true });
});

// Створення нового тесту
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
          <input type="hidden" name="csrfToken" value="${res.locals.csrfToken}">
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

// Збереження нового тесту
app.post('/admin/create-test', checkCsrfToken, checkAuth, checkAdmin, (req, res) => {
  const { testName, excelFile, timeLimit } = req.body;
  const match = excelFile.match(/^questions(\d+)\.xlsx$/);
  if (!match) {
    return res.status(400).send('Невірний формат файлу Excel');
  }
  const testNumber = match[1];
  if (testNames[testNumber]) {
    return res.status(400).send('Тест з таким номером вже існує');
  }

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
});

// Журнал дій
app.get('/admin/activity-log', checkAuth, checkAdmin, (req, res) => {
  let activities = [];
  let errorMessage = '';
  return db.collection('activity_log').find({}).sort({ timestamp: -1 }).toArray()
    .then(fetchedActivities => {
      activities = fetchedActivities;
      console.log('Fetched activities from MongoDB:', activities);

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
                      body: JSON.stringify({ csrfToken: "${res.locals.csrfToken}" }),
                      credentials: 'same-origin'
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
    })
    .catch(fetchError => {
      console.error('Ошибка при получении данных из MongoDB:', fetchError.message, fetchError.stack);
      errorMessage = `Ошибка MongoDB: ${fetchError.message}`;
      res.status(500).send(`Помилка при отриманні журналу дій: ${errorMessage}`);
    });
});

// Видалення всіх записів журналу дій
app.post('/admin/delete-activity-log', checkCsrfToken, checkAuth, checkAdmin, (req, res) => {
  console.log('Deleting all activity log entries...');
  return db.collection('activity_log').deleteMany({})
    .then(() => {
      console.log('Activity log cleared from MongoDB');
      res.json({ success: true });
    })
    .catch(error => {
      console.error('Ошибка при удалении записей журнала действий:', error.message, error.stack);
      res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
    });
});

module.exports = app;
