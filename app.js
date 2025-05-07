require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { MongoClient, ObjectId } = require('mongodb');
const bcrypt = require('bcrypt');
const fs = require('fs');
const multer = require('multer');
const nodemailer = require('nodemailer');

const app = express();

app.set('trust proxy', 1); // Довіряємо проксі Heroku

// Налаштування multer для завантаження файлів
const upload = multer({ dest: 'uploads/' });

// Налаштування nodemailer для відправки email
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'alphacentertest@gmail.com',
    pass: 'your-app-specific-password' // Замініть на пароль додатка Gmail
  }
});

const sendSuspiciousActivityEmail = async (user, activityDetails) => {
  try {
    const mailOptions = {
      from: 'alphacentertest@gmail.com',
      to: 'alphacentertest@gmail.com',
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
    console.log(`Email про підозрілу активність відправлено для користувача ${user}`);
  } catch (error) {
    console.error('Error sending suspicious activity email:', error.message, error.stack);
  }
};

// Підключення до MongoDB
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/alpha?retryWrites=true&w=majority';
const client = new MongoClient(MONGODB_URI, { connectTimeoutMS: 5000, serverSelectionTimeoutMS: 5000 });
let db;

// Кэш пользователей и вопросов
let userCache = [];
const questionsCache = {};

const connectToMongoDB = async (attempt = 1, maxAttempts = 3) => {
  try {
    console.log(`Attempting to connect to MongoDB (Attempt ${attempt} of ${maxAttempts}) with URI:`, MONGODB_URI);
    const startTime = Date.now();
    await client.connect();
    const endTime = Date.now();
    console.log('Connected to MongoDB successfully in', endTime - startTime, 'ms');
    db = client.db('alpha');
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

let isInitialized = false;
let initializationError = null;
let testNames = {
  '1': { name: 'Тест 1', timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null },
  '2': { name: 'Тест 2', timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null },
  '3': { name: 'Тест 3', timeLimit: 3600, randomQuestions: false, randomAnswers: false, questionLimit: null }
};

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(cookieParser());

// Middleware для запобігання кешуванню
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Використовуємо MongoStore для сесій
app.use(session({
  store: MongoStore.create({
    mongoUrl: MONGODB_URI,
    collectionName: 'sessions',
    ttl: 24 * 60 * 60,
    autoRemove: 'interval',
    autoRemoveInterval: 10
  }),
  secret: process.env.SESSION_SECRET || 'a1b2c3d4e5f6g7h8i9j0k1l2m3n4o5p6q7r8s9t0',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env.NODE_ENV === 'production' ? true : false,
    httpOnly: true,
    sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax',
    maxAge: 24 * 60 * 60 * 1000
  },
  name: 'connect.sid'
}));

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
      if (username && password) {
        const hashedPassword = await bcrypt.hash(password, saltRounds);
        users.push({ username, password: hashedPassword });
      }
    }
    if (users.length === 0) {
      throw new Error('Не знайдено користувачів у файлі');
    }
    await db.collection('users').deleteMany({});
    await db.collection('sessions').deleteMany({});
    console.log('Cleared all sessions after user import');
    await db.collection('users').insertMany(users);
    console.log(`Imported ${users.length} users to MongoDB with hashed passwords`);
    // Обновляем кэш
    userCache = users;
    return users.length;
  } catch (error) {
    console.error('Error importing users to MongoDB:', error.message, error.stack);
    throw error;
  }
};

const importQuestionsToMongoDB = async (filePath, testNumber) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      throw new Error('Лист "Questions" не знайдено у файлі');
    }
    const questions = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        const rowValues = row.values.slice(1);
        let questionText = rowValues[1];
        if (typeof questionText === 'object' && questionText !== null) {
          questionText = questionText.text || questionText.value || '[Невірний текст питання]';
        }
        questionText = String(questionText || '').trim();
        if (questionText === '') return;
        const picture = String(rowValues[0] || '').trim();
        let options = rowValues.slice(2, 14).filter(Boolean).map(val => String(val).trim());
        const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(val => String(val).trim());
        const type = String(rowValues[26] || 'multiple').toLowerCase();
        const points = Number(rowValues[27]) || 1;
        const variant = String(rowValues[28] || '').trim();

        if (type === 'truefalse') {
          options = ["Правда", "Неправда"];
        }

        let questionData = {
          testNumber,
          picture: picture.match(/^Picture (\d+)/i) ? `/images/Picture ${picture.match(/^Picture (\d+)/i)[1]}.png` : null,
          text: questionText,
          options,
          correctAnswers,
          type,
          points,
          variant
        };

        if (type === 'matching') {
          questionData.pairs = options.map((opt, idx) => ({
            left: opt || '',
            right: correctAnswers[idx] || ''
          })).filter(pair => pair.left && pair.right);
          if (questionData.pairs.length === 0) return;
          questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
        }

        if (type === 'fillblank') {
          questionText = questionText.replace(/\s*___\s*/g, '___');
          const blankCount = (questionText.match(/___/g) || []).length;
          if (blankCount === 0 || blankCount !== correctAnswers.length) return;
          questionData.text = questionText;
          questionData.blankCount = blankCount;
        }

        if (type === 'singlechoice') {
          if (correctAnswers.length !== 1 || options.length < 2) return;
          questionData.correctAnswer = correctAnswers[0];
        }

        questions.push(questionData);
      }
    });
    if (questions.length === 0) {
      throw new Error('Не знайдено питань у файлі');
    }
    await db.collection('questions').deleteMany({ testNumber });
    await db.collection('questions').insertMany(questions);
    console.log(`Imported ${questions.length} questions for test ${testNumber} to MongoDB`);
    // Обновляем кэш
    questionsCache[testNumber] = questions;
    return questions.length;
  } catch (error) {
    console.error('Error importing questions to MongoDB:', error.message, error.stack);
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

const loadUsersToCache = async () => {
  try {
    const startTime = Date.now();
    userCache = await db.collection('users').find({}).toArray();
    const endTime = Date.now();
    console.log(`Loaded ${userCache.length} users to cache in ${endTime - startTime} ms`);
  } catch (error) {
    console.error('Error loading users to cache:', error.message, error.stack);
  }
};

const loadQuestions = async (testNumber) => {
  try {
    const startTime = Date.now();
    // Проверяем кэш
    if (questionsCache[testNumber]) {
      const endTime = Date.now();
      console.log(`Loaded ${questionsCache[testNumber].length} questions for test ${testNumber} from cache in ${endTime - startTime} ms`);
      return questionsCache[testNumber];
    }

    const questions = await db.collection('questions').find({ testNumber: testNumber.toString() }).toArray();
    const endTime = Date.now();
    if (questions.length === 0) {
      throw new Error(`No questions found in MongoDB for test ${testNumber}`);
    }
    questionsCache[testNumber] = questions; // Сохраняем в кэш
    console.log(`Loaded ${questions.length} questions for test ${testNumber} from MongoDB in ${endTime - startTime} ms`);
    return questions;
  } catch (error) {
    console.error(`Ошибка в loadQuestions (test ${testNumber}):`, error.message, error.stack);
    throw error;
  }
};

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
  console.log('User passwords updated with hashes in', endTime - startTime, 'ms');
};

const initializeServer = async () => {
  let attempt = 1;
  const maxAttempts = 5;
  try {
    await connectToMongoDB();
    await db.collection('users').createIndex({ username: 1 }, { unique: true });
    await db.collection('questions').createIndex({ testNumber: 1, variant: 1 });
    await db.collection('test_results').createIndex({ user: 1, endTime: -1 });
    await db.collection('activity_log').createIndex({ user: 1, timestamp: -1 });
    console.log('MongoDB indexes created successfully');
    await updateUserPasswords();
    await loadUsersToCache(); // Загружаем пользователей в кэш
    isInitialized = true;
    initializationError = null;
  } catch (error) {
    console.error('Failed to initialize server:', error.message, error.stack);
    initializationError = error;
    throw error;
  }
};

(async () => {
  try {
    await initializeServer();
    app.use(ensureInitialized);
  } catch (error) {
    console.error('Failed to start server due to initialization error:', error.message, error.stack);
    process.exit(1);
  }
})();

app.get('/test-mongo', async (req, res) => {
  try {
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    await db.collection('users').findOne();
    res.json({ success: true, message: 'MongoDB connection successful' });
  } catch (error) {
    console.error('MongoDB test failed:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'MongoDB connection failed', error: error.message });
  }
});

app.get('/api/test', (req, res) => {
  console.log('Handling /api/test request...');
  res.json({ success: true, message: 'Express server is working on /api/test' });
});

// Исправленный маршрут для страницы входа с двумя полями: логин и пароль
app.get('/', (req, res) => {
  console.log('Serving index.html');
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
          input[type="text"], input[type="password"] { padding: 10px; font-size: 18px; width: 200px; margin-bottom: 10px; }
          button { padding: 10px 20px; font-size: 18px; cursor: pointer; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
          button:hover { background-color: #45a049; }
          .error { color: red; margin-top: 10px; }
          .checkbox-container { margin-bottom: 10px; }
          @media (max-width: 600px) {
            h1 { font-size: 28px; }
            input[type="text"], input[type="password"], button { font-size: 16px; width: 90%; padding: 15px; }
          }
        </style>
      </head>
      <body>
        <h1>Авторизація</h1>
        <form id="login-form" method="POST" action="/login">
          <input type="text" id="username" name="username" placeholder="Логін" required><br>
          <input type="password" id="password" name="password" placeholder="Пароль" required><br>
          <div class="checkbox-container">
            <input type="checkbox" id="show-password" onclick="togglePassword()">
            <label for="show-password">Показати пароль</label>
          </div>
          <button type="submit">Увійти</button>
        </form>
        <div id="error-message" class="error"></div>
        <script>
          console.log('Cookies before login:', document.cookie);

          document.getElementById('login-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const errorMessage = document.getElementById('error-message');

            const formData = new URLSearchParams();
            formData.append('username', username);
            formData.append('password', password);

            try {
              console.log('Sending login request...');
              const response = await fetch('/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                credentials: 'include',
                body: formData
              });

              console.log('Response received:', response);
              console.log('Response status:', response.status);
              console.log('Response headers:', [...response.headers.entries()]);

              if (!response.ok) {
                throw new Error('HTTP error! status: ' + response.status);
              }

              const setCookieHeader = response.headers.get('set-cookie');
              console.log('Response headers (set-cookie):', setCookieHeader);
              if (setCookieHeader) {
                console.log('Set-Cookie header found:', setCookieHeader);
              } else {
                console.log('No Set-Cookie header in response');
              }

              const result = await response.json();
              console.log('Parsed login response:', result);
              console.log('Cookies after login:', document.cookie);

              if (result.success) {
                console.log('Redirecting to:', result.redirect);
                window.location.href = result.redirect + '?nocache=' + Date.now();
              } else {
                console.log('Login failed with message:', result.message);
                errorMessage.textContent = result.message || 'Помилка входу';
              }
            } catch (error) {
              console.error('Error during login:', error);
              errorMessage.textContent = 'Помилка сервера: ' + error.message;
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

const logActivity = async (user, action, sessionId, ipAddress, additionalInfo = {}) => {
  try {
    const startTime = Date.now();
    const timestamp = new Date();
    const timeOffset = 3 * 60 * 60 * 1000;
    const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
    await db.collection('activity_log').insertOne({
      user,
      action,
      sessionId,
      ipAddress,
      timestamp: adjustedTimestamp.toISOString(),
      additionalInfo
    });
    const endTime = Date.now();
    console.log(`Logged activity: ${user} - ${action} at ${adjustedTimestamp}, IP: ${ipAddress}, Session: ${sessionId} in ${endTime - startTime} ms`);
  } catch (error) {
    console.error('Error logging activity:', error.message, error.stack);
  }
};

// Исправленный маршрут /login для обработки двух полей
app.post('/login', async (req, res) => {
  const startTime = Date.now();
  try {
    const { username, password } = req.body;
    console.log('Received login data:', { username, password });

    if (!username || !password) {
      console.log('Username or password not provided');
      return res.status(400).json({ success: false, message: 'Логін або пароль не вказано' });
    }

    // Ищем пользователя в кэше
    const foundUser = userCache.find(user => user.username === username);
    if (!foundUser) {
      console.log('User not found:', username);
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    // Проверяем пароль
    const passwordMatch = await bcrypt.compare(password, foundUser.password);
    if (!passwordMatch) {
      console.log('Invalid password for user:', username);
      return res.status(401).json({ success: false, message: 'Невірний логін або пароль' });
    }

    // Генерируем новый session ID
    await new Promise((resolve, reject) => {
      req.session.regenerate(err => {
        if (err) {
          console.error('Error regenerating session:', err);
          reject(err);
        } else {
          console.log('Session regenerated successfully');
          resolve();
        }
      });
    });

    req.session.user = foundUser.username;
    req.session.testVariant = Math.floor(Math.random() * 3) + 1;
    console.log(`Assigned variant ${req.session.testVariant} to user ${foundUser.username}`);
    console.log(`Session after login:`, req.session);
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    const sessionId = req.session.id;
    await logActivity(foundUser.username, 'увійшов на сайт', sessionId, ipAddress);

    console.log(`Session ID: ${sessionId}`);

    // Явно помечаем сессию как измененную
    req.session.modified = true;

    // Дебагинг заголовков ответа
    const headers = res.getHeaders();
    console.log('Response headers after session setup:', headers);
    if (headers['set-cookie']) {
      console.log('Set-Cookie details:', headers['set-cookie']);
    } else {
      console.log('No Set-Cookie header found');
    }

    if (foundUser.username === 'admin') {
      console.log('Sending response: redirect to /admin');
      res.json({ success: true, redirect: '/admin' });
    } else {
      console.log('Sending response: redirect to /select-test');
      res.json({ success: true, redirect: '/select-test' });
    }
  } catch (error) {
    console.error('Ошибка в /login:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /login executed in ${endTime - startTime} ms`);
  }
});

const checkAuth = (req, res, next) => {
  console.log(`CheckAuth: Cookies received:`, req.cookies);
  console.log(`CheckAuth: Raw cookie header:`, req.headers.cookie);
  console.log(`CheckAuth: Headers received:`, req.headers);
  console.log(`CheckAuth: Session ID from cookie:`, req.sessionID);
  console.log(`CheckAuth: Full session object:`, req.session);
  const user = req.session.user;
  console.log(`CheckAuth: user in session: ${user}, session ID: ${req.session.id}`);

  // Дебагінг: перевіряємо сесію в MongoDB
  if (req.sessionID) {
    db.collection('sessions').findOne({ _id: req.sessionID }, (err, session) => {
      if (err) {
        console.error('Error checking session in MongoDB:', err);
      } else {
        console.log('Session found in MongoDB:', session);
      }
    });
  } else {
    console.log('No session ID in request');
  }

  if (!user) {
    console.log('CheckAuth: No user in session, redirecting to /');
    return res.redirect('/');
  }
  req.user = user;
  console.log(`CheckAuth: User authenticated: ${req.user}`);
  next();
};

const checkAdmin = (req, res, next) => {
  const user = req.session.user;
  if (user !== 'admin') {
    return res.status(403).send('Доступно тільки для адміністратора (403 Forbidden)');
  }
  next();
};

app.get('/select-test', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    console.log(`Serving /select-test for user: ${req.user}`);
    if (req.user === 'admin') {
      console.log('User is admin, redirecting to /admin');
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
            ${Object.entries(testNames).map(([num, data]) => `
              <button onclick="window.location.href='/test?test=${num}'">${data.name}</button>
            `).join('')}
          </div>
          <button id="logout" onclick="logout()">Вийти</button>
          <script>
            async function logout() {
              const formData = new URLSearchParams();
              await fetch('/logout', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                credentials: 'include'
              });
              window.location.href = '/';
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /select-test executed in ${endTime - startTime} ms`);
  }
});

app.post('/logout', (req, res) => {
  const startTime = Date.now();
  try {
    const user = req.session.user;
    const sessionId = req.session.id;
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    if (user) {
      logActivity(user, 'покинув сайт', sessionId, ipAddress);
    }
    req.session.destroy(err => {
      if (err) {
        console.error('Error destroying session:', err);
        return res.status(500).json({ success: false, message: 'Помилка при виході' });
      }
      res.json({ success: true });
    });
  } finally {
    const endTime = Date.now();
    console.log(`Route /logout executed in ${endTime - startTime} ms`);
  }
});

const userTests = new Map();

const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant) => {
  const startTimeLog = Date.now();
  try {
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
    console.log('Saving result to MongoDB with answers:', result.answers);
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    const insertResult = await db.collection('test_results').insertOne(result);
    console.log(`Successfully saved result for ${user} in MongoDB with ID:`, insertResult.insertedId);
  } catch (error) {
    console.error('Ошибка сохранения в MongoDB:', error.message, error.stack);
    throw error;
  } finally {
    const endTimeLog = Date.now();
    console.log(`saveResult executed in ${endTimeLog - startTimeLog} ms`);
  }
};

app.get('/test', checkAuth, async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const testNumber = req.query.test;
    if (!testNumber || !testNames[testNumber]) {
      return res.status(400).send('Номер тесту не вказано або тест не знайдено');
    }
    let questions = await loadQuestions(testNumber);
    const userVariant = req.session.testVariant;
    questions = questions.filter(q => !q.variant || q.variant === '' || q.variant === `Variant ${userVariant}`);
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
    userTests.set(req.user, {
      testNumber,
      questions,
      answers: {},
      currentQuestion: 0,
      startTime: Date.now(),
      timeLimit: testNames[testNumber].timeLimit * 1000,
      variant: userVariant
    });
    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    const sessionId = req.session.id;
    await logActivity(req.user, `розпочав тест ${testNames[testNumber].name}`, sessionId, ipAddress);
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    console.error('Ошибка в /test:', error.message, error.stack);
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  } finally {
    const endTime = Date.now();
    console.log(`Route /test executed in ${endTime - startTime} ms`);
  }
});

app.get('/test/question', checkAuth, (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).send('Тест не розпочато');
    }
    const { questions, testNumber, answers, currentQuestion, startTime, timeLimit } = userTest;
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
            .matching-container { display: flex; justify-content: space-between; flex-wrap: wrap; }
            .matching-column { width: 45%; }
            .matching-item { border: 2px solid #ccc; padding: 10px; margin: 5px 0; border-radius: 5px; cursor: move; font-size: 16px; }
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
              .option-box, .matching-item { font-size: 18px; padding: 15px; }
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
      console.log(`Fillblank question parts for index ${index}:`, q.text.split('___'));
      const parts = q.text.split('___');
      let inputHtml = '';
      parts.forEach((part, i) => {
        inputHtml += `<span class="question-text">${part}</span>`;
        if (i < parts.length - 1) {
          const userAnswer = userAnswers[i] || '';
          inputHtml += `<input type="text" class="blank-input" id="blank_${i}" value="${userAnswer.replace(/"/g, '"')}" placeholder="Введіть відповідь">`;
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
          <input type="text" name="q${index}" id="q${index}_input" value="${userAnswer}" placeholder="Введіть відповідь" class="answer-option"><br>
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
            console.log('Starting test question script...');
            const startTime = ${startTime};
            const timeLimit = ${timeLimit};
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
            let matchingPairs = ${JSON.stringify(answers[index] || [])};

            function updateTimer() {
              try {
                console.log('Updating timer...');
                const elapsedTime = Math.floor((Date.now() - startTime) / 1000);
                const remainingTime = Math.max(0, Math.floor(timeLimit / 1000) - elapsedTime);
                const minutes = Math.floor(remainingTime / 60).toString().padStart(2, '0');
                const seconds = (remainingTime % 60).toString().padStart(2, '0');
                timerElement.textContent = 'Залишилось часу: ' + minutes + ' мм ' + seconds + ' с';
                if (remainingTime <= 0) {
                  console.log('Time is up, redirecting to /result');
                  window.location.href = '/result';
                }
              } catch (error) {
                console.error('Error in updateTimer:', error);
              }
            }
            updateTimer();
            setInterval(updateTimer, 1000);

            window.addEventListener('blur', () => {
              console.log('Window blurred');
              lastBlurTime = Date.now();
              switchCount++;
            });

            window.addEventListener('focus', () => {
              console.log('Window focused');
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

            document.addEventListener('mousemove', debounceMouseMove);
            document.addEventListener('keydown', () => {
              console.log('Key pressed');
              lastActivityTime = Date.now();
              activityCount++;
            });

            document.querySelectorAll('.option-box:not(.draggable)').forEach(box => {
              box.addEventListener('click', () => {
                console.log('Option box clicked:', box.getAttribute('data-value'));
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
                console.log('Saving answer and moving to next question:', index);
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
                console.log('Saving answer for question ' + index + ':', answers);
                const responseTime = Date.now() - questionStartTime;

                const formData = new URLSearchParams();
                formData.append('index', index);
                formData.append('answer', JSON.stringify(answers));
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  credentials: 'include',
                  body: formData
                });

                const result = await response.json();
                if (result.success) {
                  console.log('Answer saved, redirecting to next question');
                  window.location.href = '/test/question?index=' + (index + 1);
                } else {
                  console.error('Error saving answer:', result.error);
                }
              } catch (error) {
                console.error('Error in saveAndNext:', error);
              }
            }

            function showConfirm(index) {
              console.log('Showing confirmation modal for index:', index);
              document.getElementById('confirm-modal').style.display = 'block';
            }

            function hideConfirm() {
              console.log('Hiding confirmation modal');
              document.getElementById('confirm-modal').style.display = 'none';
            }

            async function finishTest(index) {
              try {
                console.log('Finishing test for index:', index);
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
                console.log('Finishing test, answer for question ' + index + ':', answers);
                const responseTime = Date.now() - questionStartTime;

                const formData = new URLSearchParams();
                formData.append('index', index);
                formData.append('answer', JSON.stringify(answers));
                formData.append('timeAway', timeAway);
                formData.append('switchCount', switchCount);
                formData.append('responseTime', responseTime);
                formData.append('activityCount', activityCount);

                const response = await fetch('/answer', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  credentials: 'include',
                  body: formData
                });

                const result = await response.json();
                if (result.success) {
                  console.log('Test finished, redirecting to /result');
                  window.location.href = '/result';
                } else {
                  console.error('Error finishing test:', result.error);
                }
              } catch (error) {
                console.error('Error in finishTest:', error);
              }
            }

            const sortable = document.getElementById('sortable-options');
            if (sortable) {
              console.log('Initializing Sortable for ordering question');
              new Sortable(sortable, { animation: 150 });
            }

            const leftColumn = document.getElementById('left-column');
            const rightColumn = document.getElementById('right-column');
            if (leftColumn && rightColumn && '${q.type}' === 'matching') {
              console.log('Initializing Sortable for matching question');
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
                console.log('Updated matching pairs:', matchingPairs);
              }

              function resetMatchingPairs() {
                matchingPairs = [];
                const rightItems = document.querySelectorAll('#right-column .droppable');
                rightItems.forEach(item => {
                  const rightValue = item.dataset.value || '';
                  item.innerHTML = rightValue;
                });
                console.log('Matching pairs reset');
              }

              const droppableItems = document.querySelectorAll('.droppable');
              if (droppableItems.length > 0) {
                console.log('Adding dragover and drop listeners to', droppableItems.length, 'droppable items');
                droppableItems.forEach(item => {
                  item.addEventListener('dragover', (e) => e.preventDefault());
                  item.addEventListener('drop', (e) => {
                    e.preventDefault();
                    const draggable = document.querySelector('.dragging');
                    if (draggable && draggable.classList.contains('draggable')) {
                      const leftValue = draggable.dataset.value || '';
                      const rightValue = item.dataset.value || '';
                      console.log('Dropped:', { leftValue, rightValue });
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
                      } else {
                        console.warn('Missing leftValue or rightValue:', { leftValue, rightValue });
                      }
                    }
                  });
                });
              } else {
                console.warn('No droppable items found for matching question');
              }
            }
            console.log('Test question script initialized successfully');
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /test/question executed in ${endTime - startTime} ms`);
  }
});

app.post('/answer', checkAuth, express.urlencoded({ extended: true }), async (req, res) => {
  const startTime = Date.now();
  try {
    if (req.user === 'admin') return res.redirect('/admin');
    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;

    console.log('Received /answer data:', req.body);

    // Проверка входных данных
    if (!index || !answer) {
      console.log('Missing required fields in /answer:', { index, answer });
      return res.status(400).json({ success: false, error: 'Необхідно надати index та answer' });
    }

    let parsedAnswer;
    try {
      // Проверяем, является ли answer строкой и пытаемся распарсить как JSON
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
      console.error('Ошибка парсинга ответа в /answer:', error.message, error.stack);
      return res.status(400).json({ success: false, error: 'Невірний формат відповіді' });
    }

    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).json({ success: false, error: 'Тест не розпочато' });
    }

    console.log(`Saving answer for question ${index}:`, parsedAnswer);
    userTest.answers[index] = parsedAnswer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = (userTest.suspiciousActivity.timeAway || 0) + (parseInt(timeAway) || 0);
    userTest.suspiciousActivity.switchCount = (userTest.suspiciousActivity.switchCount || 0) + (parseInt(switchCount) || 0);
    userTest.suspiciousActivity.responseTimes[index] = parseInt(responseTime) || 0;
    userTest.suspiciousActivity.activityCounts[index] = parseInt(activityCount) || 0;

    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка в /answer:', error.message, error.stack);
    res.status(500).json({ success: false, error: 'Помилка сервера' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /answer executed in ${endTime - startTime} ms`);
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
      if (q.type === 'multiple' && userAnswer && Array.isArray(userAnswer)) {
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\\'/g, "'"));
        const isCorrect = correctAnswers.length === userAnswers.length &&
          correctAnswers.every(val => userAnswers.includes(val)) &&
          userAnswers.every(val => correctAnswers.includes(val));
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'input' && userAnswer) {
        const normalizedUserAnswer = String(userAnswer).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        const normalizedCorrectAnswer = String(q.correctAnswers[0]).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.');
        const isCorrect = normalizedUserAnswer === normalizedCorrectAnswer;
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'ordering' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase());
        const isCorrect = userAnswers.join(',') === correctAnswers.join(',');
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'matching' && userAnswer && Array.isArray(userAnswer)) {
        const userPairs = userAnswer.map(pair => [String(pair[0]).trim().toLowerCase(), String(pair[1]).trim().toLowerCase()]);
        const correctPairs = q.correctPairs.map(pair => [String(pair[0]).trim().toLowerCase(), String(pair[1]).trim().toLowerCase()]);
        const isCorrect = userPairs.length === correctPairs.length &&
          userPairs.every(userPair => correctPairs.some(correctPair => userPair[0] === correctPair[0] && userPair[1] === correctPair[1]));
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.'));
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase().replace(/\s+/g, '').replace(',', '.'));
        console.log(`Fillblank question ${index + 1}: userAnswers=${userAnswers}, correctAnswers=${correctAnswers}`);
        const isCorrect = userAnswers.length === correctAnswers.length &&
          userAnswers.every((answer, idx) => answer === correctAnswers[idx]);
        if (isCorrect) {
          questionScore = q.points;
        }
      } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
        const correctAnswer = String(q.correctAnswer).trim().toLowerCase();
        console.log(`Single choice question ${index + 1}: userAnswers=${userAnswers}, correctAnswer=${correctAnswer}`);
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

    if (timeAwayPercent > 50 || switchCount > 5) {
      const activityDetails = {
        timeAwayPercent,
        switchCount,
        avgResponseTime,
        totalActivityCount
      };
      await sendSuspiciousActivityEmail(req.user, activityDetails);
    }

    const ipAddress = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
    const sessionId = req.session.id;
    await logActivity(req.user, `завершив тест ${testNames[testNumber].name} з результатом ${Math.round(percentage)}%`, sessionId, ipAddress, { percentage: Math.round(percentage) });

    try {
      await saveResult(req.user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant);
    } catch (error) {
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
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
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
                  header: { fontSize: 14, bold: true, margin: [0, 0, 0,                     10], lineHeight: 2 }
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
    console.log(`Route /result executed in ${endTime - startTime} ms`);
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
          console.log(`Fillblank question ${index + 1} in /results: userAnswers=${userAnswers}, correctAnswers=${correctAnswers}`);
          if (userAnswers.length === correctAnswers.length &&
              userAnswers.every((answer, idx) => answer === correctAnswers[idx])) {
            questionScore = q.points;
          }
        } else if (q.type === 'singlechoice' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase());
          const correctAnswer = String(q.correctAnswer).trim().toLowerCase();
          console.log(`Single choice question ${index + 1} in /results: userAnswers=${userAnswers}, correctAnswer=${correctAnswer}`);
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
    console.log(`Route /results executed in ${endTime - startTime} ms`);
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
              await fetch('/logout', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                credentials: 'include'
              });
              window.location.href = '/';
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin executed in ${endTime - startTime} ms`);
  }
});

app.get('/admin/users', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let users = [];
    let errorMessage = '';
    try {
      users = await db.collection('users').find({}).toArray();
      userCache = users; // Обновляем кэш
    } catch (error) {
      console.error('Error fetching users from MongoDB:', error.message, error.stack);
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
                  const response = await fetch('/admin/delete-user', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    credentials: 'include'
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні користувача: ' + result.message);
                  }
                } catch (error) {
                  alert('Помилка при видаленні користувача');
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
    console.log(`Route /admin/users executed in ${endTime - startTime} ms`);
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
            <label for="username">Ім'я користувача:</label>
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
    console.log(`Route /admin/add-user executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/add-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username, password } = req.body;
    if (!username || username.length < 3 || username.length > 50) {
      return res.status(400).send('Ім’я користувача має бути від 3 до 50 символів');
    }
    if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
      return res.status(400).send('Ім’я користувача може містити лише літери та цифри');
    }
    if (!password || password.length < 6 || password.length > 100) {
      return res.status(400).send('Пароль має бути від 6 до 100 символів');
    }
    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser) {
      return res.status(400).send('Користувач із таким ім’ям уже існує');
    }
    const saltRounds = 10;
    const hashedPassword = await bcrypt.hash(password, saltRounds);
    const newUser = { username, password: hashedPassword };
    await db.collection('users').insertOne(newUser);
    userCache.push(newUser); // Обновляем кэш
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
    console.error('Error adding user:', error.message, error.stack);
    res.status(500).send('Помилка при додаванні користувача');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/add-user (POST) executed in ${endTime - startTime} ms`);
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
    console.log(`Route /admin/edit-user executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/edit-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { oldUsername, username, password } = req.body;
    if (!username || username.length < 3 || username.length > 50) {
      return res.status(400).send('Ім’я користувача має бути від 3 до 50 символів');
    }
    if (!/^[a-zA-Z0-9а-яА-Я]+$/.test(username)) {
      return res.status(400).send('Ім’я користувача може містити лише літери та цифри');
    }
    if (password && (password.length < 6 || password.length > 100)) {
      return res.status(400).send('Пароль має бути від 6 до 100 символів');
    }
    const existingUser = await db.collection('users').findOne({ username });
    if (existingUser && username !== oldUsername) {
      return res.status(400).send('Користувач із таким ім’ям уже існує');
    }
    const updateData = { username };
    if (password) {
      const saltRounds = 10;
      const hashedPassword = await bcrypt.hash(password, saltRounds);
      updateData.password = hashedPassword;
    }
    await db.collection('users').updateOne(
      { username: oldUsername },
      { $set: updateData }
    );
    // Обновляем кэш
    const userIndex = userCache.findIndex(user => user.username === oldUsername);
    if (userIndex !== -1) {
      userCache[userIndex] = { ...userCache[userIndex], ...updateData };
    }
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
    console.error('Error editing user:', error.message, error.stack);
    res.status(500).send('Помилка при редагуванні користувача');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/edit-user (POST) executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-user', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { username } = req.body;
    await db.collection('users').deleteOne({ username });
    // Обновляем кэш
    userCache = userCache.filter(user => user.username !== username);
    res.json({ success: true });
  } catch (error) {
    console.error('Error deleting user:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні користувача' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-user executed in ${endTime - startTime} ms`);
  }
});

app.get('/admin/questions', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let questions = [];
    let errorMessage = '';
    try {
      questions = await db.collection('questions').find({}).sort({ testNumber: 1 }).toArray();
      // Обновляем кэш вопросов
      questions.forEach(q => {
        if (!questionsCache[q.testNumber]) {
          questionsCache[q.testNumber] = [];
        }
        if (!questionsCache[q.testNumber].some(cachedQ => cachedQ._id.toString() === q._id.toString())) {
          questionsCache[q.testNumber].push(q);
        }
      });
    } catch (error) {
      console.error('Error fetching questions from MongoDB:', error.message, error.stack);
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
          </style>
        </head>
        <body>
          <h1>Керування питаннями</h1>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <button class="nav-btn" onclick="window.location.href='/admin/add-question'">Додати питання</button>
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
          <script>
            async function deleteQuestion(id) {
              if (confirm('Ви впевнені, що хочете видалити це питання?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('id', id);
                  const response = await fetch('/admin/delete-question', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    credentials: 'include'
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні питання: ' + result.message);
                  }
                } catch (error) {
                  alert('Помилка при видаленні питання');
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
    console.log(`Route /admin/questions executed in ${endTime - startTime} ms`);
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
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}">${testNames[num].name}</option>`).join('')}
            </select>
            <label for="picture">Посилання на фото (опціонально):</label>
            <input type="text" id="picture" name="picture" placeholder="Введіть URL зображення">
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
            <label for="options">Варіанти відповідей (через кому):</label>
            <textarea id="options" name="options" placeholder="Введіть варіанти через кому"></textarea>
            <label for="correctAnswers">Правильні відповіді (через кому):</label>
            <textarea id="correctAnswers" name="correctAnswers" required placeholder="Введіть правильні відповіді через кому"></textarea>
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
              if (picture && !/^https?:\/\/.*\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Посилання на фото має бути дійсним URL із розширенням .jpeg, .jpg, .png або .gif';
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
    console.log(`Route /admin/add-question executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/add-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber, text, type, options, correctAnswers, points, variant, picture } = req.body;
    if (!testNumber || !text || !type || !correctAnswers) {
      return res.status(400).send('Необхідно заповнити всі обов’язкові поля');
    }
    if (text.length < 5 || text.length > 1000) {
      return res.status(400).send('Текст питання має бути від 5 до 1000 символів');
    }
    if (!['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank'].includes(type.toLowerCase())) {
      return res.status(400).send('Невірний тип питання');
    }
    const pointsNum = Number(points);
    if (!pointsNum || pointsNum < 1 || pointsNum > 100) {
      return res.status(400).send('Бали мають бути числом від 1 до 100');
    }
    if (variant && (variant.length < 1 || variant.length > 50)) {
      return res.status(400).send('Варіант має бути від 1 до 50 символів');
    }
    if (picture && !/^https?:\/\/.*\.(jpeg|jpg|png|gif)$/i.test(picture)) {
      return res.status(400).send('Посилання на фото має бути дійсним URL із розширенням .jpeg, .jpg, .png або .gif');
    }

    let questionData = {
      testNumber,
      picture: picture || '',
      text,
      type: type.toLowerCase(),
      options: options ? options.split(',').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(',').map(ans => ans.trim()).filter(Boolean),
      points: pointsNum,
      variant: variant || ''
    };

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

    const result = await db.collection('questions').insertOne(questionData);
    // Обновляем кэш
    if (!questionsCache[testNumber]) {
      questionsCache[testNumber] = [];
    }
    questionsCache[testNumber].push({ ...questionData, _id: result.insertedId });
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
    console.error('Error adding question:', error.message, error.stack);
    res.status(500).send('Помилка при додаванні питання');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/add-question (POST) executed in ${endTime - startTime} ms`);
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
          </style>
        </head>
        <body>
          <h1>Редагувати питання</h1>
          <form method="POST" action="/admin/edit-question" onsubmit="return validateForm()">
            <input type="hidden" name="id" value="${id}">
            <label for="testNumber">Номер тесту:</label>
            <select id="testNumber" name="testNumber" required>
              ${Object.keys(testNames).map(num => `<option value="${num}" ${num === question.testNumber ? 'selected' : ''}>${testNames[num].name}</option>`).join('')}
            </select>
            <label for="picture">Посилання на фото (опціонально):</label>
            <input type="text" id="picture" name="picture" value="${question.picture || ''}" placeholder="Введіть URL зображення">
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
            <label for="options">Варіанти відповідей (через кому):</label>
            <textarea id="options" name="options">${question.options.join(', ')}</textarea>
            <label for="correctAnswers">Правильні відповіді (через кому):</label>
            <textarea id="correctAnswers" name="correctAnswers" required>${question.correctAnswers.join(', ')}</textarea>
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
              if (picture && !/^https?:\/\/.*\.(jpeg|jpg|png|gif)$/i.test(picture)) {
                errorMessage.textContent = 'Посилання на фото має бути дійсним URL із розширенням .jpeg, .jpg, .png або .gif';
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
    console.log(`Route /admin/edit-question executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/edit-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id, testNumber, text, type, options, correctAnswers, points, variant, picture } = req.body;
    if (!testNumber || !text || !type || !correctAnswers) {
      return res.status(400).send('Необхідно заповнити всі обов’язкові поля');
    }
    if (text.length < 5 || text.length > 1000) {
      return res.status(400).send('Текст питання має бути від 5 до 1000 символів');
    }
    if (!['multiple', 'singlechoice', 'truefalse', 'input', 'ordering', 'matching', 'fillblank'].includes(type.toLowerCase())) {
      return res.status(400).send('Невірний тип питання');
    }
    const pointsNum = Number(points);
    if (!pointsNum || pointsNum < 1 || pointsNum > 100) {
      return res.status(400).send('Бали мають бути числом від 1 до 100');
    }
    if (variant && (variant.length < 1 || variant.length > 50)) {
      return res.status(400).send('Варіант має бути від 1 до 50 символів');
    }
    if (picture && !/^https?:\/\/.*\.(jpeg|jpg|png|gif)$/i.test(picture)) {
      return res.status(400).send('Посилання на фото має бути дійсним URL із розширенням .jpeg, .jpg, .png або .gif');
    }

    let questionData = {
      testNumber,
      picture: picture || '',
      text,
      type: type.toLowerCase(),
      options: options ? options.split(',').map(opt => opt.trim()).filter(Boolean) : [],
      correctAnswers: correctAnswers.split(',').map(ans => ans.trim()).filter(Boolean),
      points: pointsNum,
      variant: variant || ''
    };

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

    // Находим старый номер теста для вопроса
    const oldQuestion = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    const oldTestNumber = oldQuestion.testNumber;

    // Обновляем вопрос в базе
    await db.collection('questions').updateOne(
      { _id: new ObjectId(id) },
      { $set: questionData }
    );

    // Обновляем кэш
    if (questionsCache[oldTestNumber]) {
      const questionIndex = questionsCache[oldTestNumber].findIndex(q => q._id.toString() === id);
      if (questionIndex !== -1) {
        // Удаляем из старого кэша
        questionsCache[oldTestNumber].splice(questionIndex, 1);
      }
    }
    if (!questionsCache[testNumber]) {
      questionsCache[testNumber] = [];
    }
    questionsCache[testNumber].push({ ...questionData, _id: new ObjectId(id) });

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
    console.error('Error editing question:', error.message, error.stack);
    res.status(500).send('Помилка при редагуванні питання');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/edit-question (POST) executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-question', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.body;
    const question = await db.collection('questions').findOne({ _id: new ObjectId(id) });
    await db.collection('questions').deleteOne({ _id: new ObjectId(id) });
    // Обновляем кэш
    if (question && questionsCache[question.testNumber]) {
      questionsCache[question.testNumber] = questionsCache[question.testNumber].filter(q => q._id.toString() !== id);
    }
    res.json({ success: true });
  } catch (error) {
    console.error('Error deleting question:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні питання' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-question executed in ${endTime - startTime} ms`);
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
          </style>
        </head>
        <body>
          <h1>Імпорт користувачів із Excel</h1>
          <form method="POST" action="/admin/import-users" enctype="multipart/form-data">
            <label for="file">Виберіть файл users.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn">Завантажити</button>
          </form>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/import-users executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/import-users', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      return res.status(400).send('Файл не завантажено');
    }
    const filePath = req.file.path;
    const importedCount = await importUsersToMongoDB(filePath);
    fs.unlinkSync(filePath);
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
    console.error('Error importing users:', error.message, error.stack);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).send('Помилка при імпорті користувачів');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/import-users (POST) executed in ${endTime - startTime} ms`);
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
          </style>
        </head>
        <body>
          <h1>Імпорт питань із Excel</h1>
          <form method="POST" action="/admin/import-questions" enctype="multipart/form-data">
            <label for="file">Виберіть файл questions*.xlsx:</label>
            <input type="file" id="file" name="file" accept=".xlsx" required>
            <button type="submit" class="submit-btn">Завантажити</button>
          </form>
          <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/import-questions executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/import-questions', checkAuth, checkAdmin, upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      return res.status(400).send('Файл не завантажено');
    }
    const filePath = req.file.path;
    const testNumber = req.file.originalname.match(/^questions(\d+)\.xlsx$/)?.[1];
    if (!testNumber) {
      fs.unlinkSync(filePath);
      return res.status(400).send('Файл повинен мати назву у форматі questionsX.xlsx, де X — номер тесту');
    }
    const importedCount = await importQuestionsToMongoDB(filePath, testNumber);
    fs.unlinkSync(filePath);
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Питання імпортовано</title>
        </head>
        <body>
          <h1>Імпортовано ${importedCount} питань для тесту ${testNumber}</h1>
          <button onclick="window.location.href='/admin/questions'">Повернутися до списку питань</button>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error importing questions:', error.message, error.stack);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).send('Помилка при імпорті питань');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/import-questions (POST) executed in ${endTime - startTime} ms`);
  }
});

app.get('/admin/results', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let results = [];
    let errorMessage = '';
    try {
      results = await db.collection('test_results').find({}).sort({ endTime: -1 }).toArray();
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
            .delete-all-btn { background-color: #ff4d4d; color: white; padding: 10px 20px; margin: 10px 0; border: none; cursor: pointer; }
            .nav-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; background-color: #007bff; color: white; border: none; }
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
              <th>Варіант</th>
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
      adminHtml += '<tr><td colspan="12">Немає результатів</td></tr>';
    } else {
      results.forEach((r, index) => {
        const answersArray = [];
        if (r.answers) {
          Object.keys(r.answers).sort((a, b) => parseInt(a) - parseInt(b)).forEach(key => {
            const idx = parseInt(key);
            answersArray[idx] = r.answers[key];
          });
        }
        console.log(`User ${r.user} answers array:`, answersArray);
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
        adminHtml += `
          <tr>
            <td>${r.user || 'N/A'}</td>
            <td>${testNames[r.testNumber]?.name || 'N/A'}</td>
            <td>${r.variant || 'N/A'}</td>
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
          <button class="delete-all-btn" onclick="deleteAllResults()">Видалити всі результати</button>
          <script>
            async function deleteResult(id) {
              if (confirm('Ви впевнені, що хочете видалити цей результат?')) {
                try {
                  const formData = new URLSearchParams();
                  formData.append('id', id);
                  const response = await fetch('/admin/delete-result', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    credentials: 'include'
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні результату: ' + result.message);
                  }
                } catch (error) {
                  alert('Помилка при видаленні результату');
                }
              }
            }

            async function deleteAllResults() {
              if (confirm('Ви впевнені, що хочете видалити всі результати? Цю дію не можна скасувати!')) {
                try {
                  const formData = new URLSearchParams();
                  const response = await fetch('/admin/delete-all-results', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    credentials: 'include'
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні всіх результатів: ' + result.message);
                  }
                } catch (error) {
                  alert('Помилка при видаленні всіх результатів');
                }
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml.trim());
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/results executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-all-results', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    if (!db) {
      throw new Error('MongoDB connection not established');
    }
    const deleteResult = await db.collection('test_results').deleteMany({});
    console.log(`Deleted ${deleteResult.deletedCount} results from test_results collection`);
    res.json({ success: true, message: `Успішно видалено ${deleteResult.deletedCount} результатів` });
  } catch (error) {
    console.error('Ошибка при удалении всех результатов:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні всіх результатів' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-all-results executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { id } = req.body;
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении результата:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-result executed in ${endTime - startTime} ms`);
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
            ${Object.entries(testNames).map(([num, data]) => `
              <div class="test-row">
                <label for="test${num}">Назва Тесту ${num}:</label>
                <input type="text" id="test${num}" name="test${num}" value="${data.name}" required>
                <label for="time${num}">Час (сек):</label>
                <input type="number" id="time${num}" name="time${num}" value="${data.timeLimit}" required min="1">
                <label for="randomQuestions${num}">Випадковий вибір питань:</label>
                <input type="checkbox" id="randomQuestions${num}" name="randomQuestions${num}" ${data.randomQuestions ? 'checked' : ''}>
                <label for="randomAnswers${num}">Випадковий вибір відповідей:</label>
                <input type="checkbox" id="randomAnswers${num}" name="randomAnswers${num}" ${data.randomAnswers ? 'checked' : ''}>
                <label for="questionLimit${num}">Кількість питань:</label>
                <input type="number" id="questionLimit${num}" name="questionLimit${num}" value="${data.questionLimit || ''}" min="1" placeholder="Без обмеження">
                <button type="button" class="delete-btn" onclick="deleteTest('${num}')">Видалити</button>
              </div>
            `).join('')}
            <button type="submit">Зберегти</button>
          </form>
          <button onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
          <script>
            async function deleteTest(testNumber) {
              if (confirm('Ви впевнені, що хочете видалити Тест ' + testNumber + '?')) {
                const formData = new URLSearchParams();
                formData.append('testNumber', testNumber);
                await fetch('/admin/delete-test', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                  credentials: 'include'
                });
                window.location.reload();
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(html);
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/edit-tests executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/edit-tests', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    Object.keys(testNames).forEach(num => {
      const testName = req.body[`test${num}`];
      const timeLimit = req.body[`time${num}`];
      const randomQuestions = req.body[`randomQuestions${num}`] === 'on';
      const randomAnswers = req.body[`randomAnswers${num}`] === 'on';
      const questionLimit = req.body[`questionLimit${num}`] ? parseInt(req.body[`questionLimit${num}`]) : null;
      if (testName && timeLimit) {
        testNames[num] = {
          name: testName,
          timeLimit: parseInt(timeLimit) || testNames[num].timeLimit,
          randomQuestions,
          randomAnswers,
          questionLimit
        };
      }
    });
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
    console.error('Ошибка при редактировании названий тестов:', error.message, error.stack);
    res.status(500).send('Помилка при оновленні назв тестів');
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/edit-tests (POST) executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    const { testNumber } = req.body;
    if (!testNames[testNumber]) {
      return res.status(404).json({ success: false, message: 'Тест не знайдено' });
    }
    delete testNames[testNumber];
    // Удаляем вопросы из кэша
    if (questionsCache[testNumber]) {
      delete questionsCache[testNumber];
    }
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении теста:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-test executed in ${endTime - startTime} ms`);
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
            input { font-size: 24px; padding: 5px; margin: 5px; }
            select { font-size: 24px; padding: 5px; margin: 5px; }
            button { font-size: 24px; padding: 10px 20px; margin: 5px; }
          </style>
        </head>
        <body>
          <h1>Створити новий тест</h1>
          <form method="POST" action="/admin/create-test">
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
    console.log(`Route /admin/create-test executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/create-test', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
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
      questionLimit: null
    };
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
    res.status(500).send(`Помилка при створенні тесту: ${error.message}`);
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/create-test (POST) executed in ${endTime - startTime} ms`);
  }
});

app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    let activities = [];
    let errorMessage = '';
    try {
      activities = await db.collection('activity_log').find({}).sort({ timestamp: -1 }).toArray();
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
            .nav-btn, .clear-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; border: none; border-radius: 5px; }
            .clear-btn { background-color: #ff4d4d; color: white; }
            .nav-btn { background-color: #007bff; color: white; }
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
              <th>Номер сесії</th>
              <th>Дія</th>
            </tr>
    `;
    if (!activities || activities.length === 0) {
      adminHtml += '<tr><td colspan="5">Немає записів</td></tr>';
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
            <td>${activity.sessionId || 'N/A'}</td>
            <td>${actionWithInfo}</td>
          </tr>
        `;
      });
    }
    adminHtml += `
          </table>
          <button class="clear-btn" onclick="clearActivityLog()">Видалити усі записи журналу</button>
          <script>
            async function clearActivityLog() {
              if (confirm('Ви впевнені, що хочете видалити усі записи журналу дій?')) {
                try {
                  const formData = new URLSearchParams();
                  const response = await fetch('/admin/delete-activity-log', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    credentials: 'include'
                  });
                  const result = await response.json();
                  if (result.success) {
                    window.location.reload();
                  } else {
                    alert('Помилка при видаленні записів журналу: ' + result.message);
                  }
                } catch (error) {
                  alert('Помилка при видаленні записів журналу');
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
    console.log(`Route /admin/activity-log executed in ${endTime - startTime} ms`);
  }
});

app.post('/admin/delete-activity-log', checkAuth, checkAdmin, async (req, res) => {
  const startTime = Date.now();
  try {
    await db.collection('activity_log').deleteMany({});
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении записей журнала действий:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
  } finally {
    const endTime = Date.now();
    console.log(`Route /admin/delete-activity-log executed in ${endTime - startTime} ms`);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

module.exports = app;
