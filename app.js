const express = require('express');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { MongoClient } = require('mongodb');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const crypto = require('crypto');

const app = express();
const port = process.env.PORT || 3000;
const saltRounds = 10;

// MongoDB подключение
const mongoUrl = process.env.MONGODB_URI || 'mongodb://localhost:27017/alpha';
let db;
let mongoClient; // Сохраняем клиент для использования в MongoStore

MongoClient.connect(mongoUrl)
  .then(client => {
    mongoClient = client; // Сохраняем клиент
    db = client.db('alpha'); // Указываем имя базы данных
    console.log('MongoDB connected');
  })
  .catch(err => {
    console.error('MongoDB connection error:', err);
    process.exit(1);
  });

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Конфигурация сессий
app.use(session({
  store: MongoStore.create({
    client: mongoClient, // Передаём сохранённый клиент
    dbName: 'alpha', // Указываем имя базы данных
    collectionName: 'sessions' // Имя коллекции для хранения сессий
  }),
  secret: process.env.SESSION_SECRET || 'your-secret-key',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: true, // Убедись, что HTTPS включён
    httpOnly: true,
    sameSite: 'lax',
    maxAge: 24 * 60 * 60 * 1000 // 24 часа
  }
}));

// Генерация CSRF-токена
function generateCsrfToken() {
  return crypto.randomBytes(16).toString('hex');
}

// Middleware для добавления CSRF-токена в сессию
app.use((req, res, next) => {
  if (!req.session.csrfToken) {
    req.session.csrfToken = generateCsrfToken();
  }
  res.locals.csrfToken = req.session.csrfToken;
  next();
});

// Middleware для проверки CSRF-токена
function verifyCsrfToken(req, res, next) {
  const csrfToken = req.body._csrf || req.headers['x-csrf-token'];
  if (!csrfToken || csrfToken !== req.session.csrfToken) {
    return res.status(403).json({ success: false, message: 'Недійсний CSRF-токен' });
  }
  next();
}

// Rate limiting для маршрута /login
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 минут
  max: 5, // Максимум 5 попыток
  message: 'Забагато спроб входу. Спробуйте знову через 15 хвилин.'
});

// Хранилище тестов пользователей
const userTests = new Map();
const testNames = {
  1: { name: 'Тест 1' },
  2: { name: 'Тест 2' }
};

// Загрузка пользователей из Excel
async function loadUsers() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(__dirname, 'users.xlsx'));
  const worksheet = workbook.getWorksheet(1);
  const users = {};
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const username = row.getCell(1).value;
      const password = row.getCell(2).value;
      if (username && password) {
        users[username] = password;
      }
    }
  });
  return users;
}

// Хеширование пароля
async function hashPassword(password) {
  return await bcrypt.hash(password, saltRounds);
}

// Проверка пароля
async function checkPassword(inputPassword, hashedPassword) {
  return await bcrypt.compare(inputPassword, hashedPassword);
}

// Логирование активности
const logActivity = async (user, action) => {
  try {
    const timestamp = new Date();
    const timeOffset = 3 * 60 * 60 * 1000; // 3 часа в миллисекундах
    const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
    await db.collection('activity_log').insertOne({
      user,
      action,
      timestamp: adjustedTimestamp.toISOString()
    });
    console.log(`Logged activity: ${user} - ${action} at ${adjustedTimestamp}`);
  } catch (error) {
    console.error('Error logging activity:', error.message, error.stack);
  }
};

// Логирование подозрительной активности
async function logSuspiciousActivity(user, testNumber, suspiciousScore, details) {
  try {
    const timestamp = new Date();
    const timeOffset = 3 * 60 * 60 * 1000; // 3 часа в миллисекундах
    const adjustedTimestamp = new Date(timestamp.getTime() + timeOffset);
    await db.collection('suspicious_activity').insertOne({
      user,
      testNumber,
      suspiciousScore,
      details,
      timestamp: adjustedTimestamp.toISOString()
    });
    console.log(`Logged suspicious activity for ${user}: ${suspiciousScore}`);
  } catch (error) {
    console.error('Error logging suspicious activity:', error);
  }
}

// Проверка авторизации
function checkAuth(req, res, next) {
  console.log('Checking authentication...');
  console.log('Session:', req.session);
  console.log('Session ID:', req.sessionID);
  console.log('Cookies:', req.cookies);
  if (!req.session.user) {
    console.warn('User not authenticated, redirecting to login');
    return res.redirect('/');
  }
  console.log('User authenticated:', req.session.user);
  next();
}

// Проверка админ-доступа
function checkAdmin(req, res, next) {
  if (req.session.user !== 'admin') {
    console.warn('Access denied for non-admin user:', req.session.user);
    return res.status(403).send('Доступ заборонено');
  }
  next();
}

// Сохранение результата теста
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
    const durationInSeconds = duration || 1; // Избегаем деления на 0
    const timeAwayPercent = suspiciousActivity.timeAway ? 
      Math.round((suspiciousActivity.timeAway / (duration * 1000)) * 100) : 0;
    suspiciousScore += timeAwayPercent;

    const switchCount = suspiciousActivity.switchCount || 0;
    if (switchCount > totalQuestions * 3) {
      suspiciousScore += 25;
    }

    const responseTimes = suspiciousActivity.responseTimes || [];
    const avgResponseTime = responseTimes.length > 0 ? 
      (responseTimes.reduce((sum, time) => sum + (time || 0), 0) / responseTimes.length / 1000).toFixed(2) : 0;

    let typicalResponseTime = 30 * 1000;
    let typicalResponseTimeDeviation = 0;
    try {
      const allResults = await db.collection('test_results').find({}).toArray();
      if (allResults.length > 0) {
        const allResponseTimes = allResults.flatMap(r => r.suspiciousActivity.responseTimes || []);
        if (allResponseTimes.length > 0) {
          typicalResponseTime = allResponseTimes.reduce((sum, time) => sum + (time || 0), 0) / allResponseTimes.length;
          const mean = typicalResponseTime;
          typicalResponseTimeDeviation = Math.sqrt(
            allResponseTimes.reduce((sum, time) => sum + Math.pow((time || 0) - mean, 2), 0) / allResponseTimes.length
          );
        }
      }
    } catch (error) {
      console.error('Error calculating typical response time:', error);
    }
    responseTimes.forEach(time => {
      if (time < typicalResponseTime - 2 * typicalResponseTimeDeviation) {
        suspiciousScore += 15;
      } else if (time > typicalResponseTime + 2 * typicalResponseTimeDeviation) {
        suspiciousScore += 15;
      }
    });

    const activityCounts = suspiciousActivity.activityCounts || [];
    const actionsPerSecond = activityCounts.reduce((sum, count) => sum + (count || 0), 0) / durationInSeconds;
    if (actionsPerSecond > 50) {
      suspiciousScore += 20;
    }
    const avgActivityCount = activityCounts.length > 0 ? 
      (activityCounts.reduce((sum, count) => sum + (count || 0), 0) / activityCounts.length).toFixed(2) : 0;
    activityCounts.forEach((count, idx) => {
      if (count < 5 && responseTimes[idx] > 30 * 1000) {
        suspiciousScore += 10;
      }
    });

    let typicalSwitchCount = totalQuestions;
    try {
      const allResults = await db.collection('test_results').find({}).toArray();
      if (allResults.length > 0) {
        const allSwitchCounts = allResults.map(r => r.suspiciousActivity.switchCount || 0);
        typicalSwitchCount = allSwitchCounts.length > 0 ? 
          allSwitchCounts.reduce((sum, count) => sum + count, 0) / allSwitchCounts.length : typicalSwitchCount;
      }
    } catch (error) {
      console.error('Error calculating typical switch count:', error);
    }
    if (switchCount > typicalSwitchCount * 1.5) {
      suspiciousScore += 15;
    }

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

    // Логируем подозрительную активность, если подозрительность высокая
    if (suspiciousScore > 50) {
      await logSuspiciousActivity(user, testNumber, suspiciousScore, {
        timeAwayPercent,
        switchCount,
        avgResponseTime,
        actionsPerSecond
      });
    }
  } catch (error) {
    console.error('Ошибка сохранения в MongoDB:', error.message, error.stack);
    throw error;
  }
};

// Главная страница (логин)
app.get('/', (req, res) => {
  if (req.session.user) {
    if (req.session.user === 'admin') {
      return res.redirect('/admin');
    } else {
      return res.redirect('/select-test');
    }
  }
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Вхід</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
          input, button { padding: 10px; margin: 10px; font-size: 16px; }
          button { cursor: pointer; background-color: #4CAF50; color: white; border: none; border-radius: 5px; }
          button:hover { background-color: #45a049; }
        </style>
      </head>
      <body>
        <h1>Вхід</h1>
        <form method="POST" action="/login">
          <input type="hidden" name="_csrf" value="${res.locals.csrfToken}">
          <input type="password" name="password" placeholder="Введіть пароль">
          <button type="submit">Увійти</button>
        </form>
      </body>
    </html>
  `);
});

// Маршрут логина
app.post('/login', loginLimiter, verifyCsrfToken, async (req, res) => {
  try {
    console.log('Handling /login request...');
    const { password } = req.body;
    if (!password) {
      console.warn('Password not provided in /login request');
      return res.status(400).json({ success: false, message: 'Пароль не вказано' });
    }

    console.log('Loading users from Excel for authentication...');
    const validPasswords = await loadUsers();
    console.log('Checking password:', password, 'against validPasswords:', validPasswords);
    
    const user = Object.keys(validPasswords).find(async u => {
      const match = await checkPassword(password, validPasswords[u]);
      console.log(`Comparing ${u}: ${validPasswords[u]} with ${password} -> ${match}`);
      return match;
    });

    if (!user) {
      console.warn('Password not found in validPasswords');
      return res.status(401).json({ success: false, message: 'Невірний пароль' });
    }

    req.session.user = user;
    await logActivity(user, 'увійшов на сайт');
    console.log('Session after setting user:', req.session);
    console.log('Session ID after setting user:', req.sessionID);
    console.log('Cookies after setting session:', req.cookies);

    req.session.csrfToken = generateCsrfToken(); // Обновляем токен после успешного входа

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

// Маршрут выхода
app.post('/logout', verifyCsrfToken, (req, res) => {
  if (req.session.user) {
    const user = req.session.user;
    req.session.destroy(async err => {
      if (err) {
        console.error('Error destroying session:', err);
        return res.status(500).send('Помилка при виході');
      }
      await logActivity(user, 'вийшов з сайту');
      res.json({ success: true, redirect: '/' });
    });
  } else {
    res.json({ success: true, redirect: '/' });
  }
});

// Админ-панель
app.get('/admin', checkAuth, checkAdmin, (req, res) => {
  console.log('Serving /admin for user:', req.session.user);
  res.send(`
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
        <button onclick="window.location.href='/admin/results'">Перегляд результатів</button><br>
        <button onclick="window.location.href='/admin/edit-tests'">Редагувати назви тестів</button><br>
        <button onclick="window.location.href='/admin/create-test'">Створити новий тест</button><br>
        <button onclick="window.location.href='/admin/activity-log'">Журнал дій</button><br>
        <button onclick="window.location.href='/admin/suspicious-activity'">Лог підозрілої активності</button><br>
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
              window.location.href = result.redirect;
            } else {
              alert('Помилка при виході');
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Выбор теста
app.get('/select-test', checkAuth, (req, res) => {
  if (req.session.user === 'admin') return res.redirect('/admin');
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Вибір тесту</title>
        <style>
          body { font-family: Arial, sans-serif; text-align: center; padding: 50px; font-size: 24px; margin: 0; }
          h1 { font-size: 36px; margin-bottom: 20px; }
          button { padding: 15px 30px; margin: 10px; font-size: 24px; cursor: pointer; width: 300px; border: none; border-radius: 5px; background-color: #4CAF50; color: white; }
          button:hover { background-color: #45a049; }
          #logout { background-color: #ef5350; color: white; position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); width: 300px; }
          @media (max-width: 600px) {
            body { padding: 20px; padding-bottom: 80px; }
            h1 { font-size: 32px; }
            button { font-size: 20px; width: 90%; padding: 15px; }
            #logout { width: 90%; }
          }
        </style>
      </head>
      <body>
        <h1>Вибір тесту</h1>
        ${Object.keys(testNames).map(num => `
          <button onclick="startTest(${num})">${testNames[num].name}</button><br>
        `).join('')}
        <button id="logout" onclick="logout()">Вихід</button>
        <script>
          async function startTest(testNumber) {
            const response = await fetch('/start-test', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ testNumber, _csrf: "${res.locals.csrfToken}" })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = result.redirect;
            } else {
              alert('Помилка при запуску тесту');
            }
          }
          async function logout() {
            const response = await fetch('/logout', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ _csrf: "${res.locals.csrfToken}" })
            });
            const result = await response.json();
            if (result.success) {
              window.location.href = result.redirect;
            } else {
              alert('Помилка при виході');
            }
          }
        </script>
      </body>
    </html>
  `);
});

// Запуск теста
app.post('/start-test', checkAuth, verifyCsrfToken, async (req, res) => {
  if (req.session.user === 'admin') return res.status(403).send('Доступ заборонено');
  const { testNumber } = req.body;
  if (!testNames[testNumber]) {
    return res.status(400).json({ success: false, message: 'Невірний номер тесту' });
  }

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(path.join(__dirname, `test${testNumber}.xlsx`));
  } catch (error) {
    console.error(`Error reading test${testNumber}.xlsx:`, error);
    return res.status(500).json({ success: false, message: 'Помилка при завантаженні тесту' });
  }

  const worksheet = workbook.getWorksheet(1);
  const questions = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const picture = row.getCell(1).value;
      const text = row.getCell(2).value;
      const type = row.getCell(3).value;
      const options = [];
      for (let i = 4; i <= 9; i++) {
        const option = row.getCell(i).value;
        if (option) options.push(option);
      }
      const correctAnswers = [];
      for (let i = 10; i <= 15; i++) {
        const correct = row.getCell(i).value;
        if (correct) correctAnswers.push(correct);
      }
      const points = parseInt(row.getCell(16).value) || 0;
      questions.push({ picture, text, type, options, correctAnswers, points });
    }
  });

  userTests.set(req.session.user, {
    questions,
    answers: {},
    testNumber,
    startTime: Date.now(),
    timeLimit: 30 * 60 * 1000, // 30 минут
    suspiciousActivity: { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] }
  });

  await logActivity(req.session.user, `розпочав тест ${testNames[testNumber].name}`);
  req.session.csrfToken = generateCsrfToken(); // Обновляем токен
  res.json({ success: true, redirect: '/test/question?index=0' });
});

// Отображение вопроса
app.get('/test/question', checkAuth, (req, res) => {
  if (req.session.user === 'admin') return res.redirect('/admin');
  const userTest = userTests.get(req.session.user);
  if (!userTest) {
    console.warn(`Test not started for user ${req.session.user}`);
    return res.status(400).send('Тест не розпочато');
  }

  const { questions, testNumber, answers, currentQuestion, startTime, timeLimit } = userTest;
  const index = parseInt(req.query.index) || 0;

  if (index < 0 || index >= questions.length) {
    console.warn(`Invalid question index ${index} for user ${req.session.user}`);
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

          document.addEventListener('copy', () => {
            console.log('Copy event detected');
            switchCount += 5;
          });

          document.addEventListener('paste', () => {
            console.log('Paste event detected');
            switchCount += 5;
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
              if (document.querySelector('input[name="q' + index + '"]')) {
                answers = document.getElementById('q' + index + '_input').value;
              } else if (document.getElementById('sortable-options')) {
                answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
              }
              const responseTime = Date.now() - questionStartTime;
              console.log('Sending answer with data:', { index, answers, timeAway, switchCount, responseTime, activityCount });
              const response = await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount, _csrf: "${res.locals.csrfToken}" })
              });
              const result = await response.json();
              if (result.success) {
                console.log('Answer saved successfully, redirecting to next question');
                window.location.href = '/test/question?index=' + (index + 1);
              } else {
                console.error('Failed to save answer:', result);
              }
            } catch (error) {
              console.error('Error in saveAndNext:', error);
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
              if (document.querySelector('input[name="q' + index + '"]')) {
                answers = document.getElementById('q' + index + '_input').value;
              } else if (document.getElementById('sortable-options')) {
                answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
              }
              const responseTime = Date.now() - questionStartTime;
              console.log('Finishing test with data:', { index, answers, timeAway, switchCount, responseTime, activityCount });
              const response = await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount, _csrf: "${res.locals.csrfToken}" })
              });
              const result = await response.json();
              if (result.success) {
                console.log('Answer saved successfully, redirecting to result');
                hideConfirm();
                window.location.href = '/result';
              } else {
                console.error('Failed to save answer:', result);
              }
            } catch (error) {
              console.error('Error in finishTest:', error);
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

// Сохранение ответа
app.post('/answer', checkAuth, verifyCsrfToken, async (req, res) => {
  if (req.session.user === 'admin') return res.status(403).send('Доступ заборонено');
  const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;
  const userTest = userTests.get(req.session.user);
  if (!userTest) {
    return res.status(400).json({ success: false, message: 'Тест не розпочато' });
  }

  userTest.answers[index] = answer;
  userTest.suspiciousActivity.timeAway = timeAway;
  userTest.suspiciousActivity.switchCount = switchCount;
  userTest.suspiciousActivity.responseTimes[index] = responseTime;
  userTest.suspiciousActivity.activityCounts[index] = activityCount;

  await logActivity(req.session.user, `відповів на питання ${parseInt(index) + 1} тесту ${testNames[userTest.testNumber].name}`);
  req.session.csrfToken = generateCsrfToken(); // Обновляем токен
  res.json({ success: true });
});

// Результат теста
app.get('/result', checkAuth, async (req, res) => {
  if (req.session.user === 'admin') return res.redirect('/admin');
  const userTest = userTests.get(req.session.user);
  if (!userTest) {
    console.warn(`Test not started for user ${req.session.user} in /result`);
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
    await saveResult(req.session.user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage);
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

          const user = "${req.session.user}";
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

// Журнал действий
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
            <th>Час</th>
            <th>Дата</th>
          </tr>
  `;
  if (!activities || activities.length === 0) {
    adminHtml += '<tr><td colspan="4">Немає записів</td></tr>';
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

// Очистка журнала действий
app.post('/admin/delete-activity-log', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    console.log('Deleting all activity log entries...');
    await db.collection('activity_log').deleteMany({});
    console.log('Activity log cleared from MongoDB');
    req.session.csrfToken = generateCsrfToken(); // Обновляем токен
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении записей журнала действий:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
  }
});

// Просмотр результатов
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

// Удаление результата
app.post('/admin/delete-result', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  try {
    const { id } = req.body;
    console.log(`Deleting result with id ${id}...`);
    const { ObjectId } = require('mongodb');
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    console.log(`Result with id ${id} deleted from MongoDB`);
    req.session.csrfToken = generateCsrfToken(); // Обновляем токен
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении результата:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  }
});

// Лог подозрительной активности
app.get('/admin/suspicious-activity', checkAuth, checkAdmin, async (req, res) => {
  let suspiciousLogs = [];
  try {
    suspiciousLogs = await db.collection('suspicious_activity').find({}).sort({ timestamp: -1 }).toArray();
  } catch (error) {
    console.error('Error fetching suspicious activity logs:', error);
    res.status(500).send('Помилка при отриманні логу підозрілої активності');
    return;
  }

  let html = `
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Лог підозрілої активності</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          table { border-collapse: collapse; width: 100%; margin-top: 20px; }
          th, td { border: 1px solid black; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; }
          .nav-btn { padding: 10px 20px; margin: 10px 0; cursor: pointer; }
        </style>
      </head>
      <body>
        <h1>Лог підозрілої активності</h1>
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <table>
          <tr>
            <th>Користувач</th>
            <th>Тест</th>
            <th>Підозрілість (%)</th>
            <th>Деталі</th>
            <th>Час</th>
          </tr>
  `;
  if (suspiciousLogs.length === 0) {
    html += '<tr><td colspan="5">Немає записів</td></tr>';
  } else {
    suspiciousLogs.forEach(log => {
      const timestamp = new Date(log.timestamp);
      const formattedTime = `${timestamp.toLocaleTimeString('uk-UA', { hour12: false })} ${timestamp.toLocaleDateString('uk-UA')}`;
      const details = `
Время вне вкладки: ${log.details.timeAwayPercent}%
Переключения вкладок: ${log.details.switchCount}
Среднее время ответа (сек): ${log.details.avgResponseTime}
Действий в секунду: ${log.details.actionsPerSecond}
      `;
      html += `
        <tr>
          <td>${log.user}</td>
          <td>${testNames[log.testNumber]?.name || 'N/A'}</td>
          <td>${log.suspiciousScore}%</td>
          <td>${details}</td>
          <td>${formattedTime}</td>
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
});

// Создание нового теста
app.get('/admin/create-test', checkAuth, checkAdmin, (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Створити новий тест</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          input, select, button { padding: 10px; margin: 5px; font-size: 16px; }
          button { cursor: pointer; background-color: #4CAF50; color: white; border: none; border-radius: 5px; }
          .nav-btn { padding: 10px 20px; margin: 10px 0; }
          .question { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }
        </style>
      </head>
      <body>
        <h1>Створити новий тест</h1>
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <form id="createTestForm">
          <input type="hidden" name="_csrf" value="${res.locals.csrfToken}">
          <label>Назва тесту:</label>
          <input type="text" id="testName" required><br>
          <div id="questions">
            <div class="question">
              <label>Питання 1:</label><br>
              <label>Посилання на зображення:</label>
              <input type="text" name="picture_0"><br>
              <label>Текст питання:</label>
              <input type="text" name="text_0" required><br>
              <label>Тип питання:</label>
              <select name="type_0">
                <option value="multiple">Множинний вибір</option>
                <option value="input">Введення відповіді</option>
                <option value="ordering">Порядок</option>
              </select><br>
              <label>Варіанти відповідей (до 6):</label><br>
              <input type="text" name="option_0_0"><br>
              <input type="text" name="option_0_1"><br>
              <input type="text" name="option_0_2"><br>
              <input type="text" name="option_0_3"><br>
              <input type="text" name="option_0_4"><br>
              <input type="text" name="option_0_5"><br>
              <label>Правильні відповіді (до 6):</label><br>
              <input type="text" name="correct_0_0" required><br>
              <input type="text" name="correct_0_1"><br>
              <input type="text" name="correct_0_2"><br>
              <input type="text" name="correct_0_3"><br>
              <input type="text" name="correct_0_4"><br>
              <input type="text" name="correct_0_5"><br>
              <label>Бали:</label>
              <input type="number" name="points_0" value="1" required><br>
            </div>
          </div>
          <button type="button" onclick="addQuestion()">Додати питання</button>
          <button type="submit">Створити тест</button>
        </form>
        <script>
          let questionCount = 1;
          function addQuestion() {
            const questionsDiv = document.getElementById('questions');
            const newQuestion = document.createElement('div');
            newQuestion.className = 'question';
            newQuestion.innerHTML = 
              '<label>Питання ' + (questionCount + 1) + ':</label><br>' +
              '<label>Посилання на зображення:</label>' +
              '<input type="text" name="picture_' + questionCount + '"><br>' +
              '<label>Текст питання:</label>' +
              '<input type="text" name="text_' + questionCount + '" required><br>' +
              '<label>Тип питання:</label>' +
              '<select name="type_' + questionCount + '">' +
                '<option value="multiple">Множинний вибір</option>' +
                '<option value="input">Введення відповіді</option>' +
                '<option value="ordering">Порядок</option>' +
              '</select><br>' +
              '<label>Варіанти відповідей (до 6):</label><br>' +
              '<input type="text" name="option_' + questionCount + '_0"><br>' +
              '<input type="text" name="option_' + questionCount + '_1"><br>' +
              '<input type="text" name="option_' + questionCount + '_2"><br>' +
              '<input type="text" name="option_' + questionCount + '_3"><br>' +
              '<input type="text" name="option_' + questionCount + '_4"><br>' +
              '<input type="text" name="option_' + questionCount + '_5"><br>' +
              '<label>Правильні відповіді (до 6):</label><br>' +
              '<input type="text" name="correct_' + questionCount + '_0" required><br>' +
              '<input type="text" name="correct_' + questionCount + '_1"><br>' +
              '<input type="text" name="correct_' + questionCount + '_2"><br>' +
              '<input type="text" name="correct_' + questionCount + '_3"><br>' +
              '<input type="text" name="correct_' + questionCount + '_4"><br>' +
              '<input type="text" name="correct_' + questionCount + '_5"><br>' +
              '<label>Бали:</label>' +
              '<input type="number" name="points_' + questionCount + '" value="1" required><br>';
            questionsDiv.appendChild(newQuestion);
            questionCount++;
          }

          document.getElementById('createTestForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const testName = document.getElementById('testName').value;
            const questions = [];

            for (let i = 0; i < questionCount; i++) {
              const picture = formData.get('picture_' + i);
              const text = formData.get('text_' + i);
              const type = formData.get('type_' + i);
              const options = [];
              for (let j = 0; j < 6; j++) {
                const option = formData.get('option_' + i + '_' + j);
                if (option) options.push(option);
              }
              const correctAnswers = [];
              for (let j = 0; j < 6; j++) {
                const correct = formData.get('correct_' + i + '_' + j);
                if (correct) correctAnswers.push(correct);
              }
              const points = parseInt(formData.get('points_' + i)) || 0;
              questions.push({ picture, text, type, options, correctAnswers, points });
            }

            try {
              const response = await fetch('/admin/save-test', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ testName, questions })
              });
              const result = await response.json();
              if (result.success) {
                alert('Тест створено успішно!');
                window.location.href = '/admin';
              } else {
                alert('Помилка при створенні тесту: ' + result.message);
              }
            } catch (error) {
              console.error('Error creating test:', error);
              alert('Помилка при створенні тесту');
            }
          });
        </script>
      </body>
    </html>
  `);
});

// Сохранение нового теста
app.post('/admin/save-test', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  const { testName, questions } = req.body;
  if (!testName || !questions || !Array.isArray(questions) || questions.length === 0) {
    return res.status(400).json({ success: false, message: 'Невірні дані тесту' });
  }

  const testNumber = Object.keys(testNames).length + 1;
  testNames[testNumber] = { name: testName };

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Test');
  worksheet.columns = [
    { header: 'Picture', key: 'picture' },
    { header: 'Question', key: 'text' },
    { header: 'Type', key: 'type' },
    { header: 'Option 1', key: 'option1' },
    { header: 'Option 2', key: 'option2' },
    { header: 'Option 3', key: 'option3' },
    { header: 'Option 4', key: 'option4' },
    { header: 'Option 5', key: 'option5' },
    { header: 'Option 6', key: 'option6' },
    { header: 'Correct 1', key: 'correct1' },
    { header: 'Correct 2', key: 'correct2' },
    { header: 'Correct 3', key: 'correct3' },
    { header: 'Correct 4', key: 'correct4' },
    { header: 'Correct 5', key: 'correct5' },
    { header: 'Correct 6', key: 'correct6' },
    { header: 'Points', key: 'points' }
  ];

  questions.forEach(q => {
    worksheet.addRow({
      picture: q.picture || '',
      text: q.text,
      type: q.type,
      option1: q.options[0] || '',
      option2: q.options[1] || '',
      option3: q.options[2] || '',
      option4: q.options[3] || '',
      option5: q.options[4] || '',
      option6: q.options[5] || '',
      correct1: q.correctAnswers[0] || '',
      correct2: q.correctAnswers[1] || '',
      correct3: q.correctAnswers[2] || '',
      correct4: q.correctAnswers[3] || '',
      correct5: q.correctAnswers[4] || '',
      correct6: q.correctAnswers[5] || '',
      points: q.points
    });
  });

  try {
    await workbook.xlsx.writeFile(path.join(__dirname, `test${testNumber}.xlsx`));
    await logActivity(req.session.user, `створив тест ${testName}`);
    req.session.csrfToken = generateCsrfToken(); // Обновляем токен после успешного действия
    res.json({ success: true });
  } catch (error) {
    console.error('Error saving test:', error);
    res.status(500).json({ success: false, message: 'Помилка при збереженні тесту' });
  }
});

// Редактирование названий тестов
app.get('/admin/edit-tests', checkAuth, checkAdmin, (req, res) => {
  let html = `
    <!DOCTYPE html>
    <html lang="uk">
      <head>
        <meta charset="UTF-8">
        <title>Редагувати назви тестів</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          input, button { padding: 10px; margin: 5px; font-size: 16px; }
          button { cursor: pointer; background-color: #4CAF50; color: white; border: none; border-radius: 5px; }
          .nav-btn { padding: 10px 20px; margin: 10px 0; }
        </style>
      </head>
      <body>
        <h1>Редагувати назви тестів</h1>
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <form id="editTestsForm">
          <input type="hidden" name="_csrf" value="${res.locals.csrfToken}">
  `;
  Object.keys(testNames).forEach(num => {
    html += `
      <label>Тест ${num}:</label>
      <input type="text" name="test_${num}" value="${testNames[num].name}" required><br>
    `;
  });
  html += `
          <button type="submit">Зберегти</button>
        </form>
        <script>
          document.getElementById('editTestsForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(e.target);
            const data = {};
            for (let [key, value] of formData.entries()) {
              if (key.startsWith('test_')) {
                const testNumber = key.split('_')[1];
                data[testNumber] = value;
              }
            }
            try {
              const response = await fetch('/admin/save-test-names', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(data)
              });
              const result = await response.json();
              if (result.success) {
                alert('Назви тестів збережено!');
                window.location.href = '/admin';
              } else {
                alert('Помилка при збереженні назв тестів: ' + result.message);
              }
            } catch (error) {
              console.error('Error saving test names:', error);
              alert('Помилка при збереженні назв тестів');
            }
          });
        </script>
      </body>
    </html>
  `;
  res.send(html);
});

// Сохранение названий тестов
app.post('/admin/save-test-names', checkAuth, checkAdmin, verifyCsrfToken, async (req, res) => {
  const newTestNames = req.body;
  Object.keys(newTestNames).forEach(num => {
    if (testNames[num]) {
      testNames[num].name = newTestNames[num];
    }
  });
  await logActivity(req.session.user, 'оновив назви тестів');
  req.session.csrfToken = generateCsrfToken(); // Обновляем токен после успешного действия
  res.json({ success: true });
});

// Запуск сервера
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
