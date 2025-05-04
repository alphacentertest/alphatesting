require('dotenv').config();
const express = require('express');
const cookieParser = require('cookie-parser');
const path = require('path');
const ExcelJS = require('exceljs');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { MongoClient, ObjectId } = require('mongodb');
const fs = require('fs');

const app = express();

// Подключение к MongoDB
const MONGO_URL = process.env.MONGO_URL || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/testdb?retryWrites=true&w=majority';
const client = new MongoClient(MONGO_URL, { connectTimeoutMS: 5000, serverSelectionTimeoutMS: 5000 });
let db;

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
    secure: false,
    httpOnly: true,
    sameSite: 'lax',
    maxAge: 24 * 60 * 60 * 1000
  }
}));

const loadUsers = async () => {
  try {
    const filePath = path.join(__dirname, 'users.xlsx');
    console.log('Attempting to load users from:', filePath);
    if (!fs.existsSync(filePath)) {
      throw new Error(`File users.xlsx not found at path: ${filePath}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    let sheet = workbook.getWorksheet('Users') || workbook.getWorksheet('Sheet1');
    if (!sheet) {
      throw new Error('Ни один из листов ("Users" или "Sheet1") не найден');
    }
    const users = {};
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        const username = String(row.getCell(1).value || '').trim();
        const password = String(row.getCell(2).value || '').trim();
        if (username && password) {
          users[username] = password;
        }
      }
    });
    if (Object.keys(users).length === 0) {
      throw new Error('Не найдено пользователей в файле');
    }
    console.log('Loaded users from Excel:', users);
    return users;
  } catch (error) {
    console.error('Error loading users from users.xlsx:', error.message, error.stack);
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

const loadQuestions = async (testNumber) => {
  try {
    const filePath = path.join(__dirname, `questions${testNumber}.xlsx`);
    console.log(`Attempting to load questions from: ${filePath}`);
    if (!fs.existsSync(filePath)) {
      throw new Error(`File questions${testNumber}.xlsx not found at path: ${filePath}`);
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Questions');
    if (!sheet) {
      throw new Error(`Лист "Questions" не знайдено в questions${testNumber}.xlsx`);
    }
    const jsonData = [];
    sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) {
        const rowValues = row.values.slice(1);
        let questionText = rowValues[1];
        // Перевірка, чи є questionText об’єктом, і приведення до рядка
        if (typeof questionText === 'object' && questionText !== null) {
          console.warn(`Invalid question text in row ${rowNumber}: ${JSON.stringify(questionText)}. Converting to string.`);
          questionText = questionText.text || questionText.value || '[Невірний текст питання]';
        }
        questionText = String(questionText || '').trim();
        if (questionText === '') {
          console.warn(`Empty question text in row ${rowNumber}. Skipping.`);
          return;
        }
        const picture = String(rowValues[0] || '').trim();
        const options = rowValues.slice(2, 14).filter(Boolean).map(val => String(val).trim());
        const correctAnswers = rowValues.slice(14, 26).filter(Boolean).map(val => String(val).trim());
        const type = String(rowValues[26] || 'multiple').toLowerCase();
        const points = Number(rowValues[27]) || 1;
        const variant = String(rowValues[28] || '').trim();
        let questionData = {
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
          if (questionData.pairs.length === 0) {
            console.warn(`No valid pairs for matching question: ${questionText}`);
            return;
          }
          questionData.correctPairs = questionData.pairs.map(pair => [pair.left, pair.right]);
        }
        jsonData.push(questionData);
      }
    });
    if (jsonData.length === 0) {
      throw new Error(`No questions found in questions${testNumber}.xlsx`);
    }
    console.log(`Loaded questions for test ${testNumber}:`, jsonData);
    return jsonData;
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

const initializeServer = async () => {
  let attempt = 1;
  const maxAttempts = 5;
  try {
    await connectToMongoDB();
  } catch (error) {
    throw error;
  }
  while (attempt <= maxAttempts) {
    try {
      console.log(`Starting server initialization (Attempt ${attempt} of ${maxAttempts})...`);
      await loadUsers();
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

app.get('/', (req, res) => {
  console.log('Serving index.html');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const logActivity = async (user, action) => {
  try {
    const timestamp = new Date();
    const timeOffset = 3 * 60 * 60 * 1000;
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

app.post('/login', async (req, res) => {
  try {
    const { password } = req.body;
    if (!password) {
      return res.status(400).json({ success: false, message: 'Пароль не вказано' });
    }
    const validPasswords = await loadUsers();
    const user = Object.keys(validPasswords).find(u => validPasswords[u] === password);
    if (!user) {
      return res.status(401).json({ success: false, message: 'Невірний пароль' });
    }
    req.session.user = user;
    req.session.testVariant = Math.floor(Math.random() * 3) + 1; // Случайный выбор варианта (1, 2 или 3)
    console.log(`Assigned variant ${req.session.testVariant} to user ${user}`);
    await logActivity(user, 'увійшов на сайт');
    req.session.save(err => {
      if (err) {
        console.error('Error saving session in /login:', err.message, err.stack);
        return res.status(500).json({ success: false, message: 'Помилка сервера' });
      }
      if (user === 'admin') {
        res.json({ success: true, redirect: '/admin' });
      } else {
        res.json({ success: true, redirect: '/select-test' });
      }
    });
  } catch (error) {
    console.error('Ошибка в /login:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка сервера' });
  }
});

const checkAuth = (req, res, next) => {
  const user = req.session.user;
  if (!user) {
    return res.redirect('/');
  }
  req.user = user;
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
  if (req.user === 'admin') return res.redirect('/admin');
  res.send(`
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
            await fetch('/logout', { method: 'POST' });
            window.location.href = '/';
          }
        </script>
      </body>
    </html>
  `);
});

app.post('/logout', (req, res) => {
  req.session.destroy(err => {
    if (err) {
      console.error('Error destroying session:', err);
      return res.status(500).json({ success: false, message: 'Помилка при виході' });
    }
    res.json({ success: true });
  });
});

const userTests = new Map();

const saveResult = async (user, testNumber, score, totalPoints, startTime, endTime, totalClicks, correctClicks, totalQuestions, percentage, suspiciousActivity, answers, scoresPerQuestion, variant) => {
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
      variant: `Variant ${variant}` // Сохраняем вариант
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

app.get('/test', checkAuth, async (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  const testNumber = req.query.test;
  if (!testNumber || !testNames[testNumber]) {
    return res.status(400).send('Номер тесту не вказано або тест не знайдено');
  }
  try {
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
    res.redirect(`/test/question?index=0`);
  } catch (error) {
    console.error('Ошибка в /test:', error.message, error.stack);
    res.status(500).send('Помилка при завантаженні тесту: ' + error.message);
  }
});

app.get('/test/question', checkAuth, (req, res) => {
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
                         q.type === 'ordering' ? 'Розташуйте відповіді у правильній послідовності' :
                         q.type === 'matching' ? 'Складіть правильні пари, перетягуючи елементи' : '';
  html += `
          <div class="question-box">
            <h2 id="question-text">${index + 1}. ${q.text}</h2>
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
    `;
  } else if (!q.options || q.options.length === 0) {
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
          let matchingPairs = ${JSON.stringify(answers[index] || [])};

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

          document.addEventListener('mousemove', debounceMouseMove);
          document.addEventListener('keydown', () => {
            lastActivityTime = Date.now();
            activityCount++;
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
              }
              const responseTime = Date.now() - questionStartTime;
              await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount })
              });
              window.location.href = '/test/question?index=' + (index + 1);
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
            try {
              let answers = selectedOptions;
              if (document.querySelector('input[name="q' + index + '"]')) {
                answers = document.getElementById('q' + index + '_input').value;
              } else if (document.getElementById('sortable-options')) {
                answers = Array.from(document.querySelectorAll('#sortable-options .option-box')).map(el => el.dataset.value);
              } else if (document.getElementById('left-column')) {
                answers = matchingPairs;
              }
              const responseTime = Date.now() - questionStartTime;
              await fetch('/answer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ index, answer: answers, timeAway, switchCount, responseTime, activityCount })
              });
              window.location.href = '/result';
            } catch (error) {
              console.error('Error in finishTest:', error);
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
              onEnd: function(evt) {
                updateMatchingPairs();
              }
            });
            new Sortable(rightColumn, {
              group: 'matching',
              animation: 150,
              onEnd: function(evt) {
                updateMatchingPairs();
              }
            });

            function updateMatchingPairs() {
              matchingPairs = [];
              const rightItems = document.querySelectorAll('#right-column .droppable');
              rightItems.forEach(item => {
                const rightValue = item.dataset.value || '';
                const matchedText = item.querySelector('.matched')?.textContent || '';
                const leftMatch = matchedText.match(/Зіставлено: (.+)/)?.[1] || '';
                if (leftMatch && rightValue) {
                  matchingPairs.push([leftMatch, rightValue]);
                }
              });
              console.log('Updated matching pairs:', matchingPairs);
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
        </script>
      </body>
    </html>
  `;
  res.send(html);
});

app.post('/answer', checkAuth, (req, res) => {
  if (req.user === 'admin') return res.redirect('/admin');
  try {
    const { index, answer, timeAway, switchCount, responseTime, activityCount } = req.body;
    const userTest = userTests.get(req.user);
    if (!userTest) {
      return res.status(400).json({ error: 'Тест не розпочато' });
    }
    userTest.answers[index] = answer;
    userTest.suspiciousActivity = userTest.suspiciousActivity || { timeAway: 0, switchCount: 0, responseTimes: [], activityCounts: [] };
    userTest.suspiciousActivity.timeAway = (userTest.suspiciousActivity.timeAway || 0) + (timeAway || 0);
    userTest.suspiciousActivity.switchCount = (userTest.suspiciousActivity.switchCount || 0) + (switchCount || 0);
    userTest.suspiciousActivity.responseTimes[index] = responseTime || 0;
    userTest.suspiciousActivity.activityCounts[index] = activityCount || 0;
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка в /answer:', error.message, error.stack);
    res.status(500).json({ error: 'Помилка сервера' });
  }
});

app.get('/result', checkAuth, async (req, res) => {
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
});

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
      const correctAnswer = q.type === 'matching' ? q.correctPairs.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ') : q.correctAnswers.join(', ');
      const questionScore = scoresPerQuestion[index];
      const userAnswerDisplay = q.type === 'matching' && Array.isArray(userAnswer) ? userAnswer.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ') : Array.isArray(userAnswer) ? userAnswer.join(', ') : userAnswer;
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

app.get('/admin', checkAuth, checkAdmin, (req, res) => {
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
        <button id="logout" onclick="logout()">Вийти</button>
        <script>
          async function logout() {
            await fetch('/logout', { method: 'POST' });
            window.location.href = '/';
          }
        </script>
      </body>
    </html>
  `);
});

app.get('/admin/results', checkAuth, checkAdmin, async (req, res) => {
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
      const answersDisplay = answersArray.length > 0
        ? answersArray.map((a, i) => {
            if (!a) return null;
            const userAnswer = Array.isArray(a) && Array.isArray(a[0]) ? a.map(pair => `${pair[0]} -> ${pair[1]}`).join(', ') : Array.isArray(a) ? a.join(', ') : a;
            const questionScore = r.scoresPerQuestion[i] || 0;
            return `Питання ${i + 1}: ${userAnswer.replace(/\\'/g, "'")} (${questionScore} балів)`;
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
        <button class="nav-btn" onclick="window.location.href='/admin'">Повернутися до адмін-панелі</button>
        <script>
          async function deleteResult(id) {
            if (confirm('Ви впевнені, що хочете видалити цей результат?')) {
              try {
                const response = await fetch('/admin/delete-result', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ id })
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
        </script>
      </body>
    </html>
  `;
  res.send(adminHtml.trim()); // Видаляємо можливі зайві пробіли чи переноси
});

app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
  try {
    const { id } = req.body;
    await db.collection('test_results').deleteOne({ _id: new ObjectId(id) });
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении результата:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
  }
});

app.get('/admin/edit-tests', checkAuth, checkAdmin, (req, res) => {
  res.send(`
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
              await fetch('/admin/delete-test', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ testNumber })
              });
              window.location.reload();
            }
          }
        </script>
      </body>
    </html>
  `);
});

app.post('/admin/edit-tests', checkAuth, checkAdmin, (req, res) => {
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
  }
});

app.post('/admin/delete-test', checkAuth, checkAdmin, async (req, res) => {
  try {
    const { testNumber } = req.body;
    if (!testNames[testNumber]) {
      return res.status(404).json({ success: false, message: 'Тест не знайдено' });
    }
    delete testNames[testNumber];
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении теста:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні тесту' });
  }
});

app.get('/admin/create-test', checkAuth, checkAdmin, (req, res) => {
  const excelFiles = fs.readdirSync(__dirname).filter(file => file.endsWith('.xlsx') && file.startsWith('questions'));
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

app.post('/admin/create-test', checkAuth, checkAdmin, async (req, res) => {
  try {
    const { testName, excelFile, timeLimit } = req.body;
    const match = excelFile.match(/^questions(\d+)\.xlsx$/);
    if (!match) throw new Error('Невірний формат файлу Excel');
    const testNumber = match[1];
    if (testNames[testNumber]) throw new Error('Тест з таким номером вже існує');
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
  }
});

app.get('/admin/activity-log', checkAuth, checkAdmin, async (req, res) => {
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
                  headers: { 'Content-Type': 'application/json' }
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
});

app.post('/admin/delete-activity-log', checkAuth, checkAdmin, async (req, res) => {
  try {
    await db.collection('activity_log').deleteMany({});
    res.json({ success: true });
  } catch (error) {
    console.error('Ошибка при удалении записей журнала действий:', error.message, error.stack);
    res.status(500).json({ success: false, message: 'Помилка при видаленні записів журналу' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});

module.exports = app;
