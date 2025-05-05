const express = require('express');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const MongoClient = require('mongodb').MongoClient;
const cookieParser = require('cookie-parser');
const path = require('path');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;
const uri = process.env.MONGODB_URI || 'mongodb://localhost:27017/test_system';
const client = new MongoClient(uri);

let db;

const testNames = {
  1: { name: 'Тест 1', timeLimit: 15 * 60 * 1000 },
  2: { name: 'Тест 2', timeLimit: 15 * 60 * 1000 }
};

const userTests = new Map();
let users;

async function loadUsers() {
  try {
    const usersCollection = db.collection('users');
    const usersData = await usersCollection.find({}).toArray();
    const usersMap = {};
    usersData.forEach(user => {
      usersMap[user.username] = { password: user.password, role: user.role };
    });
    console.log('Loaded users from MongoDB:', usersMap);
    return usersMap;
  } catch (error) {
    console.error('Error loading users from MongoDB:', error.message, error.stack);
    throw error;
  }
}

const loadQuestions = async (testNumber) => {
  try {
    const questionsCollection = db.collection('questions');
    const questions = await questionsCollection.find({ testNumber }).toArray();
    if (questions.length === 0) {
      throw new Error(`No questions found for test ${testNumber} in MongoDB`);
    }
    console.log(`Loaded questions for test ${testNumber} from MongoDB:`, questions);
    return questions;
  } catch (error) {
    console.error(`Error in loadQuestions (test ${testNumber}):`, error.message, error.stack);
    throw error;
  }
};

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
  }
};

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

async function initialize() {
  try {
    await client.connect();
    db = client.db();
    users = await loadUsers();
  } catch (error) {
    console.error('Failed to initialize application:', error);
    process.exit(1); // Завершуємо процес із помилкою, якщо ініціалізація не вдалася
  }

  app.use(express.static('public'));
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));
  app.use(cookieParser());

  const sessionOptions = {
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: false,
    store: MongoStore.create({ client }),
    cookie: { maxAge: 24 * 60 * 60 * 1000 }
  };
  app.use(session(sessionOptions));

  app.get('/', (req, res) => {
    if (req.session.user) {
      return res.redirect('/select-test');
    }
    res.send(`
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Вхід</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; background-color: #f0f0f0; }
            input { margin: 10px; padding: 10px; width: 200px; }
            button { padding: 10px 20px; cursor: pointer; background-color: #007bff; color: white; border: none; border-radius: 5px; }
            .error { color: red; }
          </style>
        </head>
        <body>
          <h1>Вхід</h1>
          <form action="/login" method="POST">
            <input type="text" name="username" placeholder="Ім'я користувача" required><br>
            <input type="password" name="password" placeholder="Пароль" required><br>
            <button type="submit">Увійти</button>
          </form>
          <p class="error">${req.query.error || ''}</p>
        </body>
      </html>
    `);
  });

  app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
      const user = users[username];
      if (user && user.password === password) {
        req.session.user = username;
        const userDoc = await db.collection('users').findOne({ username });
        let variant = userDoc?.variant || Math.floor(Math.random() * 3) + 1;
        if (!userDoc) {
          await db.collection('users').insertOne({ username, password, role: user.role, variant });
        } else if (!userDoc.variant) {
          await db.collection('users').updateOne({ username }, { $set: { variant } });
        }
        console.log(`Assigned variant ${variant} to user ${username}`);
        await db.collection('activity_log').insertOne({
          user: username,
          action: 'увійшов на сайт',
          timestamp: new Date()
        });
        console.log(`Logged activity: ${username} - увійшов на сайт at ${new Date()}`);
        res.redirect('/select-test');
      } else {
        res.redirect('/?error=Невірне ім\'я користувача або пароль');
      }
    } catch (error) {
      console.error('Error in /login:', error);
      res.status(500).send('Помилка сервера');
    }
  });

  app.get('/logout', (req, res) => {
    req.session.destroy(() => {
      res.redirect('/');
    });
  });

  function checkAuth(req, res, next) {
    if (!req.session.user) {
      return res.redirect('/?error=Будь ласка, увійдіть');
    }
    req.user = req.session.user;
    next();
  }

  function checkAdmin(req, res, next) {
    if (users[req.session.user]?.role !== 'admin') {
      return res.redirect('/select-test');
    }
    next();
  }

  app.get('/select-test', checkAuth, (req, res) => {
    if (req.user === 'admin') return res.redirect('/admin');
    let html = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Вибір тесту</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; text-align: center; background-color: #f0f0f0; }
            button { padding: 10px 20px; margin: 10px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; background-color: #007bff; color: white; }
            .logout-btn { background-color: #ef5350; }
          </style>
        </head>
        <body>
          <h1>Вибір тесту</h1>
    `;
    Object.keys(testNames).forEach(testNumber => {
      html += `<button onclick="window.location.href='/test?test=${testNumber}'">${testNames[testNumber].name}</button>`;
    });
    html += `
          <br>
          <button class="logout-btn" onclick="window.location.href='/logout'">Вийти</button>
        </body>
      </html>
    `;
    res.send(html);
  });

  app.get('/test', checkAuth, async (req, res) => {
    if (req.user === 'admin') return res.redirect('/admin');
    const testNumber = parseInt(req.query.test);
    if (!testNames[testNumber]) {
      return res.status(400).send('Невірний номер тесту');
    }
    const userVariant = (await db.collection('users').findOne({ username: req.user }))?.variant || 1;
    let questions;
    try {
      questions = await loadQuestions(testNumber);
    } catch (error) {
      return res.status(500).send('Помилка завантаження питань');
    }
    questions = questions.filter(q => !q.variant || q.variant === '' || q.variant === `Variant ${userVariant}`);
    const timeLimit = testNames[testNumber].timeLimit;
    userTests.set(req.user, {
      questions,
      testNumber,
      answers: {},
      currentQuestion: 0,
      startTime: Date.now(),
      timeLimit,
      variant: userVariant
    });
    res.redirect(`/test/question?index=0`);
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
            .blank-input { width: 100px; margin: 0 5px; padding: 5px; border: 1px solid #ccc; border-radius: 4px; display: inline-block; }
            .question-text { display: inline; }
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
    if (q.picture) {
      html += `<img src="${q.picture}" alt="Picture" onerror="this.src='/images/placeholder.png'; console.log('Image failed to load: ${q.picture}')"><br>`;
    }
    const instructionText = q.type === 'multiple' ? 'Виберіть усі правильні відповіді' :
                           q.type === 'input' ? 'Введіть правильну відповідь' :
                           q.type === 'ordering' ? 'Розташуйте відповіді у правильній послідовності' :
                           q.type === 'matching' ? 'Складіть правильні пари, перетягуючи елементи' :
                           q.type === 'fillblank' ? 'Заповніть пропуски у реченні' : '';
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
                const questionType = '${q.type}';
                const option = box.getAttribute('data-value');
                if (questionType === 'truefalse' || questionType === 'multiple') {
                  if (questionType === 'truefalse') {
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
                console.log('Saving answer for question ' + index + ':', answers);
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
                } else if ('${q.type}' === 'fillblank') {
                  answers = [];
                  for (let i = 0; i < ${q.blankCount || 1}; i++) {
                    const input = document.getElementById('blank_' + i);
                    answers.push(input ? input.value.trim() : '');
                  }
                }
                console.log('Finishing test, answer for question ' + index + ':', answers);
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
      console.log(`Saving answer for question ${index}:`, answer);
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
      } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
        const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\s+/g, ''));
        const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase().replace(/\s+/g, ''));
        const isCorrect = userAnswers.length === correctAnswers.length &&
          userAnswers.every((answer, idx) => answer === correctAnswers[idx]);
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
        } else if (q.type === 'fillblank' && userAnswer && Array.isArray(userAnswer)) {
          const userAnswers = userAnswer.map(val => String(val).trim().toLowerCase().replace(/\s+/g, ''));
          const correctAnswers = q.correctAnswers.map(val => String(val).trim().toLowerCase().replace(/\s+/g, ''));
          if (userAnswers.length === correctAnswers.length &&
              userAnswers.every((answer, idx) => answer === correctAnswers[idx])) {
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
            <th>Ваша відповідь</th>
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
    let adminHtml = `
      <!DOCTYPE html>
      <html lang="uk">
        <head>
          <meta charset="UTF-8">
          <title>Адмін-панель</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            button { padding: 10px 20px; margin: 5px; cursor: pointer; border: none; border-radius: 5px; font-size: 16px; }
            textarea { width: 100%; height: 200px; margin-top: 10px; }
            .nav-btn { background-color: #007bff; color: white; }
            .logout-btn { background-color: #ef5350; }
          </style>
        </head>
        <body>
          <h1>Адмін-панель</h1>
          <button class="nav-btn" onclick="window.location.href='/admin/results'">Результати всіх користувачів</button>
          <button class="logout-btn" onclick="window.location.href='/logout'">Вийти</button>
          <h2>Оновлення користувачів</h2>
          <textarea id="usersInput" placeholder='Введіть користувачів у форматі JSON: [{"username": "Іваненко", "password": "pass111", "role": "user"}, ...]'></textarea>
          <button onclick="updateUsers()">Оновити користувачів</button>
          <h2>Оновлення питань</h2>
          <textarea id="questionsInput" placeholder='Введіть питання у форматі JSON: [{"testNumber": 1, "text": "Столиця України?", "options": ["Київ", "Львів"], "correctAnswers": ["Київ"], "type": "multiple", "points": 2, "variant": ""}, ...]'></textarea>
          <button onclick="updateQuestions()">Оновити питання</button>
          <script>
            async function updateUsers() {
              const usersInput = document.getElementById('usersInput').value;
              try {
                const users = JSON.parse(usersInput);
                const response = await fetch('/admin/update-users', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ users })
                });
                const result = await response.json();
                alert(result.message);
              } catch (error) {
                alert('Помилка при оновленні користувачів: ' + error.message);
              }
            }

            async function updateQuestions() {
              const questionsInput = document.getElementById('questionsInput').value;
              try {
                const questions = JSON.parse(questionsInput);
                const response = await fetch('/admin/update-questions', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ questions })
                });
                const result = await response.json();
                alert(result.message);
              } catch (error) {
                alert('Помилка при оновленні питань: ' + error.message);
              }
            }
          </script>
        </body>
      </html>
    `;
    res.send(adminHtml);
  });

  app.post('/admin/update-users', checkAuth, checkAdmin, async (req, res) => {
    try {
      const users = req.body.users;
      const usersCollection = db.collection('users');
      await usersCollection.deleteMany({});
      await usersCollection.insertMany(users);
      res.json({ success: true, message: 'Користувачі оновлені' });
    } catch (error) {
      console.error('Error updating users:', error);
      res.status(500).json({ success: false, message: 'Помилка при оновленні користувачів' });
    }
  });

  app.post('/admin/update-questions', checkAuth, checkAdmin, async (req, res) => {
    try {
      const questions = req.body.questions;
      const questionsCollection = db.collection('questions');
      await questionsCollection.deleteMany({});
      await questionsCollection.insertMany(questions);
      res.json({ success: true, message: 'Питання оновлені' });
    } catch (error) {
      console.error('Error updating questions:', error);
      res.status(500).json({ success: false, message: 'Помилка при оновленні питань' });
    }
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
  
            async function deleteAllResults() {
              if (confirm('Ви впевнені, що хочете видалити всі результати? Цю дію не можна скасувати!')) {
                try {
                  const response = await fetch('/admin/delete-all-results', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({})
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
  });
  
  app.post('/admin/delete-result', checkAuth, checkAdmin, async (req, res) => {
    try {
      const { id } = req.body;
      if (!id) {
        return res.status(400).json({ success: false, message: 'ID результату не вказано' });
      }
      const deleteResult = await db.collection('test_results').deleteOne({ _id: new MongoClient.ObjectID(id) });
      if (deleteResult.deletedCount === 0) {
        return res.status(404).json({ success: false, message: 'Результат не знайдено' });
      }
      res.json({ success: true, message: 'Результат видалено' });
    } catch (error) {
      console.error('Error deleting result:', error);
      res.status(500).json({ success: false, message: 'Помилка при видаленні результату' });
    }
  });
  
  app.post('/admin/delete-all-results', checkAuth, checkAdmin, async (req, res) => {
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
    }
  });
  
  app.listen(port, () => {
    console.log(`Server running on port ${port}`);
  });
  
  }
  
  initialize().catch(err => {
    console.error('Failed to start server:', err);
    process.exit(1);
  });
  