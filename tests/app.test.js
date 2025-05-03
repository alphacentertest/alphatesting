const request = require('supertest');
const app = require('../app');
const { MongoMemoryServer } = require('mongodb-memory-server');
const { MongoClient } = require('mongodb');
const bcrypt = require('bcrypt');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

let mongoServer;
let db;
let client;
let server;
let serverPort;

beforeAll(async () => {
  // Налаштування MongoDB в пам'яті
  mongoServer = await MongoMemoryServer.create();
  const uri = mongoServer.getUri();
  client = new MongoClient(uri);
  await client.connect();
  db = client.db('testdb');

  // Підміна глобального db у додатку
  global.db = db;

  // Створюємо фіктивний файл test-users.xlsx для тестів
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Users');
  sheet.columns = [
    { header: 'Username', key: 'username' },
    { header: 'Password', key: 'password' }
  ];
  sheet.addRow({ username: 'User1', password: 'pass111' });
  sheet.addRow({ username: 'admin', password: 'passadmin' });
  await workbook.xlsx.writeFile(path.join(__dirname, '../test-users.xlsx'));

  // Створюємо фіктивний файл test-questions1.xlsx
  const questionsWorkbook = new ExcelJS.Workbook();
  const questionSheet = questionsWorkbook.addWorksheet('Questions');
  questionSheet.columns = [
    { header: 'Picture', key: 'picture' },
    { header: 'Question', key: 'question' },
    { header: 'Option1', key: 'option1' },
    { header: 'Option2', key: 'option2' },
    { header: 'Option3', key: 'option3' },
    { header: 'Option4', key: 'option4' },
    { header: 'CorrectAnswer', key: 'correctAnswer' },
    { header: 'Type', key: 'type' },
    { header: 'Points', key: 'points' },
    { header: 'Variant', key: 'variant' }
  ];
  questionSheet.addRow({
    picture: '',
    question: 'Питання 1 (Variant 1)',
    option1: 'Відповідь 1',
    option2: 'Відповідь 2',
    option3: 'Відповідь 3',
    option4: 'Відповідь 4',
    correctAnswer: 'Відповідь 1',
    type: 'multiple',
    points: 2,
    variant: 'Variant 1'
  });
  questionSheet.addRow({
    picture: '',
    question: 'Питання 2 (Variant 2)',
    option1: 'Відповідь 1',
    option2: 'Відповідь 2',
    option3: 'Відповідь 3',
    option4: 'Відповідь 4',
    correctAnswer: 'Відповідь 1',
    type: 'multiple',
    points: 2,
    variant: 'Variant 2'
  });
  questionSheet.addRow({
    picture: '',
    question: 'Питання 3 (для всіх варіантів)',
    option1: 'Відповідь 1',
    option2: 'Відповідь 2',
    option3: 'Відповідь 3',
    option4: 'Відповідь 4',
    correctAnswer: 'Відповідь 1',
    type: 'multiple',
    points: 2,
    variant: ''
  });
  await questionsWorkbook.xlsx.writeFile(path.join(__dirname, '../test-questions1.xlsx'));

  // Запускаємо сервер і зберігаємо порт
  server = app.listen(0);
  serverPort = server.address().port;
});

afterAll(async () => {
  // Закриваємо сервер
  await new Promise(resolve => server.close(resolve));
  
  // Очистка після тестів
  await client.close();
  await mongoServer.stop();
  const usersFilePath = path.join(__dirname, '../test-users.xlsx');
  const questionsFilePath = path.join(__dirname, '../test-questions1.xlsx');
  if (fs.existsSync(usersFilePath)) {
    fs.unlinkSync(usersFilePath);
  }
  if (fs.existsSync(questionsFilePath)) {
    fs.unlinkSync(questionsFilePath);
  }
});

// Юніт-тести
describe('Unit Tests', () => {
  test('should hash password correctly', () => {
    const password = 'pass111';
    const hashedPassword = bcrypt.hashSync(password, 10);
    expect(bcrypt.compareSync(password, hashedPassword)).toBe(true);
  });

  test('should fail with incorrect password', () => {
    const password = 'pass111';
    const wrongPassword = 'pass222';
    const hashedPassword = bcrypt.hashSync(password, 10);
    expect(bcrypt.compareSync(wrongPassword, hashedPassword)).toBe(false);
  });
});

// Інтеграційні тести
describe('Integration Tests', () => {
  let agent;

  beforeEach(async () => {
    agent = request.agent(server);
    await agent.get('/');
  });

  test('should fail login with invalid password', async () => {
    const response = await agent
      .post('/login')
      .send({ password: 'wrongpass' })
      .set('Accept', 'application/json');
    expect(response.status).toBe(401);
    expect(response.body).toEqual({ success: false, message: 'Невірний пароль' });
  });

  test('should login successfully and redirect to /select-test', async () => {
    const response = await agent
      .post('/login')
      .send({ password: 'pass111' })
      .set('Accept', 'application/json');
    expect(response.status).toBe(200);
    expect(response.body).toEqual({ success: true, redirect: '/select-test' });

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(response.headers['set-cookie']).toBeDefined();
    expect(response.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));
  });

  test('should redirect to / for unauthenticated user', async () => {
    const response = await agent.get('/select-test');
    expect(response.status).toBe(302);
    expect(response.headers.location).toBe('/');
  });

  // Тест для перевірки відображення полів у /admin/edit-tests
  test('should display random order and question limit fields in /admin/edit-tests', async () => {
    // Логін як адмін
    const loginResponse = await agent
      .post('/login')
      .send({ password: 'passadmin' })
      .set('Accept', 'application/json');
    expect(loginResponse.status).toBe(200);

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(loginResponse.headers['set-cookie']).toBeDefined();
    expect(loginResponse.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));

    const response = await agent.get('/admin/edit-tests');
    expect(response.status).toBe(200);
    expect(response.text).toContain('Випадковий вибір питань');
    expect(response.text).toContain('Кількість питань');
    expect(response.text).toContain('name="random1"');
    expect(response.text).toContain('name="limit1"');
  });

  // Тест для перевірки збереження randomOrder і questionLimit
  test('should save random order and question limit in /admin/edit-tests', async () => {
    // Логін як адмін
    const loginResponse = await agent
      .post('/login')
      .send({ password: 'passadmin' })
      .set('Accept', 'application/json');
    expect(loginResponse.status).toBe(200);

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(loginResponse.headers['set-cookie']).toBeDefined();
    expect(loginResponse.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));

    const response = await agent
      .post('/admin/edit-tests')
      .send({
        test1: 'Тест 1',
        time1: 3600,
        random1: 'on',
        limit1: 2
      })
      .set('Accept', 'application/json');
    expect(response.status).toBe(200);
    expect(response.text).toContain('Назви, час, порядок та кількість питань тестів успішно оновлено');
  });

  // Тест для перевірки фільтрації питань за варіантом
  test('should filter questions by variant', async () => {
    // Логін як звичайний користувач
    const loginResponse = await agent
      .post('/login')
      .send({ password: 'pass111' })
      .set('Accept', 'application/json');
    expect(loginResponse.status).toBe(200);

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(loginResponse.headers['set-cookie']).toBeDefined();
    expect(loginResponse.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));

    // Переходимо на сторінку тесту
    const response = await agent.get('/test?test=1');
    expect(response.status).toBe(200);

    // Отримуємо variant із тестового маршруту
    const variantResponse = await agent.get('/get-test-variant');
    const variant = variantResponse.body.testVariant;

    const expectedQuestionCount = variant === 1 ? 2 : 1; // Variant 1: 2 питання (1 + для всіх), Variant 2: 1 питання, Variant 3: 1 питання
    const questionMatches = (response.text.match(/Питання \d+:/g) || []).length;
    expect(questionMatches).toBe(expectedQuestionCount);

    if (variant === 1) {
      expect(response.text).toContain('Питання 1 (Variant 1)');
      expect(response.text).toContain('Питання 3 (для всіх варіантів)');
      expect(response.text).not.toContain('Питання 2 (Variant 2)');
    } else if (variant === 2) {
      expect(response.text).toContain('Питання 2 (Variant 2)');
      expect(response.text).not.toContain('Питання 1 (Variant 1)');
    } else {
      expect(response.text).toContain('Питання 3 (для всіх варіантів)');
      expect(response.text).not.toContain('Питання 1 (Variant 1)');
      expect(response.text).not.toContain('Питання 2 (Variant 2)');
    }
  });

  // Тест для перевірки відображення варіанту у результатах
  test('should display variant in /admin/results', async () => {
    // Логін як звичайний користувач
    const loginResponseUser = await agent
      .post('/login')
      .send({ password: 'pass111' })
      .set('Accept', 'application/json');
    expect(loginResponseUser.status).toBe(200);

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(loginResponseUser.headers['set-cookie']).toBeDefined();
    expect(loginResponseUser.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));

    // Проходимо тест
    await agent.get('/test?test=1');

    // Отримуємо variant із тестового маршруту
    const variantResponse = await agent.get('/get-test-variant');
    const variant = variantResponse.body.testVariant;

    await agent
      .post('/save-result')
      .send({
        testNumber: '1',
        score: 10,
        totalPoints: 20,
        answers: { 1: 'Відповідь 1' },
        scoresPerQuestion: { 1: 10 },
        duration: 300,
        startTime: new Date().toISOString(),
        endTime: new Date().toISOString(),
        suspiciousActivity: {}
      })
      .set('Accept', 'application/json');

    // Логін як адмін
    const loginResponseAdmin = await agent
      .post('/login')
      .send({ password: 'passadmin' })
      .set('Accept', 'application/json');
    expect(loginResponseAdmin.status).toBe(200);

    // Перевіряємо, чи встановлена кука connect.sid у відповіді
    expect(loginResponseAdmin.headers['set-cookie']).toBeDefined();
    expect(loginResponseAdmin.headers['set-cookie'].some(cookie => cookie.includes('connect.sid'))).toBe(true);

    // Додаємо затримку, щоб сесія встигла зберегтися
    await new Promise(resolve => setTimeout(resolve, 100));

    const response = await agent.get('/admin/results');
    expect(response.status).toBe(200);
    expect(response.text).toContain(`<th>Варіант</th>`);
    expect(response.text).toContain(`<td>${variant}</td>`);
  });
});
