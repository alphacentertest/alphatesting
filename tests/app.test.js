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
    { header: 'Points', key: 'points' }
  ];
  questionSheet.addRow({
    picture: '',
    question: 'Тестове питання',
    option1: 'Відповідь 1',
    option2: 'Відповідь 2',
    option3: 'Відповідь 3',
    option4: 'Відповідь 4',
    correctAnswer: 'Відповідь 1',
    type: 'multiple',
    points: 2
  });
  await questionsWorkbook.xlsx.writeFile(path.join(__dirname, '../test-questions1.xlsx'));

  // Запускаємо сервер
  server = app.listen(0); // Використовуємо випадковий порт
});

afterAll(async () => {
    // Закриваємо сервер
    await new Promise(resolve => server.close(resolve));
    
    // Очистка після тестів
    await client.close();
    await mongoServer.stop();
    // Видаляємо лише тестові файли
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
  });

  test('should redirect to / for unauthenticated user', async () => {
    const response = await agent.get('/select-test');
    expect(response.status).toBe(302);
    expect(response.headers.location).toBe('/');
  });
});
