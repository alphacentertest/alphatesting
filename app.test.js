const request = require('supertest');
const app = require('./app');
const { MongoClient } = require('mongodb');

const MONGO_URL = process.env.MONGO_URL || 'mongodb+srv://romanhaleckij7:DNMaH9w2X4gel3Xc@cluster0.r93r1p8.mongodb.net/testdb?retryWrites=true&w=majority';
const client = new MongoClient(MONGO_URL);

// Налаштування перед тестами
beforeAll(async () => {
  // Переконайтеся, що сервер ініціалізовано перед тестами
  await new Promise(resolve => setTimeout(resolve, 15000));
}, 15000);

// Закриваємо з’єднання після тестів
afterAll(async () => {
  await client.close();
});

// Тести для маршруту /api/test
describe('GET /api/test', () => {
  it('should return success message', async () => {
    const res = await request(app).get('/api/test');
    expect(res.statusCode).toEqual(200);
    expect(res.body).toHaveProperty('success', true);
    expect(res.body).toHaveProperty('message', 'Express server is working on /api/test');
  });
});

// Тести для маршруту /
describe('GET /', () => {
  it('should serve login page for unauthenticated users', async () => {
    const res = await request(app).get('/');
    expect(res.statusCode).toEqual(200);
    expect(res.text).toContain('<h1>Введіть будь ласка пароль для входу (Updated Version 3)</h1>');
  });
});

// Тести для маршруту /login
describe('POST /login', () => {
  it('should return 400 for missing password', async () => {
    const res = await request(app)
      .post('/login')
      .send({ _csrf: 'invalid-token' });
    expect(res.statusCode).toEqual(400);
    expect(res.body).toHaveProperty('message', 'Пароль має бути довжиною не менше 6 символів і містити лише латинські літери та цифри');
  });

  it('should return 403 for invalid CSRF token', async () => {
    const res = await request(app)
      .post('/login')
      .send({ password: 'pass111', _csrf: 'invalid-token' });
    expect(res.statusCode).toEqual(403);
    expect(res.body).toHaveProperty('message', 'Недійсний CSRF-токен');
  });
});
