const ExcelJS = require('exceljs');
const MongoClient = require('mongodb').MongoClient;
const path = require('path');
const fs = require('fs');

const uri = process.env.MONGODB2_MONGODB_URI;
const client = new MongoClient(uri, { useUnifiedTopology: true });

async function migrateUsers() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(__dirname, 'users.xlsx'));
  const sheet = workbook.getWorksheet('Users');
  const users = [];

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > 1) {
      const username = String(row.values[1] || '').trim();
      const password = String(row.values[2] || '').trim();
      const role = username === 'admin' ? 'admin' : 'user';
      users.push({ username, password, role });
    }
  });

  await client.connect();
  const db = client.db();
  const usersCollection = db.collection('users');
  await usersCollection.deleteMany({}); // Очищаємо колекцію перед міграцією
  await usersCollection.insertMany(users);
  console.log(`Migrated ${users.length} users to MongoDB`);
}

async function migrateQuestions() {
  const questionFiles = fs.readdirSync(__dirname).filter(file => file.startsWith('questions') && file.endsWith('.xlsx'));
  const questions = [];

  for (const file of questionFiles) {
    const testNumber = parseInt(file.match(/questions(\d+)\.xlsx/)[1]);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, file));
    const sheet = workbook.getWorksheet('Questions');

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

        questions.push(questionData);
      }
    });
  }

  const questionsCollection = client.db().collection('questions');
  await questionsCollection.deleteMany({}); // Очищаємо колекцію перед міграцією
  await questionsCollection.insertMany(questions);
  console.log(`Migrated ${questions.length} questions to MongoDB`);
}

async function migrate() {
  try {
    await migrateUsers();
    await migrateQuestions();
  } catch (error) {
    console.error('Migration error:', error);
  } finally {
    await client.close();
  }
}

migrate();