const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const cors = require('cors');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static('.'));

// Инициализация БД
const db = new sqlite3.Database('samples.db');

db.serialize(() => {
  db.run(`CREATE TABLE IF NOT EXISTS samples (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    article TEXT,
    name TEXT,
    brand TEXT,
    category TEXT,
    condition TEXT,
    comment TEXT,
    photo TEXT,
    responsible TEXT,
    date DATE
  )`);
});

// Получение записей с фильтром по дате
app.get('/api/samples', (req, res) => {
  const { startDate, endDate } = req.query;
  let query = 'SELECT * FROM samples';
  const params = [];

  if (startDate || endDate) {
    query += ' WHERE';
    if (startDate) {
      query += ' date >= ?';
      params.push(startDate);
    }
    if (endDate) {
      if (startDate) query += ' AND';
      query += ' date <= ?';
      params.push(endDate);
    }
  }
  query += ' ORDER BY date DESC';

  db.all(query, params, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// Добавление записи
app.post('/api/samples', (req, res) => {
  const { article, name, brand, category, condition, comment, photo, responsible, date } = req.body;
  db.run(
    `INSERT INTO samples (article, name, brand, category, condition, comment, photo, responsible, date)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [article, name, brand, category, condition, comment, photo, responsible, date],
    function(err) {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ id: this.lastID });
    }
  );
});

// Обновление записи
app.put('/api/samples/:id', (req, res) => {
  const { id } = req.params;
  const { article, name, brand, category, condition, comment, photo, responsible, date } = req.body;
  db.run(
    `UPDATE samples SET article = ?, name = ?, brand = ?, category = ?, condition = ?, comment = ?, photo = ?, responsible = ?, date = ? WHERE id = ?`,
    [article, name, brand, category, condition, comment, photo, responsible, date, id],
    function(err) {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ changes: this.changes });
    }
  );
});

// Удаление записи
app.delete('/api/samples/:id', (req, res) => {
  const { id } = req.params;
  db.run('DELETE FROM samples WHERE id = ?', [id], function(err) {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ changes: this.changes });
  });
});

// Экспорт в Excel (с фильтрацией)
app.get('/api/export-excel', async (req, res) => {
  const { startDate, endDate } = req.query;
  let query = 'SELECT * FROM samples';
  const params = [];

  if (startDate || endDate) {
    query += ' WHERE';
    if (startDate) {
      query += ' date >= ?';
      params.push(startDate);
    }
    if (endDate) {
      if (startDate) query += ' AND';
      query += ' date <= ?';
      params.push(endDate);
    }
  }

  db.all(query, params, async (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Учёт образцов');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 5 },
      { header: 'Артикул', key: 'article', width: 15 },
      { header: 'Название', key: 'name', width: 25 },
      { header: 'ТМ', key: 'brand', width: 15 },
      { header: 'Категория', key: 'category', width: 20 },
      { header: 'Состояние', key: 'condition', width: 25 },
      { header: 'Комментарий', key: 'comment', width: 30 },
      { header: 'Фото', key: 'photo', width: 15 },
      { header: 'Ответственный', key: 'responsible', width: 20 },
      { header: 'Дата', key: 'date', width: 12 }
    ];

    rows.forEach(row => worksheet.addRow(row));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Образцы.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  });
});

// Главная страница
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
