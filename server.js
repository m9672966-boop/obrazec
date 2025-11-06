const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());

// Проверка email (опционально)
const allowedDomains = ['panna.ru', 'firma-gamma.ru', 'sb-service.ru'];

app.get('/api/export-sample-table', (req, res) => {
  const userEmail = req.headers['x-user-email'];
  if (userEmail) {
    const domain = userEmail.split('@')[1];
    if (!allowedDomains.includes(domain)) {
      return res.status(403).json({ error: 'Доступ запрещён' });
    }
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Учёт образцов');

  sheet.columns = [
    { header: '№', key: 'id', width: 5 },
    { header: 'Артикул', key: 'article', width: 15 },
    { header: 'Название', key: 'name', width: 25 },
    { header: 'Торговая марка', key: 'brand', width: 15 },
    { header: 'Категория', key: 'category', width: 20 },
    { header: 'Состояние', key: 'condition', width: 25 },
    { header: 'Комментарий', key: 'comment', width: 30 },
    { header: 'Фото', key: 'photo', width: 15 },
    { header: 'Ответственный', key: 'responsible', width: 20 },
    { header: 'Дата приёмки', key: 'date', width: 12 }
  ];

  // Можно добавить пример строки
  sheet.addRow({
    id: '',
    article: '',
    name: '',
    brand: '',
    category: 'Леонардо / Сотрудникам / Утиль',
    condition: '',
    comment: '',
    photo: '',
    responsible: '',
    date: ''
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=Таблица_учёта_образцов.xlsx');

  workbook.xlsx.write(res).then(() => res.end());
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
