const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, HeadingLevel, BorderStyle,
  PageBreak, Footer, PageNumber, VerticalAlign, LevelFormat,
} = require('docx');

const FONT = 'Arial';
const W = 9746; // A4 width 11906 minus 2 x 1080 margins
const ACCENT = '00ADD8'; // Go blue
const LIGHT = 'E6F7FB';
const GREY = 'F2F2F2';
const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };

const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const p = (text, o = {}) => new Paragraph({ children: [t(text, o.run || {})], spacing: { after: 80 }, ...o.para });

function cell(children, width, o = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    columnSpan: o.span,
    children: Array.isArray(children) ? children : [children],
  });
}

function stage(title, period, goal, tasks, checkpoint) {
  const cols = [520, 7026, 2200];
  const head = new TableRow({
    tableHeader: true,
    children: [
      cell(new Paragraph({ alignment: AlignmentType.CENTER, children: [t('✓', { bold: true, color: 'FFFFFF' })] }), cols[0], { fill: ACCENT }),
      cell(new Paragraph({ children: [t('Задача', { bold: true, color: 'FFFFFF', size: 20 })] }), cols[1], { fill: ACCENT }),
      cell(new Paragraph({ children: [t('Дата / заметка', { bold: true, color: 'FFFFFF', size: 20 })] }), cols[2], { fill: ACCENT }),
    ],
  });
  const rows = tasks.map((task, i) => {
    const isProject = task.startsWith('ПРОЕКТ');
    return new TableRow({
      cantSplit: true,
      children: [
        cell(new Paragraph({ alignment: AlignmentType.CENTER, children: [t('☐', { size: 26 })] }), cols[0]),
        cell(new Paragraph({ children: [t(task, { size: 20, bold: isProject })] }), cols[1], { fill: isProject ? LIGHT : (i % 2 ? GREY : undefined) }),
        cell(new Paragraph({ children: [t('')] }), cols[2], { fill: isProject ? LIGHT : (i % 2 ? GREY : undefined) }),
      ],
    });
  });
  return [
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t(title)] }),
    new Paragraph({ spacing: { after: 60 }, children: [t('Сроки: ', { bold: true, size: 20 }), t(period, { size: 20 })] }),
    new Paragraph({ spacing: { after: 120 }, children: [t('Цель этапа: ', { bold: true, size: 20 }), t(goal, { size: 20 })] }),
    new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols, rows: [head, ...rows] }),
    new Paragraph({
      spacing: { before: 120, after: 240 },
      border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 8 } },
      children: [t('Контрольная точка: ', { bold: true, size: 20 }), t(checkpoint, { size: 20, italics: true })],
    }),
  ];
}

const bullet = (text) => new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 60 }, children: [t(text, { size: 21 })] });
const bulletRich = (runs) => new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 60 }, children: runs });

// ---------- Cover / overview ----------
const overview = [
  new Paragraph({ heading: HeadingLevel.TITLE, children: [t('Мой путь в Go-разработку')] }),
  new Paragraph({ spacing: { after: 200 }, children: [t('Декабрь 2026 → август 2028  ·  3–4 часа в день  ·  первый язык программирования', { size: 22, color: '595959' })] }),
  new Paragraph({
    spacing: { after: 240 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: LIGHT },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: ACCENT, space: 8 } },
    children: [t('ЦЕЛЬ: ', { bold: true, size: 26 }), t('к августу 2028 года получить оффер Go-разработчика с зарплатой 250–300 тыс. ₽', { size: 26 })],
  }),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t('Карта пути')] }),
];

const mapCols = [2000, 2400, 3546, 1800];
const mapRows = [
  ['0. SQL и основы', 'дек 2026 – янв 2027', 'SQL, PostgreSQL, CS50, GitHub', '~200 ч'],
  ['1. Основы Go', 'фев – май 2027', 'Синтаксис, срезы, структуры, ошибки', '~400 ч'],
  ['2. Go глубже', 'июнь – сен 2027', 'Интерфейсы, горутины, тесты, HTTP', '~400 ч'],
  ['3. Инструменты', 'окт 2027 – янв 2028', 'Docker, Redis, очереди, gRPC', '~400 ч'],
  ['4. Витрина', 'фев – май 2028', 'Большой проект, алгоритмы, отклики', '~400 ч'],
  ['5. Финал', 'июнь – авг 2028', 'Собеседования и оффер', '~300 ч'],
];
overview.push(new Table({
  width: { size: W, type: WidthType.DXA },
  columnWidths: mapCols,
  rows: [
    new TableRow({ tableHeader: true, children: ['Этап', 'Сроки', 'Главное', 'Часы'].map((h, i) =>
      cell(new Paragraph({ children: [t(h, { bold: true, color: 'FFFFFF', size: 20 })] }), mapCols[i], { fill: ACCENT })) }),
    ...mapRows.map((r, ri) => new TableRow({ children: r.map((v, i) =>
      cell(new Paragraph({ children: [t(v, { size: 20, bold: i === 0 })] }), mapCols[i], { fill: ri % 2 ? GREY : undefined })) })),
  ],
}));

overview.push(
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t('Мой день (3–4 часа)')] }),
  bulletRich([t('~2 часа — практика: ', { bold: true, size: 21 }), t('пишу код и решаю задачи сам', { size: 21 })]),
  bulletRich([t('~1 час — теория: ', { bold: true, size: 21 }), t('книга, курс, статья', { size: 21 })]),
  bulletRich([t('~30 минут — повторение: ', { bold: true, size: 21 }), t('записываю своими словами, что понял', { size: 21 })]),
  bulletRich([t('1 день в неделю — отдых ', { bold: true, size: 21 }), t('или лёгкое повторение', { size: 21 })]),
  new Paragraph({ heading: HeadingLevel.HEADING_2, children: [t('Мои правила')] }),
  bullet('На каждый час видео или чтения — не меньше часа своего кода.'),
  bullet('Цель на неделю маленькая и конкретная, а не «выучить Go».'),
  bullet('Всё, что пишу, отправляю на GitHub — даже мелочи.'),
  bullet('Застрял больше чем на час — спрашиваю (Claude, чат, Stack Overflow).'),
  bullet('Раз в квартал проверяю прогресс. Отстаю — меняю план, а не бросаю.'),
  bullet('Отмечаю каждый день занятий в трекере в конце плана. Цепочку не рву.'),
);

// ---------- Stages ----------
const s0 = stage('Этап 0. SQL и основы мышления', 'декабрь 2026 – январь 2027',
  'уверенно писать SQL-запросы и понять, как вообще работает программирование.', [
  'w3schools: индексы, автоинкремент, даты, представления',
  'w3schools: SQL-инъекции, параметры, подготовленные запросы',
  'w3schools: упражнения, тест и задания по написанию SQL',
  'Установить PostgreSQL (или найти онлайн-песочницу)',
  'sql-ex.ru — решить 30+ задач',
  'pgexercises.com — пройти основные разделы',
  'Транзакции: BEGIN / COMMIT / ROLLBACK',
  'Нормализация таблиц: 1НФ, 2НФ, 3НФ',
  'Основы EXPLAIN: как читать план запроса',
  'CS50: первые 3 недели курса',
  'Завести аккаунт на GitHub',
  'ПРОЕКТ: спроектировать БД магазина (пользователи, заказы, товары) + 20 запросов',
  'Пройти тест по SQL у Claude и закрыть пробелы',
], 'сам проектирую 3–4 связанные таблицы и без подсказок пишу запросы с JOIN, GROUP BY, HAVING.');

const s1 = stage('Этап 1. Основы Go', 'февраль – май 2027',
  'писать простые программы на Go без подсказок.', [
  'Установить Go и редактор (VS Code или GoLand)',
  'Tour of Go (go.dev/tour) — пройти целиком',
  'Курс Яндекса «Основы Go»',
  'Go by Example — прорешать все примеры руками',
  'Типы, переменные, условия, циклы, функции',
  'Массивы, срезы (slices), мапы',
  'Структуры и методы',
  'Указатели',
  'Обработка ошибок (error, errors.Is / As)',
  'Пакеты и модули (go mod)',
  'Effective Go — прочитать',
  'Шпаргалка Go cheat sheet — распечатать / держать под рукой',
  'Exercism: 30+ задач на Go',
  'ПРОЕКТ: CLI-трекер задач с сохранением в файл',
], 'могу без интернета написать консольную программу со структурами, срезами и обработкой ошибок.');

const s2 = stage('Этап 2. Go глубже и первые проекты', 'июнь – сентябрь 2027',
  'понять то, о чём спрашивают на собеседованиях по Go, и написать первый веб-сервис.', [
  'Интерфейсы (статья «Изучаем Golang: интерфейсы»)',
  'Горутины, каналы, select',
  'sync: Mutex, WaitGroup, Once',
  'Пакет context',
  'Тесты: пакет testing, табличные тесты',
  'Learn Go with Tests — пройти основную часть',
  'Стандартная библиотека: net/http, encoding/json, io, os, time',
  'Статья про Go JSON и неправильную типизацию в API',
  'database/sql и pgx, параметризованные запросы',
  'Дженерики (generics) — основы',
  'Git: ветки, коммиты, merge, pull request',
  'ПРОЕКТ: REST API + PostgreSQL + тесты (например, сервис заметок)',
], 'объясняю, как работают горутины и каналы, и мой REST API работает с базой и покрыт тестами.');

const s3 = stage('Этап 3. Инструменты бэкенда', 'октябрь 2027 – январь 2028',
  'владеть инструментами, которые требуют в вакансиях.', [
  'Linux: терминал, файлы, процессы, права',
  'Docker и docker-compose',
  'Миграции БД (goose или golang-migrate)',
  'Redis: кэширование',
  'Kafka или RabbitMQ — основы',
  'gRPC и protobuf',
  'Логирование и метрики (Prometheus)',
  'CI: GitHub Actions (тесты при каждом коммите)',
  'Статья «Чистая архитектура и как её готовить»',
  'Книга «100 Go Mistakes and How to Avoid Them»',
  'Выписать требования из 20 вакансий Go-junior на hh.ru',
  'ПРОЕКТ: сервис в Docker с PostgreSQL, Redis и очередью',
], 'поднимаю свой проект одной командой docker-compose up, в нём есть кэш, очередь и CI.');

const s4 = stage('Этап 4. Витринный проект и подготовка', 'февраль – май 2028',
  'иметь проект, который не стыдно показать, и начать откликаться.', [
  'ПРОЕКТ-ВИТРИНА: продуманный сервис с деплоем на VPS',
  'README проекта: описание, схема, как запустить',
  'Алгоритмы: 50+ задач LeetCode (Easy / Medium)',
  'Сети: TCP/IP, HTTP, DNS — основы',
  'System design — основы',
  'Резюме + профиль на hh.ru + аккуратный GitHub',
  'Первые отклики на junior-вакансии (с февраля)',
  'Пробные собеседования (с Claude, с друзьями)',
], 'на GitHub 3+ проекта, резюме готово, прошёл первые собеседования.');

const s5 = stage('Этап 5. Финал: работа', 'июнь – август 2028',
  'получить оффер.', [
  '10+ откликов в неделю',
  'После каждого собеседования — записать вопросы, на которые не ответил',
  'Повторять частые вопросы по Go, SQL и конкурентности',
  'Подготовиться к разговору о зарплате',
  'ПРОЕКТ ЖИЗНИ: оффер 250–300 тыс. ₽',
], 'я Go-разработчик.');

// ---------- Quarterly review ----------
const reviewCols = [1900, 2600, 2600, 2646];
const quarters = ['март 2027', 'июнь 2027', 'сентябрь 2027', 'декабрь 2027', 'март 2028', 'июнь 2028', 'август 2028'];
const review = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t('Проверка раз в квартал')] }),
  p('Честно отвечаю: что получилось, что нет, что меняю в плане.', { run: { size: 20, color: '595959' } }),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: reviewCols,
    rows: [
      new TableRow({ tableHeader: true, children: ['Дата', 'Что получилось', 'Что не получилось', 'Что меняю'].map((h, i) =>
        cell(new Paragraph({ children: [t(h, { bold: true, color: 'FFFFFF', size: 20 })] }), reviewCols[i], { fill: ACCENT })) }),
      ...quarters.map(q => new TableRow({ height: { value: 1250, rule: 'atLeast' }, cantSplit: true, children: [
        cell(new Paragraph({ children: [t(q, { bold: true, size: 20 })] }), reviewCols[0], { fill: GREY }),
        cell(new Paragraph(''), reviewCols[1]), cell(new Paragraph(''), reviewCols[2]), cell(new Paragraph(''), reviewCols[3]),
      ] })),
    ],
  }),
];

// ---------- Habit tracker ----------
const months = ['Дек 2026', 'Янв 2027', 'Фев 2027', 'Мар 2027', 'Апр 2027', 'Май 2027', 'Июн 2027', 'Июл 2027',
  'Авг 2027', 'Сен 2027', 'Окт 2027', 'Ноя 2027', 'Дек 2027', 'Янв 2028', 'Фев 2028', 'Мар 2028', 'Апр 2028',
  'Май 2028', 'Июн 2028', 'Июл 2028', 'Авг 2028'];
const mCol = 1102, dCol = 278;
const trackCols = [mCol, ...Array(31).fill(dCol)];
const trackW = mCol + 31 * dCol;
const small = (s, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, children: [t(s, { size: 13, ...o })] });
const tCell = (child, width, fill) => new TableCell({
  width: { size: width, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
  margins: { top: 0, bottom: 0, left: 20, right: 20 },
  shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
  children: [child],
});
const tracker = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [t('Трекер занятий')] }),
  p('Закрашиваю или ставлю ✕ в клетке за каждый день, когда занимался. Цель — не рвать цепочку.', { run: { size: 20, color: '595959' } }),
  new Table({
    width: { size: trackW, type: WidthType.DXA },
    columnWidths: trackCols,
    rows: [
      new TableRow({ tableHeader: true, children: [
        tCell(small('Месяц', { bold: true, color: 'FFFFFF' }), mCol, ACCENT),
        ...Array.from({ length: 31 }, (_, i) => tCell(small(String(i + 1), { bold: true, color: 'FFFFFF' }), dCol, ACCENT)),
      ] }),
      ...months.map((m, ri) => new TableRow({ height: { value: 520, rule: 'exact' }, children: [
        tCell(new Paragraph({ children: [t(m, { size: 15, bold: true })] }), mCol, ri % 2 ? GREY : undefined),
        ...Array.from({ length: 31 }, () => tCell(small(''), dCol)),
      ] })),
    ],
  }),
];

const doc = new Document({
  creator: 'Go plan',
  title: 'Мой путь в Go-разработку',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', next: 'Normal',
        run: { font: FONT, size: 44, bold: true, color: '1F1F1F' }, paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 30, bold: true, color: '007D9C' }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { font: FONT, size: 25, bold: true, color: '1F1F1F' }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [t('Go-план  ·  стр. ', { size: 16, color: '808080' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children: [
      ...overview,
      new Paragraph({ children: [new PageBreak()] }),
      ...s0, ...s1, ...s2, ...s3, ...s4, ...s5,
      ...review,
      ...tracker,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => fs.writeFileSync(process.argv[2] || 'Go_plan.docx', buf));
