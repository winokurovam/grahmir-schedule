/**
 * Отдаёт сайту расписание из этой Google Таблицы в виде JSON.
 *
 * Как устроен лист «Расписание» (блоки по неделям, один под другим):
 *   строка с днями недели: КОМНАТА | ВРЕМЯ | ПОНЕДЕЛЬНИК … ВОСКРЕСЕНЬЕ;
 *   строка с датами под ними: «12 октября», «13 октября», …;
 *   по две строки на комнату: ДЕНЬ и ВЕЧЕР, название комнаты в первом столбце.
 *
 * Что значит цвет ячейки:
 *   красный      — игра отменена, слот свободен;
 *   бирюзовый    — слот свободен;
 *   розовый      — идёт набор на игру, к ней можно присоединиться;
 *   другой цвет  — слот занят (кампания, аренда);
 *   без заливки и без текста — в это время клуб не работает, на сайте слота нет.
 *
 * Время берётся из текста ячейки («19:30-23:30 …»). У пустых ячеек — из соседних
 * комнат в тот же день и то же время суток, а если там пусто — из часов работы
 * клуба: WEEKDAY_TIME для будней, WEEKEND_TIME для субботы и воскресенья.
 *
 * Пост о наборе: ссылку на пост в Telegram (t.me/…) вставьте в примечание к розовой
 * ячейке или сделайте ссылкой сам текст ячейки. На сайте появится кнопка «Пост о наборе».
 *
 * Наружу уходит только «свободно / набор / занято». Названия кампаний и аренды
 * сайт не получает, кроме названий игр, на которые идёт набор.
 */

const SHEET_NAME = 'Расписание';
const WEEKDAY_TIME = {
  'день': ['13:00', '17:00'],
  'вечер': ['19:30', '23:30'],
};
const WEEKEND_TIME = {
  'день': ['13:00', '17:00'],
  'вечер': ['18:00', '22:00'],
};

const WEEKDAYS = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];
const MONTHS = [/^янв/, /^фев/, /^мар/, /^апр/, /^ма[йя]/, /^июн/, /^июл/, /^авг/, /^сен/, /^окт/, /^ноя/, /^дек/];

function doGet() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = book.getSheetByName(SHEET_NAME) || book.getSheets()[0];
  const range = sheet.getDataRange();
  const extras = range.getNotes().map((row, r) => row.map((note, c) => `${note} ${linkUrls(range, r, c)}`));
  const body = parseSchedule(range.getDisplayValues(), range.getBackgrounds(), new Date(), extras);
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ссылки, проставленные на тексте ячейки (Ctrl+K)
let richText = null;
function linkUrls(range, r, c) {
  if (!richText) richText = range.getRichTextValues();
  const rt = richText[r][c];
  return rt ? rt.getRuns().map(run => run.getLinkUrl()).filter(Boolean).join(' ') : '';
}

function parseSchedule(values, backgrounds, now, extras) {
  const cells = [];
  const rooms = [];
  let dayCols = null;   // столбец → день недели
  let dateCols = {};    // столбец → 'ГГГГ-ММ-ДД'
  let room = null;

  values.forEach((raw, r) => {
    const row = raw.map(v => String(v).trim());

    const weekdays = {};
    row.forEach((v, c) => {
      const i = WEEKDAYS.indexOf(v.toLowerCase());
      if (i >= 0) weekdays[c] = i;
    });
    if (Object.keys(weekdays).length >= 3) {
      dayCols = weekdays;
      dateCols = {};
      room = null;
      return;
    }
    if (!dayCols) return;

    const dates = {};
    Object.keys(dayCols).forEach(c => {
      const d = parseDate(row[c], now);
      if (d) dates[c] = d;
    });
    if (Object.keys(dates).length >= 3) {
      dateCols = dates;
      return;
    }

    const firstDayCol = Math.min.apply(null, Object.keys(dayCols).map(Number));
    const period = row.slice(0, firstDayCol).map(v => v.toLowerCase()).find(v => WEEKDAY_TIME[v]);
    if (!period) return;

    if (row[0]) {
      room = row[0];
      if (!rooms.some(x => x.name === room)) rooms.push({ name: room, color: backgrounds[r][0] });
    }
    if (!room) return;

    Object.keys(dateCols).forEach(c => {
      cells.push({
        date: dateCols[c],
        weekday: dayCols[c],
        room,
        period,
        text: row[c],
        kind: colorKind(backgrounds[r][c]),
        time: parseTime(row[c]),
        post: telegramLink(extras ? extras[r][c] : ''),
      });
    });
  });

  const slots = [];
  cells.forEach(cell => {
    let status;
    if (cell.kind === 'red' || cell.kind === 'cyan') status = 'free';
    else if (cell.kind === 'pink') status = 'recruit';
    else if (cell.text) status = 'busy';
    else return;

    const hours = cell.weekday >= 5 ? WEEKEND_TIME : WEEKDAY_TIME;
    const time = cell.time || timeFromNeighbours(cells, cell) || hours[cell.period];
    const slot = { date: cell.date, start: time[0], end: time[1], room: cell.room, status };
    if (status === 'recruit') {
      slot.title = gameTitle(cell.text);
      if (cell.post) slot.post = cell.post;
    }
    slots.push(slot);
  });

  slots.sort((a, b) => (a.date + a.start + a.room).localeCompare(b.date + b.start + b.room));
  return { updated: now.toISOString(), rooms, slots };
}

function timeFromNeighbours(cells, cell) {
  const same = cells.find(x => x.time && x.date === cell.date && x.period === cell.period);
  return same ? same.time : null;
}

// «19:30-23:30 …», «18:00 - 22:00 …», «18-22» → ['19:30', '23:30']
function parseTime(text) {
  const m = String(text).match(/(\d{1,2})(?:[:.](\d{2}))?\s*[-–—]\s*(\d{1,2})(?:[:.](\d{2}))?/);
  if (!m) return null;
  const pad = (h, min) => `${String(h).padStart(2, '0')}:${min || '00'}`;
  return [pad(m[1], m[2]), pad(m[3], m[4])];
}

// Первая ссылка на Telegram в тексте: «t.me/grahmir/123» → 'https://t.me/grahmir/123'.
// Остальной текст примечаний на сайт не уходит.
function telegramLink(text) {
  const m = String(text || '').match(/(?:https?:\/\/)?(?:t\.me|telegram\.me)\/[\w\/+-]+/i);
  if (!m) return null;
  return m[0].replace(/^(?:https?:\/\/)?/i, 'https://');
}

// Название игры без времени: «13:00 - 17:00 Campaign Beksultan» → «Campaign Beksultan»
function gameTitle(text) {
  return String(text).replace(/^\s*\d{1,2}(?:[:.]\d{2})?\s*[-–—]\s*\d{1,2}(?:[:.]\d{2})?\s*/, '').trim();
}

// «12 октября», «12.10», «12.10.2026» → '2026-10-12'. Год без указания — ближайший к сегодня.
function parseDate(text, now) {
  const s = String(text).trim().toLowerCase();
  let day, month, year = null;
  let m = s.match(/^(\d{1,2})\s+([а-яё]+)/);
  if (m) {
    day = Number(m[1]);
    month = MONTHS.findIndex(re => re.test(m[2]));
  } else if ((m = s.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?$/))) {
    day = Number(m[1]);
    month = Number(m[2]) - 1;
    if (m[3]) year = Number(m[3].length === 2 ? '20' + m[3] : m[3]);
  } else {
    return null;
  }
  if (month < 0 || month > 11 || day < 1 || day > 31) return null;

  if (year === null) {
    const y = now.getFullYear();
    year = [y - 1, y, y + 1].reduce((best, cand) =>
      Math.abs(new Date(cand, month, day) - now) < Math.abs(new Date(best, month, day) - now) ? cand : best);
  }
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

// Цвет заливки → 'red' | 'pink' | 'cyan' | 'other' | 'none'
function colorKind(hex) {
  const m = String(hex).match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return 'none';
  const [r, g, b] = m.slice(1).map(x => parseInt(x, 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  if (max === 0 || d / max < 0.12) return 'none';  // белый, серый, чёрный
  let hue;
  if (max === r) hue = 60 * (((g - b) / d) % 6);
  else if (max === g) hue = 60 * ((b - r) / d + 2);
  else hue = 60 * ((r - g) / d + 4);
  if (hue < 0) hue += 360;
  if (hue < 12 || hue >= 345) return 'red';
  if (hue >= 280) return 'pink';
  if (hue >= 165 && hue < 200) return 'cyan';
  return 'other';
}
