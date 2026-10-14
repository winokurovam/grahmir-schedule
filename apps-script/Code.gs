/**
 * Отдаёт сайту расписание из этой Google Таблицы в виде JSON.
 *
 * Как устроен каждый лист (один лист = один стол или зал):
 *   A1 — пусто; B1…H1 — дни недели (Пн, Вт, Ср, Чт, Пт, Сб, Вс);
 *   A2, A3, … — время слота, например «18:00–22:00» или «18-22»;
 *   в ячейке название компании — слот занят; пустая ячейка — свободен;
 *   ячейка залита красным — бронь отменена, снаружи слот свободен.
 * Листы, имя которых начинается с «_», сайт не показывает.
 *
 * Наружу уходит только «свободно / занято»: названия компаний сайт не видит.
 */

const DAY_NAMES = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'];

function doGet() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(sheet => !sheet.getName().startsWith('_'));

  const body = {
    updated: new Date().toISOString(),
    rooms: sheets.map(readSheet),
  };

  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet(sheet) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();

  const dayColumns = [];
  values[0].forEach((title, col) => {
    const day = DAY_NAMES.indexOf(String(title).trim().toLowerCase().slice(0, 2));
    if (col > 0 && day >= 0) dayColumns.push({ col, day });
  });

  const slots = [];
  for (let row = 1; row < values.length; row++) {
    const time = parseTime(values[row][0]);
    if (!time) continue;

    dayColumns.forEach(({ col, day }) => {
      const booked = String(values[row][col]).trim() !== '';
      const cancelled = isRed(backgrounds[row][col]);
      slots.push({ day, start: time.start, end: time.end, free: !booked || cancelled });
    });
  }

  return { name: sheet.getName(), slots };
}

// «18:00–22:00», «18-22», «9:30 — 13:30» → { start: '18:00', end: '22:00' }
function parseTime(text) {
  const m = String(text).match(/(\d{1,2})(?:[:.](\d{2}))?\s*[-–—]\s*(\d{1,2})(?:[:.](\d{2}))?/);
  if (!m) return null;
  const pad = (h, min) => `${String(h).padStart(2, '0')}:${min || '00'}`;
  return { start: pad(m[1], m[2]), end: pad(m[3], m[4]) };
}

// Любой оттенок красного из палитры Google Таблиц, включая светло-красный.
// Оранжевый, жёлтый и розово-фиолетовый красным не считаются.
function isRed(hex) {
  const m = String(hex).match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return false;
  const [r, g, b] = m.slice(1).map(x => parseInt(x, 16));
  return r >= 150 && r - g >= 35 && r - b >= 35 && Math.abs(g - b) <= 60;
}
