/**
 * Приём заявок с лендинга: запись в Google Таблицу + сообщение в Telegram.
 *
 * Настройки (Project Settings → Script properties):
 *   TELEGRAM_TOKEN  — токен бота от @BotFather
 *   TELEGRAM_CHAT   — id чата, куда слать заявки (ваш личный или группы)
 *   SHEET_ID        — id Google Таблицы (из адреса между /d/ и /edit)
 *   FOLDER_ID       — id папки Google Диска для фото (необязательно)
 *   MANAGER_KEY     — пароль для панели менеджера manager.html (длинная случайная строка)
 *
 * Разверните как веб-приложение: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. URL вставьте в CONFIG.endpoint в index.html.
 */

const HEADERS = ['Дата', 'Тип', 'Город', 'Площадь, м²', 'Состояние', 'Створок окон',
  'Желаемая дата', 'Имя', 'Телефон', 'Комментарий', 'Ориентир', 'Фото', 'UTM', 'Страница', 'Статус',
  'Заметка', 'Сумма заказа'];
const STATUS_COL = 15, NOTE_COL = 16, AMOUNT_COL = 17;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action) return manager(data);
    if (data.website) return json({ ok: true }); // бот

    const props = PropertiesService.getScriptProperties();
    const photos = (data.photos || []).slice(0, 5).map(function (p, i) {
      return Utilities.newBlob(Utilities.base64Decode(p.data), p.type || 'image/jpeg', p.name || ('photo' + (i + 1) + '.jpg'));
    });

    let photoLinks = [];
    const folderId = props.getProperty('FOLDER_ID');
    if (folderId && photos.length) {
      const folder = DriveApp.getFolderById(folderId);
      const stamp = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd_HH-mm');
      photoLinks = photos.map(function (b, i) {
        return folder.createFile(b.copyBlob().setName(stamp + '_' + clean(data.phone) + '_' + (i + 1) + '.jpg')).getUrl();
      });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheets()[0];
      if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
      sheet.appendRow([new Date(), data.type, data.city, data.area, data.condition, data.windows,
        data.date, data.name, "'" + data.phone, data.comment, data.estimate,
        photoLinks.join('\n') || (photos.length ? photos.length + ' фото в Telegram' : ''),
        JSON.stringify(data.utm || {}), data.page, 'Новая']);
    } finally {
      lock.releaseLock();
    }

    notify(props, data, photos);
    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  }
}

/** Панель менеджера: список заявок и смена статуса. Доступ — по MANAGER_KEY. */
function manager(data) {
  const key = (PropertiesService.getScriptProperties().getProperty('MANAGER_KEY') || '').trim();
  if (!key) return json({ ok: false, error: 'MANAGER_KEY не задан в свойствах скрипта' });
  if (String(data.key || '') !== key) {
    Utilities.sleep(800);
    return json({ ok: false, error: 'Неверный ключ', auth: false });
  }
  const sheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheets()[0];
  ensureHeaders(sheet);

  if (data.action === 'list') {
    const last = sheet.getLastRow();
    if (last < 2) return json({ ok: true, leads: [] });
    const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
    const leads = rows.map(function (r, i) {
      let utm = {};
      try { utm = JSON.parse(r[12] || '{}'); } catch (err) {}
      return {
        row: i + 2, date: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        type: r[1], city: r[2], area: r[3], condition: r[4], windows: r[5], wishDate: fmt(r[6]),
        name: r[7], phone: String(r[8]).replace(/^'/, ''), comment: r[9], estimate: r[10],
        photos: String(r[11] || ''), utm: utm, page: r[13],
        status: r[14] || 'Новая', note: r[15], amount: r[16]
      };
    }).filter(function (l) { return l.date || l.phone; });
    return json({ ok: true, leads: leads });
  }

  if (data.action === 'update') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const row = Number(data.row);
      if (!(row >= 2 && row <= sheet.getLastRow())) return json({ ok: false, error: 'Нет такой строки' });
      const d = sheet.getRange(row, 1).getValue();
      const stamp = d instanceof Date ? d.toISOString() : String(d);
      if (data.date && stamp !== data.date) {
        return json({ ok: false, error: 'Таблица изменилась (сортировка или удаление строк). Обновите список.' });
      }
      if (data.status != null) sheet.getRange(row, STATUS_COL).setValue(String(data.status));
      if (data.note != null) sheet.getRange(row, NOTE_COL).setValue(String(data.note));
      if (data.amount != null) sheet.getRange(row, AMOUNT_COL).setValue(data.amount === '' ? '' : Number(data.amount) || '');
      return json({ ok: true });
    } finally {
      lock.releaseLock();
    }
  }
  return json({ ok: false, error: 'Неизвестное действие' });
}

function ensureHeaders(sheet) {
  const width = sheet.getLastColumn();
  if (sheet.getLastRow() === 0) return sheet.appendRow(HEADERS);
  if (width < HEADERS.length) {
    sheet.getRange(1, width + 1, 1, HEADERS.length - width).setValues([HEADERS.slice(width)]);
  }
}

function fmt(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Europe/Moscow', 'dd.MM.yyyy') : String(v || '');
}

function notify(props, d, photos) {
  const token = props.getProperty('TELEGRAM_TOKEN');
  const chat = props.getProperty('TELEGRAM_CHAT');
  if (!token || !chat) return;
  const api = 'https://api.telegram.org/bot' + token + '/';
  const utm = d.utm && d.utm.utm_source ? '\nИсточник: ' + esc(d.utm.utm_source + ' / ' + (d.utm.utm_campaign || '')) : '';
  const text = '🧹 <b>Новая заявка — ' + esc(d.type) + '</b>\n' +
    '📍 ' + esc(d.city) + ', ' + esc(d.area) + ' м²' +
    (d.condition ? ', ' + esc(d.condition) : '') + (d.windows ? ', окон: ' + esc(d.windows) : '') + '\n' +
    (d.date ? '📅 ' + esc(d.date) + '\n' : '') +
    '👤 ' + esc(d.name) + '\n📞 ' + esc(d.phone) + '\n' +
    (d.comment ? '💬 ' + esc(d.comment) + '\n' : '') +
    '💰 Ориентир: ' + esc(d.estimate) + utm +
    '\n\n⏱ Связаться в течение 5 минут';
  UrlFetchApp.fetch(api + 'sendMessage', {
    method: 'post', muteHttpExceptions: true,
    payload: { chat_id: chat, text: text, parse_mode: 'HTML' }
  });
  if (photos.length) {
    const payload = { chat_id: chat, media: JSON.stringify(photos.map(function (_, i) {
      return { type: 'photo', media: 'attach://p' + i };
    })) };
    photos.forEach(function (b, i) { payload['p' + i] = b; });
    UrlFetchApp.fetch(api + 'sendMediaGroup', { method: 'post', payload: payload, muteHttpExceptions: true });
  }
}

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function clean(s) { return String(s || '').replace(/\D/g, ''); }
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Запустите вручную из редактора: проверит настройки и Telegram, результат — в «Журнале выполнения». */
function testNotify() {
  const props = PropertiesService.getScriptProperties();
  const token = (props.getProperty('TELEGRAM_TOKEN') || '').trim();
  const chat = (props.getProperty('TELEGRAM_CHAT') || '').trim();
  const sheetId = (props.getProperty('SHEET_ID') || '').trim();
  console.log('Свойства скрипта: ' + JSON.stringify(props.getKeys()));
  console.log('TELEGRAM_TOKEN: ' + (token ? 'есть, начинается с ' + token.split(':')[0] : 'НЕТ'));
  console.log('TELEGRAM_CHAT: ' + (chat || 'НЕТ'));
  console.log('SHEET_ID: ' + (sheetId || 'НЕТ'));
  if (!token || !chat) return console.log('Не хватает TELEGRAM_TOKEN или TELEGRAM_CHAT — проверьте названия свойств.');

  const me = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getMe', { muteHttpExceptions: true });
  console.log('Проверка токена: ' + me.getContentText());
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', muteHttpExceptions: true,
    payload: { chat_id: chat, text: '🧹 Тест: заявки ЧистоСпейс будут приходить сюда' }
  });
  console.log('Отправка сообщения: ' + r.getContentText());

  if (sheetId) {
    try { console.log('Таблица найдена: ' + SpreadsheetApp.openById(sheetId).getName()); }
    catch (err) { console.log('Таблица НЕ открывается: ' + err); }
  }
}
