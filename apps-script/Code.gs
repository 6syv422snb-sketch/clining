/**
 * Приём заявок с лендинга: запись в Google Таблицу + сообщение в Telegram.
 *
 * Настройки (Project Settings → Script properties):
 *   TELEGRAM_TOKEN  — токен бота от @BotFather
 *   TELEGRAM_CHAT   — id чата, куда слать заявки (ваш личный или группы)
 *   SHEET_ID        — id Google Таблицы (из адреса между /d/ и /edit)
 *   FOLDER_ID       — id папки Google Диска для фото (необязательно)
 *   WEBAPP_URL      — URL веб-приложения (…/exec), нужен для setupManager (необязательно,
 *                     если не задан — берётся из ScriptApp.getService().getUrl())
 *
 * Разверните как веб-приложение: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. URL вставьте в CONFIG.endpoint в index.html.
 *
 * Менеджер в Telegram: после публикации один раз запустите setupManager — под каждой заявкой
 * появятся кнопки статусов (пишутся в столбец «Статус»), а в чате заработают /new и /stat.
 */

const HEADERS = ['Дата', 'Тип', 'Город', 'Площадь, м²', 'Состояние', 'Створок окон',
  'Желаемая дата', 'Имя', 'Телефон', 'Комментарий', 'Ориентир', 'Фото', 'UTM', 'Страница', 'Статус',
  'Менеджер', 'Статус изменён'];
const COL = { date: 1, type: 2, city: 3, area: 4, name: 8, phone: 9, estimate: 11, status: 15, manager: 16, changed: 17 };
const TZ = 'Europe/Moscow';

/** Кнопки под заявкой: код (в callback_data) → [подпись кнопки, значение в столбце «Статус»]. */
const STATUSES = {
  work: ['🙋 Взял', 'В работе'],
  price: ['💬 Назвал цену', 'Цена названа'],
  expensive: ['💸 Дорого', 'Дорого'],
  order: ['✅ Заказ', 'Заказ'],
  noanswer: ['📵 Не дозвон', 'Не дозвонился'],
  lost: ['✖️ Отказ', 'Отказ']
};

function doPost(e) {
  if (e && e.parameter && e.parameter.tg) return telegramUpdate(e);
  try {
    const data = JSON.parse(e.postData.contents);
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

    let row = 0;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = leadSheet(props);
      if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
      sheet.appendRow([new Date(), data.type, data.city, data.area, data.condition, data.windows,
        data.date, data.name, "'" + data.phone, data.comment, data.estimate,
        photoLinks.join('\n') || (photos.length ? photos.length + ' фото в Telegram' : ''),
        JSON.stringify(data.utm || {}), data.page, 'Новая']);
      row = sheet.getLastRow();
    } finally {
      lock.releaseLock();
    }

    notify(props, data, photos, row);
    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: String(err) });
  }
}

function notify(props, d, photos, row) {
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
    payload: { chat_id: chat, text: text, parse_mode: 'HTML',
      reply_markup: row ? JSON.stringify(statusKeyboard(row, d.phone, '')) : '' }
  });
  if (photos.length) {
    const payload = { chat_id: chat, media: JSON.stringify(photos.map(function (_, i) {
      return { type: 'photo', media: 'attach://p' + i };
    })) };
    photos.forEach(function (b, i) { payload['p' + i] = b; });
    UrlFetchApp.fetch(api + 'sendMediaGroup', { method: 'post', payload: payload, muteHttpExceptions: true });
  }
}

/* ───────────────────────── Менеджер в Telegram ───────────────────────── */

function leadSheet(props) {
  return SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheets()[0];
}

function statusKeyboard(row, phone, current) {
  const ref = row + ':' + clean(phone).slice(-10);
  const buttons = Object.keys(STATUSES).map(function (code) {
    const label = STATUSES[code][0];
    return { text: STATUSES[code][1] === current ? '● ' + label : label, callback_data: 's:' + code + ':' + ref };
  });
  return { inline_keyboard: [buttons.slice(0, 3), buttons.slice(3)] };
}

/** Обработка webhook от Telegram. Отвечаем через HtmlService: он отдаёт 200 без редиректа,
 *  иначе Telegram считает доставку неудачной и повторяет её. */
function telegramUpdate(e) {
  const ok = HtmlService.createHtmlOutput('ok');
  try {
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty('TG_SECRET') || e.parameter.tg !== props.getProperty('TG_SECRET')) return ok;
    const u = JSON.parse(e.postData.contents);
    const chat = String(props.getProperty('TELEGRAM_CHAT')).trim();
    if (u.callback_query) onStatusButton(props, chat, u.callback_query);
    else if (u.message && u.message.text && String(u.message.chat.id) === chat) onCommand(props, chat, u.message);
  } catch (err) {
    console.error(err);
  }
  return ok;
}

function onStatusButton(props, chat, q) {
  const api = 'https://api.telegram.org/bot' + props.getProperty('TELEGRAM_TOKEN') + '/';
  const m = /^s:(\w+):(\d+):(\d*)$/.exec(q.data || '');
  if (!q.message || String(q.message.chat.id) !== chat || !m || !STATUSES[m[1]]) {
    return tg(api, 'answerCallbackQuery', { callback_query_id: q.id, text: 'Нет доступа' });
  }
  const status = STATUSES[m[1]][1];
  const who = [q.from.first_name, q.from.last_name].filter(Boolean).join(' ') || q.from.username || String(q.from.id);
  const now = new Date();

  let row;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = leadSheet(props);
    row = findLeadRow(sheet, Number(m[2]), m[3]);
    if (row) {
      if (!sheet.getRange(1, COL.manager).getValue()) {
        sheet.getRange(1, COL.manager, 1, 2).setValues([[HEADERS[COL.manager - 1], HEADERS[COL.changed - 1]]]);
      }
      sheet.getRange(row, COL.status, 1, 3).setValues([[status, who, now]]);
    }
  } finally {
    lock.releaseLock();
  }
  if (!row) return tg(api, 'answerCallbackQuery', { callback_query_id: q.id, text: 'Заявка не найдена в таблице', show_alert: true });

  // Дописываем строку статуса в исходное сообщение, сохраняя его форматирование.
  const base = String(q.message.text || '').split('\n\n🏷 ')[0];
  const entities = (q.message.entities || []).filter(function (en) { return en.offset + en.length <= base.length; });
  tg(api, 'editMessageText', {
    chat_id: chat, message_id: q.message.message_id,
    text: base + '\n\n🏷 ' + status + ' · ' + who + ', ' + Utilities.formatDate(now, TZ, 'dd.MM HH:mm'),
    entities: JSON.stringify(entities),
    reply_markup: JSON.stringify(statusKeyboard(row, m[3], status))
  });
  tg(api, 'answerCallbackQuery', { callback_query_id: q.id, text: 'Статус: ' + status });
}

/** Строка из кнопки, если телефон в ней совпадает; иначе (таблицу сортировали) — последняя заявка с этим телефоном. */
function findLeadRow(sheet, row, phone10) {
  const last = sheet.getLastRow();
  const phoneOf = function (v) { return clean(v).slice(-10); };
  if (row >= 2 && row <= last && phoneOf(sheet.getRange(row, COL.phone).getValue()) === phone10) return row;
  if (!phone10 || last < 2) return 0;
  const phones = sheet.getRange(2, COL.phone, last - 1, 1).getValues();
  for (let i = phones.length - 1; i >= 0; i--) if (phoneOf(phones[i][0]) === phone10) return i + 2;
  return 0;
}

function onCommand(props, chat, msg) {
  const cmd = msg.text.trim().split(/[\s@]/)[0].toLowerCase();
  let text;
  if (cmd === '/new') text = newLeadsReport(props);
  else if (cmd === '/stat') text = statReport(props);
  else if (cmd === '/help' || cmd === '/start') text = 'Команды:\n/new — необработанные заявки\n/stat — сводка за 7 и 30 дней\n\n' +
    'Статус заявки меняйте кнопками под сообщением — он сразу попадёт в таблицу.';
  else return;
  tg('https://api.telegram.org/bot' + props.getProperty('TELEGRAM_TOKEN') + '/', 'sendMessage',
    { chat_id: chat, text: text, parse_mode: 'HTML', disable_web_page_preview: true });
}

function leadRows(props) {
  const sheet = leadSheet(props);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, COL.status).getValues().map(function (r, i) {
    return { row: i + 2, date: r[COL.date - 1] instanceof Date ? r[COL.date - 1] : null, type: r[COL.type - 1],
      city: r[COL.city - 1], area: r[COL.area - 1], name: r[COL.name - 1], phone: String(r[COL.phone - 1]),
      estimate: r[COL.estimate - 1], status: String(r[COL.status - 1] || 'Новая') };
  });
}

function newLeadsReport(props) {
  const rows = leadRows(props).filter(function (r) { return r.status === 'Новая' || r.status === 'Не дозвонился'; });
  if (!rows.length) return '👌 Необработанных заявок нет';
  const now = Date.now();
  return '<b>Ждут ответа: ' + rows.length + '</b>\n\n' + rows.slice(-20).reverse().map(function (r) {
    const mins = r.date ? Math.round((now - r.date.getTime()) / 60000) : null;
    const age = mins == null ? '' : mins < 60 ? mins + ' мин' : mins < 1440 ? Math.round(mins / 60) + ' ч' : Math.round(mins / 1440) + ' дн';
    return (r.status === 'Новая' ? '🆕 ' : '📵 ') + esc(r.name) + ' · ' + esc(r.phone) + '\n' +
      '    ' + esc(r.type) + ', ' + esc(r.city) + ', ' + esc(r.area) + ' м²' + (age ? ' · ' + age + ' назад' : '');
  }).join('\n') + (rows.length > 20 ? '\n\n…показаны последние 20' : '');
}

function statReport(props) {
  const rows = leadRows(props);
  const period = function (days) {
    const from = Date.now() - days * 86400000;
    const list = rows.filter(function (r) { return r.date && r.date.getTime() >= from; });
    const by = {};
    list.forEach(function (r) { by[r.status] = (by[r.status] || 0) + 1; });
    const n = list.length;
    const pct = function (k) { return n ? Math.round((by[k] || 0) * 100 / n) + '%' : '—'; };
    const types = {};
    list.forEach(function (r) {
      const t = r.type || '—';
      types[t] = types[t] || { all: 0, order: 0 };
      types[t].all++;
      if (r.status === 'Заказ') types[t].order++;
    });
    return '<b>' + days + ' дней: ' + n + ' заявок</b>\n' +
      Object.keys(by).sort(function (a, b) { return by[b] - by[a]; }).map(function (k) { return '  ' + esc(k) + ': ' + by[k]; }).join('\n') +
      (n ? '\n  Конверсия в заказ: ' + pct('Заказ') + ' · «дорого»: ' + pct('Дорого') : '') +
      (Object.keys(types).length > 1 ? '\n  По типу: ' + Object.keys(types).map(function (t) {
        return esc(t) + ' ' + types[t].order + '/' + types[t].all;
      }).join(', ') : '');
  };
  return '📊 <b>ЧистоСпейс — сводка</b>\n\n' + period(7) + '\n\n' + period(30);
}

function tg(api, method, payload) {
  const r = UrlFetchApp.fetch(api + method, { method: 'post', payload: payload, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) console.log(method + ': ' + r.getContentText());
  return r;
}

/** Запустите вручную ОДИН раз после публикации веб-приложения (и после смены URL):
 *  подключает webhook бота к этому скрипту и регистрирует команды /new, /stat. */
function setupManager() {
  const props = PropertiesService.getScriptProperties();
  const token = (props.getProperty('TELEGRAM_TOKEN') || '').trim();
  const url = (props.getProperty('WEBAPP_URL') || ScriptApp.getService().getUrl() || '').trim();
  if (!token) return console.log('Нет TELEGRAM_TOKEN');
  if (!/\/exec$/.test(url)) return console.log('Нужен URL веб-приложения …/exec — задайте свойство WEBAPP_URL. Сейчас: ' + url);
  let secret = props.getProperty('TG_SECRET');
  if (!secret) props.setProperty('TG_SECRET', secret = Utilities.getUuid().replace(/-/g, ''));
  const api = 'https://api.telegram.org/bot' + token + '/';
  console.log('setWebhook: ' + tg(api, 'setWebhook', {
    url: url + '?tg=' + secret, allowed_updates: JSON.stringify(['message', 'callback_query']), drop_pending_updates: 'true'
  }).getContentText());
  console.log('setMyCommands: ' + tg(api, 'setMyCommands', { commands: JSON.stringify([
    { command: 'new', description: 'Необработанные заявки' },
    { command: 'stat', description: 'Сводка за 7 и 30 дней' },
    { command: 'help', description: 'Как работать с заявками' }
  ]) }).getContentText());
  console.log('getWebhookInfo: ' + UrlFetchApp.fetch(api + 'getWebhookInfo').getContentText());
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
