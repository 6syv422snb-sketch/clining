/**
 * Приём заявок с лендинга: запись в Google Таблицу + сообщение в Telegram.
 *
 * Настройки (Project Settings → Script properties):
 *   TELEGRAM_TOKEN  — токен бота от @BotFather
 *   TELEGRAM_CHAT   — id чата, куда слать заявки (ваш личный или группы)
 *   SHEET_ID        — id Google Таблицы (из адреса между /d/ и /edit)
 *   FOLDER_ID       — id папки Google Диска для фото (необязательно)
 *
 * Разверните как веб-приложение: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. URL вставьте в CONFIG.endpoint в index.html.
 */

const HEADERS = ['Дата', 'Тип', 'Город', 'Площадь, м²', 'Состояние', 'Створок окон',
  'Желаемая дата', 'Имя', 'Телефон', 'Комментарий', 'Ориентир', 'Фото', 'UTM', 'Страница', 'Статус'];

function doPost(e) {
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
    '\n\n⏱ Перезвонить в течение 5 минут';
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

/** Запустите вручную один раз из редактора, чтобы проверить Telegram и выдать разрешения. */
function testNotify() {
  notify(PropertiesService.getScriptProperties(), {
    type: 'Тест', city: 'Санкт-Петербург', area: 60, name: 'Проверка', phone: '+7 000 000-00-00', estimate: 'от 10 800 ₽'
  }, []);
}
