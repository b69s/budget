// === Синхронизация бюджета: вставить в Extensions → Apps Script твоей таблицы ===
// 1. Замени SECRET на свою случайную строку (та же строка вводится в приложении)
// 2. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
// 3. Скопируй URL деплоя (…/exec) в настройки приложения

const SECRET = 'CHANGE_ME_random_string';
const SHEET_NAME = 'Журнал';
const ID_COLUMN = 6; // колонка F — служебный ID записи для защиты от дублей

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) return json({ ok: false, error: 'bad secret' });

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: 'sheet not found' });

    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1
      ? new Set(sheet.getRange(2, ID_COLUMN, lastRow - 1, 1).getValues().flat().map(String))
      : new Set();

    const added = [];
    for (const en of body.entries || []) {
      if (existing.has(String(en.id))) { added.push(en.id); continue; } // уже есть — считаем успехом
      sheet.appendRow([
        new Date(en.date + 'T00:00:00'),
        en.type === 'inc' ? 'Доход' : 'Расход',
        en.cat,
        Number(en.val),
        en.note || '',
        String(en.id)
      ]);
      added.push(en.id);
    }
    return json({ ok: true, synced: added });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
