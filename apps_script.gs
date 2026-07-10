// === Синхронизация бюджета (v2, двусторонняя) ===
// 1. Замени SECRET на свою случайную строку (та же вводится в приложении)
// 2. Deploy → New deployment → Web app → Execute as: Me → Access: Anyone
// 3. URL деплоя (…/exec) → в настройки приложения
// При обновлении кода: Deploy → Manage deployments → карандаш → New version

const SECRET = 'CHANGE_ME_random_string';
const SHEET_NAME = 'Журнал';
const ID_COL = 6; // колонка F — ID записи

// --- Приём записей из приложения ---
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) return json({ ok: false, error: 'bad secret' });
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: 'sheet not found' });

    const lastRow = sheet.getLastRow();
    const existing = lastRow > 1
      ? new Set(sheet.getRange(2, ID_COL, lastRow - 1, 1).getValues().flat().map(String))
      : new Set();

    const added = [];
    for (const en of body.entries || []) {
      if (!existing.has(String(en.id))) {
        sheet.appendRow([
          new Date(en.date + 'T00:00:00'),
          en.type === 'inc' ? 'Доход' : 'Расход',
          en.cat, Number(en.val), en.note || '', String(en.id)
        ]);
      }
      added.push(en.id);
    }
    return json({ ok: true, synced: added });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

// --- Отдача всего журнала в приложение ---
function doGet(e) {
  try {
    if (e.parameter.secret !== SECRET) return json({ ok: false, error: 'bad secret' });
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: 'sheet not found' });

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return json({ ok: true, entries: [] });

    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const range = sheet.getRange(2, 1, lastRow - 1, ID_COL);
    const rows = range.getValues();
    const out = [];
    let dirty = false;

    rows.forEach((r, i) => {
      const [date, typ, cat, val, note] = r;
      let id = String(r[5] || '').trim();
      if (!date || !val) return; // пустые строки пропускаем
      if (!id) { // строка добавлена руками в таблице — присваиваем ID
        id = 'sh' + Date.now() + i;
        rows[i][5] = id;
        dirty = true;
      }
      out.push({
        id: id,
        date: date instanceof Date ? Utilities.formatDate(date, tz, 'yyyy-MM-dd') : String(date),
        type: String(typ).trim() === 'Доход' ? 'inc' : 'exp',
        cat: String(cat),
        val: Number(val),
        note: String(note || '')
      });
    });

    if (dirty) range.setValues(rows); // дописываем ID новым ручным строкам
    return json({ ok: true, entries: out });
  } catch (err) { return json({ ok: false, error: String(err) }); }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
