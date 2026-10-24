/**
 * Анкета гостей → Google-таблица.
 * Вставьте этот код в Google-таблице: Расширения → Apps Script,
 * затем «Начать развертывание → Новое развертывание → Веб-приложение»,
 * «Запуск от имени: я», «У кого есть доступ: все».
 * Каждый ответ записывается одной строкой на листе «Ответы».
 */
var SHEET_NAME = 'Ответы';
var HEADERS = [
  'Дата ответа', 'Приглашение', 'Ответ', 'Контактное лицо', 'Контакт',
  'Взрослых', 'Взрослые гости (латиницей)', 'Питание взрослых',
  'Детей', 'Дети (латиницей, возраст)', 'Питание детей',
  '7 июля — свадьба', '8 июля — White Pool Party',
  'Трансфер из Тбилиси', 'Дата приезда в Тбилиси', 'Пожелание / комментарий'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var yes = d.attending === 'yes';
    var adults = yes ? (d.adults || []) : [];
    var kids = yes ? (d.kids || []) : [];
    var food = function (p) { return p.name + ' — ' + (p.food || 'Без ограничений') + (p.allergy ? ' (' + p.allergy + ')' : ''); };
    var row = [
      new Date(),
      clean_(d.invite),
      yes ? 'Будем' : 'Не сможем',
      clean_(d.contactName),
      clean_(d.contact),
      adults.length,
      clean_(adults.map(function (a) { return a.name; }).join('\n')),
      clean_(adults.map(food).join('\n')),
      kids.length,
      clean_(kids.map(function (k) { return k.name + ' — ' + k.age; }).join('\n')),
      clean_(kids.map(food).join('\n')),
      yes ? (d.day1 ? 'Да' : 'Нет') : '',
      yes ? (d.day2 ? 'Да' : 'Нет') : '',
      yes ? (d.transfer === 'need' ? 'Нужен' : 'Доберутся сами') : '',
      clean_(d.arrival),
      clean_(d.wish)
    ];
    sheet_().appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Анкета работает' });
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F1E8DA').setWrap(true);
    sh.setColumnWidths(1, HEADERS.length, 190);
  }
  return sh;
}

// Text only, trimmed, and never treated as a formula by Sheets
function clean_(v) {
  v = v == null ? '' : String(v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
