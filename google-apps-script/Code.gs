/**
 * RDNK Slovakia Quiz — Google Apps Script (doPost)
 *
 * Установка:
 * 1. Откройте таблицу → Расширения → Apps Script
 * 2. Замените код этим файлом
 * 3. Создайте листы (или скрипт создаст сам):
 *    - "Tier-1 Leads 🟢"
 *    - "Tier-2 Leads 🟡"
 *    - "Tier-3 Leads 🔴"
 * 4. Развернуть → Новое развертывание → Веб-приложение
 *    (выполняется от имени вас, доступ: «У всех»)
 * 5. Если URL изменился — обновите WEB_APP_URL в v2/index.html
 *
 * Совместимость:
 * - Старый квиз (v1) без quizVersion/tier → пишет в активный лист, как раньше
 * - Новый квиз (v2) → раскладывает лиды по Tier-листам
 */

var TIER_SHEETS = {
  1: 'Tier-1 Leads 🟢',
  2: 'Tier-2 Leads 🟡',
  3: 'Tier-3 Leads 🔴'
};

var TIER_HEADERS = [
  'Дата и время',
  'Имя',
  'Телефон / Telegram',
  'Ветка',
  'Tier',
  'Когда выехали',
  'Возраст на момент выезда',
  'Пересечение границы',
  'Семейное положение',
  'Страна проживания (план)',
  'Документы в ЕС',
  'Основной запрос',
  'Путь'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // --- Legacy quiz (v1): не меняем поведение ---
    if (!isV2Lead(data)) {
      return handleLegacyLead(data);
    }

    // --- Quiz v2: распределение по Tier-листам ---
    return handleV2Lead(data);
  } catch (error) {
    Logger.log(error.toString());
    return jsonResponse({ result: 'error', error: error.toString() });
  }
}

function isV2Lead(data) {
  return data.quizVersion === 'v2' || data.tier === 1 || data.tier === 2 || data.tier === 3 || data.tier === '1' || data.tier === '2' || data.tier === '3';
}

function handleLegacyLead(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timestamp = new Date();

  sheet.appendRow([
    timestamp,
    data.name || '',
    data.phone || '',
    data.telegram || '',
    data.livingCountry || '',
    data.citizenship || '',
    data.goal || ''
  ]);

  return jsonResponse({ result: 'success', mode: 'legacy' });
}

function handleV2Lead(data) {
  var resolved = resolveTier(data);
  var sheet = getOrCreateTierSheet(resolved.tier);
  var timestamp = new Date();
  var contact = data.contact || data.phone || data.telegram || '';

  sheet.appendRow([
    timestamp,
    data.name || '',
    contact,
    resolved.branch,
    'Tier-' + resolved.tier,
    data.departure_date || '',
    data.age_at_departure || '',
    data.border_crossing || '',
    data.family_status || '',
    data.destination_country || '',
    data.eu_documents || '',
    data.main_request || '',
    data.path || ''
  ]);

  return jsonResponse({
    result: 'success',
    mode: 'v2',
    tier: resolved.tier,
    branch: resolved.branch,
    sheet: TIER_SHEETS[resolved.tier]
  });
}

/**
 * Определяет Tier на сервере (на случай, если клиент не прислал tier).
 *
 * Tier-1: основная ветка + «Страна ЕС»; ветка «а»
 * Tier-2: ветка «б»; основная ветка + «Страна вне ЕС»
 * Tier-3: ветка «в» (интересуюсь); ветка «г»
 */
function resolveTier(data) {
  var tierFromClient = parseInt(data.tier, 10);
  var branchFromClient = data.branch || '';

  if (tierFromClient === 1 || tierFromClient === 2 || tierFromClient === 3) {
    return {
      tier: tierFromClient,
      branch: branchFromClient || inferBranch(data)
    };
  }

  var branch = inferBranch(data);
  var tier = 3;

  if (branch === 'main' || branch === 'а') {
    tier = 1;
  } else if (branch === 'main-non-eu' || branch === 'б') {
    tier = 2;
  } else {
    tier = 3; // в, г, unknown
  }

  return { tier: tier, branch: branch };
}

function inferBranch(data) {
  var path = String(data.path || '');
  var curiosity = 'Просто интересуюсь, узнаю варианты, стоимость';

  // Основная ветка: есть вопрос страны
  if (path.indexOf('q5') !== -1 || data.destination_country) {
    if (data.destination_country === 'Страна вне ЕС') {
      return 'main-non-eu';
    }
    return 'main';
  }

  // Ветка б
  if (path.indexOf('q3b') !== -1 || path.indexOf('q4b') !== -1) {
    return 'б';
  }

  // Ветка а
  if (path.indexOf('q3a') !== -1) {
    return 'а';
  }

  // Ветка г (прошли семейное положение после 3г)
  if (path.indexOf('q4g') !== -1) {
    return 'г';
  }

  // Ветка в: интерес с 3г сразу на контакты (или старый путь с 2a)
  if (
    data.main_request === curiosity ||
    (path.indexOf('q3g') !== -1 && path.indexOf('q4g') === -1)
  ) {
    return 'в';
  }

  return branchFromFallback(data);
}

function branchFromFallback(data) {
  if (data.eu_documents === 'Без статуса' || data.eu_documents === 'Ничего нету') {
    return 'а';
  }
  if (data.age_at_departure && String(data.age_at_departure).indexOf('до 23') === 0) {
    return 'б';
  }
  if (data.eu_documents === 'Защита' || data.eu_documents === 'ВНЖ') {
    return 'г';
  }
  return 'unknown';
}

function getOrCreateTierSheet(tier) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = TIER_SHEETS[tier] || TIER_SHEETS[3];
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  ensureHeaders(sheet);
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(TIER_HEADERS);
    sheet.getRange(1, 1, 1, TIER_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }

  var firstCell = String(sheet.getRange(1, 1).getValue() || '');
  if (firstCell !== TIER_HEADERS[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, TIER_HEADERS.length).setValues([TIER_HEADERS]);
    sheet.getRange(1, 1, 1, TIER_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Ручная проверка из редактора Apps Script:
 * выполнить testDoPost() и посмотреть листы.
 */
function testDoPost() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        quizVersion: 'v2',
        tier: 1,
        branch: 'main',
        name: 'Test Lead',
        contact: '+380000000000',
        departure_date: 'после 05.08.2026',
        age_at_departure: 'больше 23 лет',
        border_crossing: 'Транспорт',
        family_status: 'Один',
        destination_country: 'Страна ЕС',
        path: 'q1 → q2 → q3 → q4 → q5 → contacts'
      })
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}
