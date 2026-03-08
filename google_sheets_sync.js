/**
 * Google Apps Script — Webhook для синхронізації з Telegram ботом
 * 
 * ВСТАНОВЛЕННЯ:
 * 1. Відкрий Google Sheets (створи нову таблицю)
 * 2. Розширення → Apps Script
 * 3. Вставте цей код і збережи
 * 4. Запусти → Розгорнути → Нове розгортання
 *    - Тип: Веб-застосунок
 *    - Виконувати від: Мене
 *    - Доступ: Усі
 * 5. Скопіюй URL і встанови як SHEETS_WEBHOOK у Railway
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet(data.year + "_" + getMonthName(data.month));
    
    appendPayment(sheet, data);
    
    return ContentService
      .createTextOutput(JSON.stringify({status: "ok"}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // Заголовки
    const headers = ["Точка", "Сума (₴)", "Дата", "Місяць", "Рік", "Нотатка", "Час запису"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Форматування заголовків
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1a1a2e");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 160);
  }
  
  return sheet;
}

function appendPayment(sheet, data) {
  const MONTHS_UA = ["Січень","Лютий","Березень","Квітень","Травень","Червень",
                     "Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"];
  
  // Перевіряємо чи вже є такий запис (дедублікація)
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][0] === data.point && 
        existing[i][3] === data.month && 
        existing[i][4] === data.year) {
      // Оновлюємо існуючий рядок
      sheet.getRange(i + 1, 1, 1, 7).setValues([[
        data.point,
        data.amount,
        data.date,
        MONTHS_UA[data.month - 1],
        data.year,
        data.note || "",
        new Date().toLocaleString("uk-UA")
      ]]);
      colorRow(sheet, i + 1, true);
      return;
    }
  }
  
  // Додаємо новий рядок
  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 7).setValues([[
    data.point,
    data.amount,
    data.date,
    MONTHS_UA[data.month - 1],
    data.year,
    data.note || "",
    new Date().toLocaleString("uk-UA")
  ]]);
  
  colorRow(sheet, lastRow, data.paid);
  
  // Форматування суми
  sheet.getRange(lastRow, 2).setNumberFormat('#,##0 [$₴]');
}

function colorRow(sheet, row, isPaid) {
  const color = isPaid ? "#d4edda" : "#f8d7da";
  sheet.getRange(row, 1, 1, 7).setBackground(color);
}

function getMonthName(month) {
  const MONTHS = ["sichen","lytyy","berezen","kviten","traven","cherven",
                  "lypen","serpen","veresen","zhovten","lystopad","hruden"];
  return MONTHS[month - 1] || month;
}

/**
 * Тестова функція — запусти вручну щоб перевірити
 */
function testWebhook() {
  const testData = {
    point: "ТЦ Глобус",
    amount: 3500,
    date: "01.03.2025",
    month: 3,
    year: 2025,
    note: "Тест",
    paid: true
  };
  
  const sheet = getOrCreateSheet("2025_berezen");
  appendPayment(sheet, testData);
  Logger.log("Test completed!");
}
