function startProject() {
  return;
}
/** 
 * Триггер, который каждое утро открывает список недостач 
*/
function createAllTriggers() {
  ScriptApp.newTrigger("showLowStockDialog") 
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}
/**validateSaleData
 * Триггер, который срабатывает при редактировании листа
 * Автоматически обрабатывает новые записи в журнале продаж
 */
function onEditTrigger(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow(),
        col = range.getColumn();

  if (row === 1) return;

  if (sheet.getName() === "Журнал продаж") processSaleRow(sheet, row, col);
  return;
}
/**
 * Обрабатывает строку продажи: рассчитывает сумму, проверяет наличие, обновляет остаток
 */
function processSaleRow(sheet, row, col) {

  const rowRange = sheet.getRange(row - 1, 1, 1, sheet.getLastColumn()); // диапазон размером в одну строку (предыдущую по отношению к row)
  if (rowRange.isBlank()) return; // проверяем, есть ли значения в строке

  // значения ячеек в записываем в переменные
  const dataValue = sheet.getRange(row, 1).getValue();
  const articleCode = sheet.getRange(row, 2).getValue();
  const amountValue = sheet.getRange(row, 3).getValue();
  const summValue = sheet.getRange(row, 4).getValue();
  const statusValue = sheet.getRange(row, 5).getValue();

  if (!articleCode) {
    console.log("Код продукта не найден");
    return;
  }

  const product = findProductByArticle(articleCode);
  if (!product || !product.price) {
    Logger.log("Продукт или его цена не найдены"); 
    mySetCellColor(sheet, row, "warning");
    return;
    }

  if (amountValue && product.price) {
    const total = amountValue * product.price;
    sheet.getRange(row, 4).setValue(total);
  }

  if (col === 3 && amountValue > 0) {
    const result = myUpdateProductStock(articleCode, amountValue);

    if (result === "success") {
      mySetCellColor(sheet, row, "success");
    } else if (result === "low") {
      mySetCellColor(sheet, row, "low");
      findManager();
    } else {
      mySetCellColor(sheet, row, "error");
      findManager();
    }
  }
}
/**
 * Обновляет остаток товара в базе
 * возвращает значения для выставления статуса.
 */
function myUpdateProductStock(articleCode, quantity) {

  const product = findProductByArticle(articleCode);

  if (!product) {
    Logger.log("Товар с артикулом " + articleCode + " не найден");
    return "not_found";
  }

  const stock = Number(product.stock);
  const saleQuantity = Number(quantity);
  const minStock = Number(product.minStock);

  if (isNaN(stock) || isNaN(saleQuantity)) {
    Logger.log("Ошибка: значения не являются числами");
    return "error";
  }

  // Проверяем, хватает ли товара для продажи
  if (stock < saleQuantity) {
    Logger.log("Недостаточно товара. Доступно: " + stock + ", требуется: " + saleQuantity);
    return "error";
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName("База товаров");

  const newStock = stock - saleQuantity;

  productsSheet.getRange(product.row, 5).setValue(newStock);

  // Товар продан, но остаток стал ниже минимального
  if (newStock < minStock) {
    return "low";
  }

  return "success";
}
/**
 * Проверяет товары с низким остатком и показывает уведомление
 * Можно запустить вручную из меню или настроить триггер по времени
 */
function showLowStockDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesLedger = ss.getSheetByName("База товаров");

  if (!salesLedger) {
    SpreadsheetApp.getUi().alert("Лист 'База товаров' не найден!");
    return;
  }

  const lastRow = salesLedger.getLastRow();
  if (lastRow < 2) return;

  const data = salesLedger.getRange(2, 1, lastRow - 1, 6).getValues();

  // строки html-таблицы напрямую в цикле
  let tableRowsHtml = "";
  let hasItems = false;

  for (let i = 0; i < data.length; i++) {
    const stock = data[i][4];    // Столбец E остаток
    const minStock = data[i][5]; // Столбец F мин. остаток

    if (stock !== "" && minStock !== "" && stock <= minStock) {
      hasItems = true;
      tableRowsHtml += "<tr>";
      for (let j = 0; j < data[i].length; j++) {
        tableRowsHtml += `<td style="padding: 6px 10px; border: 1px solid #dadce0;">${data[i][j]}</td>`;
      }
      tableRowsHtml += "</tr>";
    }
  }

  if (!hasItems) {
    tableRowsHtml = '<tr><td colspan="6" style="padding: 10px; text-align: center;">Все товары в достаточном количестве</td></tr>';
  }

  // итоговая html-страница
  const htmlContent = `
    <div style="font-family: Arial, sans-serif; padding: 5px;">
      <table style="border-collapse: collapse; width: 100%; font-size: 13px; margin-bottom: 15px;">
        <thead>
          <tr style="background-color: #f1f3f4; text-align: left; font-weight: bold;">
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Артикул</th>
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Наименование</th>
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Категория</th>
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Цена</th>
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Остаток</th>
            <th style="padding: 8px 10px; border: 1px solid #dadce0;">Мин. остаток</th>
          </tr>
        </thead>
        <tbody>
          ${tableRowsHtml}
        </tbody>
      </table>
      
      <div style="text-align: right;">
        <button onclick="google.script.host.close()" 
                style="padding: 8px 16px; background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">
          Закрыть
        </button>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(700)
      .setHeight(450); // Уменьшено до комфортной высоты экрана

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Недостаток на складе');
}
/**
 * экспорт статистики за месяц
 */
function exportMonthlySalesStats(year, month) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Журнал продаж");

  if (!salesSheet) {
    SpreadsheetApp.getUi().alert("Лист 'Журнал продаж' не найден!");
    return;
  }

  const lastRow = salesSheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("В журнале продаж нет данных.");
    return;
  }

  // Получаем данные:
  // A — дата
  // D — сумма
  // E — статус
  const data = salesSheet
    .getRange(2, 1, lastRow - 1, 5)
    .getValues();

  let totalAmount = 0;
  let salesCount = 0;

  for (let i = 0; i < data.length; i++) {

    const rawDate = data[i][0];
    const summ = data[i][3];
    const status = data[i][4];

    if (!rawDate || status !== "Выполнено") {
      continue;
    }

    const date = rawDate instanceof Date
      ? rawDate
      : new Date(rawDate);

    if (isNaN(date.getTime())) {
      continue;
    }

    const saleYear = date.getFullYear();
    const saleMonth = date.getMonth() + 1;

    if (saleYear === Number(year) && saleMonth === Number(month)) {

      totalAmount += Number(summ) || 0;
      salesCount++;
    }
  }

  // Получаем или создаём лист статистики
  let statsSheet = ss.getSheetByName("Месячная статистика");

  if (!statsSheet) {
    statsSheet = ss.insertSheet("Месячная статистика");

    statsSheet.appendRow([
      "Год",
      "Месяц",
      "Количество продаж",
      "Общая выручка"
    ]);
  }

  // Записываем результат
  statsSheet.appendRow([
    year,
    month,
    salesCount,
    totalAmount
  ]);

  // Форматируем сумму
  const lastStatsRow = statsSheet.getLastRow();

  statsSheet
    .getRange(lastStatsRow, 4)
    .setNumberFormat('#,##0.00 ₽');

  SpreadsheetApp.getUi().alert(
    `Статистика за ${month}.${year} экспортирована.\n\n` +
    `Продаж: ${salesCount}\n` +
    `Выручка: ${totalAmount.toLocaleString('ru-RU')} ₽`
  );
}
/**
 * функция меню - просит пользователя ввести месяц
 */
function exportMonthlyStatsMenu() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Экспорт статистики",
    "Введите месяц в формате ММ.ГГГГ\nНапример: 10.2026",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const value = response.getResponseText().trim();
  const parts = value.split(".");

  if (parts.length !== 2) {
    ui.alert("Неверный формат. Используйте ММ.ГГГГ");
    return;
  }

  const month = Number(parts[0]);
  const year = Number(parts[1]);

  if (
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12 ||
    !Number.isInteger(year)
  ) {
    ui.alert("Некорректный месяц или год.");
    return;
  }

  exportMonthlySalesStats(year, month);
}
/**
 * Получает статистику продаж за указанную дату
 * @param {string} dateString - Дата в формате "DD.MM.YYYY" (например, "15.10.2026")
 * @return {Object} - Объект с общей суммой и количеством продаж
 */
function getDailySales(dateString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Журнал продаж");

  if (!salesSheet) {
    Logger.log("Лист 'Журнал продаж' не найден!");
    return { date: dateString, totalAmount: 0, successCount: 0 };
  }

  const lastRow = salesSheet.getLastRow();
  if (lastRow < 2) {
    return { date: dateString, totalAmount: 0, successCount: 0 };
  }

  // Забираем данные таблицы: Дата (A), Сумма (D), Статус (E)
  const data = salesSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  let totalAmount = 0;
  let successCount = 0;

  for (let i = 0; i < data.length; i++) {
    const rawDate = data[i][0];   // Столбец A (Дата)
    const summ = data[i][3];      // Столбец D (Сумма)
    const status = data[i][4];    // Столбец E (Статус)

    if (!rawDate) continue;

    // Приводим дату из ячейки к формату DD.MM.YYYY
    const formattedCellDate = formatDateToString(rawDate);

    // Сравниваем даты и проверяем, что продажа выполнена
    if (formattedCellDate === dateString) {
      if (status === "Выполнено") {
        totalAmount += Number(summ) || 0;
        successCount++;
      }
    }
  }

  return {
    date: dateString,
    totalAmount: totalAmount,
    successCount: successCount
  };
}
/**
 * Показывает статистику продаж за сегодняшний день в всплывающем окне
 */
function showTodayStats() {
  // Получаем сегодняшнюю дату в формате DD.MM.YYYY
  const todayStr = formatDateToString(new Date());

  // Получаем агрегированные данные
  const stats = getDailySales(todayStr);

  // Форматируем сумму с разделением тысяч для красоты
  const formattedAmount = stats.totalAmount.toLocaleString('ru-RU', {
    style: 'currency',
    currency: 'RUB'
  });

  const message = `Статистика за сегодня (${stats.date}):\n\n` +
                  `Успешных продаж: ${stats.successCount}\n` +
                  `Выручка: ${formattedAmount}`;

  SpreadsheetApp.getUi().alert("Дневной отчёт", message, SpreadsheetApp.getUi().ButtonSet.OK);
}
/**
 * Вспомогательная функция для приведения объекта Date или строки к формату "DD.MM.YYYY"
 */
function formatDateToString(dateVal) {
  if (dateVal instanceof Date) {
    const day = String(dateVal.getDate()).padStart(2, '0');
    const month = String(dateVal.getMonth() + 1).padStart(2, '0');
    const year = dateVal.getFullYear();
    return `${day}.${month}.${year}`;
  }
  return String(dateVal).trim();
}
/**
 * Создает пользовательское меню в Google Sheets
 * Вызывается автоматически при открытии документа
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Отчёты")
    .addItem("Показать продажи за сегодня", "showTodayStats")
    .addItem("Показать недостачу", "showLowStockDialog")
    .addItem("Экспорт статистики за месяц", "exportMonthlyStatsMenu")
    .addToUi();
}

function findProductByArticle(articleCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName("База товаров");
  
  if (!productsSheet) {
    Logger.log("Лист 'База товаров' не найден!");
    return null;
  }
  
  const data = productsSheet.getDataRange().getValues();
  
  // Пропускаем заголовок (первая строка)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === articleCode) {
      return {
        row: i + 1, // +1 потому что индексация в Sheets начинается с 1
        article: data[i][0],
        name: data[i][1],
        category: data[i][2],
        price: data[i][3],
        stock: data[i][4],
        minStock: data[i][5]
      };
      break;
    }
  }
  Logger.log("Товар с артикулом " + articleCode + " не найден");
  return null;
}
/**
 * Покрас статуса
 */
function mySetCellColor(sheet, row, status) {

  const range = sheet.getRange(Number(row), 5);

  if (status === "success") {
    range.setBackground("#d4edda");
    range.setFontColor("#155724");
    range.setValue("Выполнено");
  } else if (status === "error") {
    range.setBackground("#f8d7da");
    range.setFontColor("#721c24");
    range.setValue("Недостаточно товара");
  } else if (status === "low") {
    range.setBackground("#f8d7da");
    range.setFontColor("#721c24");
    range.setValue("Низкий остаток");
  } else {
    range.setBackground("#fff3cd");
    range.setFontColor("#856404");
    range.setValue("Ключа не существует");
  }
}
/**
 * искать менеджера, готовому принимать сообщения на почту. Сообщение о недостаче товара
 */
function findManager() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetManagers = ss.getSheetByName("Менеджеры");

  if (!sheetManagers) return;

  const lastRow = sheetManagers.getLastRow();
  if (lastRow < 2) return;

  const data = sheetManagers.getRange(2, 1, lastRow - 1, 3).getValues();
  let email = "";
  for (let i = 0; i < data.length; i++) { // отправляем только тому, кто принимает соообщения на почту
    if (data[i][2] === "да") { 
      email = data[i][1];
      break; 
    }
  }
  const recipient = email;
  const subject = 'Письмо от компании "..."';
  const body = "Доброе утро! Для склада требуется закупка!\n\nМожете ознакомиться с таблицей";

  sendMessage(recipient, subject, body);
}
/**
 * Отправка сообщения
 */
function sendMessage(recipient, subject ,body) {
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });
    console.log("Письмо отправлено");
    Logger.log("Письмо успешно отправлено на адрес: " + recipient);
    return true;

  } catch (error) {
    Logger.log("Ошибка при отправке письма: " + error.toString());
    return false;
  }
}