function startProgramm() {
  return;
}

function createAllTriggers() {
  
  ScriptApp.newTrigger("updatePeriod") // обновление дней в статусе
    .timeBased()
    .everyDays(1)
    .atHour(15)
    .create();

  ScriptApp.newTrigger("archiveClosedLeads") // перенос в архив
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
}

function updatePeriod() {
  console.log("trigger start");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadSheet = ss.getSheetByName("ЛИДЫ");

  if (!leadSheet) return;

  const lastRow = leadSheet.getLastRow();
  if (lastRow < 2) return; 

  const rangeStatus = leadSheet.getRange(2, 6, lastRow - 1, 1);
  const dataStatus = rangeStatus.getValues();
  const range = leadSheet.getRange(2, 11, lastRow - 1, 1);
  const data = range.getValues();

  for (let i = 0; i < data.length; i++) {
    if (dataStatus[i][0]) {
      const currentValue = Number(data[i][0]) || 0;
      data[i][0] = currentValue + 1;
    }            
  }

  range.setValues(data);
  console.log("trigger finished");
}

function archiveClosedLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadSheet = ss.getSheetByName("ЛИДЫ");
  const archiveSheet = ss.getSheetByName("Архив_Лиды"); 

  if (!leadSheet || !archiveSheet) return;

  const lastRow = leadSheet.getLastRow();
  if (lastRow < 2) return;

  const leadRange = leadSheet.getRange(2, 1, lastRow - 1, 11);
  const data = leadRange.getValues();

  const leadsToArchive = [];
  const rowsToDelete = []; 

  for (let i = 0; i < data.length; i++) {
    const status = data[i][5];     
    const daysInWork = data[i][10]; 

    if ((status === "Сделка закрыта" && daysInWork >= 30) || status === "Отказ") {
      leadsToArchive.push(data[i]);
      rowsToDelete.push(i + 2); 
    }
  }

  if (leadsToArchive.length > 0) {
    const archiveLastRow = archiveSheet.getLastRow();
    archiveSheet.getRange(archiveLastRow + 1, 1, leadsToArchive.length, 11).setValues(leadsToArchive);

    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      leadSheet.deleteRow(rowsToDelete[j]);
    }

    console.log(`Успешно перенесено строк в архив: ${leadsToArchive.length}`);
  } else {
    console.log("Нет подходящих сделок для архивации.");
  }
}

function onEditTrigger(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== "ЛИДЫ") return;

  const row = range.getRow();

  if (row === 1) return;

  const rowRange = sheet.getRange(row - 1, 1, 1, sheet.getLastColumn());
  if (rowRange.isBlank()) return; // если строка сверху пустая - не добавляем;

  //const firstColumn = sheet.getRange(row, 1).getValue() // A дата
  const secondColumn = sheet.getRange(row, 2).getValue() // B имя
  const thirdColumn = sheet.getRange(row, 3).getValue() // C телефон
  const fourthColumn = sheet.getRange(row, 5).getValue() // E источник


  const email = sheet.getRange(row, 4).getValue() // D (почта)
  const managerValue = sheet.getRange(row, 7).getValue(); // G (Менеджер)
  const statusValue = sheet.getRange(row, 6).getValue();  // F (Статус)
  const daysWorked = sheet.getRange(row, 11).getValue(); // K (Дни в работе)

  if (!statusValue || !managerValue || !daysWorked) {
    if (!statusValue) {
      sheet.getRange(row, 6).setValue("Новый"); // статус
    }
    if (!daysWorked) {
      sheet.getRange(row, 11).setValue(1);
    }
    if (!managerValue) {
      const assignedManager = TaskProvider(); 
      if (assignedManager) {
        sheet.getRange(row, 7).setValue(assignedManager);
      }
    }
  }

  if (range.getColumn() === 6 && secondColumn && thirdColumn && fourthColumn && email && managerValue && statusValue === "В работе") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const managersSheet = sheet.getSheetByName("Менеджеры");
    if (!managersSheet) return;

    const lastRow = managersSheet.getLastRow();
    if (lastRow < 2) return;

    const data = managersSheet.getRange(2, 1, lastRow - 1, 3).getValues();

    let replyTo = "";
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === managerValue) {
        replyTo = data[i][1];
        break; 
      }
    }

    const recipient = email;
    const subject = 'Письмо от компании "..."';
  
    const body = `Компания приветствует вас, ${secondColumn}!\nВаша заявка в работе. Менеджер ${managerValue} свяжется с вами в ближайшее время.\n\nПожалуйста, если есть вопросы, не отвечайте на это сообщение прямо. Можете задать вопрос менеджеру на почту ${replyTo}`;

    sendMessage(recipient, subject, body);
  }
}

function TaskProvider() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const managersSheet = ss.getSheetByName("Менеджеры");
  if (!managersSheet) return "";

  const lastRow = managersSheet.getLastRow();
  if (lastRow < 2) return "";

  const data = managersSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  let minDeals = Infinity;
  let managerName = "";

  for (let i = 0; i < data.length; i++) {
    const name = data[i][0]; 
    const dealsCount = data[i][2]; 

    if (dealsCount < minDeals && name !== "") {
      minDeals = dealsCount;
      managerName = name;
    }
  }
  return managerName;
}

function sendMessage(recipient, subject ,body) {
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });
    console.log("Письмо отправлено");
    Logger.log("Письмо успешно отправлено на адрес: " + recipient);
    return true;

  } catch (error) {
    Logger.log("Ошибка при отправке письма: " + error.toString());
    return false;
  }
}





















