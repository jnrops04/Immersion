

const MAX_ATTEMPTS = 3;
const LOG_FILE_NAME = "Employee Onboarding Logs";

function setup() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  PropertiesService
    .getScriptProperties()
    .setProperty(
      "MAIN_SPREADSHEET_ID",
      ss.getId()
    );

  initializeMainSheet(
    ss.getActiveSheet()
  );

  getLogSheet();

  createEditTrigger();

  createSyncTrigger();
}

function createEditTrigger() {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === "onEdit");

  if (!exists) {
    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  initializeMainSheet(sheet);

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  for (let row = startRow; row < startRow + numRows; row++) {
    processEmployeeRow(sheet, row);
  }
}

function initializeMainSheet(sheet) {

  const headers = [[
    "Name",
    "ID",
    "Email",
    "Status",
    "Task List URL",
    "Task List ID"
  ]];

  const firstRow =
    sheet.getRange(1,1,1,6).getValues()[0];

  if (firstRow.every(v => v === "")) {

    sheet.getRange(1,1,1,6)
      .setValues(headers);

    sheet.setFrozenRows(1);

    sheet.autoResizeColumns(
      1,
      6
    );
  }
}

function processEmployeeRow(sheet, row) {

  if (row < 2) return;

  const name = sheet.getRange(row,1).getValue().toString().trim();
  const id = sheet.getRange(row,2).getValue().toString().trim();
  const email = sheet.getRange(row,3).getValue().toString().trim();

  if (!name || !id || !email) {

    sheet
      .getRange(row, 1, 1, 6)
      .setBackground(null);

    sheet
      .getRange(row, 4)
      .clearContent();

    return;
  }

  const statusCell = sheet.getRange(row,4);
  const currentStatus = statusCell.getValue();

  if (
    currentStatus === "Sent" ||
    currentStatus === "Max Attempts Reached"
  ) return;

  if (!isValidEmail(email)) {
    statusCell.setValue("Failed - Wrong Email Format");
    setRowColor(sheet,row,"Failed - Wrong Email Format");

    logAttempt(
      name,id,email,
      "Failed",0,
      "Invalid Email Format"
    );
    return;
  }

  const duplicateType =
    checkDuplicateEmployee(
      sheet,
      row,
      id,
      email
    );

  if (duplicateType) {

    statusCell.setValue("Duplicate");
    setRowColor(sheet,row,"Duplicate");

    let reason;

    switch (duplicateType) {

      case "ID_AND_EMAIL":
        reason =
          "Duplicate ID and Email found";
        break;

      case "ID":
        reason =
          "Duplicate ID found";
        break;

      case "EMAIL":
        reason =
          "Duplicate Email found";
        break;
    }

    if (
      !duplicateAlreadyLogged(
        id,
        reason
      )
    ) {

      logAttempt(
        name,
        id,
        email,
        "Duplicate",
        0,
        reason
      );
    }

    return;
  }

  let attempts = getAttemptCount(id);

  if (attempts >= MAX_ATTEMPTS) {
    statusCell.setValue("Max Attempts Reached");
    setRowColor(sheet,row,"Max Attempts Reached");

    logAttempt(
      name,id,email,
      "Max Attempts Reached",
      attempts,
      "Maximum email attempts exceeded"
    );

    return;
  }

  try {

    const taskList =
      createEmployeeTaskList(
        name,id,email
      );
    
    createTaskLogEntries(
      name,
      id,
      email
    );

    sheet.getRange(row,5)
      .setValue(taskList.url);

    sheet.getRange(row,6)
      .setValue(taskList.id);

    const emailBody =
`Hello ${name},

Welcome to the company.

Your Employee ID is:

${id}

Please complete your onboarding checklist using the link below:

${taskList.url}

You have been granted editor access to the checklist.

Regards,
HR Team`;

    MailApp.sendEmail(
      email,
      "Welcome to the Company",
      emailBody
    );

    statusCell.setValue("Sent");
    setRowColor(sheet,row,"Sent");

    logAttempt(
      name,id,email,
      "Success",
      attempts + 1,
      "N/A"
    );

  } catch(err) {

    attempts++;

    if (attempts >= MAX_ATTEMPTS) {
      statusCell.setValue("Max Attempts Reached");
      setRowColor(sheet,row,"Max Attempts Reached");
    } else {
      statusCell.setValue("Failed");
      setRowColor(sheet,row,"Failed");
    }

    logAttempt(
      name,id,email,
      "Failed",
      attempts,
      err.message
    );
  }
}

function createEmployeeTaskList(name,id,employeeEmail) {

  const fileName = `${name}_${id}_tasklist`;

  const folder = getTasklistsFolder();

  const existingFiles = folder.getFilesByName(fileName);

  let ss;

  if (existingFiles.hasNext()) {
    ss = SpreadsheetApp.openById(
      existingFiles.next().getId()
    );
  } else {

    ss = SpreadsheetApp.create(fileName);

    const folder =
      getTasklistsFolder();

    const file =
      DriveApp.getFileById(
        ss.getId()
      );

    file.moveTo(folder);

    const taskSheet =
      ss.getSheets()[0];

    taskSheet.getRange("A1").setValue("Task");
    taskSheet.getRange("B1").setValue("Completed");

    const tasks = [
      ["Apply For Locker"],
      ["Get ID"],
      ["Familiarize with colleagues"]
    ];

    taskSheet.getRange(2,1,tasks.length,1)
      .setValues(tasks);

    taskSheet.getRange(2,2,tasks.length,1)
      .insertCheckboxes();

    taskSheet.autoResizeColumns(1,2);
  }

  try {
    DriveApp.getFileById(ss.getId())
      .addEditor(employeeEmail);
  } catch(err) {
    Logger.log(err.message);
  }

  return {
    url: ss.getUrl(),
    id: ss.getId()
  };
}

function isValidEmail(email) {
  return /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[A-Za-z]{2,}$/
    .test(email);
}


function getAttemptCount(employeeId) {

  const data =
    getLogSheet()
      .getDataRange()
      .getValues();

  let maxAttempts = 0;

  for (let i=1;i<data.length;i++) {

    if (data[i][2] == employeeId) {

      maxAttempts = Math.max(
        maxAttempts,
        Number(data[i][5]) || 0
      );
    }
  }

  return maxAttempts;
}

function logAttempt(name,id,email,status,attempts,error) {

  getLogSheet().appendRow([
    new Date(),
    name,
    id,
    email,
    status,
    attempts,
    error
  ]);
}

function getLogSheet() {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const props =
      PropertiesService.getScriptProperties();

    let logSpreadsheetId =
      props.getProperty("LOG_SPREADSHEET_ID");

    let logSS;

    if (!logSpreadsheetId) {

      logSS =
        SpreadsheetApp.create(
          LOG_FILE_NAME
        );

      logSpreadsheetId =
        logSS.getId();

      props.setProperty(
        "LOG_SPREADSHEET_ID",
        logSpreadsheetId
      );

    } else {

      try {

        logSS =
          SpreadsheetApp.openById(
            logSpreadsheetId
          );

      } catch(err) {

        logSS =
          SpreadsheetApp.create(
            LOG_FILE_NAME
          );

        logSpreadsheetId =
          logSS.getId();

        props.setProperty(
          "LOG_SPREADSHEET_ID",
          logSpreadsheetId
        );
      }
    }

    let sheet =
      logSS.getSheetByName("Logs");

    if (!sheet) {

      sheet = logSS.getSheets()[0];
      sheet.setName("Logs");
    }


    let taskLogs =
      logSS.getSheetByName(
        "Task Logs"
      );

    if (!taskLogs) {

      taskLogs =
        logSS.insertSheet(
          "Task Logs"
        );

      taskLogs.getRange(
        1,
        1,
        1,
        6
      ).setValues([[
        "Name",
        "ID",
        "Email",
        "Status",
        "Task Name",
        "Task Number"
      ]]);

      taskLogs.setFrozenRows(1);
      taskLogs.autoResizeColumns(1,6);
    }

    const headerRow =
      sheet.getRange(1,1,1,7)
        .getValues()[0];

    if (
      headerRow.join("") === ""
    ) {

      sheet.getRange(1,1,1,7)
        .setValues([[
          "Timestamp",
          "Name",
          "ID",
          "Email",
          "Status",
          "Attempts",
          "Error"
        ]]);

        sheet.autoResizeColumns(1,7);
    }

    return sheet;

  } finally {

    lock.releaseLock();
  }
}

function setRowColor(sheet,row,status) {

  const rowRange =
    sheet.getRange(row,1,1,6);

  switch(status) {

    case "Sent":
      rowRange.setBackground("#d9ead3");
      break;

    case "Failed":
    case "Failed - Wrong Email Format":
    case "Max Attempts Reached":
      rowRange.setBackground("#f4cccc");
      break;

    case "Duplicate":
      rowRange.setBackground("#fff2cc");
      break;

    default:
      rowRange.setBackground(null);
  }
}

function createTaskLogEntries(
  name,
  id,
  email
) {

  const props =
    PropertiesService
      .getScriptProperties();

  const logId =
    props.getProperty(
      "LOG_SPREADSHEET_ID"
    );

  const logSS =
    SpreadsheetApp.openById(
      logId
    );

  const taskLogs =
    logSS.getSheetByName(
      "Task Logs"
    );

  const existingData =
    taskLogs.getDataRange()
      .getValues();

  for (let i = 1; i < existingData.length; i++) {

    if (
      existingData[i][1].toString() === id
    ) {
      return;
    }
  }

  const rows = [
    [
      name,
      id,
      email,
      "Todo",
      "Apply For Locker",
      1
    ],
    [
      name,
      id,
      email,
      "Todo",
      "Get ID",
      2
    ],
    [
      name,
      id,
      email,
      "Todo",
      "Familiarize with colleagues",
      3
    ]
  ];

  const startRow =
    taskLogs.getLastRow() + 1;

  taskLogs.getRange(
    startRow,
    1,
    rows.length,
    6
  ).setValues(rows);

  taskLogs.getRange(
    startRow,
    1,
    rows.length,
    6
  ).setBackground(
    "#d9eaf7"
  );
}

function createSyncTrigger() {

  const exists =
    ScriptApp
      .getProjectTriggers()
      .some(
        t =>
          t.getHandlerFunction() ===
          "syncTaskLogs"
      );

  if (!exists) {

    ScriptApp
      .newTrigger(
        "syncTaskLogs"
      )
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

function syncTaskLogs() {

  const props =
    PropertiesService
      .getScriptProperties();

  const logId =
    props.getProperty(
      "LOG_SPREADSHEET_ID"
    );

  if (!logId) return;

  const logSS =
    SpreadsheetApp
      .openById(logId);

  const taskLogs =
    logSS.getSheetByName(
      "Task Logs"
    );

  if (!taskLogs) return;

  const mainSpreadsheetId =
    props.getProperty(
      "MAIN_SPREADSHEET_ID"
    );

  if (!mainSpreadsheetId)
    return;

  const mainSS =
    SpreadsheetApp
      .openById(
        mainSpreadsheetId
      );

  const mainSheet =
    mainSS.getSheets()[0];

  const employees =
    mainSheet
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < employees.length;
    i++
  ) {

    const employeeId =
      employees[i][1];

    const taskListId =
      employees[i][5];

    if (!taskListId) continue;

    try {

      const employeeSS =
        SpreadsheetApp
          .openById(
            taskListId
          );

      const taskSheet =
        employeeSS
          .getSheets()[0];

      const taskData =
        taskSheet
          .getRange(
            2,
            1,
            3,
            2
          )
          .getValues();

      const taskLogData =
        taskLogs
          .getDataRange()
          .getValues();

      for (
        let taskIndex = 0;
        taskIndex < taskData.length;
        taskIndex++
      ) {

        const completed =
          taskData[taskIndex][1];

        for (
          let logRow = 1;
          logRow < taskLogData.length;
          logRow++
        ) {

          if (
            taskLogData[logRow][1].toString() ===
              employeeId.toString() &&
            Number(
              taskLogData[logRow][5]
            ) ===
              taskIndex + 1
          ) {

            if (completed === true) {

              taskLogs
                .getRange(
                  logRow + 1,
                  4
                )
                .setValue(
                  "Checked"
                );

              taskLogs
                .getRange(
                  logRow + 1,
                  1,
                  1,
                  6
                )
                .setBackground(
                  "#d9ead3"
                );

            } else {

              taskLogs
                .getRange(
                  logRow + 1,
                  4
                )
                .setValue(
                  "Todo"
                );

              taskLogs
                .getRange(
                  logRow + 1,
                  1,
                  1,
                  6
                )
                .setBackground(
                  "#d9eaf7"
                );
            }
          }
        }
      }

    } catch(err) {

      Logger.log(
        "Sync Error: " +
        err.message
      );
    }
  }
}

function getTasklistsFolder() {

  const folders =
    DriveApp.getFoldersByName(
      "Tasklists"
    );

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(
    "Tasklists"
  );
}

function checkDuplicateEmployee(
  sheet,
  currentRow,
  id,
  email
) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const data =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      3
    ).getValues();

  for (let i = 0; i < data.length; i++) {

    const rowNum = i + 2;

    if (rowNum === currentRow)
      continue;

    const existingId =
      data[i][1]
        .toString()
        .trim();

    const existingEmail =
      data[i][2]
        .toString()
        .trim()
        .toLowerCase();

    const idMatch =
      existingId === id;

    const emailMatch =
      existingEmail ===
      email.toLowerCase();

    if (
      idMatch &&
      emailMatch
    ) {
      return "ID_AND_EMAIL";
    }

    if (idMatch) {
      return "ID";
    }

    if (emailMatch) {
      return "EMAIL";
    }
  }

  return null;
}

function duplicateAlreadyLogged(
  id,
  reason
) {

  const data =
    getLogSheet()
      .getDataRange()
      .getValues();

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      data[i][2].toString() === id &&
      data[i][4] === "Duplicate" &&
      data[i][6] === reason
    ) {
      return true;
    }
  }

  return false;
}