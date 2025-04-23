function assignLeadPenalized(emailsPenalized) {
    const dbUsers = SpreadsheetApp.openById("1qmOSNXbz2KQxqbQLF-V0uoFUds_EoA55RyJgQQqVCQA");
    const dataSheet = dbUsers.getSheetByName("Data");
    let userData = dataSheet.getRange("A1:K" + dataSheet.getLastRow()).getDisplayValues();
    userData = userData.filter(row => {
      return !emailsPenalized.includes(row[2])
    })
    let bestAgent = null;
    let highestEffectiveness = -1;
    let bestSortingKey = Number.POSITIVE_INFINITY;
    let equity = 0.5;
    for (let i = 0; i < userData.length; i++) {
      let id = userData[i][0];
      let agentName = userData[i][1];
      let email = userData[i][2];
      let novelty = userData[i][4];
      let totalCapacity = Number(userData[i][5]);
      let totalInProcess = Number(userData[i][7]);
      let effectivenessStr = userData[i][10] ? userData[i][10].toString().replace("%", "").trim() : "0";
      let effectiveness = Number(effectivenessStr);
      if (!novelty || novelty.toString().trim() === "") {
        let sortingKey = (-(1 - equity) * totalCapacity) + (equity * totalInProcess);
        if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)) {
          bestSortingKey = sortingKey;
          highestEffectiveness = effectiveness;
          bestAgent = {
            name: agentName,
            email: email,
          };
        }
      }
    }
    if (bestAgent !== null) {
      Logger.log("The most available agent is:");
      Logger.log(bestAgent);
      return bestAgent;
    } else {
      Logger.log("No agents available.");
      return null;
    }
  }
  
  
  function reassignLead() {
    const sheet = SheetWebAppGestion;
    const sheetUsers = SpreadsheetApp.openById("1vnC-1VugKccEd8JKS6WpuWGzKTADdCzOx4YV92formY").getSheetByName("Tabla Estado Usuarios")
    let dataRange = sheet.getRange("A2:BL" + sheet.getLastRow()).getValues()
    dataRange = dataRange.filter(function (row) {
      return row[8] === "Pendiente Gestión";
    });
    dataRange = dataRange.map(row => {
      return [row[0], row[3], row[1], row[63], row[33], row[32], row[27], row[29]]
    })
    let arrayReassigned = [];
    let emailsPenalized = [];
    for (let i = 0; i < dataRange.length; i++) {
      let date = dataRange[i][0];
      let currentEmail = dataRange[i][1];
      let uniqueId = dataRange[i][2];
      let reassignmentDate = dataRange[i][3];
      let nit = dataRange[i][4];
      let complexName = dataRange[i][5];
      let adminName = dataRange[i][6];
      let adminPhone = dataRange[i][7];
      let now = new Date();
  
      if (reassignmentDate instanceof Date) {
        let workingHoursDifference = countWorkingHours(reassignmentDate, now);
        if (workingHoursDifference > 4) {
          arrayReassigned.push([date, currentEmail, uniqueId, workingHoursDifference, nit, adminName, complexName, adminPhone]);
          emailsPenalized.push(currentEmail);
        }
      } else if (!reassignmentDate && date instanceof Date) {
        let workingHoursDifference = countWorkingHours(date, now);
        if (workingHoursDifference > 4) {
          arrayReassigned.push([date, currentEmail, uniqueId, workingHoursDifference, nit, adminName, complexName, adminPhone]);
          emailsPenalized.push(currentEmail);
        }
      }
    }
    arrayReassigned.forEach(row => {
      let uniqueId = row[2]
      let currentEmail = row[1]
      let adminPhone = row[7]
      let adminName = row[6]
      let complexName = row[5]
      let nit = row[4]
      let assignedEmail = assignLeadPenalized(emailsPenalized).email.toLowerCase().trim();
      let reassignedDate = new Date();
      console.log("The assigned email is: " + assignedEmail)
      let rowWebhook = sheetUsers.getRange("B:B").createTextFinder(assignedEmail).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
      let webHook = sheetUsers.getRange(rowWebhook, 10).getValues();
      let advisorName = sheetUsers.getRange(rowWebhook, 1).getValue().split(" ")[0].toLowerCase().replace(/^./, c => c.toUpperCase());
      let sheetRow = sheet.getRange("B:B").createTextFinder(uniqueId).matchEntireCell(true).ignoreDiacritics(true).findPrevious().getRow();
      sheet.getRange(sheetRow, 4).setValue(assignedEmail);
      sheet.getRange(sheetRow, 64).setValue(reassignedDate);
      sendWebHookReasigned(webHook, complexName, nit, adminName, adminPhone, advisorName)
    })
  }
  
  function countWorkingHours(startDate, endDate) {
    let hoursDifference = 0;
    while (startDate < endDate) {
      let nextHour = new Date(startDate.getTime() + 60 * 60 * 1000)
      if (isWorkingHour(startDate)) {
        hoursDifference += 1;
      }
      startDate = nextHour;
    }
    return hoursDifference;
  }
  
  
  function isWorkingHour(currentDate) {
    const dayOfWeek = currentDate.getDay();
    const currentHour = currentDate.getHours();
    let isWeekdayValid;
    let isHolidayValid;
    let isHourValid;
    const colombiaHolidays = [
      new Date(currentDate.getFullYear(), 0, 1),
      new Date(currentDate.getFullYear(), 0, 6),
      new Date(currentDate.getFullYear(), 2, 24),
      new Date(currentDate.getFullYear(), 3, 17),
      new Date(currentDate.getFullYear(), 3, 18),
      new Date(currentDate.getFullYear(), 4, 1),
      new Date(currentDate.getFullYear(), 5, 2),
      new Date(currentDate.getFullYear(), 5, 23),
      new Date(currentDate.getFullYear(), 5, 30),
      new Date(currentDate.getFullYear(), 6, 20),  // Independence day was missing
      new Date(currentDate.getFullYear(), 7, 7),
      new Date(currentDate.getFullYear(), 7, 18),
      new Date(currentDate.getFullYear(), 9, 13),
      new Date(currentDate.getFullYear(), 10, 3),
      new Date(currentDate.getFullYear(), 10, 17),
      new Date(currentDate.getFullYear(), 11, 8),
      new Date(currentDate.getFullYear(), 11, 25)
    ];
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      isWeekdayValid = true;
    }
    for (let i = 0; i < colombiaHolidays.length; i++) {
      const holiday = colombiaHolidays[i];
      if (currentDate.getFullYear() === holiday.getFullYear() && currentDate.getMonth() === holiday.getMonth() && currentDate.getDate() === holiday.getDate()) {
        isHolidayValid = true;
        break;
      } else {
        isHolidayValid = false;
        break;
      }
    }
    if (currentHour >= 8 && currentHour <= 17) {
      isHourValid = true;
    } else {
      isHourValid = false;
    }
    if (isWeekdayValid === true && isHolidayValid === false && isHourValid === true) {
      return true
    } else
      return false
  }
  