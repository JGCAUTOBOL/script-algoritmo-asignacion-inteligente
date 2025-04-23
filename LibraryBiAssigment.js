const dbUsers = SpreadsheetApp.openById("1effw7xHg2YAUbfaxl9_o94-KTXnSSCLI3HOAHOmqmMs");
const dataSheet = dbUsers.getSheetByName("Data");

function AssignLead() {
  let userData = dataSheet.getRange("A1:K" + dataSheet.getLastRow()).getDisplayValues();
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
      let availability = totalCapacity - totalInProcess;
      let sortingKey = (-(1 - equity) * totalCapacity) + (equity * totalInProcess);

      if (sortingKey < bestSortingKey || (sortingKey === bestSortingKey && effectiveness > highestEffectiveness)) {
        bestSortingKey = sortingKey;
        highestEffectiveness = effectiveness;

        bestAgent = {
          name: agentName,
          email: email
        };
      }
    }
  }

  if (bestAgent !== null) {
    console.log("The most available agent is: " + bestAgent);
    return bestAgent;
  } else {
    Logger.log("No agents available.");
    return null;
  }
}

