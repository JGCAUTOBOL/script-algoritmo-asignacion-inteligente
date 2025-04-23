function onOpen() {
    var Menu = SpreadsheetApp.getUi().createMenu("Actualizacion Novedades").addItem("Actualizar Comerciales", "updateUsers").addToUi();
}
function updateUsers() {
    const managementTable = SpreadsheetApp.openById("1qmOSNXbz2KQxqbQLF-V0uoFUds_EoA55RyJgQQqVCQA").getSheetByName("Data")
    let userStatusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tabla Estado Usuarios")
    let allUsers = userStatusSheet.getRange("A2:J" + userStatusSheet.getLastRow()).getValues();
    let userUnable = managementTable.getRange("A2:E" + managementTable.getLastRow()).getValues();
    let usersMap = {};
    userUnable.forEach((row, index) => {
        let email = row[2];
        if (email) {
            usersMap[email] = index + 2;
        }
    });
    allUsers.forEach(row => {
        let name = row[0]
        let email = row[1]
        let status = row[2]

        if (usersMap[email]) {
            let fila = usersMap[email] + 2;
            if (status != "Activo") {
                managementTable.getRange(fila, 5).setValue("novedad")
            } else {
                managementTable.getRange(fila, 5).setValue("")

            }
        } else {
            let lastValue = managementTable.getRange(managementTable.getLastRow(), 1).getValue();
            let nuevoNumero = (isNaN(lastValue) || lastValue === "") ? 1 : lastValue + 1;

            name = name.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
            if (status != "Activo") {
                managementTable.appendRow([nuevoNumero, name, email, "Nuevo", "Novedad"]);
            } else {
                managementTable.appendRow([nuevoNumero, name, email, "Nuevo", ""]);
            }
            console.log("Added " + name);
            let formulaRange = managementTable.getRange("F2:K2");
            formulaRange.copyTo(managementTable.getRange(managementTable.getLastRow(), 6, 1, 6), { contentsOnly: false });
        }
    });
}
