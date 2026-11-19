function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet("Architect");
    sheet.getRange("B1").setValue("✍️ HackPropagation");
    sheet.getRange("G4").setValue("Mode");
    sheet.getRange("H4").setValue("Epochs ↓");
    const arrowrange = sheet.getRangeByIndexes(4, 1, 50, 5);
    arrowrange.setValue("");
    const messagerane = sheet.getRangeByIndexes(5, 6, 50, 2);
    messagerane.setValue("");
    sheet.getCell(4, 7).setValue(1);
    sheet.getRange("G5").setControl({ type: ExcelScript.CellControlType.empty });
    sheet.getCell(4, 6).setValue("Backpropagation");
    const sheets = workbook.getWorksheets();
    for (let i = sheets.length - 1; i >= 0; i--) {
        const sheet = sheets[i];
        const name = sheet.getName();
        if (name === "Blueprint" || name === "Blueprint_inference" || /^Blueprint_ep\d+$/.test(name)) {
            sheet.delete();
        }
    }
    let dashboardSheet = workbook.getWorksheet("Dashboard");
    if (dashboardSheet) {
        dashboardSheet.delete();
    }
}
