Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }
    document.getElementById("formatCells").onclick = setHeight;
  }
});

const formatCells = () => {
  Excel.run(async function (context) {
    let range = context.workbook.getSelectedRange();
    range.format.horizontalAlignment = "Center";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
  });
};

const setHeight = async () => {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("B2:C5");
    range.load("dataValidation");
    await context.sync();

    range.dataValidation.clear();

    range.dataValidation.rule = {
      wholeNumber: {
        formula1: 0,
        operator: Excel.DataValidationOperator.greaterThan,
      },
    };
  }).catch(function (error) {
    console.log("Error: " + error);
  });
};
