/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }
    document.getElementById("random-number").onclick = splitName;
    document.getElementById("open-dialog").onclick = openDialog;
  }
});

async function generateRandomNumber() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "columnCount", "getCell"]);
      await context.sync();

      const numRows = range.rowCount;
      const numCols = range.columnCount;

      // Loop through each cell in the range and set its value
      for (let i = 0; i < numRows; i++) {
        for (let j = 0; j < numCols; j++) {
          const randomNumber = createRand();
          const cell = range.getCell(i, j);
          cell.values = [[randomNumber]];
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

const createRand = () => {
  return Math.random();
};
let dialog = null;

function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },

    // TODO2: Add callback parameter.
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

const splitName = () => {
  Excel.run(async function (context) {
    let sheet = context.workbook.worksheets.getItem("Sheet1");
    let field1 = sheet.getCell(2, 3);
    field1.load("values");
    let field2 = sheet.getCell(2, 4);
    field2.load("values");
    let field3 = sheet.getCell(2, 5);
    field3.load("values");
    let activeCell = context.workbook.getActiveCell();
    activeCell.load("values");
    await context.sync();
    const cellValueAsArray = activeCell.values[0];
    const fullName = cellValueAsArray[0];
    const nameParts = fullName.split(" ");

    if (nameParts.length > 2) {
      field1.values = nameParts[0];
      field2.values = nameParts[1];
      field3.values = nameParts[2];
    } else {
      field1.values = nameParts[0];
      field2.values = "";
      field3.values = nameParts[1];
    }

    // console.log(nameParts);
  }).catch(function (error) {
    console.log("Error: " + error);
  });
};
