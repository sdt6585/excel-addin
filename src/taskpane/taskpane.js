Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Instead of waiting for a button click, call your initialization function directly.
    initializeEventHandler();
  }
});

// Global variable and constants.
let previousRowCount = 0;
const sheetName = "Sexuality";
const tableName = "SexualityTracker";

// Your initialization function that registers the event handler.
async function initializeEventHandler() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const table = sheet.tables.getItem(tableName);

      table.load("rows/count");
      await context.sync();
      previousRowCount = table.rows.count;

      // Register the event handler for changes.
      sheet.onChanged.add(handleWorksheetChanged);
      await context.sync();

      console.log("Event handler registered on sheet:", sheetName, "with initial row count:", previousRowCount);
    });
  } catch (error) {
    handleError(error)
  }
}

// Your event handler that updates new rows.
async function handleWorksheetChanged(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const table = sheet.tables.getItem(tableName);

      table.load("rows/count");
      await context.sync();

      const currentRowCount = table.rows.count;

      if (currentRowCount > previousRowCount) {
        const newRowsCount = currentRowCount - previousRowCount;
        console.log(`${newRowsCount} new row(s) added.`);

        const dataBodyRange = table.getDataBodyRange();
        dataBodyRange.load(["values", "rowCount"]);
        await context.sync();

        const now = new Date().toLocaleString();
        const values = dataBodyRange.values;
        for (let i = previousRowCount; i < currentRowCount; i++) {
          values[i][0] = now; // Update the first column.
        }

        dataBodyRange.values = values;
        previousRowCount = currentRowCount;
        await context.sync();
        console.log("Updated new rows with date/time:", now);
      } else {
        previousRowCount = currentRowCount;
      }
    });
  } catch (error) {
    handleError(error);
  }
}

/**
 * Error Handling function
 * @param {Error} error
 */
function handleError(error) {
  // Log error for debugging.
  console.error(error);

  // Update the error message div.
  const errorDiv = document.getElementById("error-message");
  errorDiv.innerText = `Error: ${error.message || error}`;
  errorDiv.style.display = "block";

  // Optionally, show a pop-up alert.
  alert(`An error occurred: ${error.message || error}`);
}
