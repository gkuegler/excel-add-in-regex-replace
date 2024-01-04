/**
 * API REFERENCE
 * https://learn.microsoft.com/en-us/javascript/api/excel?view=excel-js-preview
 */

// TODO: enable every reporting dialogbox
// TODO: enable match options and modes
// TODO: fixed display alignment issues
// TODO: implement custom undo

// All the controls except the 'match mode' have been implemented.

function replaceTableData(table, data) {
  // Testing data.
  // const data = [
  // ["A1", "existing", "replacement"],m
  // ["B1", "existing", "replacement"],
  // ];

  // If there are no data to add. Clear the table and return.
  if (data.length < 1) {
    table.tBodies[0].remove(); // 1 DOM reflow 1 repaint
    return;
  }
  const rows = data.length;
  const columns = data[0].length;
  const tB = document.createElement("tbody");
  const frag = document.createDocumentFragment();

  for (let r = 0; r < rows; r++) {
    let row = tB.insertRow();
    for (let c = 0; c < columns; c++) {
      row.insertCell().textContent = data[r][c];
    }
  }
  // Remove the existing table body.
  if (table.tBodies.length > 0) {
    table.tBodies[0].remove(); // 1 DOM reflow 1 repaint
  }

  // Add the table body.
  if (data.length > 0) {
    frag.append(tB);
    table.append(frag); // 1 DOM reflow 1 repaint
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Present to the user that an error method.
    document.getElementById("sideload-msg").style.display = "none";

    // Assign event handlers
    document.getElementById("search").onclick = () => tryCatch(onSearch);
    document.getElementById("replace").onclick = () => tryCatch(onReplace);

    document.getElementById("app-body").style.display = "flex";
  }
});

async function onSearch() {
  await searchOrReplace(true);
}

async function onReplace() {
  await searchOrReplace(false);
}

async function searchOrReplace(searchOnly) {
  await Excel.run(async (context) => {
    var searchPattern = document.getElementById("search_ptrn").value;
    var replacePattern = document.getElementById("replace_ptrn").value;
    var useRegex = document.getElementById("useRegex").checked;
    var caseInsensitive = document.getElementById("caseInsensitive").checked;
    var useActiveSelection = document.getElementById("useActiveSelection").checked;

    // Find the match mode radio box election.
    var radios = document.getElementsByName("matchMode");
    var matchMode;

    for (var i = 0, length = radios.length; i < length; i++) {
      if (radios[i].checked) {
        // only one radio can be logically checked, don't check the rest
        matchMode = radios[i].value;
        break;
      }
    }

    var flags = "";

    if (caseInsensitive) {
      flags += "i";
    }

    var regexPatternFull = useRegex ? new RegExp(searchPattern, flags) : searchPattern;

    var range = useActiveSelection
      ? context.workbook.getSelectedRange()
      : context.workbook.worksheets.getActiveWorksheet().getUsedRange();

    // Need to load the Excel proxy object property.
    range.load("address");
    range.load("rowCount");
    range.load("columnCount");
    range.load("values");

    // Need to call the sync after the load function
    // before I can access the proxy object property.
    await context.sync();

    // Apparently 'forEach' not a function
    // range.forEach(row => {})

    var values = range.values;
    var newHtmlTableRows = [];
    for (var i = 0; i < range.rowCount; i++) {
      for (var j = 0; j < range.columnCount; j++) {
        const existing = values[i][j];
        // If cell is string type that is not empty.
        if ((typeof existing === "string") & (existing != "")) {
          const replacement = existing.replace(regexPatternFull, replacePattern);

          // Find if the cell's contents would be replaced.
          if (replacement != existing) {
            // Add to the output table.
            newHtmlTableRows.push(["A1", existing, replacement]);
            values[i][j] = replacement;
          }
        }
      }
    }

    replaceTableData(document.querySelector("table"), newHtmlTableRows);

    // This is the only syntax I could get to work for
    // updating the values in the selected range.
    if (!searchOnly) {
      range.values = values;
    }

    // The context.sync method sends
    // all queued commands to Excel for execution.
    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your
    // add-in's UI.
    console.error(error);
  }
}
