function getTableHtml(labels_data, template_name) {
  let htmlTemplate = HtmlService.createTemplateFromFile(template_name);
  htmlTemplate.labels_data = labels_data;
  let htmlBody = htmlTemplate.evaluate().getContent();
  return htmlBody;
}

function toInt(v) {
  const n = Number(v); // converts strings, undefined → NaN, etc.
  return Number.isNaN(n) ? 0 : n;
}

function set_data(headers, data, target_sheet){
  target_sheet.clear();
  target_sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  target_sheet.getRange(2, 1, data.length, data[0].length).setValues(data)
  applyZebraStriping(target_sheet.getDataRange())
}

function applyZebraStriping(range) {
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const blue = "#cfe6f7";
  const gray = "#f2f2f2";

  // Prepare 2D array of background colors
  const backgrounds = [];

  for (let row = 0; row < numRows; row++) {
    const rowColor = row % 2 === 0 ? blue : gray;
    backgrounds.push(Array(numCols).fill(rowColor));
  }

  // Apply the colors to the range
  range.setBackgrounds(backgrounds);
}

//Function Button
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Agency')
    .addItem('Generate Postal Shipping Label', 'generate_postal_shipping_label')
    .addItem('Generate Roadways Shipping Label', 'generate_roadways_shipping_label')
    .addItem('Generate Door Delivery Label', 'create_door_delivery_labels')
    .addItem('Generate All', 'generate_all_shipping_labels')
    .addItem('Magazine Details Form', 'show_form')
    .addItem('Generate Postal Invoice', 'create_postal_invoice')
    .addToUi();
}

function get_formatted_ss_no(ss_no) {
  ss_no = ss_no.trim()
  let formatted_ss_no = "";
  let is_subscr_no_started = false;

  for (let i = 0; i < ss_no.length; i++) {
    let char = ss_no[i];

    if (is_subscr_no_started) {
      formatted_ss_no += char;
    } else {
      if (/\d/.test(char)) {
        if (char !== "0") {
          is_subscr_no_started = true;
          formatted_ss_no += char;
        }
        // if it's "0", skip it (leading zero)
      } else if (char === "-") {
        formatted_ss_no += " ";
      } else {
        formatted_ss_no += char;
      }
    }
  }

  return formatted_ss_no;
}

function is_sheet_exists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName) !== null;
}

function getSheet(sheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    Logger.log(`Sheet "${sheetName}" already exists.`);
    return sheet;
  }
  // doesn’t exist → create it
  return ss.insertSheet(sheetName);
}
