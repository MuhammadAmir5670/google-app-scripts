///////////////////////////////////////////////////////////////////////////////////////////////
// BEGIN EDITS ////////////////////////////////////////////////////////////////////////////////

const TEMPLATE_FILE_ID = "TARGET_DOCX_ID";
const DESTINATION_FOLDER_ID = "TARGET_DRIVE_FOLDER_ID";
const CURRENCY_SIGN = "$";
const HOURS_PER_DAY = 8;

// END EDITS //////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////
// WARNING: EDITING ANYTHING BELOW THIS LINE WILL CHANGE THE BEHAVIOR OF THE SCRIPT. //////////
// DO SO AT YOUR OWN RISK.//// ////////////////////////////////////////////////////////////////
// ----------------------------------------------------------------------------------------- //

// Converts a float to a string value in the desired currency format
function toCurrency(num) {
  var fmt = Number(num).toFixed(0);
  return `${CURRENCY_SIGN}${fmt}`;
}

// Returns the date object
function toDate(dt_string) {
  var millis = Date.parse(dt_string);
  var date = new Date(millis);

  return date;
}

// Format datetimes to: MM/DD/YYYY
function toDateFmt(dt_string) {
  var date = toDate(dt_string);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the date in MM/DD/YYYY format
  return `${month}/${day}/${year}`;
}

// Converts the string into ALL_CAPS_SNAKE_CASE
function toSnakeCase(str) {
  return str
    .replace(/([a-z])([A-Z])/g, "$1_$2") // Insert underscore between camelCase words
    .replace(/\s+/g, "_") // Replace spaces with underscores
    .replace(/[^a-zA-Z0-9_]/g, "") // Remove non-alphanumeric characters (except underscores)
    .toUpperCase(); // Convert to uppercase
}

// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
  // Set temporary variables to hold prices and data.
  var total = 0;
  var responseData = {};

  // Iterate through all of our response data and add the keys (headers)
  // and values (data) to the response dictionary object.
  for (var i = 0; i < values.length; i++) {
    // Extract the key and value
    var key = header[i].toLowerCase();
    var value = values[i];

    // If we have a price, add it to the running total and format it to the
    // desired currency.
    if (key.includes("price")) {
      total += value;
      // format it to the desired currency.
    } else if (key.includes("date")) {
      value = toDate(value);
    }

    // Add the key/value data pair to the response dictionary.
    responseData[key] = value;
  }

  // Once all data is added, we'll adjust the total
  responseData["total"] = toCurrency(total);

  return responseData;
}

// Maps response variables to variables used template
function prepareTemplateData(startDate, endDate, responseData) {
  const templateData = {};

  // Convert keys from responseData to SNAKE_CASE and map to templateData
  Object.keys(responseData).forEach((key) => {
    const snakeCaseKey = toSnakeCase(key);
    templateData[snakeCaseKey] = responseData[key];
  });

  // Template defaults
  templateData["CURRENT_DATE"] = toDateFmt(new Date());
  templateData["PRICE_PER_DAY"] = toCurrency(
    responseData.price * HOURS_PER_DAY,
  );
  templateData["HOURS_PER_DAY"] = HOURS_PER_DAY;

  // Generate billing day descriptions
  var billingDays = generateBillingDays(startDate, endDate, responseData.price);
  var numberOfBillingDays = billingDays.length;
  templateData["NUMBER_OF_BILLING_DAYS"] = numberOfBillingDays;
  templateData["BILLING_DAYS"] = billingDays;

  // Calculating Total
  templateData["TOTAL"] = toCurrency(
    responseData.price * HOURS_PER_DAY * numberOfBillingDays,
  );

  return templateData;
}

function generateBillingDays(startDate, endDate, price) {
  const billingDays = [];
  let currentDate = new Date(startDate);
  let billingDayNumber = 1;

  while (currentDate <= endDate) {
    let dayOfWeek = currentDate.getDay(); // 0 = Sunday, 6 = Saturday
    let currentBillingDay = {};

    // Skip Saturday (6) and Sunday (0)
    if (dayOfWeek != 0 && dayOfWeek != 6) {
      // Add billing day description
      currentBillingDay[`BILLING_DAY_${billingDayNumber}`] =
        toDateFmt(currentDate);

      // Add billing day price and hours
      currentBillingDay[`BILLING_DAY_${billingDayNumber}_PRICE`] = toCurrency(
        price * HOURS_PER_DAY,
      );
      currentBillingDay[`BILLING_DAY_${billingDayNumber}_HOURS`] =
        HOURS_PER_DAY;

      billingDays.push(currentBillingDay);
      billingDayNumber++;
    }

    // Move to the next day
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return billingDays;
}

// Helper function to inject data into the template
function populateTemplate(document, templateData) {
  // Get the document header and body (which contains the text we'll be replacing).
  var documentHeader = document.getHeader();
  var documentBody = document.getBody();

  function replaceTemplateText(data) {
    for (var key in data) {
      const matchText = `{{${key}}}`;
      const value = data[key];

      console.log(matchText, value);

      documentHeader.replaceText(matchText, value);
      documentBody.replaceText(matchText, value);
    }
  }

  // Populate root variables
  replaceTemplateText(templateData);

  // Populate billing variables
  templateData.BILLING_DAYS.forEach((billingDetails, index) => {
    replaceTemplateText(billingDetails);
  });
}

function appendBillingRows(document, billingData) {
  const body = document.getBody();

  // Locate the first table in the document
  const tables = body.getTables();
  if (tables.length < 2) {
    throw new Error("The document does not contain a second table!");
  }
  const table = tables[1]; // Use the second table
  const rowOffsetIndex = 1;

  // Append rows for each billing day
  billingData.forEach((entry, index) => {
    // Construct placeholders for the row
    let billingDayKey = `{{BILLING_DAY_${index + 1}}}`;
    let billingDayValue = `{{BILLING_DAY_${index + 1}_PRICE}}/{{BILLING_DAY_${
      index + 1
    }_HOURS}}`;

    // Append a new row to the table
    let row = table.insertTableRow(rowOffsetIndex + index);
    let descriptionCell = row.appendTableCell(billingDayKey);
    let amountCell = row.appendTableCell(billingDayValue);

    // Set row and cell formatting
    row.setFontFamily("Century Gothic");
    row.setForegroundColor("#404040");
    row.setFontSize(18);

    // Align right cell text to right
    amountCell.getChild(0).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  });
}

// Function to populate the template form
function createDocFromForm() {
  // Get active sheet and tab of our response data spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow() - 1;

  // Get the data from the spreadsheet.
  var range = sheet.getDataRange();

  // Identify the most recent entry and save the data in a variable.
  var data = range.getValues()[lastRow];

  // Extract the headers of the response data to automate string replacement in our template.
  var headers = range.getValues()[0];

  // Parse the form data.
  var responseData = parseFormData(data, headers);
  var startDate = responseData["start date"];
  var endDate = responseData["end date"];
  // Prepare template data
  var templateData = prepareTemplateData(startDate, endDate, responseData);
  console.log(templateData);

  // Retreive the template file and destination folder.
  var templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  var targetFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

  // Copy the template file so we can populate it with our data.
  // The name of the file will be the company name and the invoice number in the format: DATE_COMPANY_NUMBER
  var filename = `Amir's Invoice for | ${toDateFmt(startDate)} - ${toDateFmt(
    endDate,
  )}`;
  var documentCopy = templateFile.makeCopy(filename, targetFolder);

  // Open the copy.
  var document = DocumentApp.openById(documentCopy.getId());

  // Add billing rows to the template
  appendBillingRows(document, templateData.BILLING_DAYS);

  // Populate the template with our form responses and save the file.
  populateTemplate(document, templateData);
  document.saveAndClose();
}
