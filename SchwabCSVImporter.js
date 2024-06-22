// This file can be used by people receiving ESPP and RSU in Europe (to ease tax process)
//
// It takes as input a file input.csv that must be placed in your google drive folder.
// On Schwab site you can download a CSV file for a specific date range
// It open the CSV, clean up numbers, add 3 columns and add the googlefinance formulas to get the converted rate (Euros).
// the google finance function allows only the use of close with USDEUR pair, so I can't propose an optimised convertion rate with the lowest of the day.

function importCSVFromDrive() {
  const fileId = 'input.csv';  // CSV file from Schwab site, to place in your google drive folder to import
  const sheetName = 'Feuille 1';  // Replace with the name of the sheet where you want to import the data. This sheet must exist.

  // Get the CSV file from Google Drive
  const files = DriveApp.getFilesByName(fileId);
  const file = files.next();
  const csvData = file.getBlob().getDataAsString();
  
  // Parse the CSV data
  const data = CSVToArray(csvData);

  // Insert 3 new columns for Euro converted rates
  data.forEach(row => {
    row.splice(11, 0, ''); // SalePrice (converted)
    row.splice(21, 0, ''); // VestFairMarketValue (converted)
    row.splice(23, 0, ''); // Gross Proceeds (converted)
    row.splice(24, 0, ''); // Capital gain and loss (converted)
  });

  const numColumns = data[0].length;

  // To repeat the date in first column to ease the googlefinance formula
  var lastDateSeen = "";
  
  // index for date formula
  var DateIdx = 1;

  // required as some cells can be empty
  const normalizedData = data.map(row=> { 
    
    //repeat the date in first column
    if (row[0] != undefined) {
      // we are in a single date range here
      lastDateSeen = row[0];      
    } else {
      row[0] = lastDateSeen;
    }
    
    // need a CSV with dividends to complete
    // if RSU Sale or ESPP Sale    
    if (row[8] == "RS" || row[8] == "ESPP") {    
      
      //remove $ sign  
      row[10] = row[10].replace(/[^0-9.-]+/g, "");
      // replace dot with comma, use it if located in France
      row[10] = row[10].replace(/\./g, ",");

      //SubscriptionFairMarketValue      
      if (row[13]) {
        row[13] = row[13].replace(/[^0-9.-]+/g, "");
        row[13] = row[13].replace(/\./g, ",");
      }

      //PurchasePrice
      if (row[15]) {
        row[15] = row[15].replace(/[^0-9.-]+/g, "");
        row[15] = row[15].replace(/\./g, ",");
      }

      // VestFairMarketValue
      if (row[20]) {
        row[20] = row[20].replace(/[^0-9.-]+/g, "");
        row[20] = row[20].replace(/\./g, ",");
      }

      // GrossProceeds
      if (row[22]) {
        row[22] = row[22].replace(/[^0-9.-]+/g, "");
        row[22] = row[22].replace(/\./g, ",");
      } 

      // SalePrice 
      row[11] = "=INDEX(GOOGLEFINANCE(\"currency:USDEUR\";\"close\";DATE(RIGHT(A"+DateIdx+";4);LEFT(A"+DateIdx+";2);MID(A"+DateIdx+";4;2)));2;2) * K"+DateIdx;    

      // VestFairMarketValue 
      row[21] = "=INDEX(GOOGLEFINANCE(\"currency:USDEUR\";\"close\";DATE(RIGHT(A"+DateIdx+";4);LEFT(A"+DateIdx+";2);MID(A"+DateIdx+";4;2)));2;2) * U"+DateIdx;

      // Gross Proceeds
      row[23] = "=INDEX(GOOGLEFINANCE(\"currency:USDEUR\";\"close\";DATE(RIGHT(A"+DateIdx+";4);LEFT(A"+DateIdx+";2);MID(A"+DateIdx+";4;2)));2;2) * W"+DateIdx;

      // Capital gain and loss
      if (row[8] == "RS")
        row[24] = "=X" + DateIdx + " - ( J" + DateIdx + " * V" + DateIdx + ")";
      if (row[8] == "ESPP")
        row[24] = "=X" + DateIdx + " - ( J" + DateIdx + " * P" + DateIdx + " * "+ "INDEX(GOOGLEFINANCE(\"currency:USDEUR\";\"close\";DATE(RIGHT(A"+DateIdx+";4);LEFT(A"+DateIdx+";2);MID(A"+DateIdx+";4;2)));2;2))";
    }

    while (row.length < numColumns) {
      row.push('');
    }
    DateIdx++;
    return row;
  });

  // Open the Google Sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // Clear any existing content in the sheet
  sheet.clear();

  // Import the CSV data into the Google Sheet
  sheet.getRange(1, 1, normalizedData.length, numColumns).setValues(normalizedData);

  // Add title to the new columns
  sheet.getRange(1, 12).setValue('1 Share in Euros');
  sheet.getRange(1, 22).setValue('VestFairMarketValue in Euros');
  sheet.getRange(1, 24).setValue('Gross proceeds in Euros');
  sheet.getRange(1, 25).setValue('Plus ou moins Value in Euros');
}

/**
 * Helper function to parse CSV data into a 2D array
 * @param {string} strData - The CSV data as a string
 * @param {string} [strDelimiter=','] - The delimiter used in the CSV file (default is comma)
 * @return {Array} - The parsed CSV data as a 2D array
 */
function CSVToArray(strData, strDelimiter) {
  strDelimiter = (strDelimiter || ",");
  const objPattern = new RegExp((
    "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
    "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
    "([^\"\\" + strDelimiter + "\\r\\n]*))"
  ), "gi");

  let arrData = [[]];
  let arrMatches = null;

  while (arrMatches = objPattern.exec(strData)) {
    const strMatchedDelimiter = arrMatches[1];
    if (strMatchedDelimiter.length && strMatchedDelimiter !== strDelimiter) {
      arrData.push([]);
    }

    const strMatchedValue = arrMatches[2] ?
      arrMatches[2].replace(new RegExp("\"\"", "g"), "\"") :
      arrMatches[3];

    arrData[arrData.length - 1].push(strMatchedValue);
  }

  return arrData;
}
