function TrackPortfolio() {
  
  // Constants
  const SHEET_NAME = "PortfolioTracker"; // Active sheet
  const ROW_INIT = 11; // Where to start writing dates and prices
  const COL_INIT = 1;
  
  // Load sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  // Load startDate, endDate, and interval (hyperparams)
  var startDate = new Date(sheet.getRange("B2").getValue());
  var endDate = new Date(sheet.getRange("B3").getValue());
  var interval = sheet.getRange("B4").getValue();
  
  // Clear previous data from sheet - avoid text overlap
  clearSheetFromRow(sheet, ROW_INIT, COL_INIT);

  // Create date array and populate dates in the sheet
  var datesArray = createDateArray(startDate, endDate, interval);
  var numDates = datesArray.length;
  sheet.getRange(ROW_INIT, COL_INIT, numDates).setHorizontalAlignment("center").setValues(datesArray);

  // Load tickers and generate price formulas
  var tickers = loadTickers(sheet);
  var numTickers = tickers.length;
  var arrayPrices = createPriceArray(sheet, tickers, numDates, ROW_INIT);

  // Write ticker prices to the sheet
  sheet.getRange(ROW_INIT, COL_INIT + 2, numDates, numTickers).setFontColor("#c0bfbf").setHorizontalAlignment("center").setValues(arrayPrices);

  // Calculate portfolio value and write formulas to the sheet
  var workingCells = sheet.getRange(ROW_INIT + 1, COL_INIT, numDates - 1, numTickers + 2);
  var a1NotationMatrix = convertRangeToA1NotationMatrix(workingCells);
  var allFormulas = calculatePortfolioFormulas(a1NotationMatrix, tickers);

  // Write the portfolio header
  sheet.getRange(ROW_INIT, COL_INIT + 1).setHorizontalAlignment("center").setValue("PORTFOLIO");
  
  // Write portfolio values to the sheet
  sheet.getRange(ROW_INIT + 1, COL_INIT + 1, a1NotationMatrix.length).setHorizontalAlignment("center").setValue(allFormulas);
}

/**
 * Clears the content, formatting, and data from the sheet starting from a given row.
 */
function clearSheetFromRow(sheet, rowInit, colInit) {
  // Find last row and column
  var lastRow = sheet.getLastRow(); 
  var lastColumn = sheet.getLastColumn();
  // Ensure there is actually something..
  if (lastRow > rowInit) {
    var cleaningRange = sheet.getRange(rowInit, colInit, lastRow - rowInit + 1, lastColumn);
    cleaningRange.clear();
  }
}

/**
 * Creates an array of dates between startDate and endDate with a given interval.
 */
function createDateArray(startDate, endDate, interval) {
  var datesArray = [["DATE"]]; // Initialize with header
                               // To write content on the sheet we need a 2d-array
  var currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    var currentDateString = `=DATE(${Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy;M;d")})`;
    datesArray.push([currentDateString]);
    currentDate.setDate(currentDate.getDate() + interval);
  }
  return datesArray;
}

/**
 * Loads tickers from the sheet, removing any spaces.
 */
function loadTickers(sheet) {
  // This is a single String (e.g., "AMS:IWDA; LON:EIMI")
  var tickersMerged = sheet.getRange("B1").getValue(); 
  // Remove space and split with semicolon
  return tickersMerged.replace(/\s/g, "").split(";"); 
}

/**
 * Creates a 2D array with price formulas for each ticker and date.
 */
function createPriceArray(sheet, tickers, numDates, rowInit) {
  var arrayPrices = Array.from({ length: numDates }, () => Array(tickers.length).fill(null));
  
  // Set the tickers header
  // e.g., [[AMS:IWDA,LON:EIMI],[[null,null],..,[[null,null]]
  for (var tickerID = 0; tickerID < tickers.length; tickerID++) {
    arrayPrices[0][tickerID] = tickers[tickerID];
  }

  // Populate price formulas
  for (var dateID = 1; dateID < numDates; dateID++) {
    var currentDateCell = sheet.getRange(rowInit + dateID, 1).getA1Notation();
    
    for (var tickerID = 0; tickerID < tickers.length; tickerID++) {
      // Handle weekends
      // .. GOOGLEFINANCE already handles weekends. However, if END_DATE=TODAY() is on a weekend, it fails..
      var closestWeekDay = `IF(WEEKDAY(${currentDateCell}; 2) > 5; WORKDAY(${currentDateCell}; -1); ${currentDateCell})`;
      var currentFormula = `=INDEX(GOOGLEFINANCE("${tickers[tickerID]}"; "CLOSE"; ${closestWeekDay}); 2; 2)`;
      arrayPrices[dateID][tickerID] = currentFormula;
    }
  }

  return arrayPrices;
}

/**
 * Converts a given range into a 2D array of A1 notations (i.e., (2,2) = B2).
 */
function convertRangeToA1NotationMatrix(range) {
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var a1NotationMatrix = Array.from({ length: numRows }, () => Array(numCols).fill(null));

  for (var row = 1; row <= numRows; row++) {
    for (var col = 1; col <= numCols; col++) {
      var cell = range.getCell(row, col);
      a1NotationMatrix[row - 1][col - 1] = cell.getA1Notation();
    }
  }
  return a1NotationMatrix;
}

/**
 * Calculates the portfolio value for each date.
 */
function calculatePortfolioFormulas(a1NotationMatrix, tickers) {
  var allFormulas = new Array(a1NotationMatrix.length);

  for (var dateID = 0; dateID < a1NotationMatrix.length; dateID++) {
    var dateCell = a1NotationMatrix[dateID][0];
    var fullCellFormula = "="; // Init formula

    // Iterate over the tickers. 
    // ..start from 2 as we first have Date and Portfolio
    for (var tickerID = 2; tickerID < a1NotationMatrix[0].length; tickerID++) {
      var cellPrice = a1NotationMatrix[dateID][tickerID];
      var currTicker = tickers[tickerID - 2]; // remove the shift

      // Return the value of the current "ticker" in your portfolio at the current "date"
      var cellFormula = `IFNA(INDEX(FILTER(Orders!$H$4:$H; Orders!$B$4:$B = MAX(FILTER(Orders!$B$4:$B; Orders!$B$4:$B < ${dateCell}; Orders!$C$4:$C = "${currTicker}")); Orders!$C$4:$C = "${currTicker}"); COUNT(FILTER(Orders!$I$4:$I; Orders!$B$4:$B = MAX(FILTER(Orders!$B$4:$B; Orders!$B$4:$B < ${dateCell}; Orders!$C$4:$C = "${currTicker}")); Orders!$C$4:$C = "${currTicker}")));0) * ${cellPrice}`;

      // Sum all the individual contributions
      fullCellFormula = `${fullCellFormula} + ${cellFormula}`;
    }

    // Store final formula
    allFormulas[dateID] = fullCellFormula;
  }

  return allFormulas;
}