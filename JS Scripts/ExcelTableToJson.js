function convertExcelDataToJSON(excelDataRows, columnHeaders) {
  var jsonArray = [];

  // Helper function to check if a variable is an array
  function checkIsArray(variable) {
    return Object.prototype.toString.call(variable) === "[object Array]";
  }

  // Validate excelDataRows: it should be an array of arrays
  if (
    !excelDataRows ||
    !checkIsArray(excelDataRows) ||
    excelDataRows.length === 0
  ) {
    // Return an empty JSON array string if input is invalid or empty
    return JSON.stringify([]);
  }

  var numDataRows = excelDataRows.length;
  var numColsInFirstDataRow = 0;
  if (excelDataRows[0] && checkIsArray(excelDataRows[0])) {
    numColsInFirstDataRow = excelDataRows[0].length;
  }

  var keys;
  // Determine column keys for the JSON objects
  if (
    columnHeaders &&
    checkIsArray(columnHeaders) &&
    columnHeaders.length > 0
  ) {
    // Use provided columnHeaders
    keys = columnHeaders;
  } else {
    // Generate generic keys (e.g., column1, column2, ...) if no headers are provided
    // or if the first data row is empty/invalid for inferring column count
    keys = [];
    var colsToGenerateKeysFor = numColsInFirstDataRow;
    // If columnHeaders were expected but not valid, and numColsInFirstDataRow is 0,
    // this loop won't run, resulting in empty objects if numDataRows > 0.
    for (var j = 0; j < colsToGenerateKeysFor; j++) {
      keys.push("column" + (j + 1));
    }
  }

  // Iterate over each row in excelDataRows
  for (var i = 0; i < numDataRows; i++) {
    var currentRowData = excelDataRows[i];
    var rowObject = {};

    if (checkIsArray(currentRowData)) {
      // Map data from the current row to the keys
      for (var k = 0; k < keys.length; k++) {
        // If currentRowData has fewer items than keys, remaining keys get null
        rowObject[keys[k]] =
          k < currentRowData.length ? currentRowData[k] : null;
      }
    } else {
      // If a row in excelDataRows is not an array (e.g., null or other type),
      // create an object with null values for all keys.
      // This ensures the output JSON array maintains an object for each input "row".
      if (keys.length > 0) {
        for (var k_null = 0; k_null < keys.length; k_null++) {
          rowObject[keys[k_null]] = null;
        }
      }
      // If keys is also empty (e.g. bad input, no headers), rowObject will be {}
    }
    jsonArray.push(rowObject);
  }

  // Convert the array of objects into a JSON string (pretty-printed with 2 spaces for indentation)
  return JSON.stringify(jsonArray, null, 2);
}

var ExcelData = [
  ["Row1Value1", "Row1Value2", "Row1Value3", "Row1Value4"],
  ["Row2Value1", "Row2Value2", "Row2Value3", "Row2Value4"],
  // ... (7 more rows) ...
  ["Row9Value1", "Row9Value2", "Row9Value3", "Row9Value4"],
];

var mySpecificHeaders = ["Header1", "Header2", "Header3", "Header4"]; // Replace with your actual headers
var jsonOutput = convertExcelDataToJSON(ExcelData, mySpecificHeaders);
WScript.Echo(jsonOutput);
