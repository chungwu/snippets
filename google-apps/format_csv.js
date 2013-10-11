/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function transformRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var numColumns = rows.getNumColumns();
  var values = rows.getValues();
  
  for (var i = 0; i < numRows; i++) {
    var row = values[i];
    for (var c = 0; c < row.length; c++) {
      var cell = row[c];
      
      if (typeof cell != "string" || cell.indexOf("*") != 0) {
        continue;
      }
      
      // Getting the range is an expensive function, so we only call 
      // it if we know we need to transform data
      var range = sheet.getRange(i+1, c+1);
      var value = range.getValue();
      while (value.indexOf('*') == 0) {
        newValue = value.substring(2)
        
        if (value.indexOf('*F') == 0) {
          // Formulas have to be handled specially
          range.setFormula(newValue);
        } else {      
          range.setValue(newValue);
          updateCellStyle(value, range);
        }
        value = newValue;
      }
      
    }
  }
}

function updateCellStyle(value, range) {
  if (value.indexOf('*T') == 0) {
    // Title
    range.setFontWeight("bold");       
  } else if (value.indexOf('*H') == 0) {
    // Header
    range.setBackgroundColor("#efefef");
  } else if (value.indexOf('*D') == 0) {
    // Dollar amount
    range.setNumberFormat("$0");
  } else if (value.indexOf('*P') == 0) {
    // Percent
    range.setNumberFormat("0.00%");        
  } else if (value.indexOf('*N') == 0) {
    // Simple number
    range.setNumberFormat("0");        
  }  
}