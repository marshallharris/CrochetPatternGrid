// To use this script:
// Open your Google Sheets document.
// Click on "Extensions" in the top menu.
// Select "Apps Script."
// Paste the provided code into the script editor.
// Save the script.
// Run the colorCellsBasedOnRGB function by clicking the play button in the toolbar.

function colorCellsBasedOnRGB() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the active sheet
  var sheet = spreadsheet.getActiveSheet();
  
  // Get the range of cells in the sheet
  var range = sheet.getDataRange();
  
  // Get the values in the range
  var values = range.getValues();
  
  // Iterate through each cell in the range
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      // Get the RGB value from the cell
      var rgbValue = values[i][j];
      
      // Check if the cell contains a valid RGB value
      if (isValidRGB(rgbValue)) {
        // Split the RGB value into individual components
        var rgbComponents = rgbValue.split(',');
        
        // Convert the components to integers
        var red = parseInt(rgbComponents[0]);
        var green = parseInt(rgbComponents[1]);
        var blue = parseInt(rgbComponents[2]);
        
        // Set the background color of the cell
        range.getCell(i + 1, j + 1).setBackgroundRGB(red, green, blue);
      }
    }
  }
}

// Helper function to check if a given string is a valid RGB value
function isValidRGB(rgbValue) {
  // Use a regular expression to match the RGB format (e.g., "255, 0, 0")
  var rgbRegex = /^(\d{1,3},\s*\d{1,3},\s*\d{1,3})$/;
  
  return rgbRegex.test(rgbValue);
}
