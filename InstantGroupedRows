function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Group Controls')
      .addItem('Show Grouped Rows', 'expandAll')
      .addItem('Hide Grouped Rows', 'collapseAll')
      .addToUi();
}

function expandAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.showRows(1, sheet.getMaxRows());
}

function collapseAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var lastRow = range.getLastRow();
  
  // Get array of row depths all at once to minimize API calls
  var depths = Array.from({length: lastRow}, (_, i) => sheet.getRowGroupDepth(i + 1));
  
  // Hide all grouped rows in one operation
  sheet.hideRows(1, lastRow);
  
  // Show all non-grouped rows in one operation
  depths.forEach((depth, index) => {
    if (depth === 0) {
      sheet.showRows(index + 1, 1);
    }
  });
}
