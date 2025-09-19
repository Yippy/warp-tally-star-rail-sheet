/*
 * Version 1.03 made by yippym - 2025-09-17 01:00
 * https://github.com/Yippy/warp-tally-star-rail-sheet
 */
function getSourceDocument() {
  // Due to the nature of the document, when new contents is being added to the source. It would be disabled from access, which this function will try and load a message or backup document for the user.
  var sheetRedirectSource = SpreadsheetApp.openById(WARP_TALLY_SHEET_SOURCE_REDIRECT_ID);
  var isSourceAvailable = sheetRedirectSource.getRange("B6").getValue();
  var sheetSource = null;
  if (isSourceAvailable == 'YES') {
    var sheetSourceId = sheetRedirectSource.getRange("B8").getValue();
    sheetSource = SpreadsheetApp.openById(sheetSourceId);
  } else {
    var isBackupAvailable = sheetRedirectSource.getRange("F6").getValue();
    if (isBackupAvailable == 'YES') {
      var sheetBackupId = sheetRedirectSource.getRange("F8").getValue();
      sheetSource = SpreadsheetApp.openById(sheetBackupId);
      var showBackupMessage = sheetRedirectSource.getRange("F10").getValue();
      if (showBackupMessage == 'YES') {
        displayBackup();
      }
    } else {
      displayMaintenance();
    }
  }
  return sheetSource;
}