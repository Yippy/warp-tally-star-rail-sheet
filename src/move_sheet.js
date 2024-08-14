/*
 * Version 1.02 made by yippym - 2024-08-14 19:00
 * https://github.com/Yippy/warp-tally-star-rail-sheet
 */
function moveToSettingsSheet() {
  moveToSheetByName(WARP_TALLY_SETTINGS_SHEET_NAME);
}

function moveToDashboardSheet() {
  moveToSheetByName(WARP_TALLY_DASHBOARD_SHEET_NAME);
}

function moveToCharacterEventWarpHistorySheet() {
  moveToSheetByName(WARP_TALLY_CHARACTER_EVENT_WARP_SHEET_NAME);
}

function moveToStellarWarpHistorySheet() {
  moveToSheetByName(WARP_TALLY_STELLAR_WARP_SHEET_NAME);
}

function moveToLightConeEventWarpHistorySheet() {
  moveToSheetByName(WARP_TALLY_LIGHT_CONE_EVENT_WARP_SHEET_NAME);
}

function moveToDepartureWarpHistorySheet() {
  moveToSheetByName(WARP_TALLY_DEPARTURE_WARP_SHEET_NAME);
}

function moveToChangelogSheet() {
  moveToSheetByName(WARP_TALLY_CHANGELOG_SHEET_NAME);
}

function moveToPityCheckerSheet() {
  moveToSheetByName(WARP_TALLY_PITY_CHECKER_SHEET_NAME);
}

function moveToEventsSheet() {
  moveToSheetByName(WARP_TALLY_EVENTS_SHEET_NAME);
}

function moveToCharactersSheet() {
  moveToSheetByName(WARP_TALLY_CHARACTERS_SHEET_NAME);
}

function moveToLightConesSheet() {
  moveToSheetByName(WARP_TALLY_LIGHT_CONES_SHEET_NAME);
}

function moveToResultsSheet() {
  moveToSheetByName(WARP_TALLY_RESULTS_SHEET_NAME);
}

function moveToReadmeSheet() {
  moveToSheetByName(WARP_TALLY_README_SHEET_NAME);
}

function moveToShardCalculatorSheet() {
  moveToSheetByName(WARP_TALLY_SHARD_CALCULATOR_SHEET_NAME);
}

function moveToSheetByName(nameOfSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    var settingsForOptionalSheet = SETTINGS_FOR_OPTIONAL_SHEET[nameOfSheet];
    if (settingsForOptionalSheet) {
      var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WARP_TALLY_SETTINGS_SHEET_NAME);
      if (settingsSheet) {
        var settingOption = settingsForOptionalSheet["setting_option"];
        if (!settingsSheet.getRange(settingOption).getValue()) {
          displayUserAlert("Optional Sheet", nameOfSheet+" has been disabled within Settings, enable this sheet at cell '"+settingOption+"', and run 'Update Items'",  SpreadsheetApp.getUi().ButtonSet.OK)
        }
      }
    }
    title = "Error";
    message = "Unable to find sheet named '"+nameOfSheet+"'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}