/*
 * Version 1.00 made by yippym - 2023-05-02 19:00
 * https://github.com/Yippy/warp-tally-star-rail-sheet
 */
function onEdit(e) {
    const sheet = e.range.getSheet(); 
    if(sheet.getName() == WARP_TALLY_CHARACTERS_SHEET_NAME || sheet.getName() == WARP_TALLY_LIGHT_CONES_SHEET_NAME) {
        if (e.value == "TRUE") {
            var allowableColumns = sheet.getRange(1,12).getValue();
            allowableColumns = String(allowableColumns).split(",");
            if (allowableColumns.includes(String(e.range.columnStart))) {
                sheet.getRange(e.range.rowStart, e.range.columnStart).setValue(false);
                var characterName = sheet.getRange(e.range.rowStart, e.range.columnStart+1).getValue();
                var indexRow = sheet.getRange(1,e.range.columnStart).getValue();
                sheet.getRange(indexRow, e.range.columnStart).setValue(characterName);
            }
        }
    }
}