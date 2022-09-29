
// Given a cell reference and sheet, return a string containing all rich text from that cell
// The cell can have one or more rich text values
// The "cell" parameter can be for example of the format "G2"
function _getRichTextValue(cell, sheet) {
  
    console.log("Cell to get rich text = " + cell);
    
    // Each "run" is similar to a word in the cell
    var runs = sheet.getRange(cell).getRichTextValue().getRuns();
    
    var s = "";
    var i = 0;
  
    runs.forEach(function(item) {
      
      // Print the links each on a new line and avoid null values
      if(item.getLinkUrl() !== null ) {
        
        // Format like a website link
        s += '<a href="' + item.getLinkUrl() + '">' + item.getText() + '</a>'
  
        if(i < runs.length-1) {
          s += "<br>";
        }
      }
  
      i++;
  
    })
    return s;
  }
  
  
  // Copies new entries in the source sheet that are not in the target sheet based on the Response Number (column A in both sheets)
  // I've kept logging statements through this function to help troubleshoot
  function copyNewEntries() {
  
    // Setup our spreadsheet references
    const sh = SpreadsheetApp.getActiveSpreadsheet();
    const source = sh.getSheetByName("Responses");
    const target = sh.getSheetByName("FormattedResponses");
  
    // Get the last row in the target spreadsheet so we know where to start in the source sheet. If no rows then go with 1.
    var targetLastRow = target.getLastRow();
    
    console.log("Last row in target is: " + targetLastRow);
  
    var lastTargetEntryID = 0;
  
    // // Get the last entry ID in the target sheet (either a positive number or 0 if no entries found)
    if(targetLastRow > 0) {
      // Get the Entry ID from the last row, first column in the sheet
      lastTargetEntryID = target.getRange(targetLastRow, 1).getValue();
      // If the latest row containes anything other than a number, use 0
      lastTargetEntryID = (!isNaN(lastTargetEntryID) ? lastTargetEntryID : 0);  
    } 
    
    console.log("lastTargetEntryID = " + lastTargetEntryID);
  
    // Set the ID in source sheet that is one more than last entry ID from target sheet
    var firstSourceEntryIdNotInTarget = lastTargetEntryID + 1;
    console.log("firstSourceEntryIdNotInTarget = " + firstSourceEntryIdNotInTarget);
  
    // Find the row in the source sheet, first column, using the latest target ID + 1
    var match = source.getRange(1,1,source.getLastRow()).createTextFinder(firstSourceEntryIdNotInTarget).findNext();
  
    // If a matching row exists
    if(match != null) {
      firstSourceEntryRowNotInTarget = match.getRow();
      console.log("firstSourceEntryRowNotInTarget is = " + firstSourceEntryRowNotInTarget);
  
      // // Copy entries from the source sheet (that are AFTER the last entry ID in target sheet and up to the last row in the source sheet)
      var data = source.getRange(firstSourceEntryRowNotInTarget, 1, source.getLastRow()-(firstSourceEntryRowNotInTarget-1), 8).getValues();
  
      console.log(data);
      var i = firstSourceEntryRowNotInTarget;
  
      data.forEach(function(row) {
        var richText = _getRichTextValue("G"+i, source);
        console.log("richText: " + richText);
        row[row.length] = richText;
        i++;
      })
    
      console.log("Data copied from source sheet: " + data);
      
      // Write data to the target sheet
      target.getRange(target.getLastRow()+1,1,data.length,data[0].length).setValues(data);
  
    } else {
      console.log("Source sheet does not have entry ID = " + firstSourceEntryIdNotInTarget);
    }
    
    console.log("Wrote " + ((data != null && data.length > 0) ? data.length: 0) + " entries to target sheet");
  }