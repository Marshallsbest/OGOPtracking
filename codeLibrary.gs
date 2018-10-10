function userPrompt(title,question,error) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var salePrompt = ui.prompt(title,
                             question, 
                             ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  var button = salePrompt.getSelectedButton();
  var saleText = salePrompt.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    return saleText;  
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get the name. Data will not be moved');
    return null;
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog, please restart process')
    return null;
  }
}

function setBgc(rng, row, col, sht, clr, msg){
  var ss = SpreadsheetApp.getActive();
  var confirm = ss.getSheetByName(sht).getRange(row,col) || ss.getRangeByName(rng)
  var oldClr = confirm.getBackground();
  var oldValue = confirm.getValue();
  confirm.setBackground(clr); 
  confirm.setValue(msg);
  Utilities.sleep(2000);
  confirm.setBackground(oldClr)
  confirm.setValue(oldValue);
}

function log(text,val){
Logger.log(text,val);
console.log(text,val);
return
};
