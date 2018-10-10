//  Name: Submit webforms with Data from Google Sheets */
/*	About:
 *      Written by: Luke 
 *		Date Started: Sept 22 2018
 *		Last Edit: Sept 23 2018
 *      Credit: Autohotkey.com & 
 *      Tutorials: CivReborn on Youtube amoung many others from the excellent Community
 *	Program Description: 
 *		This program with retrieve text from a Google Spreadsheet and paste it in to a web form.
 *      This is written with spoecific webpages in mind and is by no means universal
 *      This is only added publicly as a point of reference for others who may be doing something simalar
 *   Program Requirements:
 *      Autohotkey V2 or Latest https://autohotkey.com/download
 */
 
//  Start the Process 

/*  Instrucitons, 
 *   Open the window spy found in the context menu of your Autohotkey taskbar Icon
 *   click on the fields you want to use and make not eof the ClassNN and Window Titles 
 *   Use the gathered values in the appropriate place in the code (Copy and Paste Locations)
 *   Desired Goals
 *   To track Sales Performance via the Cost vs Sale amount For
 *   A: Used Items purchased for the purpose of sale 
 *   Make Tags for sales items from home and fgrom purchases
 *   requested layout:
 *   One Working Sheet to track Cost and Sale Amounts with fields 
 *   Cost, Price, Name, Category, Markdown, Sold, Commission, Earned Amount, Profit 
 *   One Working Sheet strictly for tag info printing with fields
 *   Name, Price, Category, Size, MArkdown, Sold, Commission, Earned Amount, Profit 
 *  
 */

// -- GLOBAL VARIBLES START -- //

var s = SpreadsheetApp();
var ss = s.getActive();

// -- GLOBAL VARIABLES END -- //

// -- CREATE USER MENU -- // 

function onOpen(e){
  SpreadsheetApp.getUi()
  .createMenu('Michelles Menu')
  .addItem('Update categpries', 'updateCategories')
  .addItem('Update sizes', 'updateSizes')
  .addItem('Move Sold Items', 'itemsComplete')
  .addItem('Remove Sold Items', 'removeItems')
  .addToUi();
};

// -- GET THE SUBMITTED PRODUCT DATA -- //   

function onFormSubmit(e){
  Logger.log(e);
  var name = e.namedValues.Name;
  Logger.log(name);
  var cost = e.namedValues.Cost;
  Logger.log(cost);
  var category = e.namedValues.Category;
  Logger.log(category);
  var price = e.namedValues.Price;
  Logger.log(price);
  var size = e.namedValues.Size;
  Logger.log(size);
  var item = e.namedValues.Item
  Logger.log(item);
  var targetSheet;
  var colNum;
  targetSheet = item
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(targetSheet);
  var rowNum = sheet.getLastRow();
  var insertRow = sheet.insertRowAfter(rowNum);
  rowNum++;
  sheet.getRange(rowNum, 1).setValue(name);
  sheet.getRange(rowNum, 2).setValue(cost);
  sheet.getRange(rowNum, 3).setValue(price);
  sheet.getRange(rowNum, 4).setValue(category);
  sheet.getRange(rowNum, 5).setValue(size);
  var check = sheet.getRange(rowNum,6,1,8);
  ss.getRange('CHECKBOX').copyTo(check,SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION)
 }
///////////////// call your form and connect to the drop-down item

function updateCategories(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var catId = ss.getRangeByName('categoryID').getValue();
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var formId = form.getId();
  var categoryId = catId;
  var categoryList = form.getItemById(categoryId).asListItem();
  var categorySheet = ss.getSheetByName("Categories");
  var categoryLength  = ss.getRangeByName("categoryLength").getValue();
  // grab the values in the first column of the sheet - use 2 to skip header row 
  var categories = categorySheet.getRange(2, 4,categoryLength,1).getValues();
  var categoryNames = [];
  // convert the array ignoring empty cells
  for(var i = 0; i < categories.length; i++){
    if(categories[i][0] != ""){
      categoryNames[i] = categories[i][0];
      // populate the drop-down with the array data
      categoryList.setChoiceValues(categoryNames);
    }
  }
  ss.toast("OK Size Question has had the "+categories.length+" options you selected added to the form");
}

///////////////////////////////////////////////////// - Update Sizes

function updateSizes(){
  var ss = SpreadsheetApp.getActive();
  var theForm = ss.getRangeByName('FormID').getValue();
  var form = FormApp.openById(theForm);
  var sizeId = ss.getRangeByName('sizesID').getValue();
  var sizesList = form.getItemById(sizeId).asListItem();
  // identify the sheet where the data resides needed to populate the drop-down
  var sizeSheet = ss.getSheetByName("Sizes");
  var Length  = ss.getRangeByName("sizeLength").getValue();
  // grab the values in the first column of the sheet - use 2 to skip header row 
  var sizes = sizeSheet.getRange(2, 4,Length,1).getValues();
  var sizeNames = [];
  // convert the array ignoring empty cells
  for(var i = 0; i < sizes.length; i++){
    if(sizes[i][0] != ""){
      sizeNames[i] = sizes[i][0];
      // populate the drop-down with the array data
      sizesList.setChoiceValues(sizeNames);
    };
  };
  ss.toast("OK Size Question has had the"+sizes.length+" options you selected added to the form");
  
  setBgc(rng, 3, 3, "Welcome", "green", "Sizes Changes")
  
  };

//////////////////////////////////////////////////////// -- Move Sold Lines
function itemsComplete(){
  
  var saleName = userPrompt("This will move all fields marked as sold to a group you name","Please enter a Sale Name at least 3 letters long");
 Logger.log("Sale Name: ", saleName)
 if(!saleName){
  return};
  
 var seasonName = userPrompt("This will move all fields marked as sold to a group you name","Please enter a Season Name at least 3 letters long");
  if(!seasonName){
  return};
var lastRow = soldSheet.getLastRow();
  var headers = ss.getRangeByName('HEADING').getFormulas();
 
  var firstRow = soldSheet.insertRowsAfter(lastRow+1)
   
  var headingRow = soldSheet.getRange(firstRow,1,1,11)
  var hRow = headingRow.getRow();
  var sName = saleName;
  var seaName =  seasonName;
  headingRow.setFormulas(headers);  
  soldSheet.getRange(hRow,1).setValue(sName);
  soldSheet.getRange(hRow,2).setValue(seaName);
  
moveItems("Purchased")
moveItems("Personal")
};

function moveItems(sheetName){
var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var soldSheet = ss.getSheetByName('Sold');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var header = data.shift();
  // ref Martin Hawksleys post https://mashe.hawksey.info/?p=17869/#comment-184945
  var object = data.map(function(row) {  
    var nextRowObject = header.reduce(function(accumulator, currentValue, currentIndex) {
      accumulator[currentValue] = row[currentIndex];      
      return accumulator;
    }, {}) // Use {} here rather than initialAccumulatorValue (see next post comment)
    return nextRowObject;
  });
  
  
  var soldData = new Array();
  
  for(var i = 0; i<object.length; i++){
    var row = object[i];
    Logger.log("Before if",row.Sold);
    if(row.Sold){
      Logger.log("After if",row.Sold);
      soldData[i].push(row.Name ,row.Cost, row.Price, row.Amount,row.Earned, row.Profit, row.Category, row.Size);
      soldSheet.appendRow(soldData);
      //   return soldData
    }  
     };
  
  var itemRange = soldSheet.getRange(firstRow+1, 3, soldData.length,soldData[0].length);
  itemRange.setValues(soldData);
  
  
  
 
};  

function onEdit(e){
if(e.range.getA1Notation() == 'B3') {
    if (/^\w+$/.test(e.value)) {        
      eval(e.value)();
      e.range.clear();
      var sheet = SpreadsheetApp.getActiveSheet();
        sheet.getRange(3,3).setValue('If statement ran');
    }
  }
}

