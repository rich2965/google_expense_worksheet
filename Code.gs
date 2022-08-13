/**
 * Code for expensive splitter
 */


//Triggered by the Generate button, creates all the columns with the participant names in the header
function generateExpenseReportPrompt() {
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Is this for a restaurant bill?', ui.ButtonSet.YES_NO_CANCEL);

// Process the user's response.
if (response == ui.Button.YES) {
  showTaxPrompt();
  showTipPrompt();
  generateRestaurantReport();
  } else if (response == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Action canceled');
  } else if (response == ui.Button.NO) {
    // User clicked X in the title bar.
    generateReport();
  }
  expenseCalculatorGenerator();
}

//=========================================================================================================================================================================================

function showTaxPrompt() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
  var ui = SpreadsheetApp.getUi(); 

  var result = ui.prompt(
    'Please enter tax %',
    'Standard 11.5% Applied if blank:',
      ui.ButtonSet.OK);

  // Process the user's response.
  var button = result.getSelectedButton();
  // Checking if NULL
  if (result.getResponseText() == '') {
    var text = '11.5';
    }
  else {
    var text = result.getResponseText()
    };
    
  if (button == ui.Button.OK) {
    // User clicked "OK".
    
   var finalText = text.replace('%','') 
   sheet.getRange(2,6).setValue(Number(finalText)*.01)
  }
  else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
  }
}
//=========================================================================================================================================================================================

function showTipPrompt() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
  var ui = SpreadsheetApp.getUi(); 

  var result = ui.prompt(
    'Please enter tip %',
    'Standard 18% Applied if blank:',
      ui.ButtonSet.OK);

  // Process the user's response.
  var button = result.getSelectedButton();
  // Checking if NULL
  if (result.getResponseText() == '') {
    var text = '18';
    }
  else {
    var text = result.getResponseText()
    };
  
  if (button == ui.Button.OK) {
    // User clicked "OK".
    
   var finalText = text.replace('%','') 
   sheet.getRange(2,7).setValue(Number(finalText)*.01)
  }
  else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
  }
}


//=========================================================================================================================================================================================



function generateRestaurantReport() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Browser.msgBox('Error', 'You need to enter at least 2 participants.',Browser.Buttons.OK);
    return
  }
  var total_participants = sheet.getRange(2,1,lastRow-1).getValues();
  var num = total_participants.filter(String).length;

   //Browser.msgBox('Made it here' + num + total_participants);
    
  if (num > 2) {
    expensesheet.insertColumnsAfter(5,num-2);
  }
  
    
  //Population of the participant names that were entered in the Setup tab 
  var participants = [ total_participants];
  
  expensesheet.getRange(2,4,1,num).setValues(participants).setFontWeight('bold');
  var lastRow = expensesheet.getLastRow();
  var lastColumn = expensesheet.getLastColumn();
  //This for loop creates the sum formulas for the bottom row 
  for (i = 4; i < lastColumn-1; i++) {
    var expenseRows =  expensesheet.getRange(3,i,lastRow-3).getA1Notation();
    var expenseColumns =  expensesheet.getRange(3,lastColumn,lastRow-3).getA1Notation();
    expensesheet.getRange(lastRow,i).setFormula("=sumifs(" + expenseColumns + "," + expenseRows +",TRUE)");
  }
  /*expensesheet.autoResizeColumn(4);*/
  
  expensesheet.getRange(3,1,lastRow-3).setValue(sheet.getRange(2,1).getValue())
  
  
  //Switches view to the Expenses Overview Tab
    expensesheet.activate();
}

//=========================================================================================================================================================================================


function generateReport() {
   
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  //This sets the Tax and Tip to 0%
  sheet.getRange(2,7).setValue(0)
  sheet.getRange(2,6).setValue(0)
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Browser.msgBox('Error', 'You need to enter at least 2 participants.',Browser.Buttons.OK);
    return
  }
  var total_participants = sheet.getRange(2,1,lastRow-1).getValues();
  var num = total_participants.filter(String).length;

   //Browser.msgBox('Made it here' + num + total_participants);
    
  if (num > 2) {
    expensesheet.insertColumnsAfter(5,num-2);
  }
  
    
  //Population of the participant names that were entered in the Setup tab 
  var participants = [ total_participants];
  
  expensesheet.getRange(2,4,1,num).setValues(participants).setFontWeight('bold');
  var lastRow = expensesheet.getLastRow();
  var lastColumn = expensesheet.getLastColumn();
  //This for loop creates the sum formulas for the bottom row 
  for (i = 4; i < lastColumn-1; i++) {
    var expenseRows =  expensesheet.getRange(3,i,lastRow-3).getA1Notation();
    var expenseColumns =  expensesheet.getRange(3,lastColumn,lastRow-3).getA1Notation();
    expensesheet.getRange(lastRow,i).setFormula("=sumifs(" + expenseColumns + "," + expenseRows +",TRUE)");
  }
    //Switches view to the Expenses Overview Tab
    expensesheet.activate();
  //expensesheet.autoResizeColumns(4,lastColumn-5);
}



//=========================================================================================================================================================================================

//Looks at the last row of the Expenses and adds one more. 
function addExpenseRow() {
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  var lastRow = expensesheet.getLastRow();
  var lastColumn = expensesheet.getLastColumn();
  expensesheet.insertRowsAfter(lastRow-1,1);
  //These copy over the values from the last 2 columns down to the new row
  expensesheet.getRange(lastRow-2,lastColumn-1).copyTo(expensesheet.getRange(lastRow,lastColumn-1));
  expensesheet.getRange(lastRow-2,lastColumn).copyTo(expensesheet.getRange(lastRow,lastColumn));
  //This for loop creates the sum formulas for the bottom row 
  for (i = 4; i < lastColumn-1; i++) {
    var expenseRows =  expensesheet.getRange(3,i,lastRow-2).getA1Notation();
    var expenseColumns =  expensesheet.getRange(3,lastColumn,lastRow-2).getA1Notation();
    expensesheet.getRange(lastRow+1,i).setFormula("=sumifs(" + expenseColumns + "," + expenseRows +",TRUE)");
  };
  var totalRows =  expensesheet.getRange(3,lastColumn-1,lastRow-2).getA1Notation();
  expensesheet.getRange(lastRow+1,lastColumn-1).setFormula("=sum("+totalRows+")");
  expensesheet.getRange(lastRow-1,1).copyTo(expensesheet.getRange(lastRow,1));
}


//=========================================================================================================================================================================================


function showAddPersonPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Trying to enter another payer?',
      'Please enter the name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    
    
    addPerson(text);
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Action canceled');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
  }
}


//=========================================================================================================================================================================================

function addPerson (payer) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  var lastColumn = expensesheet.getLastColumn();
  var lastRow = expensesheet.getLastRow();
  expensesheet.insertColumnsAfter(lastColumn-2,1);
  expensesheet.getRange(2,lastColumn-1).setValue( payer) ;
  var expenseRows =  expensesheet.getRange(3,lastColumn-1,lastRow-3).getA1Notation();
  var expenseColumns =  expensesheet.getRange(3,lastColumn+1,lastRow-3).getA1Notation();
  expensesheet.getRange(lastRow,lastColumn-1).setFormula("=sumifs(" + expenseColumns + "," + expenseRows +",TRUE)");
  var payerlastRow = sheet.getLastRow();
  sheet.getRange(payerlastRow+1,1).setValue( payer) 
}


//=========================================================================================================================================================================================

function restartSheet () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup')
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  var calculatesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalculateSheet');
  var lastColumn = expensesheet.getLastColumn();
  var lastRow = expensesheet.getLastRow();
  expensesheet.getRange(3,1,lastRow-3,3 ).setValue('');
  expensesheet.getRange(3,4,lastRow-3,lastColumn-5 ).setValue("FALSE");
  if (lastColumn > 7)
  {
    var columnsToRemove = (lastColumn - 7);
    expensesheet.deleteColumns(6,columnsToRemove);
  }
  if (lastRow > 12)
  {
    var rowsToRemove = (lastRow - 12);
    expensesheet.deleteRows(12,rowsToRemove);
  }
  var sheetlastRow = sheet.getLastRow();
  var total_participants = sheet.getRange(2,1,sheetlastRow-1).setValue('');
  calculatesheet.clear();
  sheet.activate();
}


//=========================================================================================================================================================================================

function showRestartPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to start over?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    restartSheet();
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Action canceled');
  }
}

//=========================================================================================================================================================================================

function clearAll(){
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  var lastColumn = expensesheet.getLastColumn();
  var lastRow = expensesheet.getLastRow();
  expensesheet.getRange(3,1,lastRow-3,3 ).setValue('');
  expensesheet.getRange(3,4,lastRow-3,lastColumn-5 ).setValue("FALSE");
}

//=========================================================================================================================================================================================

function clearChecks(){
  var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
  var lastColumn = expensesheet.getLastColumn();
  var lastRow = expensesheet.getLastRow();
  expensesheet.getRange(3,4,lastRow-3,lastColumn-5 ).setValue("FALSE");
}


//=========================================================================================================================================================================================

function showParticipantPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Please enter Payer name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    
    expenseCalculatorGenerator();
    expenseCalculator(text);
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('Action canceled');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
  }
}
//=========================================================================================================================================================================================

function expenseCalculatorGenerator (){
 var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setup');
 var calculatesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalculateSheet');
 var expensesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses Overview');
 
 var lastRow = sheet.getLastRow();
 var expensesheetlastRow= expensesheet.getLastRow();
  var expensesheetlastColumn = expensesheet.getLastColumn();

 var total_participants = sheet.getRange(2,1,lastRow-1).getValues();
 var num = total_participants.filter(String).length;

   //Browser.msgBox('Made it here' + num + total_participants);
    
  //if (num > 2) {
  //  calculatesheet.insertColumnsAfter(5,num-2);
  //}
  
    
  //Population of the participant names that were entered in the Setup tab 
  var participants = [ total_participants];
  calculatesheet.getRange(2,1,num).setValues(total_participants);
  calculatesheet.getRange(1,2,1,num).setValues(participants).setFontWeight('bold');
  

  var firstColumn = expensesheet.getRange(1,1,expensesheetlastRow).getA1Notation();
  var lastColumn = expensesheet.getRange(1,expensesheetlastColumn,expensesheetlastRow).getA1Notation();
  
//Iterating through rows and columns to add in the sum formulas  
  for (row = 2; row <num+2; row++){
    for (column = 2; column < num+2; column++) {
      var columnParticipant = expensesheet.getRange(1,2+column,expensesheetlastRow).getA1Notation();
      calculatesheet.getRange(row,column).setFormula("=sumifs('Expenses Overview'!" + lastColumn + ",'Expenses Overview'!"+ firstColumn + " ,$A$"+row +",'Expenses Overview'!"+ columnParticipant +",TRUE)");
    }
  }

}

//=========================================================================================================================================================================================
function expenseCalculator(name){
  var calculatesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalculateSheet');
  var data = calculatesheet.getDataRange().getValues();
  var display = []
  //var Participant = sheet.getRange("C2").getValue();
  for(var i = 0; i<data.length;i++){
    if(data[i][0] == name){ //[1] because column B
       //Browser.msgBox('Found your name in row ' + i);
      for (var x = 1; x<data.length;x++) {
        var money_diff = data[i][x] - data[x][i] 
        if (money_diff > 0) {
          display.push(data[x][0] + ' owes you $' + money_diff.toFixed(2) + '\\n' )
        } else if (money_diff < 0) {
           display.push('You owe ' +data[x][0] + ' $' +  Math.abs(money_diff.toFixed(2)) +'\\n' )
        } else {continue}
       
    }
    //  for (var y = 1; y<data.length;y++) {
      //  Browser.msgBox('Found money you owe ' + data[y][i] );
     }
  }

}
//=========================================================================================================================================================================================


function expenseCalculator_Total(name){
    var calculatesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalculateSheet');
    var expensesTotalsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ExpensesTotal');
    expensesTotalsheet.clear();
    var data = calculatesheet.getDataRange().getValues();
    var display = []
    //var Participant = sheet.getRange("C2").getValue();
    for(var x = 1; x<data.length+1;x++){
        if (data[0][x]) {
        expensesTotalsheet.getRange(1,x).setValue(data[0][x]).setBackground("lightblue").setFontWeight("bold")
        }
        for (var y = 1; y<data.length;y++) {
          var money_diff = data[y][x-1] - data[x-1][y] 
          if (money_diff > 0) {
            expensesTotalsheet.getRange(x,y).setValue(data[x-1][0] + ' owes you $' + money_diff.toFixed(2)  ).setBackground("green").setFontColor('white')
          } else if (money_diff < 0) {
             expensesTotalsheet.getRange(x,y).setValue('You owe ' +data[x-1][0] + ' $' +  Math.abs(money_diff.toFixed(2))).setBackground("red").setFontColor('white')
          } 
          else {continue}
          //else {expensesTotalsheet.getRange(x,y).setValue('0')}
         
        }
      //expensesTotalsheet.getRange(1,2,1,data.length).setValues(display)



    }

    for (var x = 1; x<data.length+1;x++){
    expensesTotalsheet.getRange(2,x,data.length-1,1).sort(x);
    expensesTotalsheet.autoResizeColumn(x);
    }

    expensesTotalsheet.activate()
    //Browser.msgBox( display.join(""))


}

//=========================================================================================================================================================================================

