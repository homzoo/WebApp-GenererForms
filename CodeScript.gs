function Generer() {
  //Connexion au spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var range = sheet.getDataRange();
  
  var data = range.getValues();
  var numberRows = range.getNumRows();
  var numberColumns = range.getNumColumns();
  var firstRow = 1;
  var form = FormApp.openById('1hXD7p162fQJauCu9uVtzKYuK4wHDdXtT2vdEaw_3-TE'); //Connexion au form via sa clé
   
  var items = form.getItems();
  while(items.length > 0){
    form.deleteItem(items.pop());
  }
  
  
 for(var i=0;i<numberRows;i++){
  var questionType = data[i][0]; 
  if (questionType==''){
     continue;
  }
  else if(questionType=='TEXT'){
   form.addTextItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2])
     .setRequired(true);   
  } 
  else if(questionType=='PARAGRAPH'){
   form.addParagraphTextItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2])
     .setRequired(true);
  }
  else if(questionType=='CHOICE'){
  var rowLength = data[i].length;
  var currentRow = firstRow+i;
  var currentRangeValues = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentRow,1,1,rowLength).getValues();
  var getSheetRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
  var numberOfColumnsSheet = getSheetRange.getNumColumns();
  var numberOfOptionsInCurrentRow = numberOfColumnsSheet;
  var lastColumnInRange = String.fromCharCode(64 + (numberOfOptionsInCurrentRow));
  var range_string = 'E' + currentRow + ":" + lastColumnInRange + currentRow;
  var optionsArray = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range_string).getValues();
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
        choicesForQuestion.push(optionsArray[0][j]);
        }
  form.addMultipleChoiceItem()
    .setTitle(data[i][1]) 
    .setHelpText(data[i][2])
    .setChoiceValues(choicesForQuestion)
    .setRequired(true); 
  }
  else if(questionType=='CHECKBOX'){
  var rowLength = data[i].length;
  var currentRow = firstRow+i;
  var currentRangeValues = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentRow,1,1,rowLength).getValues();
  var getSheetRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
  var numberOfColumnsSheet = getSheetRange.getNumColumns();
  var numberOfOptionsInCurrentRow = numberOfColumnsSheet;
  var lastColumnInRange = String.fromCharCode(64 + (numberOfOptionsInCurrentRow));
  var range_string = 'E' + currentRow + ":" + lastColumnInRange + currentRow;
  var optionsArray = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range_string).getValues();
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
        choicesForQuestion.push(optionsArray[0][j]);
        }
  form.addCheckboxItem()
    .setTitle(data[i][1]) 
    .setHelpText(data[i][2])
    .setChoiceValues(choicesForQuestion)
    .setRequired(true);
  }
  else if(questionType=='LIST'){
  var rowLength = data[i].length;
  var currentRow = firstRow+i;
  var currentRangeValues = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentRow,1,1,rowLength).getValues();
  var getSheetRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
  var numberOfColumnsSheet = getSheetRange.getNumColumns();
  var numberOfOptionsInCurrentRow = numberOfColumnsSheet;
  var lastColumnInRange = String.fromCharCode(64 + (numberOfOptionsInCurrentRow));
  var range_string = 'E' + currentRow + ":" + lastColumnInRange + currentRow;
  var optionsArray = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range_string).getValues();
  var choicesForQuestion =[];
    for (var j=0;j<optionsArray[0].length;j++){
        choicesForQuestion.push(optionsArray[0][j]);
        }
  form.addListItem()
    .setTitle(data[i][1]) 
    .setHelpText(data[i][2])
    .setChoiceValues(choicesForQuestion)
    .setRequired(true);
  }
  else if(questionType=='GRID'){
  var rowLength = data[i].length;
  var currentRow = firstRow+i;
  var currentRangeValues = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentRow,1,1,rowLength).getValues();
  var getSheetRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
  var numberOfColumnsSheet = getSheetRange.getNumColumns();
  var numberOfOptionsInCurrentRow = numberOfColumnsSheet;
  var lastColumnInRange = String.fromCharCode(64 + (numberOfOptionsInCurrentRow));
  var range_string = 'E' + currentRow + ":" + lastColumnInRange + currentRow;
  var optionsArray = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range_string).getValues();
  var rowTitles =[];
    for (var j=0;j<optionsArray[0].length;j++){
        rowTitles.push(optionsArray[0][j]);
        }
  var rowLength = data[i+1].length;
  var currentRow = firstRow+i+1;
  var currentRangeValues = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(currentRow,1,1,rowLength).getValues();
  var getSheetRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange();
  var numberOfColumnsSheet = getSheetRange.getNumColumns();
  var numberOfOptionsInCurrentRow = numberOfColumnsSheet;
  var lastColumnInRange = String.fromCharCode(64 + (numberOfOptionsInCurrentRow));
  var range_string = 'E' + currentRow + ":" + lastColumnInRange + currentRow;
  var optionsArray = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(range_string).getValues();
  var columnTitles =[];
    for (var j=0;j<optionsArray[0].length;j++){
        columnTitles.push(optionsArray[0][j]);
        }  
  form.addGridItem()
    .setTitle(data[i][1]) 
    .setHelpText(data[i][2])
    .setRows(rowTitles)
    .setColumns(columnTitles)
    .setRequired(false);
  } 
   else if(questionType=='IMAGE'){    
   var A1 = data[i][3];
   var A1String = A1.toString();  
     
   var img = UrlFetchApp.fetch(A1String); 
   form.addImageItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2])
     .setImage(img);
  }
  else if(questionType=='VIDEO'){ 
   var A1 = data[i][3];
   var A1String = A1.toString();  
    
   form.addVideoItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2])
    .setVideoUrl(A1String);
  }
  else if(questionType=='PAGE'){
   form.addPageBreakItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2]);   
  } 
  else if(questionType=='SECTION'){
   form.addSectionHeaderItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2]);   
  }
  else if(questionType=='TIME'){
   form.addTimeItem()
     .setTitle(data[i][1]) 
     .setHelpText(data[i][2]);   
  }
  else{
    continue;
  }
 } 
}

