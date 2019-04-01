function traceDependents() {
  var dependents = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var currentCell = ss.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var listOfNamedRanges = [];
  var allDependentRefs = [];
  
//  var range = ss.getDataRange();

  
  var regex = new RegExp("\\b" + currentCellRef + "\\b");
  var output = "Dependent of cell " + currentCellRef + ":\n";
//  var formulas = range.getFormulas();
  for (var s = 0; s < sheets.length; s++) {
    var dependentRefs = [];
    var dependents = [];
    var sheet = sheets[s];
    var range = sheet.getDataRange();
    var formulas = range.getFormulas();
    for(var i =0; i < formulas.length; i++) {
      var row = formulas[i];
      for (var j = 0; j < row.length; j++) {
        var cellFormula = row[j].replace(/\$/g,"");
        if (regex.test(cellFormula)) {
          dependents.push([i,j]);
        }
      }
    }
    for (var k = 0; k < dependents.length; k++) {
      var rowNum = dependents[k][0] + 1;
      var colNum = dependents[k][1] + 1;
      var cell = range.getCell(rowNum, colNum);
      var cellRef = cell.getA1Notation();
      dependentRefs.push(cell.getSheet().getName() + ': ' + cellRef + '\n');
    }
    allDependentRefs = allDependentRefs.concat(dependentRefs);
  }
  
  if (allDependentRefs.length > 0) {
    output += allDependentRefs.join(" \n");
  } else {
    output += " None";
  }
  currentCell.setNote(output);  
}



function traceDependentsOfNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var namedRanges = ss.getNamedRanges();
  var currentCell = ss.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var listOfNamedRanges = [];
  var allDependentRefs = [];
  for (var q = 0; q < namedRanges.length; q++) {
    var namedRange = namedRanges[q];
    listOfNamedRanges.push([namedRange.getName(),namedRange.getRange().getSheet().getSheetName() + ': ' +  namedRange.getRange().getA1Notation()]);
  }
  
  var namedRangeQuestion = SpreadsheetApp.getUi().prompt('User input required', 'Enter named range from the list below:\n' + listOfNamedRanges, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  var currentNamedRange = namedRangeQuestion.getResponseText();
  var regex = new RegExp("\\b" + currentNamedRange + "\\b");
  var output = "Dependents of " + currentNamedRange + ":\n";

  
  for (var s = 0; s < sheets.length; s++) {
    var dependentRefs = [];
    var dependents = [];
    var sheet = sheets[s];
    var range = sheet.getDataRange();
    Logger.log('data range is = ' + range.getSheet().getName() + ': ' + range.getA1Notation());
    var formulas = range.getFormulas();
    for (var i = 0; i < formulas.length; i++) {
    var row = formulas[i];
      for (var j = 0; j < row.length; j++) {
        var cellFormula = row[j].replace(/\$/g,"");
        if (regex.test(cellFormula)) {
          dependents.push([i,j]);
        }
      }
    }
    
    for (var k = 0; k < dependents.length; k++) {
      var rowNum = dependents[k][0] + 1;
      var colNum = dependents[k][1] + 1;
      var cell = range.getCell(rowNum, colNum);
      var cellRef = cell.getA1Notation();
      dependentRefs.push(cell.getSheet().getName() + ': ' + cellRef + '\n');
    }
    allDependentRefs = allDependentRefs.concat(dependentRefs);
  }
  
  if (allDependentRefs.length > 0) {
    output += allDependentRefs.join(" \n");
  } else {
    output += " None";
  }
  
  currentCell.setNote(output);  
}



function checkDataValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
}
