/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [{
        name: 'Fill Net Values',
        functionName: 'fillNetValues'
    },
    {
        name: 'Goto last cell',
        functionName: 'gotoLastCell'
    },
    
    {
        name: 'testMethod',
        functionName: 'testMethod'
    }
    ];
    spreadsheet.addMenu('Custom functions', menuItems);
}
function testMethod()      
{ 
// Some how it is not copying values from B43 to other cells.. It is only showing loading...
// Need to check this...

// Don't know how it started working automatically again..
    copyValues('B43', 'C43')
}
function getColumns(sheetName)
{
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var columns = new Array(26);
    // Iterate from A1 to Z1
    for(i=1;i<=26;i++)
    {
         columns[i-1] = settingsSheet.getRange(1, i).getValue();
    }
    return columns;
}
function getRows(sheetName)
{
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var rows = new Array(60);
    // Iterate from A1 to Z1
    for(i=1;i<=60;i++)
    {
         rows[i-1] = settingsSheet.getRange(i, 1).getValue();
    }
    return rows;
}
function gotoLastCell()
{
    var dateColumn = findColumn("Date");
    // This last row variable is used for writing on to cell.
    var lastRow = findEmptyRow(dateColumn + '1:' + dateColumn + '1000');
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName('Consolidated FDs');
    
    settingsSheet.getRange(dateColumn + lastRow).activate();
}



var columns = getColumns('Consolidated FDs');
var rows = getRows('Consolidated FDs');
function fillNetValues() {
    var EmiColumn = findColumn("EMI");
    var TaxSaving = findColumn("Tax Saving");
    var TotalWithdrawableBalanceColumn = findColumn("Total Withdrawable Balance");
    var MutualFundValueColumn = findColumn("Mutual Fund Value");

    var TotalBalanceColumn = findColumn("Total Balance");
    var totalRetirementFundRow = findRowNumberInConsolidatedFD('Total Retirement Fund');
    var totalRetirementFundColumn = findColumn("Total Retirement Fund");
    var retirementContriRequiredColumn = findColumn("Retirement contribution Required");
    var dateColumn = findColumn("Date");

    // This last row variable is used for writing on to cell.
    var lastRow = findEmptyRow(dateColumn + '1:' + dateColumn + '1000');
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName('Consolidated FDs');
    // These three variables are used for reading.
    var firstRow = findRowNumberInConsolidatedFD('Tax Saving FD');
    var mutualFundRow = findRowNumberInConsolidatedFD('Mutual Funds Total Value');
    var secondRow = firstRow + 1;
    var thirdRow = firstRow + 2;

    /*var lastFilledDate = settingsSheet.getRange(dateColumn + (lastRow - 1));
    var milliseconds = (new Date() - lastFilledDate.getValue());
    diffInDays = milliseconds / (1000 * 60 * 60 * 12);

    // If this value has been filled, then just return...
    if (diffInDays < 1) {
        //return;
    }*/
    settingsSheet.getRange(dateColumn + lastRow).setValues([
        [new Date()]
    ]);
    settingsSheet.getRange(dateColumn + lastRow + ':' + TaxSaving + lastRow).setBorder(true, true, true, true, true, true);
    // Change the second parameter, if you insert new columns.
    var background = settingsSheet.getRange(lastRow - 2, 12, 1, 4).getBackground();
    var numberFormat = settingsSheet.getRange(lastRow - 2, 12, 4).getNumberFormat();
    settingsSheet.getRange(TotalWithdrawableBalanceColumn + lastRow + ':' + TaxSaving + lastRow).setBackground(background);
    settingsSheet.getRange(TotalWithdrawableBalanceColumn + lastRow + ':' + TaxSaving + lastRow).setNumberFormat(numberFormat);



    copyValues('B' + firstRow, TaxSaving + lastRow); // Tax Saving FD
    copyValues('B' + secondRow, TotalWithdrawableBalanceColumn + lastRow); // Non Tax Saving FD
    copyValues('B' + thirdRow, TotalBalanceColumn + lastRow); // Total
    copyValues('B' + mutualFundRow, MutualFundValueColumn+lastRow);// Total Mutual Fund Value

    numberFormat = settingsSheet.getRange(totalRetirementFundColumn + (lastRow - 1)).getNumberFormat();
    settingsSheet.getRange(totalRetirementFundColumn + lastRow).setNumberFormat(numberFormat);
    copyValues('B' + totalRetirementFundRow, totalRetirementFundColumn + lastRow); // Total Retirement Fund

    // Do not use copyValues function here, because it reads only from Consolidated FDs sheet.
    var amortizationSchedule = spreadsheet.getSheetByName('Amortization Schedule');
    row = amortizationSchedule.getRange('B6');

    row.copyTo(settingsSheet.getRange(EmiColumn + lastRow), {
        contentsOnly: true
    }); // EMI

    var retirementPlan = spreadsheet.getSheetByName('Retirement Plan');

    rowNum = findRowNumber("A1:A60", "Monthly contribution required", 'Retirement Plan');
    row = retirementPlan.getRange('G' + rowNum);//Retirement Contribution + Child Edu Contribution
    numberFormat = settingsSheet.getRange(retirementContriRequiredColumn + (lastRow - 1)).getNumberFormat();
    settingsSheet.getRange(retirementContriRequiredColumn + lastRow).setNumberFormat(numberFormat);

    row.copyTo(settingsSheet.getRange(retirementContriRequiredColumn + lastRow), {
        contentsOnly: true
    }); //Retirement contribution required
   
    SpreadsheetApp.flush();
}



function copyValues(srcCell, targetCell) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName('Consolidated FDs');
    var row = settingsSheet.getRange(srcCell);
    var targetRange = settingsSheet.getRange(targetCell);
    // Set value  is no longer working, now we can use copyTo function.
//    settingsSheet.getRange(targetCell).setValue(row.getValues());
    row.copyTo(targetRange, {
        contentsOnly: true
    });
    SpreadsheetApp.flush();
}

function findEmptyRow(range) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName('Consolidated FDs');
    var results = settingsSheet.getRange(range).getValues();
    var lastRow = 1;
    var foundEmptyRow = false;
    for (var i = 1; i < results.length; i++) {
            if (results[i][0].toString().length == 0) {
                foundEmptyRow = true;
                break;
            }
        if (foundEmptyRow == true)
            break;
    }
    lastRow = i + 1;
    return lastRow;
}

function findEmptyRowFromSheet(range, sheetName) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var results = settingsSheet.getRange(range).getValues();
    var lastRow = 1;
    var foundEmptyRow = false;
    for (var i = 0; i < results.length; i++) {
            if (results[i][0].toString().length == 0) {
                foundEmptyRow = true;
                break;
            }
            if (foundEmptyRow == true)
              break;
    }
    lastRow = i + 1;
    return lastRow;
}
function findRowNumber(range, text, sheetName) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var results = settingsSheet.getRange(range).getValues();
    for (var i = 0; i < results.length; i++) {
        if (results[i][0].toString() == text) {
            return i + 1;
        }
    }
}
/*
 * This function will find the text in first column of the Consolidated FDs sheet
 */
function findRowNumberInConsolidatedFD(text) {
    for (var i = 0; i < rows.length; i++) {
        if (rows[i].toString() == text) {
            return i + 1;
        }
    }
}
/*function findColumn(range, text, sheetName) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var results = settingsSheet.getRange(range).getValues();
    for (var i = 0; i < 26; i++) {
        if (results[0][i].toString() == text) {
            return columnToLetter(i + 1);
        }
    }
}*/

function findColumn(text) {
    for (var i = 0; i < 26; i++) {
        if (columns[i].toString() == text) {
            return columnToLetter(i + 1);
        }
    }
}

function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

/*

function EvaluateWhereToPutExtraInvestment(sheetName) {
    var spreadsheet = SpreadsheetApp.getActive();
    var settingsSheet = spreadsheet.getSheetByName(sheetName);
    var amortization = spreadsheet.getSheetByName("Amortization Schedule");
    //amortization.getRange("J5").setFormula("=\"Yes\"");
    amortization.getRange("J5").setValue("No");
    SpreadsheetApp.flush();
    // Case 1: Put additional Investment in Retirement

    var range = findRowNumber("A1:A50", "Any additional investment", "Retirement Plan");
    var cell = settingsSheet.getRange("C" + range);
    var interestRowNum = findRowNumber("A1:A50", "Annual Interest assumed", "Retirement Plan");
    var yearsToRetireRowNum = findRowNumber("A1:A50", "Years to retire", "Retirement Plan");
    var monthlyContributionRequiredRowNum = findRowNumber("D1:D50", "Monthly contribution required", "Retirement Plan");

    cell.setFormula("='Amortization Schedule'!$G$14");
    SpreadsheetApp.flush();
    var range2 = findRowNumber("D1:D50", "Any additional investment", "Retirement Plan");
    var cell2 = settingsSheet.getRange("F" + range2);
    cell2.setFormula("=0");
    SpreadsheetApp.flush();
    settingsSheet.getRange("B24").setFormula("=B" + monthlyContributionRequiredRowNum + "*B" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("B24").copyTo(settingsSheet.getRange("B24"), {
        contentsOnly: true
    });
    SpreadsheetApp.flush();
    settingsSheet.getRange("E24").setFormula("=E" + monthlyContributionRequiredRowNum + "*E" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("E24").copyTo(settingsSheet.getRange("E24"), {
        contentsOnly: true
    });
    amortization.getRange("C13").copyTo(settingsSheet.getRange("G24"), {
        contentsOnly: true
    });
    SpreadsheetApp.flush();
    // Case 2: Push your extra investment to Child marriage.
    cell.setFormula("=0");
    SpreadsheetApp.flush();
    var interestRowNum2 = findRowNumber("A1:A50", "Annual Interest assumed", "Retirement Plan");
    var yearsToRetireRowNum2 = findRowNumber("A1:A50", "Years to retire", "Retirement Plan");
    cell2.setFormula("='Amortization Schedule'!$G$14");
    settingsSheet.getRange("B29").setFormula("=B" + monthlyContributionRequiredRowNum + "*B" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("B29").copyTo(settingsSheet.getRange("B29"), {
        contentsOnly: true
    });
    settingsSheet.getRange("E29").setFormula("=E" + monthlyContributionRequiredRowNum + "*E" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("E29").copyTo(settingsSheet.getRange("E29"), {
        contentsOnly: true
    });
    amortization.getRange("C13").copyTo(settingsSheet.getRange("G29"), {
        contentsOnly: true
    });

    // Case 3: If I make the downpayment for home loan
    cell.setFormula("=0");
    cell2.setFormula("=0");
    SpreadsheetApp.flush();
    amortization.getRange("J5").setValue("No");
    settingsSheet.getRange("B34").setFormula("=B" + monthlyContributionRequiredRowNum + "*B" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("B34").copyTo(settingsSheet.getRange("B34"), {
        contentsOnly: true
    });
    settingsSheet.getRange("E34").setFormula("=E" + monthlyContributionRequiredRowNum + "*E" + yearsToRetireRowNum + "*12");
    SpreadsheetApp.flush();
    settingsSheet.getRange("E34").copyTo(settingsSheet.getRange("E34"), {
        contentsOnly: true
    });
    amortization.getRange("C13").copyTo(settingsSheet.getRange("G34"), {
        contentsOnly: true
    });

    // After everything is done, restore the actual values
    cell.setFormula("='Amortization Schedule'!$G$14");
    cell2.setFormula("=0");
    
    amortization.getRange("J5").setValue("No");
    SpreadsheetApp.flush();
}

*/
