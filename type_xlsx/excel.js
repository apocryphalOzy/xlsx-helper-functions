//https://www.npmjs.com/package/xlsx
//Docs: https://github.com/SheetJS/js-xlsx/blob/master/README.md
const XLSX = require("xlsx");
const fs = require("fs");
const http = require("http");
const url = require("url");
//const XLSXPOPULATE = require('xlsx-populate');

//Object to hold custom variables
//TODO this should be a class that extends XLSX
const excel = {};
excel.xlsx = XLSX;

//Create a new workbook
//Returns workbook object.
//Note: Does not have any sheets
excel.createNewWorkbook = function() {
  return this.xlsx.utils.book_new();
};

//Create new worksheet with data
//Note: Creates worksheet object, but does not add to a workbook
/* 'data' arg structure':

      let worksheet_data = [
        [ "S", "h", "e", "e", "t", "J", "S" ],
        [  1 ,  2 ,  3 ,  4 ,  5 ]
      ];

*/
excel.createNewWorksheet = function(worksheetData) {
  return this.xlsx.utils.aoa_to_sheet(worksheetData);
};

//Create and add new sheet to workbook
excel.addNewWorksheet = function(workbook, worksheetData, worksheetName) {
  const worksheet = this.createNewWorksheet(worksheetData);
  return this.xlsx.utils.book_append_sheet(workbook, worksheet, worksheetName);
};

//Read from workbook
//Returns workbook object
excel.readWorkbook = function(workbookPath) {
  const xlsx = this.xlsx;
  console.log("Validating file path...");
  let pathValidation = /^[a-z]:((\\|\/)[a-z0-9\s_@\-^!#$%&+={}\[\]]+)+\.xlsx$/i;
  if (pathValidation.test(workbookPath)) {
    //validates path
    console.log("Waiting...");
    console.log("Path extension validated!");
    return xlsx.readFile(workbookPath);
  } else {
    return console.log("That file path is not a validated xlsx extension.");
  }
};

//Write to workbook
excel.writeWorkbook = function(workbook, filePath) {
  return this.xlsx.writeFile(workbook, filePath);
};

//Get list of worksheet names
excel.getWorksheetNames = function(workbook) {
  return workbook.SheetNames;
};

//Make available to other files
module.exports = excel;

/* ------------------- HELPER FUNCTIONS ------------------- */
//Test file path
let createTestFilePath = function(fileName, extensionType) {
  const testFilesDirectory = __dirname + "\\testfiles\\";
  const pathStartStub = testFilesDirectory + fileName;
  const extensionPart =
    extensionType && extensionType.indexOf(".") !== -1
      ? extensionType
      : "." + extensionType;
  return pathStartStub + extensionPart;
};

//Determines sheet cell range starting at 'A1' notation
excel.cellRange = function(worksheetObj) {
  return worksheetObj["!ref"];
};

// Sum value of two cell objects
excel.sumCell = function(worksheetObj, cell1, cell2) {
  if (cell1 || cell2 === undefined) {
    console.log("This cell does not exist");
  } else {
    let cellObj1 = worksheetObj[cell1];
    let cellObj2 = worksheetObj[cell2];
    let cellValue1 = cellObj1.v;
    let cellValue2 = cellObj2.v;
    let total = cellValue1 + cellValue2;
    return total;
  }
};

// Output values of specified cell range
excel.cellRangeObjects = function(
  range1,
  range2,
  worksheetObj,
  controlOutputFormat
) {
  worksheetObj["!ref"] = range1 + "+" + range2;
  return this.xlsx.utils.sheet_to_json(worksheetObj, controlOutputFormat);
};

// Produces HTML output of excel table and saves in file location

//TODO: https://nodejs.org/api/console.html#console_console_table_tabulardata_properties

excel.sheetToHTML = function(fileName, worksheetObj) {
  const fileCreatedPath = createTestFilePath(fileName, "html");
  fs.writeFile(
    fileCreatedPath,
    this.xlsx.utils.sheet_to_html(worksheetObj),
    err => {
      if (err) throw err;
      console.log("File saved");
    }
  );
};

// Create new worksheet on workbook
excel.createWorksheet = function(workbookName, aoaData, worksheetName) {
  const testWorkbookPath = createTestFilePath(workbookName, "xlsx");
  const workbook = excel.readWorkbook(testWorkbookPath);
  excel.addNewWorksheet(workbook, aoaData, worksheetName);
  excel.writeWorkbook(workbook, testWorkbookPath);
  console.log("WORKSHEET CREATED :D");
};

// Return raw value of cell object
excel.rawValue = function(worksheetObj, cellObject) {
  return worksheetObj[cellObject].v;
};

//output excel data to console
excel.table = function(worksheetObj) {
  return console.table(worksheetObj);
};

/*--------TODOS--------*/
// Commit to every change on file, wait to push as bundle

//build a table of excel in node console
//Reurn a specified string
//Return a count and location of where strings are located in excel
// Create system that produces dummy data for excel worksheet

// const getColumnAsString = new Promise(function(resolve, reject) {
//   let data = '';
//   try {
//     for(let row in worksheet) {
//       let rowValue = worksheet[row].v;
//       let matchArray = rowValue ? rowValue.match(/\.(.*\..*)/i) : undefined;
//       let parsedValue = matchArray && matchArray[1] ? matchArray[1] : rowValue;
//       data = data + parsedValue + "\r\n";
//     }
//     resolve(data);
//   } catch(error) {
//     reject(error);
//   }
// });

// getColumnAsString
//   .catch(function(error) {
//     throw error;
//   })
//   .then(function(data) {
//   const textFilePath = createTestFilePath('data', 'txt');
//   fs.writeFileSync(textFilePath, data);
//   console.log("DONE");
// });
