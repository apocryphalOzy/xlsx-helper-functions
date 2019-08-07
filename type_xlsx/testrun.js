//Get helpers scripts
const excel = require("./excel.js");
const fs = require("fs");
let horspool = require("./boyer-moore");

//Test file path
function createTestFilePath(fileName, extensionType) {
  const testFilesDirectory = __dirname + "\\testfiles\\";
  const pathStartStub = testFilesDirectory + fileName;
  const extensionPart =
    extensionType && extensionType.indexOf(".") !== -1
      ? extensionType
      : "." + extensionType;
  return pathStartStub + extensionPart;
}

//Create test workbook file name
const testWorkbookPath = createTestFilePath("testName", "xlsx");

//Create test workbook
//const workbook = excel.createNewWorkbook();
//excel.addNewWorksheet(workbook, [['test_column']], 'test_name');
//excel.writeWorkbook(workbook, testWorkbookPath);

//get workbook
const workbook = excel.readWorkbook(testWorkbookPath);
const worksheet = workbook.Sheets["test_name"];
//console.log(excel)
//console.log(worksheet)
//console.log(workbook)

//console.log(excel.sumCell(worksheet, 'G1', 'F2'))

//console.log(excel.cellRangeObjects("A1", "A4", worksheet, {header:1}))

//console.log(excel.xlsx.utils.sheet_to_html(worksheet))

//excel.sheetToHTML('writtenFile', worksheet)

// let arrayOfArraysData = [['test_column'], ['This cell is below A1 cell']]
// excel.createWorksheet('testName', arrayOfArraysData, 'AnotherSheetName')

//console.log(excel.rawValue(worksheet, 'A1'))

//excel.table(worksheet);
const arrayOfObj = Object.entries(worksheet).map(e => ({ [e[0]]: e[1] }));
let haystack = new Buffer(arrayOfObj);
let needle = new Buffer("john");

let index = horspool(haystack, needle);

console.log(index);
