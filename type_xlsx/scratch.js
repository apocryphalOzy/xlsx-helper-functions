//Get helpers scripts
const excel = require('./excel.js');
const fs = require('fs');

//Test file path
function createTestFilePath(fileName, extensionType) {
  const testFilesDirectory = __dirname + '\\testfiles\\';
  const pathStartStub = testFilesDirectory + fileName;
  const extensionPart = extensionType && extensionType.indexOf(".") !== -1 ? extensionType : "." + extensionType;
  return pathStartStub + extensionPart;
}

//Create test workbook file name
const testWorkbookPath = createTestFilePath('testName', 'xlsx');

//Create test workbook
// const workbook = excel.createNewWorkbook();
// excel.addNewWorksheet(workbook, [['test_column']], 'test_name');
// excel.writeWorkbook(workbook, testWorkbookPath);

//get workbook
const workbook = excel.readWorkbook(testWorkbookPath);

//Get test worksheet
const worksheet = workbook.Sheets['test_name'];

const getColumnAsString = new Promise(function(resolve, reject) {
  let data = '';
  try {
    for(let row in worksheet) {
      let rowValue = worksheet[row].v;
      let matchArray = rowValue ? rowValue.match(/\.(.*\..*)/i) : undefined;
      let parsedValue = matchArray && matchArray[1] ? matchArray[1] : rowValue;
      data = data + parsedValue + "\r\n";
    }
    resolve(data);
  } catch(error) {
    reject(error);
  }
});

getColumnAsString
  .catch(function(error) {
    throw error;
  })
  .then(function(data) {
  const textFilePath = createTestFilePath('data', 'txt');
  fs.writeFileSync(textFilePath, data);
  console.log("DONE");
});


//console.log(worksheet);


/*----------------- HELPER FUNCTIONS -----------------*/

//Counts how many objects exist within worksheet
function objectLength(obj) {
  var result = 0;
  for (var prop in obj) {
    if (obj.hasOwnProperty(prop)) {
      result++;
    }
  }
  return result;
}
//objectLength(worksheet);


//Determines sheet cell range starting at 'A1' notation 

let sheetColumn = function(worksheetObj){
  return worksheetObj['!ref']
}
//console.log(sheetRange(worksheet))


//Object representing page margins
/* Set worksheet sheet to "normal" */
//ws["!margins"]={left:0.7, right:0.7, top:0.75,bottom:0.75,header:0.3,footer:0.3}
/* Set worksheet sheet to "wide" */
//ws["!margins"]={left:1.0, right:1.0, top:1.0, bottom:1.0, header:0.5,footer:0.5}
/* Set worksheet sheet to "narrow" */
//ws["!margins"]={left:0.25,right:0.25,top:0.75,bottom:0.75,header:0.3,footer:0.3}
let pageMargins = function(worksheetObj){
  return worksheetObj['!margins']
}
//console.log(pageMargins(worksheet))


//Returns stated cell object from worksheet
let cellObj = function(worksheetObj, cellNumber){
  return worksheetObj[cellNumber]
}
//console.log(cellObj(worksheet, 'A2'))


// //Filter by domain name extension, '.com', '.net', '.org'
// let filterDomainExtensions = function(worksheetObj) {
//   let filteredDomain = worksheetObj.filter(function(elem){ 
//     return elem.v.toLowerCase() 
// })
// }


//console.log(workbook['!protect'])













