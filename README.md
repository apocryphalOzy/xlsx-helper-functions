# Helper functions using SheetJS
The objective of this project is to use excel without having to be within excel

## Authors
- Jeremy B.

### How to run
I use node to execute code 

# Working with the Workbook


### Load in xlsx module
```
const XLSX = require('xlsx');
```

### Object to hold custom variables
```
const excel = {};
excel.xlsx = XLSX;
```

### Create new workbook property that returns workbook object
```
excel.createNewWorkbook = function() {
  return this.xlsx.utils.book_new();
};
```

### Takes an array of arrays of JS values and returns a worksheet resembling the input data
```
excel.createNewWorksheet = function(worksheetData) {
  return this.xlsx.utils.aoa_to_sheet(worksheetData);
};
```

### Create and add new sheet to workbook
```
excel.addNewWorksheet = function(workbook, worksheetData, worksheetName) {
  const worksheet = this.createNewWorksheet(worksheetData);
  return this.xlsx.utils.book_append_sheet(workbook, worksheet, worksheetName);
};
```

### Validate file path, read from workbook and return workbook object
```
excel.readWorkbook = function(workbookPath) {
  const xlsx = this.xlsx;
  console.log('Validating file path...');
  if(/^[a-z]:((\\|\/)[a-z0-9\s_@\-^!#$%&+={}\[\]]+)+\.xlsx$/i.test(workbookPath)){
    //TODO should validate path
    console.log('Waiting...');
    console.log('Path extension validated!');
    return xlsx.readFile(workbookPath);
  }else{
    return console.log('That file path is not a validated xlsx extension.');
  };

};
```

### Write to workbook
```
excel.writeWorkbook = function(workbook, filePath) {
  return this.xlsx.writeFile(workbook, filePath);
};
```

### Get list of worksheet names
```
excel.getWorksheetNames = function(workbook) {
  return workbook.SheetNames;
};
```

### Make available to other files
```
module.exports = excel;
```

### Determines sheet cell range starting at 'A1' notation
```
excel.cellRange = function(worksheetObj){
  return worksheetObj['!ref']
}
```

### Sum value of two cell objects
```
excel.sumCell = function(worksheetObj, cell1, cell2){
  if (cell1 || cell2 === undefined){
    console.log('That cell does not exist')
  }else{
    let cellObj1 = worksheetObj[cell1] 
    let cellObj2 = worksheetObj[cell2]
    let cellValue1 = cellObj1.v
    let cellValue2 = cellObj2.v
    let total = cellValue1 + cellValue2
    return total
  }
}
```

### Output values of specified cell range
```
excel.specificCellRange = function(range1, range2, worksheetObj, controlOutputFormat) {
  worksheetObj['!ref'] = range1 + "+" + range2;
  return this.xlsx.utils.sheet_to_json(worksheetObj, controlOutputFormat)
  }
```

### Produces HTML output of excel table and saves in file location
```
excel.sheetToHTML = function(fileName, worksheetObj) {
  function createTestFilePath(fileName, extensionType) {
    const testFilesDirectory = __dirname + '\\testfiles\\';
    const pathStartStub = testFilesDirectory + fileName;
    const extensionPart = extensionType && extensionType.indexOf(".") !== -1 ? extensionType : "." + extensionType;
    return pathStartStub + extensionPart;
  }
  const fileCreatedPath = createTestFilePath(fileName, 'html')
  fs.writeFile(fileCreatedPath, this.xlsx.utils.sheet_to_html(worksheetObj), (err) => {
    if (err) throw err;
    console.log('File saved');
  });
};
```

### Creates new worksheet on workbook
```
excel.createWorksheet = function(workbookName, aoaData, worksheetName) {
    const testWorkbookPath = createTestFilePath(workbookName,'xlsx');
    const workbook = excel.readWorkbook(testWorkbookPath);
    excel.addNewWorksheet(workbook, aoaData, worksheetName);
    excel.writeWorkbook(workbook, testWorkbookPath);
    console.log('WORKSHEET CREATED :D');
  

};

```

### Return raw value of cell object
```
excel.rawValue = function(worksheetObj, cellObject){
  return worksheetObj[cellObject].v;
};
```