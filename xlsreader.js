var path = require('path')
var fs = require('fs')
var Excel = require('exceljs');
var readWorkbook = new Excel.Workbook();
var writeWorkbook = new Excel.Workbook();
var writeSheet = writeWorkbook.addWorksheet('Sheet 1');
writeSheet.columns = [
  { header: 'Question', key: 'Question', width: 10},
  { header: 'Feature Status', key: 'Feature_Status'},
  { header: 'Remark', key: 'Remark', width: 10},
  { header: 'Doc Ref', key: 'Doc_Ref', width: 10},
  { header: 'RFP Name', key: 'RFP_Name', width: 10},
  { header: 'Product Version', key: 'Product_Version', width: 20}
];

function writeToXlsx() {
  var filepath = '/home/athul/Downloads/rfp/new_wb.xlsx'
  writeWorkbook.xlsx.writeFile(filepath)
  .then(function() {
      console.log('Done');
  });
}

function extractFromFile(fileConfigs) {
  var fileCounter = 0;
  fileConfigs.forEach(function(thisFile, i) {
    readWorkbook.xlsx.readFile(thisFile.file).then(function() {
      fileCounter++;
      var filename = path.basename(thisFile.file);
      if(thisFile.sheets.length > 0) {
        thisFile.sheets.forEach(function(thisSheet, j) {
          worksheet = readWorkbook.getWorksheet(thisSheet.id);
          worksheet.eachRow(function(row, rowNumber) {
            if(rowNumber > thisSheet.header) {
              var newRow = {}
              for(var key in thisSheet.columns) {
                newRow[key] = row.getCell(thisSheet.columns[key]).value;
              }
              newRow['RFP_Name'] = filename;
              writeSheet.addRow(newRow);
            } else if (rowNumber == thisSheet.header) {
              var headerRow = row.values;
              thisSheet.columns = divide_and_rule(headerRow, thisSheet.columns);
            }
          });
        });
      } else {
        thisSheet = thisFile.sheets;
        var columnsSet = false;
        readWorkbook.eachSheet(function(worksheet, sheetId) {
          worksheet.eachRow(function(row, rowNumber) {
            if(rowNumber > thisSheet.header) {
              var newRow = {}
              for(var key in thisSheet.columns) {
                newRow[key] = row.getCell(thisSheet.columns[key]).value;
              }
              newRow['RFP_Name'] = filename;
              writeSheet.addRow(newRow);
            } else if (rowNumber == thisSheet.header && !columnsSet) {
              var headerRow = row.values;
              thisSheet.columns = divide_and_rule(headerRow, thisSheet.columns);
              columnsSet = true;
            }
          });
        });
      }
      if(fileCounter == fileConfigs.length) {
        writeToXlsx()
      }
    });
  });
}

function divide_and_rule(input_string, input_obj) {
  let output_obj = {}
  for(var key in input_obj) {
      output_obj[key] = input_string.indexOf(input_obj[key]);
  }
  return output_obj;
}

extractFromFile(configs)
