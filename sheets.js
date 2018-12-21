var Excel = require('exceljs');
var path = require('path')
var workbook = new Excel.Workbook();

exports.getSheetNames = function (file, res) {
  var counter = 0;
    workbook.xlsx.readFile(file).then(function() {
      var filename = path.basename(file);
      counter++;
      var fileDetails = {
        'filename': file,
        id: '5666823336886272', // bad-code alert
        name: filename
      };
      var sheets = [];
      workbook.eachSheet(function(worksheet) {
        sheets.push(worksheet.name);
      });
      fileDetails.sheets = sheets;    
      res.write(JSON.stringify(fileDetails));
      res.end();
    });  
}

exports.extractQuestions = function(fileConfig, res) {
  var questions = []
  workbook.xlsx.readFile(fileConfig.file).then(function() {
    if(fileConfig.sheets.length > 0) {
      fileConfig.sheets.forEach(function(sheetConfig) {
        if(!sheetConfig.skippable){
          worksheet = workbook.getWorksheet(sheetConfig.name);
          extractQuestionsFromSheet(worksheet, sheetConfig, questions);
        }
      });
    }
    else {
      workbook.eachSheet(function(worksheet) {
        extractQuestionsFromSheet(worksheet, fileConfig.sheets, questions);
      });
    }
    res.write(JSON.stringify(questions));
    res.end();
  });
}

function extractQuestionsFromSheet(worksheet, sheetConfig, questions) {
  worksheet.eachRow(function(row, rowNumber) {
    if(rowNumber > sheetConfig.header_index) {
      questions.push(row.getCell(sheetConfig.question).value);            
    } else if (rowNumber == sheetConfig.header_index) {
      sheetConfig.question = divide_and_rule(row.values, sheetConfig.question);
    }
  });
}

function divide_and_rule(input_string, input_obj) {
  return input_string.indexOf(input_obj);
}
