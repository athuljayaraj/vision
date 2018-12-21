var excel = require('exceljs');
var path = require('path');
const fs = require('fs');
var workbook = new excel.Workbook();

exports.getSheetNames = function (file, id, callback) {  
  workbook.xlsx.readFile(file).then(function() {
    var filename = path.basename(file);
    var fileDetails = {
      id: id,
      filename:file,
      name:filename
    };
    var sheets = [];
    workbook.eachSheet(function(worksheet) {
      sheets.push(worksheet.name);
    });
    fileDetails.sheets = sheets;    
    if(!callback.isCollapsed)
      callback(JSON.stringify(fileDetails));
  });
};

exports.extractQuestions = function(file_name, fileConfig, file_buffer, templateRefId, callback) { 
  if (fs.existsSync(file_name)) {
    extract(file_name, fileConfig, templateRefId, callback);      
  } else {
    var wstream = fs.createWriteStream(file_name);
    wstream.write(file_buffer);
    wstream.end(function () { 
      extract(file_name, fileConfig, templateRefId, callback);
    });   
  } 
};

function extract(file_name, fileConfig, templateRefId, callback) {
  var questions = [];
  workbook.xlsx.readFile(file_name).then(function() {
    if(fileConfig.length > 0) {
      fileConfig.forEach(function(sheetConfig) {
        if(sheetConfig.skippable == 'false') {
          var worksheet = workbook.getWorksheet(sheetConfig.sheet_name);
          extractQuestionsFromSheet(worksheet, sheetConfig, questions, templateRefId);
        }
      });
    }
    else {
      workbook.eachSheet(function(worksheet) {
        extractQuestionsFromSheet(worksheet, fileConfig, questions, templateRefId);
      });
    }
    if(!callback.isCollapsed)
      callback(questions);
    });
}

function extractQuestionsFromSheet(worksheet, sheetConfig, questions, templateRefId) {
  worksheet.eachRow(function(row, rowNumber) {
    if(rowNumber > sheetConfig.header_index) {
      if(sheetConfig.sheet_name == 'No Technical')
        console.log(rowNumber, row.getCell(sheetConfig.columns.question));      
      let questionObj = row.getCell(sheetConfig.columns.question).value;
      let questionString = questionObj;
      if(questionObj.richText) {
        questionString = '';
        questionObj.richText.forEach(function(obj) {
          questionString +=obj.text; 
        });
        console.log(questionString);
      }
      var q = {
        'templateRefId': templateRefId,
        'sheetName': sheetConfig.sheet_name,
        'rowNumber': rowNumber,
        'question': questionString
      };
      questions.push(q);
    } else if (rowNumber == sheetConfig.header_index) {
      sheetConfig.columns.question = divide_and_rule(row.values, sheetConfig.columns.question);      
    }
  });
}

function divide_and_rule(input_string, input_obj) {
  return input_string.indexOf(input_obj);
}

workbook.xlsx.readFile('Anexo D RFP_Servicios Fijos AMX-Inteligencia Comercial ENGLISH_Response_13 4.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(5);
        row.getCell(1).value = 5; // A5's value set to 5
        row.commit();
        return workbook.xlsx.writeFile('new.xlsx');
    })