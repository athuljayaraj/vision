var fs = require('fs');
var http = require('http');
var formidable = require('formidable');
const excelutils = require('./excelutilsTest.js');

var testconfig = {
  file: '/home/athul/Downloads/rfp/Anexo D RFP_Servicios Fijos AMX-Inteligencia Comercial ENGLISH_Response_13 4.xlsx',
  sheets: [
    {
      name: 'Functional',
      skippable: false,
      header_index: 9,
      question: 'Requirement'
    },  {
      name: 'Technical',
      skippable: false,
      header_index: 9,
      question: 'Requirement'
    },  {
      name: 'No Technical',
      skippable: false,
      header_index: 9,
      question: 'Requirement'
    }
  ]
}
doOnComplete = function(questionObject) {
  // let result = questionObject.map(a => a.rowNumber);
  // result.forEach(element => {
  //   console.log(element);  
  // });
  console.log(questionObject.length);
  
  let stringToWrite = questionObject.map(a => a.question + '\n');
  fileList = ['file1'];
  for (var thisFile in fileList) {
    fs.writeFile('/tmp/' + fileList[thisFile], stringToWrite, function(err) {
      if(err) {
        return console.log(err);
      }    
      console.log("The file was saved!");
    }); 
  }
}


http.createServer(function (req, res) {
  var form = new formidable.IncomingForm();
  res.writeHead(200, {'Content-Type': 'application/json', "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "*" });
  if (req.url == '/fileupload') {    
    form.parse(req, function (err, fields, files) {
      var oldpath = files.filetoupload.path;

      var newpath = '/home/athul/uploads/' + files.filetoupload.name;
      // Read the file
       fs.readFile(oldpath, function (err, data) {
           if (err) throw err;

           // Write the file
           fs.writeFile(newpath, data, function (err) {
               if (err) throw err;
             });

           // Delete the file
           fs.unlink(oldpath, function (err) {
               if (err) throw err;
           });
       });
       console.log('File successfully uploaded');
       res.write('File successfully uploaded');
       res.end();
 });
  } else if (req.url == '/submit') {
    let sheets = JSON.parse('[{ "sheet_name": "Sheet1", "skippable": "true", "header_index": "9", "columns": {  "question": "Requirement",  "feature_status": "Compliance (Total,  Partial, Does not meet)",  "remark": "DETAILED response",  "doc_ref": "Link to documentation that support the response (required)" }}, { "sheet_name": "Functional", "skippable": "false", "header_index": "9", "columns": {  "question": "Requirement",  "feature_status": "Compliance (Total,  Partial, Does not meet)",  "remark": "DETAILED response",  "doc_ref": "Link to documentation that support the response (required)" }}, { "sheet_name": "Technical", "skippable": "false", "header_index": "9", "columns": {  "question": "Requirement",  "feature_status": "Compliance (Total,  Partial, Does not meet)",  "remark": "DETAILED response",  "doc_ref": "Link to documentation that support the response (required)" }}, { "sheet_name": "No Technical", "skippable": "false", "header_index": "9", "columns": {  "question": "Requirement",  "feature_status": "Compliance (Total,  Partial, Does not meet)",  "remark": "DETAILED response",  "doc_ref": "Link to documentation that support the response (required)" }}]');
    // let sheets =  JSON.parse('[{ "sheet_name": "Sheet1", "skippable": "false", "header_index": "9", "columns": {  "question": "Requirement",  "feature_status": "Compliance (Total,  Partial, Does not meet)",  "remark": "DETAILED response",  "doc_ref": "Link to documentation that support the response (required)" }}]');
    let fileName = '/home/athul/Downloads/rfp/Anexo D RFP_Servicios Fijos AMX-Inteligencia Comercial ENGLISH_Response_13 4.xlsx'
    
    fs.readFile(fileName, function(err, data) {    console.log("reached readfile");
    
    excelutils.extractQuestions('/tmp/fileName2.xlsx', sheets, data, '22345567778889', null, doOnComplete);
    res.end();
  });

  }  else {
    res.writeHead(200, {'Content-Type': 'text/html'});
    res.write('<form action="fileupload" method="post" enctype="multipart/form-data">');
    res.write('<input type="file" name="filetoupload"><br>');
    res.write('<input type="submit">');
    res.write('</form>');
    return res.end();
  }
}).listen(8080);
