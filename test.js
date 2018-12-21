const excelutils = require('./excelutils.js');
var fs = require('fs');

doOnComplete = function(questionObject) {
  fs.writeFile("/tmp/test", questionObject, function(err) {
    if(err) {
      return console.log(err);
    }    
    console.log("The file was saved!");
  }); 
}

sheets = JSON.parse('[{"sheet_name":"Sheet1","skippable":"true","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"Functional","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"Technical","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"No Technical","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"Sheet1","skippable":"true","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"Functional","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"Technical","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}},{"sheet_name":"No Technical","skippable":"false","header_index":"9","columns":{"question":"Requirement","feature_status":"Compliance (Total,  Partial, Does not meet)","remark":"DETAILED response","doc_ref":"Link to documentation that support the response (required)"}}]');
excelutils.extractQuestions('Anexo D RFP_Servicios Fijos AMX-Inteligencia Comercial ENGLISH_Response_13 4.xlsx',sheets,'','22345567778889', doOnComplete);
