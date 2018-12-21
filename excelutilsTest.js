var Excel = require('exceljs');
var path = require('path')
var workbook = new Excel.Workbook();
const fs = require('fs');
const async = require('async');




exports.extractQuestions =  function(file_name,fileConfig, file_buffer, templateRefId, aidatastore, maincallback) {
   
        tasks = [function isExist(callback){
          callback(null,file_name,fileConfig, file_buffer, templateRefId,aidatastore);

                  // aidatastore.is_rfp_exist(templateRefId).then(results=>{
                  //   entities = results[0];
                  //   console.log("entities",entities);
                  //     if (entities && entities.length>0){
                  //       console.log("Already parsed the RFP with Id ",templateRefId);
                  //       maincallback();
                       
                  //   }else{
                  //     callback(null,file_name,fileConfig, templateRefId,aidatastore);
                  //   }

                  // });

                },

         function execute (file_name,fileConfig, file_buffer, templateRefId,aidatastore,callback){
               console.log("Start Process",templateRefId)
              if (fs.existsSync(file_name)) {
                  console.log("file exist");
                  callback(null,file_name,fileConfig, templateRefId,aidatastore)
                  //extract(file_name,fileConfig, templateRefId,aidatastore);      
              } else {
                var wstream = fs.createWriteStream(file_name);
                    
                    wstream.write(file_buffer);
                    wstream.end(function () { 
                      // console.log("write to file ", file_name);
                      callback(null,file_name,fileConfig, templateRefId,aidatastore)
                      // extract(callback,file_name,fileConfig, templateRefId,aidatastore);
                    });

                // aidatastore.getRFPRequest(templateRefId).then(results =>{
                //   const task = results[0];
                //   // console.log(task);
                //   if (task){
                //     file_buffer = task.inputFile;
                      
                //   }else{
                //     // console.log("template ref not found");
                //     callback(null);
                //   }

                // });    

            }
        },
        function extract(file_name, fileConfig, templateRefId,aidatastore,callback) {
          var questions = []
          // console.log("file name",file_name);
          workbook.xlsx.readFile(file_name).then(function() {
            if(fileConfig.length > 0) {
              fileConfig.forEach(function(sheetConfig) {
                 // console.log("sheet iteration", sheetConfig.sheet_name);
                if(sheetConfig.skippable == 'false') {
                  var worksheet = workbook.getWorksheet(sheetConfig.sheet_name);
                  //callback(null,worksheet, fileConfig, questions, templateRefId);
                  extractQuestionsFromSheet(worksheet, sheetConfig, questions, templateRefId);
                }
              });
            }
            else {
              workbook.eachSheet(function(worksheet) {
                //callback(null,worksheet, fileConfig, questions, templateRefId);
                extractQuestionsFromSheet(worksheet, fileConfig, questions, templateRefId);
              });
            }
            
            fs.unlink(file_name,function(err){
                if(err) console.log(err);     
                console.log('file deleted successfully',file_name);
                callback(null,questions,aidatastore,templateRefId);              
           });  
            
          
          });
        },
        function save(questions,aidatastore,templateRefId,callback){
          console.log("Total questions",questions.length);
          // aidatastore.saveRecord(questions,templateRefId);
        }
        
        ];

        async.waterfall(tasks, (err, results) => {
         console.log("completed ");
            if (err) {
              console.log(err);
              return next(err);
            }
            maincallback();
      });

  
    };


function extractQuestionsFromSheet(worksheet, sheetConfig, questions, templateRefId) {
  worksheet.eachRow(function(row, rowNumber) {
    if(rowNumber > sheetConfig.header_index) {
      let questionObj = row.getCell(sheetConfig.columns.question).value;
      let questionString = questionObj;
      if(questionObj.richText) {
        questionString = '';
        questionObj.richText.forEach(function(obj) {
          questionString += obj.text; 
        });
      }      
      var q = {
        templateRefId: templateRefId,
        question: questionString,
        sheetName: sheetConfig.sheet_name,
        rowNumber: rowNumber ,  
        requestSource:'RFP'   
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
