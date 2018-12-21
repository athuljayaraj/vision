var http = require('http'),
    inspect = require('util').inspect;
var fs = require('fs');
var exceljs = require('exceljs');
var Busboy = require('busboy');
var sheetUtil = require('./sheets.js');

http.createServer(function(req, res) {
  if (req.method === 'POST') {
    var busboy = new Busboy({ headers: req.headers });
    busboy.on('file', function(fieldname, file, filename, encoding, mimetype) {
      console.log('File [' + fieldname + ']: filename: ' + filename);
      var path = '/tmp/' + filename;
	    var wstream = fs.createWriteStream(path);
	    var data_full= [];

      file.on('data', function(data) {
		data_full.push(data)
		console.log('File [' + fieldname + '] got ' + data.length + ' bytes');
      });

      file.on('end', function() {
        console.log('File [' + fieldname + '] Finished');
		databuffer = Buffer.concat(data_full);
		wstream.write(databuffer);
		wstream.end(function () {
      sheetUtil.getSheetNames([path], res);
      console.log('done');
		});
      });
    });
    busboy.on('field', function(fieldname, val, fieldnameTruncated, valTruncated) {
      console.log('Field [' + fieldname + ']: value: ' + inspect(val));
    });
    busboy.on('finish', function() {

    });
    req.pipe(busboy);
  } else if (req.method === 'GET') {
    res.writeHead(200, { Connection: 'close' });
    res.end('<html><head></head><body>\
               <form method="POST" enctype="multipart/form-data">\
                <input type="text" name="textfield"><br />\
                <select name="selectfield">\
                  <option value="1">1</option>\
                  <option value="10">10</option>\
                  <option value="100">100</option>\
                  <option value="9001">9001</option>\
                </select><br />\
                <input type="checkbox" name="checkfield">Node.js rules!<br />\
                <input type="submit">\
<input type="file" name="fileupload">\
              </form>\
            </body></html>');
  }
}).listen(8000, function() {
  console.log('Listening for requests');
});
