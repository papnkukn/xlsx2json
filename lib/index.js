var fs = require('fs');
var path = require('path');
var unzip = require('unzip');
var xml2js = require('xml2js').parseString;
var MemoryStream = require('memory-stream');

/** Converts XLSX file to JSON object */
function xlsx2json(file, options, callback) {
  var worksheets = { };
  
  //Shift arguments
  if (typeof options == "function" && typeof callback == "undefined") {
    callback = options;
    options = null;
  }
  
  if (typeof callback != "function") {
    callback = function(error, result) { }; //Dummy
  }
  
  options = options || { };
 
  try {
    if (typeof file != "string" || file.length == 0) {
      throw new Error("Please check the 'file' argument!");
    }
   
    fs.createReadStream(file)
    .pipe(unzip.Parse())
    .on('error', function (error) {
      if (options.verbose) {
        console.error(error);
      }
    })
    .on('entry', function (entry) {
      if (options.verbose) {
        //console.log(entry.path);  
      }
      
      //Worksheet names
      if (entry.path.indexOf("xl/workbook.xml") == 0) {
        if (options.verbose) {
          console.log("Found wl/workbook.xml");
        }
        var ws = new MemoryStream();
        ws.on('finish', function() {
          var xml = ws.toString();
          xml2js(xml, function(error, result) {
            try {
              if (error) {
                throw error;
              }
              
              sheets = result.workbook.sheets[0].sheet;
              if (options.verbose) {
                console.log("Found " + sheets.length + " worksheets");
              }
            }
            catch (e) {
              if (options.verbose) {
                console.error(e);
              }
              var func = callback;
              callback = function(error, result) { };
              func(e);
            }
          });
          
        });      
        entry.pipe(ws);
        return;
      }
      
      //Worksheet
      var match = /^xl\/worksheets\/sheet(\d+)\.xml$/g.exec(entry.path);
      if (match) {
        var sheetId = parseInt(match[1]);
        
        if (options.verbose) {
          console.log("Found xl/worksheets/sheet" + sheetId + ".xml");
        }
        
        //Define the worksheet record
        worksheets[sheetId] = {
          sheetId: sheetId,
          //file: "sheet" + sheetId,
          name: "Sheet " + sheetId,
          data: [ ]
        };
        
        //Find the sheet name
        for (var i = 0; i < sheets.length; i++) {
          //if (sheets[i].$.sheetId == sheetId) {
          if (i + 1 == sheetId) {
            worksheets[sheetId].name = sheets[i].$.name;
            break;
          }
        }

        //Parse the worksheet file
        var ws = new MemoryStream();
        ws.on('finish', function() {
          var xml = ws.toString();
          xml2js(xml, function(error, result) {
            if (typeof result.worksheet.sheetData[0] != "object") {
              result.worksheet.sheetData[0] = { row: [ ] };
            }
            var rows = result.worksheet.sheetData[0].row;
            if (options.verbose) {
              console.log("Found " + rows.length + " rows on sheet " + sheetId + " [" + worksheets[sheetId].name + "]");
            }
            for (var i = 0; i < rows.length; i++) {
              var r = parseInt(rows[i].$.r);
              var cells = rows[i].c;
              if (!cells) continue;
              for (var j = 0; j < cells.length; j++) {
                var cell = cells[j];
                var item = { };
                var formula = (cell.f && cell.f.length && !cell.f[0].$ ? cell.f[0] : undefined);
                var value = (cell.v && cell.v.length ? cell.v[0] : undefined);
                if (formula || value) {
                  item.cell = cell.$.r;
                  if (formula) {
                    item.formula = "=" + formula;
                  }
                  if (value) {
                    item.value = value;
                  }
                  worksheets[sheetId].data.push(item);
                }
              }
            }
          });
        });      
        entry.pipe(ws);
        return;
      }
      
      entry.autodrain();
    })
    .on('close', function() { //close, finish
      if (options.verbose) {
        console.log("Finished reading document");
      }
        
      //Convert dictionary to sorted array
      var sheets = [ ];
      var keys = Object.keys(worksheets);
      for (var i = 0; i < keys.length; i++) {
        var worksheet = worksheets[keys[i]]; //Sorted as in Excel
        //var worksheet = worksheets[(i + 1).toString()]; //Sorted by sheetId
        sheets.push(worksheet);
      }
      
      if (keys.length == 0 && options.verbose) {
        console.log("No records found. May not be an XLSX file.");
      }
      
      var result = {
        worksheets: sheets
      };
      callback(null, result);
    });
  }
  catch (e) {
    callback(e);
  }
};

module.exports = xlsx2json;