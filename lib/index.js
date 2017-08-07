var fs = require('fs');
var path = require('path');
var unzip = require('unzip');
var xml2js = require('xml2js').parseString;
var MemoryStream = require('memory-stream');

Array.prototype.contains = function(obj) {
  var i = this.length;
  while (i--) {
    if (this[i] === obj) {
      return true;
    }
  }
};

var Excel = { };

//Gets a column letter from a number, e.g. 1 = A, 2 = B, 27 = AA, etc.
Excel.colname = function(column) {
  var columnString = "";
  var columnNumber = column;
  while (columnNumber > 0) {
    var currentLetterNumber = (columnNumber - 1) % 26;
    var currentLetter = String.fromCharCode(currentLetterNumber + 65);
    columnString = currentLetter + columnString;
    columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
  }
  return columnString;
}

//Gets a column number from a letter, e.g. A = 1, B = 2, AA = 27, etc.
Excel.colnum = function(column) {
  var retVal = 0;
  var col = column.toUpperCase();
  for (var iChar = col.length - 1; iChar >= 0; iChar--) {
    var colPiece = col[iChar];
    var colNum = colPiece.charCodeAt(0) - 64;
    retVal = retVal + colNum * Math.floor(Math.pow(26, col.length - (iChar + 1)));
  }
  return retVal;
}

Excel.name = function(pos) {
  if (arguments.length == 0) throw new Error("Formula error");
  return Excel.colname(pos.column) + pos.row;
};

Excel.pos = function(name) {
  if (arguments.length == 0) throw new Error("Formula error");
  var m = /^\$?([A-Z]+)\$?([0-9]+)$/.exec(name);
  if (!m) throw new Error("Argument");
  var column = Excel.colnum(m[1]);
  var row = parseInt(m[2]);
  return { row: row, column: column };
};

Excel.range = function(range) {
  var parts = range.split(":");
  var start = Excel.pos(parts[0]);
  var stop = Excel.pos(parts[1]);
  
  var result = [ ];
  for (var r = Math.min(start.row, stop.row); r <= Math.max(start.row, stop.row); r++) {
    for (var c = Math.min(start.column, stop.column); c <= Math.max(start.column, stop.column); c++) {
      var cell = { row: r, column: c };
      cell.name = Excel.name(cell);
      result.push(cell);
    }
  }
  return result;
};

Excel.offset = function(range) {
  var parts = range.split(":");
  var start = Excel.pos(parts[0]);
  var stop = Excel.pos(parts[1]);

  var result = {
    r1: start.row,
    c1: start.column,
    r2: stop.row,
    c2: stop.column
  };
  
  //Distance between rows and columns
  result.dr = result.r2 - result.r1;
  result.dc = result.c2 - result.c1;
  
  return result;
};

function zeropad(num, size) {
  var s = num+"";
  while (s.length < size) s = "0" + s;
  return s;
}

function getCells(formula) {
  var cells = [ ];
  var regex = /('?([\w\d\s]+)'?\!)?(\$?[A-Z]+\$?[0-9]+)/g;
  
  var match;
  while (match = regex.exec(formula)) {
    if (match.index === regex.lastIndex) {
      regex.lastIndex++;
    }
  
    var cell = match[0];
    if (cell == "LOG10") {
      continue;
    }
    cells.push(cell);
  }
  
  return cells;
}

function shiftCells(cells, deltaRows, deltaColumns) {
  var result = [ ];
  for (var i = 0; i < cells.length; i++) {
    var cell = cells[i];
    if (cell.indexOf('!') > 0) {
      continue; //Skip cross-sheet references
    }
    
    var pcell = Excel.pos(cell);
    
    var lockColumn = cell.indexOf('$') == 0; //e.g. $A1
    var lockRow = cell.indexOf('$') > 0; //e.g. A$1
    
    var shifted = Excel.name({
      row: pcell.row + (lockRow ? 0 : deltaRows),
      column: pcell.column + (lockColumn ? 0 : deltaColumns),
    });
    
    var item = {
      original: cell,
      shifted: shifted
    };
    
    result.push(item);
  }
  return result;
}

function shiftFormula(originalFormula, deltaRows, deltaColumns) {
  var formula = "" + originalFormula; //Copy string
  
  var originalCells = getCells(originalFormula);
  var shiftedCells = shiftCells(originalCells, deltaRows, deltaColumns);
  //console.log("Shifted cells", deltaRows, deltaColumns, shiftedCells);
  
  //Replace original cell names with constants
  for (var i = 0; i < shiftedCells.length; i++) {
    var regex = new RegExp('\\b' + shiftedCells[i].original + '\\b', 'g');
    formula = formula.replace(regex, "SHIFTED_CELL_" + zeropad(i, 4));
  }
  
  //Replace constants with shifted cell names
  for (var i = 0; i < shiftedCells.length; i++) {
    var regex = new RegExp("SHIFTED_CELL_" + zeropad(i, 4), 'g');
    formula = formula.replace(regex, shiftedCells[i].shifted);
  }
  
  return formula;
}

/** Converts XLSX file to JSON object */
function xlsx2json(file, options, callback) {
  //Private members
  var worksheets = { };
  var strings = [ ];
  var styles = [ ];
  var table = [ ];
  var after = [ ];
  
  //Shift arguments
  if (typeof options == "function" && typeof callback == "undefined") {
    callback = options;
    options = null;
  }
  
  if (typeof callback != "function") {
    callback = function(error, result) { }; //Dummy
  }
  
  //Default options
  options = options || { };
  
  //Converts Excel value
  function getValueType(t) {
    switch (t) {
      case "s":
      case "str":
      case "inlineStr":
        return "string";
        
      case "b":
        return "bool";
        
      case "d":
        return "date";
        
      case "e":
        return "error";
        
      case "n":
        return "number";
    }
    
    return undefined; //General
  }
  
  //Formats a value to display to the user
  function getValueToDisplay(value, style) {
    var match, format = style.formatCode;
    
    //Format number: 0
    if (format == "0") {
      return "" + Math.round(parseFloat(value));
    }
    
    //Format number: 0.00, 0.000, 0.0000, etc.
    match = /^0\.(0+)$/g.exec(format);
    if (match) {
      var digits = match[1].length;
      return "" + parseFloat(value).toFixed(digits);
    }
    
    //Exponential number: 0.00, 0.000, 0.0000, etc.
    match = /^0\.(0+)[Ee]([\+\-]0+)$/g.exec(format);
    if (match) {
      var digits = match[1].length;
      //var powerDigits = match[2].length;
      return "" + parseFloat(value).toExponential(digits).toString().toUpperCase() + "%";
    }
    
    //Format percent: 0%
    if (format == "0%") {
      return "" + Math.round(parseFloat(100 * value)) + "%";
    }
    
    //Format percent: 0.00%, 0.000%, 0.0000%, etc.
    match = /^0\.(0+)\%$/g.exec(format);
    if (match) {
      var digits = match[1].length;
      return "" + parseFloat(100 * value).toFixed(digits);
    }
  }
  
  //Reads an xml entry and responds with json object
  function readEntry(entry, done) {
    if (options.verbose) {
      console.log("Found " + entry.path);
    }
    
    var ws = new MemoryStream();
    ws.on('finish', function() {
      var xml = ws.toString();
      xml2js(xml, function(error, result) {
        try {
          if (error) {
            throw error;
          }
          
          done(result);
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
  }
  
  //Process the worksheet
  function readWorksheet(ws, done) {
    var sheetId = ws.sheetId;
    var result = ws.result;
    
    var sharedFormulas = { };
    
    //Define the worksheet record
    worksheets[sheetId] = {
      sheetId: sheetId,
      //file: "sheet" + sheetId,
      name: "Sheet " + sheetId,
      data: [ ]
    };
    
    //Only specific worksheet
    //if (options.sheet != worksheets[sheetId].name) {
    //  return done();
    //}
    
    //Find the sheet name
    for (var i = 0; i < sheets.length; i++) {
      //if (sheetId == sheets[i].$.sheetId) {
      if (sheetId == i + 1) {
        worksheets[sheetId].name = sheets[i].$.name;
        break;
      }
    }
    
    if (typeof result.worksheet.sheetData[0] != "object") {
      result.worksheet.sheetData[0] = { row: [ ] };
    }
    
    var rows = result.worksheet.sheetData[0].row;
    if (options.verbose) {
      console.log("Found " + rows.length + " rows on sheet " + sheetId + " [" + worksheets[sheetId].name + "]");
    }
    
    //For each row
    for (var i = 0; i < rows.length; i++) {
      var r = parseInt(rows[i].$.r);
      var cells = rows[i].c;
      if (!cells) continue;
      
      //For each cell in a row
      for (var j = 0; j < cells.length; j++) {
        var cell = cells[j];
        
        var name = cell.$.r;
        var pos = Excel.pos(name);
        //if (pos.row <= capacity.row && pos.column <= capacity.column) {
        
        //Resolve formula
        var formula = undefined;
        var f = cell.f;
        if (f && f.length) { //Expected array
          f = f[0];
          if (typeof f == "string") {
            formula = f;
          }
          else if (f.$) {
            var attr = f.$;
            if (attr.t == "shared") { //Shared formula
              if (attr.ref) {
                sharedFormulas[attr.si] = { formula: f._, ref: attr.ref };
              }
              if (f._) {
                formula = f._;
              }
              else if (sharedFormulas[attr.si]) {
                var o = sharedFormulas[attr.si].ref.split(':')[0];
                var t = name;
                var offset = Excel.offset(o + ':' + t);
                formula = shiftFormula(sharedFormulas[attr.si].formula, offset.dr, offset.dc);
                if (options.debug) {
                  console.log("Shifting formula", '=' + sharedFormulas[attr.si].formula, "to", '=' + formula, "by offset", "R" + offset.dr + ",", "C" + offset.dc);
                }
              }
            }
            else if (f._) {
              formula = f._;
            }
          }
          
          if (!formula && options.verbose) {
            console.log("Unknown formula [" + worksheets[sheetId].name + "." + name + "]:", cell.f);
          }
        }
        
        //Resolve constant value
        var value = undefined;
        var v = cell.v;
        if (v && v.length) {
          v = v[0];
          if (typeof v == "string") {
            value = v;
          }
          else if (v["_"]) {
            value = v["_"];
          }
          else if (options.verbose) {
            console.log("Unknown value [" + worksheets[sheetId].name + "." + name + "]:", cell.v);
          }
        }
        
        if (formula || value) {
          
          var item = { };
          var type = getValueType(cell.$.t);
          var style = cell.$.s ? styles[parseInt(cell.$.s)] : null;
          
          //Shared string
          if (cell.$.t == "s") {
            var index = parseInt(value);
            value = strings[index];
          }
          
          item.cell = cell.$.r;
          item.type = type;
          if (formula) {
            item.formula = "=" + formula;
          }
          if (value) {
            item.value = value;
          }
          if (value && style && style.formatCode) {
            item.display = getValueToDisplay(value, style);
            item.format = style.formatCode;
          }
          worksheets[sheetId].data.push(item);
        }
      }
    }
    
    done();
  }
  
  //Conversion complete, let's invoke the callback function
  function readComplete() {
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
  }
 
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
      
      //Worksheet names
      if (entry.path.indexOf("xl/workbook.xml") == 0) {
        readEntry(entry, function(result) {
          sheets = result.workbook.sheets[0].sheet;
          if (options.verbose) {
            console.log("Found " + sheets.length + " worksheets");
          }
        });
        return;
      }
      
      //Shared strings
      if (entry.path.indexOf("xl/sharedStrings.xml") == 0) {
        readEntry(entry, function(result) {
          var length = result.sst.si.length;
          if (options.verbose) {
            console.log("Found " + length + "  shared strings");
          }
          
          for (var i = 0; i < length; i++) {
            if (result.sst.si[i].t) {
              var text = result.sst.si[i].t[0];
              strings.push(text);
            }
          }
        });
        return;
      }
      
      //Styles
      if (entry.path.indexOf("xl/styles.xml") == 0) {
        readEntry(entry, function(result) {
          var formats = { };
          
          //Built-in formats
          var i = 0;
          formats[++i] = { numFmtId: 1, formatCode: "0" };
          formats[++i] = { numFmtId: 2, formatCode: "0.00" };
          formats[++i] = { numFmtId: 3, formatCode: "#,##0" };
          formats[++i] = { numFmtId: 4, formatCode: "#,##0.00" };
          formats[++i] = { numFmtId: 5, formatCode: "$#,##0_);($#,##0)" };
          formats[++i] = { numFmtId: 6, formatCode: "$#,##0_);[Red]($#,##0)" };
          formats[++i] = { numFmtId: 7, formatCode: "$#,##0.00_);($#,##0.00)" };
          formats[++i] = { numFmtId: 8, formatCode: "$#,##0.00_);[Red]($#,##0.00)" };
          formats[++i] = { numFmtId: 9, formatCode: "0%" };
          formats[++i] = { numFmtId: 10, formatCode: "0.00%" };
          formats[++i] = { numFmtId: 11, formatCode: "0.00E+00" };
          formats[++i] = { numFmtId: 12, formatCode: "# ?/?" };
          formats[++i] = { numFmtId: 13, formatCode: "# ??/??" };
          formats[++i] = { numFmtId: 14, formatCode: "m/d/yyyy" };
          formats[++i] = { numFmtId: 15, formatCode: "d-mmm-yy" };
          formats[++i] = { numFmtId: 16, formatCode: "d-mmm" };
          formats[++i] = { numFmtId: 17, formatCode: "mmm-yy" };
          formats[++i] = { numFmtId: 18, formatCode: "h:mm AM/PM" };
          formats[++i] = { numFmtId: 19, formatCode: "h:mm:ss AM/PM" };
          formats[++i] = { numFmtId: 20, formatCode: "h:mm" };
          formats[++i] = { numFmtId: 21, formatCode: "h:mm:ss" };
          formats[++i] = { numFmtId: 22, formatCode: "m/d/yyyy h:mm" };
          formats[++i] = { numFmtId: 37, formatCode: "#,##0_);(#,##0)" };
          formats[++i] = { numFmtId: 38, formatCode: "#,##0_);[Red](#,##0)" };
          formats[++i] = { numFmtId: 39, formatCode: "#,##0.00_);(#,##0.00)" };
          formats[++i] = { numFmtId: 40, formatCode: "#,##0.00_);[Red](#,##0.00)" };
          formats[++i] = { numFmtId: 45, formatCode: "mm:ss" };
          formats[++i] = { numFmtId: 46, formatCode: "[h]:mm:ss" };
          formats[++i] = { numFmtId: 47, formatCode: "mm:ss.0" };
          formats[++i] = { numFmtId: 48, formatCode: "##0.0E+0" };
          formats[++i] = { numFmtId: 49, formatCode: "@" };

          //Additional formats
          var numFmts = result.styleSheet.numFmts[0].numFmt;
          for (var i = 0; i < numFmts.length; i++) {
            formats[numFmts[i].$.numFmtId] = numFmts[i].$;
          }
          
          var xfs = result.styleSheet.cellXfs[0].xf;
          var length = xfs.length;
          if (options.verbose) {
            console.log("Found " + length + " value formats");
          }
          
          for (var i = 0; i < length; i++) {
            var style = xfs[i].$;
            var format = formats[style.numFmtId];
            if (format) {
              style.formatCode = format.formatCode;
            }
            styles.push(style);
          }
        });
        return;
      }
      
      //Worksheet
      var match = /^xl\/worksheets\/sheet(\d+)\.xml$/g.exec(entry.path);
      if (match) {
        var sheetId = parseInt(match[1]);
        readEntry(entry, function(result) {
          //Process worksheet after the shared strings, styles, etc. are done
          after.push({
            sheetId: sheetId,
            entry: entry,
            result: result
          });
        });
        return;
      }
      
      entry.autodrain();
    })
    .on('close', function() {
    
      //Read worksheets
      var i = 0;
      (function next() {
        var item = after[i++];
        if (!item) return readComplete();
        readWorksheet(item, function() {
          next();
        });
      })();
      
    });
  }
  catch (e) {
    callback(e);
  }
};

module.exports = xlsx2json;