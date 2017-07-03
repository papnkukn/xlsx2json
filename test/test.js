var fs = require('fs');
var path = require('path');
var unzip = require('unzip');
var xlsx2json = require('../lib/index.js');

exports["Read entries"] = function(test) {
  var entries = [ ];
  var file = path.resolve(__dirname, 'fixture.xlsx');
  fs.createReadStream(file)
  .pipe(unzip.Parse())
  .on('entry', function (entry) {
    entries.push(entry.path);
  })
  .on('error', function (error) {
    test.ok(false, error.message);
  })
  .on('close', function () {
    test.ok(entries.length > 0, "Missing entries!")
    test.done();
  });
};
  
exports["Convert XLSX file"] = function(test) {
  var file = path.resolve(__dirname, 'fixture.xlsx');
  xlsx2json(file, { verbose: true }, function(error, result) {
    if (error) {
      test.ok(false, error.message);
    }
    else {
      test.ok(result.worksheets.length == 3, "Expected 3 worksheets");
      test.ok(result.worksheets[0].name == "Blank", "First worksheet should be named 'Blank'");
      test.ok(result.worksheets[1].name == "Calculate", "First worksheet should be named 'Calculate'");
      test.ok(result.worksheets[2].name == "Variables", "First worksheet should be named 'Variables'");
      test.ok(result.worksheets[0].data.length == 0, "Expected no data for worksheet 'Blank'");
      test.ok(result.worksheets[1].data.length == 4, "Expected 4 cells for worksheet 'Calculate'");
      console.log(result.worksheets[1].data);
      console.log(result.worksheets[2].data);
    }
    test.done();
  });
};
