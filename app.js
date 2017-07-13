var os = require('os');
var fs = require('fs');
var path = require('path');

var xlsx2json = require('./lib/index.js');

//Default configuration
var config = {
  force: false,
  verbose: process.env.NODE_VERBOSE == "true" || process.env.NODE_VERBOSE == "1",
  source: undefined,
  target: undefined
};

//Prints help message
function help() {
  console.log("Usage:");
  console.log("  xslx2json [options] <input.xlsx> <output.json>");
  console.log("");
  console.log("Options:");
  console.log("  --force         Force overwrite file");
  console.log("  --help          Print this message");
  console.log("  --verbose       Enable detailed logging");
  console.log("  --version       Print version number");
  console.log("");
  console.log("Examples:");
  console.log("  xlsx2json --version");
  console.log("  xlsx2json --verbose file.xlsx");
  console.log("  xlsx2json --force file.xlsx file.json");
}

//Command line interface
var args = process.argv.slice(2);
for (var i = 0; i < args.length; i++) {
  switch (args[i]) {
    case "--help":
      help();
      process.exit(0);
      break;
      
    case "-f":
    case "--force":
      config.force = true;
      break;
      
    case "--verbose":
      config.verbose = true;
      break;
    
    case "-v":    
    case "--version":
      console.log(require('./package.json').version);
      process.exit(0);
      break;
      
    default:
      if (args[i].indexOf('-') == 0) {
        console.error("Unknown command line argument: " + args[i]);
        process.exit(2);
      }
      else if (!config.source) {
        config.source = args[i];
      }
      else if (!config.target) {
        config.target = args[i];
      }
      else {
        console.error("Too many arguments: " + args[i]);
        process.exit(2);
      }
      break;
  }
}

//Check if the source file argument is defined
if (!config.source) {
  console.error("Source file not defined!");
  process.exit(2);
}

//Check if the source file exists
if (!fs.existsSync(config.source)) {
  console.error("File not found: " + config.source);
  process.exit(2);
}

//if (!config.target) {
//  config.target = config.source + ".json";
//}

//Check if output file is ready to overwrite
if (config.target && !config.force && fs.existsSync(config.target)) {
  console.error("File already exists: " + config.source);
  process.exit(3);
}

//Convert XLSX to JSON
xlsx2json(config.source, { verbose: config.verbose }, function(error, result) {
  if (error) {
    console.error(error);
    process.exit(1);
    return;
  }
  
  var json = JSON.stringify(result, " ", 2);
  if (!config.target) {
    //Print to console
    console.log(json);
    if (config.verbose) {
      console.log("Done!");
    }
  }
  else {
    //Write to file
    fs.writeFileSync(config.target, json);
    console.log("Done!");
  }
  process.exit(0);
});