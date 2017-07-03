## Introduction

Excel XLSX document to JSON converter as a Node.js library and a command line utility.

## Getting Started

Install the package:
```bash
npm install -g node-xlsx2json
```

Convert the document:
```bash
xlsx2json --verbose path/to/sample.xlsx sample.json
```

## Command Line

```
Usage:
  xslx2json [options] <input.xlsx> <output.json>

Options:
  -f              Force overwrite
  --help          Print this message
  --verbose       Enable detailed logging
  --version, -v   Print version number

Examples:
  xlsx2json --version
  xlsx2json --verbose file.xlsx
  xlsx2json -f file.xlsx file.json
```

## Using as library

```javascript
var xlsx2json = require('node-xlsx2json');
xlsx2json('path/to/sample.xlsx', function(error, result) {
  if (error) return console.error(error);
  console.log(result);
});
```

## Example of output

```json
{
  "worksheets": [
    {
      "sheetId": 1,
      "name": "Blank",
      "data": []
    },
    {
      "sheetId": 2,
      "name": "Calculate",
      "data": [
        {
          "cell": "A1",
          "value": "1"
        },
        {
          "cell": "A2",
          "value": "2"
        },
        {
          "cell": "A3",
          "formula": "=SUM(A1:A2)",
          "value": "3"
        },
        {
          "cell": "C3",
          "formula": "=2 * (PI() / 4)",
          "value": "1.5707963267948966"
        }
      ]
    },
    {
      "sheetId": 3,
      "name": "Variables",
      "data": [
        {
          "cell": "A1",
          "value": "0"
        },
        {
          "cell": "A2",
          "value": "1234"
        },
        {
          "cell": "A3",
          "value": "1"
        },
        {
          "cell": "A4",
          "value": "10"
        },
        {
          "cell": "A5",
          "value": "0.99"
        },
        {
          "cell": "A6",
          "value": "0.5625"
        }
      ]
    }
  ]
}
```