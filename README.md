# excel-to-json

Converting xlsx/xls/csv file buffer to json buffer using nodejs.

## Install

```
  npm install @krvinay/excel_to_json
```

## Usage

```javascript
  excel_json = require("@krvinay/excel_to_json");
  fs = require("fs");
  excel_json(fs.readFileSync("<excel file>"), function(err, result) {
    if(err) {
      console.error(err.message);
    }else {
      console.log(result);
    }
  });
```