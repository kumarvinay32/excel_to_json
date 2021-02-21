const xlsx = require("xlsx");
const xlsjs = require('xlsjs');
const cvcsv = require('csv');
const path = require('path');

module.exports = excel_to_json;

function excel_to_json(input, callback) {
    try {
        if (!input.hasOwnProperty('buffer')) {
            throw new Error("Buffer not found.");
        }
        if (['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'].indexOf(input.mimetype) == -1) {
            throw new Error('Invalid File, Only Excel file supported.');
        }
        let csv = '';
        if (path.extname(input.originalname) == '.csv') {
            csv = Buffer.from(input.buffer).toString('utf-8');
        } else if (path.extname(input.originalname) == '.xlsx') {
            let wb = xlsx.read(input.buffer, { type: 'buffer' });
            let ws = wb.Sheets[wb.SheetNames[0]];
            csv = xlsx.utils.make_csv(ws);
        } else {
            let wb = xlsjs.read(input.buffer, { type: 'buffer' });
            let ws = wb.Sheets[wb.SheetNames[0]];
            csv = xlsx.utils.make_csv(ws);
        }
        if (!csv) {
            throw new Error("Invalid File.");
        }
        let record = [];
        let header = [];
        cvcsv().from.string(csv).transform(function (row) {
            row.unshift(row.pop());
            return row;
        }).on('record', function (row, index) {
            if (index === 0) {
                header = row;
            } else {
                var obj = {};
                header.forEach(function (column, index) {
                    var key = column.trim();
                    obj[key] = row[index].trim();
                })
                record.push(obj);
            }
        }).on('end', function () {
            callback(null, record);
        }).on('error', function (error) {
            callback(error);
        });
    } catch (error) {
        callback(error);
    }
}