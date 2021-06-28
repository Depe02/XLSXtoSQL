const ExcelJS = require('exceljs');

const express = require('express');
const app = express();
const port = 8080;

app.post('/', async (req, res) => {
    //read file
    var wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.query.filename);
    var ws1 = wb.worksheets[0];

    //get rows
    var result = "";
    var nameTab;
    var nameCol;
    var i;
    var arrCol = new Array(); //array column names
    var arrDef = new Array(); //array default
    ws1.eachRow({includeEmpty: true}, function (row, rowNumber) {
        if (rowNumber == 1) { //column names
            row.eachCell(function (cell, colNumber) {
                arrCol.push(cell.value);
            });
            //get template
            var ws2 = wb.worksheets[1];
            ws2.eachRow(function (row, rowNumber) {
                if (rowNumber == 1) { //table name
                    row.eachCell(function (cell, colNumber) {
                        if (colNumber == 1) {
                            nameTab = cell.value;
                        }
                    });
                } else if (rowNumber > 2) {
                    row.eachCell(function (cell, colNumber) {
                        switch (colNumber) {
                            case 1: //column name
                                nameCol = cell.value;
                                break;
                            case 2: //db name
                                i = arrCol.indexOf(nameCol);
                                if (i != -1) { //value found
                                    arrCol[i] = cell.value;
                                }
                                break;
                            case 3: //type
                                switch (cell.value.toLowerCase()) {
                                    case "int":
                                    case "float":
                                    case "decimal":
                                        arrDef[i] = 0;
                                        break;
                                    case "string":
                                    case "varchar":
                                        arrDef[i] = "";
                                        break;
                                    case "date":
                                        arrDef[i] = "1970-01-01";
                                        break;
                                }
                                break;
                            case 4: //default
                                arrDef[i] = cell.value;
                        }
                    });
                }
            });
        } else {
            var arrRow = new Array();
            arrCol.forEach(function (item, index) {
                arrRow.push(arrDef[index]);
            });
            row.eachCell({includeEmpty: true}, function (cell, colNumber) {
                if (typeof cell.value != 'object') {
                    arrRow[colNumber - 1] = cell.value
                } else if (cell.value != null) { //type formula
                    arrRow[colNumber - 1] = cell.value.result;
                }
            });
            result += 'INSERT INTO ' + nameTab + '(';
            arrCol.forEach(function (item, index) {
                result += item + ',';
            });
            result = result.slice(0, -1); //remove ','
            result += ') VALUES('
            arrRow.forEach(function (item, index) {
                result += item + ',';
            });
            result = result.slice(0, -1); //remove ','
            result += ');'
            result += '<br><br>';
        }

    });

    res.writeHead(200, {'Content-Type': 'text/html'});
    res.write(result);
    res.end();
});

app.listen(port, () => {
    console.log(`Listening on port ${port}!`)
});