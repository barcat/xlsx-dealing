module.exports = (function() {
    'use strict';
    const XLSX = require('XLSX');

    const _header = ['Id',
        'ShopId',
        'ShopProductId',
        'OfferName',
        'ProductDescription',
        'CategoryId',
        'CategoryName',
        'ShopCategoryName',
        'Price',
        'ShopProductUrl',
        'AttributeName1',
        'AttributeValue1',
        'AttributeName2',
        'AttributeValue2',
        'AttributeName3',
        'AttributeValue3',
        'AttributeName4',
        'AttributeValue4',
        'AttributeName5',
        'AttributeValue5',
        'AttributeName6',
        'AttributeValue6',
        'AttributeName7',
        'AttributeValue7',
        'AttributeName8',
        'AttributeValue8',
        'AttributeName9',
        'AttributeValue9',
        'AttributeName10',
        'AttributeValue10',
        'ReasonsOfHiding',
        'newName'
    ];

    const _workbook = function() {
        return {
            SheetNames: [],
            Sheets: {}
        };
    };

    const _datenum = function(v, date1904) {
        if (date1904) {
            v += 1462;
        }
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    };

    const _sheetFromArrayOfArrays = function(data) {
        var ws = {};
        var range = {
            s: {
                c: 10000000,
                r: 10000000
            },
            e: {
                c: 0,
                r: 0
            }
        };
        for (var R = 0; R !== data.length; ++R) {
            for (var C = 0; C !== data[R].length; ++C) {
                if (range.s.r > R) {
                    range.s.r = R;
                }
                if (range.s.c > C) {
                    range.s.c = C;
                }
                if (range.e.r < R) {
                    range.e.r = R;
                }
                if (range.e.c < C) {
                    range.e.c = C;
                }
                var cell = {
                    v: data[R][C]
                };
                if (cell.v === null) {
                    continue;
                }
                var cell_ref = XLSX.utils.encode_cell({
                    c: C,
                    r: R
                });
                if (typeof cell.v === 'number') {
                    cell.t = 'n';
                } else if (typeof cell.v === 'boolean') {
                    cell.t = 'b';
                } else if (cell.v instanceof Date) {
                    cell.t = 'n';
                    cell.z = XLSX.SSF._table[14];
                    cell.v = _datenum(cell.v);
                } else {
                    cell.t = 's';
                    ws[cell_ref] = cell;
                }
            }
        }
        if (range.s.c < 10000000) {
            ws['!ref'] = XLSX.utils.encode_range(range);
        }
        return ws;
    };

    const read = function(path) {
        let workbook = XLSX.readFile(path);
        let sheets = {};
        for (let name of workbook.Props.SheetNames) {
            sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
        }
        return sheets;
    };

    const write = function(json, path, header) {
        if (header === undefined) {
            header = _header;
        }

        const wb = _workbook();
        console.log(json);
        for (let sName in json) {
            wb.SheetNames.push(sName);
            wb.Sheets.Content = _sheetFromArrayOfArrays(json[sName]);
        }

        XLSX.writeFile(wb, path);
    };

    return {
        read: read,
        write: write
    };
})();
