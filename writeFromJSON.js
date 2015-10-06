/*jshint camelcase: false */
module.exports = (function() {
    'use strict';
    const XLSX = require('xlsx');

    const head = ['Id',
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

    function workbook() {
        return {
            SheetNames: [],
            Sheets: {}
        };
    }

    function datenum(v, date1904) {
        if (date1904) {
            v += 1462;
        }
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    }

    function sheetFromArrayOfArrays(data) {
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
                    cell.v = datenum(cell.v);
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
    }

    function write(source, path, sName) {
        let fbw = workbook();
        let arr = source.map(x => {
            let a = [];
            for (let v in x) {
                a.push(x[v]);
            }
            return a;
        });
        let workSheet = sheetFromArrayOfArrays(arr);

        fbw.SheetNames.push(sName);
        fbw.Sheets[sName] = workSheet;
        XLSX.writeFile(fbw, path);
    }

    function addHader(header, json) {
        let arr = [];
        arr.push(header);

        for (let i of json) {
            //console.log('>>>>>',i);
            let fild = i;
            let a = [];
            for (let v of header) {
                a.push(fild[v]);
            }
            arr.push(a);
        }
        return arr;
    }

    function writeConten(source, path) {
        let fbw = workbook();
        let arr = addHader(head, source);


        let workSheet = sheetFromArrayOfArrays(arr);

        fbw.SheetNames.push('Content');
        fbw.Sheets.Content = workSheet;
        XLSX.writeFile(fbw, path);
    }

    return {
        write: write,
        writeConten: writeConten
    };
})();
