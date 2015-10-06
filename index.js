//jshint esnext:true
/*jshint camelcase: false */
(function() {
    'use strict';
    const XLSX = require('xlsx');
    const JSON = require('./writeFromJSON');
    const home = require('os-homedir');
    const str = require('./stringManipulations');

    const path = home() + '/Desktop/out.xlsx';
    const file = XLSX.readFile('test.xlsx');
    const json = XLSX.utils.sheet_to_json(file.Sheets.Content);

    const nJson = json.map(x => {
        x.newName = 'new ' + str.corectName(x.OfferName, x.CategoryName);
        return x;
    });

    JSON.writeConten(nJson, path);

})();
