//jshint esnext:true
(function() {
    'use strict';
    const XLSX = require('xlsx');
    const JSON = require('./writeFromJSON');
    const home = require('os-homedir');

    const path = home() + '/Desktop/out.xlsx';
    const file = XLSX.readFile('test.xlsx');
    const json = XLSX.utils.sheet_to_json(file.Sheets.Content);
    const nJson = json.map(x => {
        x.newName = 'new ' + x.OfferName;
        return x;
    });

    JSON.writeConten(nJson, path);

})();
