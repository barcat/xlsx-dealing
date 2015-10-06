//jshint esnext:true
/*jshint camelcase: false */
module.exports = (function() {
    'use strict';
    const XLSX = require('xlsx');
    // const home = require('os-homedir');

    const file = XLSX.readFile('lista.xlsx');
    const listJson = XLSX.utils.sheet_to_json(file.Sheets.Val);

    function repalceWords(str, orginalPhraze, endPhraze) {
        return str.replace(orginalPhraze, endPhraze);
    }
    //
    // function setOneWord(wordsArr, chosenWord) {
    //
    // }

    function corectName(orginalName, cName) {
        let res;
        for (let obj of listJson) {
            if (cName === obj.cName) {
                res = repalceWords(orginalName.toLowerCase(), obj.arg1.toLowerCase(), obj.arg2);
            }
        }
        return res;
    }

    return {
        corectName: corectName
    };

})();
