//jshint esnext:true
(function() {
    'use strict';
    const excel = require('./js/excel');
    const osHomedir = require('os-homedir');

    const workbook = excel.read('test.xlsx');

    if (workbook.Content !== undefined) {
        //workbook.Content.map(x => console.log(x.OfferName));

        const contentArr = workbook.Content.reduce((p, c) => {
            let a = [];
            for (let name in c) {
                a.push(c[name]);
            }

            p.push(a);
            return p;
        }, []);

        const wb = {};
        wb.Content = contentArr;

        excel.write(wb, osHomedir() + '/desktop/out.xlsx');

    } else {
        console.log('brak arkusza Content');
    }

})();
