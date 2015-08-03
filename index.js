var xlsx = require('xlsx');

/**
 * Converts an excel file to JSON
 * @param {string} file the name of the excel file
 * @param {string} sheet the name of the sheet to convert
 * @returns {Array}
 */
function parseExcel(file, sheet) {

    var workbook = xlsx.readFile(file);
    var worksheet = workbook.Sheets[sheet];

    var header = {}, rows = {};
    for (z in worksheet) {
        if (z[0] === '!') continue;

        var zPattern = new RegExp(/^([A-Z]+)([0-9]+)$/);
        var zMatch = z.match(zPattern);

        var x = zMatch[1];
        var y = parseInt(zMatch[2]);

        if (y == 1) {
            header[x] = worksheet[z].v;
        }
    }

    for (z in worksheet) {
        if (z[0] === '!') continue;

        var zPattern = new RegExp(/^([A-Z]+)([0-9]+)$/);
        var zMatch = z.match(zPattern);

        var x = zMatch[1];
        var y = parseInt(zMatch[2]);

        if (y != 1) {
            if (typeof rows[y - 1] == 'undefined') rows[y - 1] = {};
            rows[y - 1][header[x]] = worksheet[z].v;
        }

    }

    var arr = Object.keys(rows).map(function (k) {
        return rows[k]
    });

    return arr;

}

module.exports = {
    parseExcel: parseExcel
}