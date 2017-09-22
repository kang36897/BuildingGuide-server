var parseXlsx = require("excel");


function convertArrayToObject(data) {
    var temp = [];

    var keys = data[0];

    for (var i = 1; i < data.length; i++) {
        var item = {};

        var values = data[i];

        for (var j = 0; j < keys.length; j++) {
            var key = keys[j];
            var v = values[j].trim();

            item[key] = v;
        }

        console.log(item);
        temp.push(item);
    }

    return temp;
}


var excelHandler = {

    importData: function (excelPath, callback) {

        try {
            parseXlsx(excelPath, function (err, data) {
                if (err) {
                    callback(err);
                    return;
                }
                // data is an array of arrays

                console.log(data[0]);
                callback(null, convertArrayToObject(data));
            });
        } catch (e) {
            callback(e);
        }

    }
};


module.exports = excelHandler;
