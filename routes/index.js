var express = require('express');
var router = express.Router();
var excel = require('../utils/excel_handler');
var path = require("path");
var fs = require("fs");

/* GET home page. */
router.get('/', function (req, res, next) {
    res.render('index', {title: 'Express'});
});


router.get("/getArrangement", function (req, res) {

    try {
        var resourceFile = "chores/arrangement.xlsx";
        var excelPath = path.resolve(resourceFile);
        if (!fs.existsSync(excelPath)) {
            throw new Error("file not exist error");
        }

        excel.importData(excelPath, function (err, data) {
            if (err) {
                res.json({
                    "error": "failed to parse excel",
                    "message": "There is something wrong of your excel format."
                });
                return;
            }

            res.json({"error": null, "message": "success", "data": data});


        });

    } catch (e) {
        res.json({"error": e.message, "message": "file: " + resourceFile + " not exist"});
    }


});

module.exports = router;
