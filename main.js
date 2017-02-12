/**
 * Created by amit on 2/12/17.
 */
var fs = require('fs');
var XLSX = require('xlsx');
var moment = require('moment');

var excelFile = "/Users/amit/Downloads/Book1.xlsx";
var numDays = 3;

fs.readFile(excelFile, function (err, buffer) {
    if (err) throw err;
    /* convert data to binary string */
    var data = new Uint8Array(buffer);
    var arr = new Array();
    for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
    var bstr = arr.join("");
    /* Call XLSX */
    var workbook = XLSX.read(bstr, {type:"binary"});
    var first_sheet_name = workbook.SheetNames[0];
    var address_of_cell = 'A1';

    /* Get worksheet */
    var worksheet = workbook.Sheets[first_sheet_name];
    var json = XLSX.utils.sheet_to_json(worksheet);
    for(var rowNum in json){
        var dueDate = moment(json[rowNum]['Date Due'], "M/DD/YY");
        var withinNumDays = dueDate.clone().subtract(numDays, 'days').startOf('day');
        if (withinNumDays.isAfter(withinNumDays)){
            console.log(dueDate+" is within "+numDays+" days");
        }
        console.log(dueDate);
    }
});