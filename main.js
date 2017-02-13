/**
 * Created by amit on 2/12/17.
 */
'use strict';

var fs = require('fs');
var XLSX = require('xlsx');
var moment = require('moment');
var sendpulse = require("sendpulse");
var config = require("./config.json");
var winston = require('winston');

winston.add(winston.transports.File, { filename: 'trace.log' });
winston.level = 'info';

sendpulse.init(config.email.sendpulse.key, config.email.sendpulse.secret);

fs.readFile(config.excelFile, function (err, buffer) {
    if (err) throw err;

    var itemsWithDueDtWithin3Days = "";

    /* Read file */
    var data = new Uint8Array(buffer);
    var arr = new Array();
    for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
    var bstr = arr.join("");

    /* Convert XLSX to JSON object, and read the first worksheet */
    var workbook = XLSX.read(bstr, {type:"binary"});
    var first_sheet_name = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[first_sheet_name];
    winston.log('debug', 'Name of worksheet'+first_sheet_name);
    var json = XLSX.utils.sheet_to_json(worksheet);
    winston.log('debug', 'JSON Data: '+JSON.stringify(json));
    /* Today and the date after 3 days. */
    var today = moment();
    var withinNumDays = today.clone().add(config.noOfDays, 'days').startOf('day');

    /* Loop through the list of items */
    for(var rowNum in json){
        winston.log('debug', 'Using dueDate : '+json[rowNum][config.dueDateColumn]);
        var dueDate = moment(json[rowNum][config.dueDateColumn], "MM/DD/YY");
        /* Check if the due date is within 3 days */
        if (dueDate.isBetween(today, withinNumDays)){
            itemsWithDueDtWithin3Days += "<br>"+json[rowNum]["Task Overview"]+"("+json[rowNum][config.dueDateColumn]+")</br>";
            winston.log('debug', 'Date compare is : '+dueDate.format('MM/DD/YY')+' between '+today.format('MM/DD/YY')+' and '+withinNumDays.format('MM/DD/YY')+'..Yes');
        }
        else{
            winston.log('debug', 'Date compare is : '+dueDate.format('MM/DD/YY')+' between '+today.format('MM/DD/YY')+' and '+withinNumDays.format('MM/DD/YY')+'..No');
        }
    }

    /* Send email */
    if (itemsWithDueDtWithin3Days.length > 0){
        var email = {
            "html" : "<pre>"+itemsWithDueDtWithin3Days+"</pre>",
            "text" : "Your email text version goes here",
            "subject" : "Tasks due in the next "+config.noOfDays+" days.",
            "from" : config.email.from,
            "to" : config.email.to
        };

        var answerGetter = function answerGetter(data){
            winston.log('info', 'Email sent result '+JSON.stringify(data));
        }
        sendpulse.smtpSendMail(answerGetter,email);
    }
    else{
        winston.log('info', 'No records with due date in the next '+config.noOfDays+ ' days.');
    }
});