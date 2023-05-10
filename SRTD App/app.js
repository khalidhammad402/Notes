const express = require("express");
const bodyParser = require("body-parser");
const fs = require("fs");
const watch = require("node-watch");
const socketIo = require("socket.io");
const excelJs = require("exceljs");

const app = express();
app.use(bodyParser.urlencoded({extended: true}));
app.set('view engine', 'ejs');
app.use(express.static('public'));

const server = app.listen(3000, function(req, res){
    console.log("server started on port 3000");
})

const io = socketIo(server);

const folderpath = __dirname+"/applications";
const workbook = new excelJs.Workbook();

watch(folderpath, { recursive: true }, function(evt, name) {
    var pdf_name = name.slice(__dirname.length + 14);
    pdf_name = pdf_name.substring(0, pdf_name.length-4);
    if(evt=="update" && pdf_name!=".DS_S"){
        details = pdf_name.split('_');
        workbook.csv.readFile('db.csv').then(function() {
            const worksheet = workbook.getWorksheet(1);
            worksheet.addRow(details);
            return workbook.csv.writeFile('db.csv');
        }).then(function() {
            console.log('Row added to csv file');
            io.emit('fileChange', { data: details});
        }).catch(function(error) {
            console.log('Error:', error);
        });
    }
});

app.get("/", function(req, res){
    database = [];
    workbook.csv.readFile('db.csv').then(() => {
        const worksheet = workbook.getWorksheet(1);
        worksheet.eachRow((row, rowNumber) => {
            details = []
            row.eachCell((cell, colNumber) => {
                details.push(cell.value);
            });
            database.push(details);
        });
    }).then(function() {
        res.render('index', {data: database});
    }).catch((error) => {
        console.log(error);
    });
});

