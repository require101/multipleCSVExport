var express = require('express');
var router = express.Router();
var xlsx = require('xlsx');
var fs = require('fs');

var names = [`Batman`, `Robin`, `nightwing`];

var data = {
  "sapa": [
    {
      "id": 4,
      "name": "Ready to Wear (RTW)",
      "gross_sales_weighted": 3375300.1999999997,
      "area_weighted": 39211.70008255585,
      "store_23_gross_sales_weighted": 1752961.8,
      "store_23_area_weighted": 8866,
      "store_11_gross_sales_weighted": 1622338.3999999997,
      "store_11_area_weighted": 10297.4,
      "$$treeLevel": 0
    },
    {
      "id": 4,
      "name": "Ready to Wear (RTW)",
      "gross_sales_weighted": 3375300.2,
      "area_weighted": 39211.700082555835,
      "store_23_gross_sales_weighted": 1752961.8,
      "store_23_area_weighted": 8866,
      "store_11_gross_sales_weighted": 1622338.3999999997,
      "store_11_area_weighted": 10297.4,
      "$$treeLevel": 1
    },
    {
      "id": 10,
      "name": "Coats/Swim",
      "gross_sales_weighted": 215867.6,
      "area_weighted": 1816.1863235791438,
      "store_23_gross_sales_weighted": 75892.8,
      "store_23_area_weighted": 487.4,
      "store_11_gross_sales_weighted": 139974.8,
      "store_11_area_weighted": 400.2,
      "$$treeLevel": 2
    } ]
  }


router.get('/', function(req, res, next) {
  res.render('index');
});

router.get('/testexport', function(req, res, next){
  console.log(`Got route`);
  intializeWorkbook();
  for (i = 0; i < names.length; i++){
    organizeData(data, names[i]);
  }
  xlsx.writeFile(workbook, 'temp/tempExcel.xlsx');
 // reads the file at the temporary location
   fs.readFile('temp/tempExcel.xlsx', function(err, file){
    res.setHeader('Content-disposition', 'attachment; filename=export.xlsx');
    res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    var filestream = fs.createReadStream('temp/tempExcel.xlsx');
    filestream.pipe(res)
   });
});

function organizeData(data, name){
  if (name === 'Batman'){
    var newData = [];
    for (let record of data.sapa) {
      let newRecord = {
        id: record.id,
        name: record.name,
        gross_sales_weighted: record.gross_sales_weighted
      };
      newData.push(newRecord)
    }
    var ws = xlsx.utils.json_to_sheet(newData);
    workbook.SheetNames.push(name);
    workbook.Sheets[name] = ws;
  }
  else if (name === 'Robin'){
    var newData = [];
    for (let record of data.sapa){
      let newRecord = {
        id : record.id,
        name : record.name
      };
      newData.push(newRecord);
    }
    var ws = xlsx.utils.json_to_sheet(newData);
    workbook.SheetNames.push('Robin');
    workbook.Sheets['Robin'] = ws;
  } else {
    var newData = [];
    for (let record of data.sapa){
      let newRecord = {
        id : record.id,
        gross_sales_weighted : record.gross_sales_weighted
      }
      newData.push(newRecord);
    }
    var ws = xlsx.utils.json_to_sheet(newData);
    workbook.SheetNames.push('nightwing');
    workbook.Sheets['nightwing'] = ws;
  }
}

function intializeWorkbook(){
  workbook = xlsx.read("");
  workbook.SheetNames = [];
}

module.exports = router;
