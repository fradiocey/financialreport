
//var azure = require('azure-storage');
var express = require('express');
var router = express.Router();
var async = require("async");
var sql = require("seriate");
var azure = require('azure-storage');
var moment = require('moment');


console.log("OK");

var http = require('http');
var port = process.env.port || 1337;


var excelbuilder = require('msexcel-builder');

var totalexcel = 0;
var totalsheet = 0;
var RevenueValueJSON = [];

var datedb = [];
var dateArr = [];



// Change the config settings to match your 
// SQL Server and database

var config = {  
    "server": "bjic5w3pth.database.windows.net",
    "user": "bpldbsa@bjic5w3pth",
    "password": "Password999",
    "database": "bpldb2",
    "procName": "ups_SummaryList",
    "options" : {encrypt: true,
                "requestTimeout": 90000}  
};

sql.setDefaultConfig( config );



//create worksheet
var workbook1 = excelbuilder.createWorkbook('./Summary_List', 'Summary List for Financial Report.xlsx');
var curr = 0;
function createWorkSheet(id,totalexcel,workcode,workname,c,d,callback1){
var sysdb = [];

var conarr = [];
var revarr = [];


var sheetname = workcode +"-"+ workname;
    var sheet1 = workbook1.createSheet(sheetname, 100, 120);
     
    // Fill some data
    sheet1.set(1+1, 1+2, 'Date');
    //sheet1.set(2, 1, 'Date DB');
    sheet1.set(2+1, 1+2, 'Revenue \n (RM)');
    sheet1.set(3+1, 1+2, 'Steel \n (RM)');
    sheet1.set(4+1, 1+2, 'Concrete \n (RM)');
    sheet1.set(5+1, 1+2, 'Total \n (RM)');


    // type center
  sheet1.align(1+1, 1+2, 'center');
  sheet1.align(2+1, 1+2, 'center');
  sheet1.align(3+1, 1+2, 'center');
  sheet1.align(4+1, 1+2, 'center');
  sheet1.align(5+1, 1+2, 'center');

  // type border
  sheet1.border(1+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(2+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(3+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(4+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(5+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});

  sheet1.width(1+1, '15');
  sheet1.width(2+1, '15');
  sheet1.width(3+1, '15');
  sheet1.width(4+1, '15');
  sheet1.width(5+1, '15');



    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth();
    var date = today.getDate();
    var dayslists = [];

    for(var i=0; i<30; i++){
      var day=new Date(year, month - 1, date + i+1);
     
      var datefield = (moment(day).format('MMM DD YYYY'))
      dayslists.push(datefield)
   
    }

    for (var u = 0; u < dayslists.length; u++){
    sheet1.set(1+1, u+2+2, dayslists[u]);
     // type border
  sheet1.border(1+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(2+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(3+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(4+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(5+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
    sysdb.push(dayslists[u])
    }





var startdate = moment( moment().subtract(30, 'days') ).format("YYYY-MM-DD")
var enddate = moment().format("YYYY-MM-DD")

//get steelrate
sql.execute( { 
 query: "select DISTINCT steelrate from BplPjtFinancialRate where Project_Id = "+id+""
} ).then( function( result ) {

//console.log(result)

})


// get date 
sql.execute( { 
 query: "select b.id, b.pileno, b.CageInstallStartTimeStamp, b.mainreinforcementcontent, b.helicalreinforcementcontent from bplpile as b where project_id = "+id+" and CAST(b.CageInstallStartTimeStamp as date) BETWEEN  '"+startdate+"' AND '"+enddate+"' and (MainReinforcementContent is not null or helicalreinforcementcontent is not null) order by b.CageInstallStartTimeStamp"
} ).then( function( result ) {


// get result
for (var t = 0; t < result.length; t++){
var str1 = result[t].CageInstallStartTimeStamp;
var str = str1.toString()
var datares = str.slice(4,15);
datedb.push(datares)
var datefielddb = (moment(str1).format('YYYY-MM-DD'));
dateArr.push(datefielddb) 
var b =0

 //get concrete & Revenue
sql.execute( { 
 query: "execute dbo.usp_CountRevenue "+id+",'"+dateArr[t]+"' "
} ).then( function( result1 ) {
//RevenueValueJSON.push({date:datefielddb,revenue:result[0].RevenueValue,concrete:result[0].ConcreteValue})
//RevenueValueJSON.push({date:datefielddb})

b++ 

 //console.log(b +"=="+ dateArr.length +"complete")
   
if (b == result.length){
  //console.log(revarr)
console.log(datedb)






  for (var g = 0; g <sysdb.length; g++ ){
  for (var h = 0; h < datedb.length; h++){
//console.log(sysdb[g] +"=="+ datedb[h])

//console.log(h)
    if (sysdb[g] == datedb[h]){
   


    //sheet1.set(2, g+2, datedb[h]);
    sheet1.set(2+1, g+2+2, parseFloat(revarr[h]).toFixed(2));
    sheet1.set(4+1, g+2+2, parseFloat(conarr[h]).toFixed(2));
 
    

    }
    }
    }

// Save it 
  workbook1.save(function (err) {
        if (err)
        throw err;
      else
      console.log('congratulations, your workbook created');
 
    
  }); 

 
  revarr.length = 0
  conarr.length = 0
}
else{
  
  revarr.push(result1[0].RevenueValue)
  conarr.push(result1[0].ConcreteValue)
  //console.log(revarr)
 
}



 // get worksheet


})

}




//






})

callback1();

}



// main
sql.execute( { 
 query: "SELECT id, ProjectCode, ProjectName, ConcreteCover, SteelDiameter FROM BplProject where id in (5,7)" 
} ).then( function( result ) {
  totalexcel = result.length;

  
  var y = 0;
  var loopsheet = function(result){
    var id = result[y].id;
   
    createWorkSheet(result[y].id,totalexcel,result[y].ProjectCode,result[y].ProjectName,result[y].ConcreteCover,result[y].SteelDiameter,function(){
  
      y++
     
      if (y < result.length){
      loopsheet(result);
      }else{
        console.log("Completed - End");
        
        setTimeout(function() {
          console.log("Calling To Donwnload")
  CallToDownload();
}, 30000);
    
      }
      
    })
  }
  //start loopWorkbook
  loopsheet(result) 

  
}, function( err ) {
        console.log( "Something bad happened here:", err );
    } );



module.exports = router;
