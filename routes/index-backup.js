var express = require('express');
var router = express.Router();
var async = require("async");

var sql = require("seriate");
var prodID = 5;
var PileDim = 880;
var operation = 0;

//var excelbuilder = require('msexcel-builder');
var excelbuilder = require('msexcel-builder-colorfix-intfix');
var workbookName = [];
var workbookID = [];
var pilediameter = [];

var totalexcel = 0;
var totalsheet = 0;

// Change the config settings to match your 
// SQL Server and database

// Create a new workbook file in current working-path
  //var workbook = excelbuilder.createWorkbook('./', 'Summary_List.xlsx')



var config = {  
    "server": "bjic5w3pth.database.windows.net",
    "user": "bpldbsa@bjic5w3pth",
    "password": "Password999",
    "database": "bpldb2",
    "procName": "ups_SummaryList",
    "options" : {encrypt: true}  
};

sql.setDefaultConfig( config );

/* template execute
sql.execute( {  
} ).then( function( result ) {
  
}, function( err ) {
        console.log( "Something bad happened:", err );
    } );
*/

  /*for (var j = 0; j < totalexcel; j++){
  var workbook1 = excelbuilder.createWorkbook('./', 'Summary List for'+workbookID[j]+' - '+workbookName[j]+'.xlsx');
  }*/


//get file name
sql.execute( { 
  query: "SELECT id, ProjectCode, ProjectName FROM BplProject " 
} ).then( function( result ) {
  //console.log(result)
  totalexcel = result.length;
  for (var i=0; i < totalexcel; i++){
  workbookID.push(result[i].ProjectCode);
  workbookName.push(result[i].ProjectName);
  }

  
}, function( err ) {
        console.log( "Something bad happened:", err );
    } );


var workbook1 = excelbuilder.createWorkbook('./', 'Summary List for.xlsx');
// get pile diameter
sql.execute( {  
  query: "SELECT distinct pilediameter FROM bplpile where Project_Id = 5 order by PileDiameter asc " 
} ).then( function( result ) {
  console.log(result)

  totalsheet = result.length;
  var x = 0;
  for (var i = 0; i < totalsheet; i++){
  pilediameter.push(result[i].pilediameter)
  }
  //;

  //for (var i = 0; i < totalsheet; i++){  
  
    // getExcel(result[i].pilediameter,i);
  //}
  var i = 0;
  var loopSheet = function(result){
    getExcel(result[i].pilediameter,i,function(){
      i++
      if (i < result.length){
        loopSheet(result);
      }

    })
  }
  // start loopSheet
  loopSheet(result)



}, function( err ) {
        console.log( "Something bad happened:", err );
    } );



// execute strdproc 
function getExcel(pileno,curno,callback){
console.log(pileno)
sql.execute( {      
        query: "execute dbo.usp_SummaryList 5,"+pileno+""
    } ).then( function( result ) {
 
 console.log("Here"+pileno)
  var totaldata = result.length+5;
  //for (var k=0; k < totalsheet; k++){
  
  var sheet1 = workbook1.createSheet(pileno, 25, totaldata)
  //}
    // header
  sheet1.set(2, 1, 'Summary List - Project Name');
  //type
  sheet1.set(6, 3, 'General Details');
  sheet1.set(11, 3, 'Pile Details');
  sheet1.set(17, 3, 'Steel Cage');
  sheet1.set(21, 3, 'Concrete');
  // type center
  sheet1.align(6, 3, 'center');
  sheet1.align(11, 3, 'center');
  sheet1.align(17, 3, 'center');
  sheet1.align(21, 3, 'center');
  //merge table
  sheet1.merge({col:6,row:3},{col:9,row:3});
  sheet1.merge({col:11,row:3},{col:15,row:3});
  sheet1.merge({col:17,row:3},{col:19,row:3});
  sheet1.merge({col:21,row:3},{col:25,row:3});
  // type border
  sheet1.border(6, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(7, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(8, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(9, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(6, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(7, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(8, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(9, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
  
  sheet1.border(11, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(12, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(13, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(14, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(15, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(11, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(12, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(13, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(14, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(15, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(17, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(18, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(19, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
 
  sheet1.border(17, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(18, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(19, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(21, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(22, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(23, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(24, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(25, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  
  sheet1.border(21, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(22, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(23, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(24, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(25, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
 

  // Fill some data
  sheet1.set(1, 4, 'Pile No.');
  sheet1.set(2, 4, 'Rig Number (last rig to operate on pile)');
  sheet1.set(3, 4, 'Boring Start Date');
  sheet1.set(4, 4, 'Concrete Start Date');
  //gap
  sheet1.set(5, 4, '');
  //end gap
  sheet1.set(6, 4, 'Platform Level (mRL)');
  sheet1.set(7, 4, 'Cut-off Level (mRL)');
  sheet1.set(8, 4, 'Hit Rock Level (mRL)');
  sheet1.set(9, 4, 'Toe Level (mRL)');
   //gap
  sheet1.set(10, 4, '');
  //end gap
  sheet1.set(11, 4, 'Bored Depth PPL (m)');
  sheet1.set(12, 4, 'Pile Length (m)');
  sheet1.set(13, 4, 'Cavity (m)');
  sheet1.set(14, 4, 'Total Rock Coring (m)');
  sheet1.set(15, 4, 'Rock Socket (m)');
    //gap
  sheet1.set(16, 4, '');
  //end gap
  sheet1.set(17, 4, 'Reinforcement Contain');
  sheet1.set(18, 4, 'Helical/Spiral');
  sheet1.set(19, 4, 'Cage Length');
     //gap
  sheet1.set(20, 4, '');
  //end gap
  sheet1.set(21, 4, 'Theoretical');
  sheet1.set(22, 4, 'Actual');
  sheet1.set(23, 4, 'Wastage (%)');
  sheet1.set(24, 4, 'Grade');
  sheet1.set(25, 4, 'DO Number');
// end create a new worksheet

// wrap header true
sheet1.wrap(1, 4, 'true');
sheet1.wrap(2, 4, 'true');
sheet1.wrap(3, 4, 'true');
sheet1.wrap(4, 4, 'true');
sheet1.wrap(5, 4, 'true');
sheet1.wrap(6, 4, 'true');
sheet1.wrap(7, 4, 'true');
sheet1.wrap(8, 4, 'true');
sheet1.wrap(9, 4, 'true');
sheet1.wrap(10, 4, 'true');
sheet1.wrap(11, 4, 'true');
sheet1.wrap(12, 4, 'true');
sheet1.wrap(13, 4, 'true');
sheet1.wrap(14, 4, 'true');
sheet1.wrap(15, 4, 'true');
sheet1.wrap(16, 4, 'true');
sheet1.wrap(17, 4, 'true');
sheet1.wrap(18, 4, 'true');
sheet1.wrap(19, 4, 'true');
sheet1.wrap(20, 4, 'true');
sheet1.wrap(21, 4, 'true');
sheet1.wrap(22, 4, 'true');
sheet1.wrap(23, 4, 'true');
sheet1.wrap(24, 4, 'true');
sheet1.wrap(25, 4, 'true');

// header center

sheet1.align(1, 4, 'center');
sheet1.align(2, 4, 'center');
sheet1.align(3, 4, 'center');
sheet1.align(4, 4, 'center');
sheet1.align(5, 4, 'center');
sheet1.align(6, 4, 'center');
sheet1.align(7, 4, 'center');
sheet1.align(8, 4, 'center');
sheet1.align(9, 4, 'center');
sheet1.align(10, 4, 'center');
sheet1.align(11, 4, 'center');
sheet1.align(12, 4, 'center');
sheet1.align(13, 4, 'center');
sheet1.align(14, 4, 'center');
sheet1.align(15, 4, 'center');
sheet1.align(16, 4, 'center');
sheet1.align(17, 4, 'center');
sheet1.align(18, 4, 'center');
sheet1.align(19, 4, 'center');
sheet1.align(20, 4, 'center');
sheet1.align(21, 4, 'center');
sheet1.align(22, 4, 'center');
sheet1.align(23, 4, 'center');
sheet1.align(24, 4, 'center');
sheet1.align(25, 4, 'center');

for (var i = 5; i < totaldata; i++){
   
sheet1.set(1, i, result[i-5].v1);
sheet1.set(2, i, result[i-5].v2);
sheet1.set(3, i, result[i-5].v3);
sheet1.set(4, i, result[i-5].v4);
sheet1.set(5, i, "");
sheet1.set(6, i, result[i-5].v5);
sheet1.set(7, i, result[i-5].v6);
sheet1.set(8, i, result[i-5].v7);
sheet1.set(9, i, result[i-5].v8);
sheet1.set(10, i, "");
sheet1.set(11, i, result[i-5].v9);
sheet1.set(12, i, result[i-5].v10);
sheet1.set(13, i, result[i-5].v11);
sheet1.set(14, i, result[i-5].v12);
sheet1.set(15, i, result[i-5].v13);
sheet1.set(16, i, "");
sheet1.set(17, i, result[i-5].v14);
sheet1.set(18, i, result[i-5].v15);
sheet1.set(19, i, result[i-5].v16);
sheet1.set(20, i, "");
sheet1.set(21, i, result[i-5].v17);
sheet1.set(22, i, result[i-5].v18);
sheet1.set(23, i, result[i-5].v19);
sheet1.set(24, i, result[i-5].v20);
sheet1.set(25, i, result[i-5].v21);


// wrap true
sheet1.wrap(1, i, 'true');
sheet1.wrap(2, i, 'true');
sheet1.wrap(3, i, 'true');
sheet1.wrap(4, i, 'true');
sheet1.wrap(5, i, 'true');
sheet1.wrap(6, i, 'true');
sheet1.wrap(7, i, 'true');
sheet1.wrap(8, i, 'true');
sheet1.wrap(9, i, 'true');
sheet1.wrap(10, i, 'true');
sheet1.wrap(11, i, 'true');
sheet1.wrap(12, i, 'true');
sheet1.wrap(13, i, 'true');
sheet1.wrap(14, i, 'true');
sheet1.wrap(15, i, 'true');
sheet1.wrap(16, i, 'true');
sheet1.wrap(17, i, 'true');
sheet1.wrap(18, i, 'true');
sheet1.wrap(19, i, 'true');
sheet1.wrap(20, i, 'true');
sheet1.wrap(21, i, 'true');
sheet1.wrap(22, i, 'true');
sheet1.wrap(23, i, 'true');
sheet1.wrap(24, i, 'true');
sheet1.wrap(25, i, 'true');


    
}

sheet1.width(1, '10');
sheet1.width(2, '10');
sheet1.width(3, '20');
sheet1.width(4, '20');
sheet1.width(5, '5');
sheet1.width(6, '10');
sheet1.width(7, '10');
sheet1.width(8, '10');
sheet1.width(9, '10');
sheet1.width(10, '5');
sheet1.width(11, '10');
sheet1.width(12, '10');
sheet1.width(13, '10');
sheet1.width(14, '10');
sheet1.width(15, '10');
sheet1.width(16, '5');
sheet1.width(17, '10');
sheet1.width(18, '10');
sheet1.width(19, '10');
sheet1.width(20, '5');
sheet1.width(21, '10');
sheet1.width(22, '10');
sheet1.width(23, '10');
sheet1.width(24, '10');
sheet1.width(25, '10');

sheet1.merge({col:1,row:1},{col:23,row:1});
callback();

 


if (curno == parseInt(totalsheet-1) ){
   workbook1.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  });  
}  
    
    }, function( err ) {
        console.log( "Something bad happened:", err );
    } );
//    
}//get excelbuilder
// end function 





router.get('/', function(req, res) {
  res.render('index', { title: 'Express',test: 'Node JS'});
 
})

module.exports = router;
