//var azure = require('azure-storage');
var express = require('express');
var router = express.Router();
var async = require("async");
var sql = require("seriate");
var azure = require('azure-storage');

console.log("OK");

var http = require('http');
var port = process.env.port || 1337;

var prodID = 5;
var PileDim = 880;
var operation = 0;

var excelbuilder = require('msexcel-builder');
//var excelbuilder = require('msexcel-builder-colorfix-intfix');
var workbookName = [];
var workbookID = [];
var pilediameter = [];
var workbookname = [];
var excellist = "";
var newRCArr = [];
var newHLArr = [];
var uniqueRCArr = [];
var uniqueHLArr = [];
var newArrayMB = [];


var totalexcel = 0;
var totalsheet = 0;
var completed = 0;


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


 var workbook1 = "";
//get file name
sql.execute( { 
 query: "SELECT id, ProjectCode, ProjectName, ConcreteCover, SteelDiameter FROM BplProject where ProjectStatus IS NULL" 
} ).then( function( result ) {
  totalexcel = result.length;

  
  var y = 0;
  var loopWorkbook = function(result){
    var projCode = result[y].ProjectCode;
     //console.log ("RUN WORKBOOK"+y) 
    getPile(result[y].id,result[y].ProjectCode,result[y].ProjectName,result[y].ConcreteCover,result[y].SteelDiameter,function(){
    
      y++
      if (y < result.length){
      loopWorkbook(result);
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
  loopWorkbook(result) 

  
}, function( err ) {
        console.log( "Something bad happened here:", err );
    } );


function getPile(id,workcode,workname,c,d,callback1){
workbook1 = excelbuilder.createWorkbook('./Summary_List', 'Summary List for '+workcode+'-'+workname+'.xlsx');
workbookname.push('Summary List for '+workcode+'-'+workname+'.xlsx')
console.log(workbook1)
// get pile diameter
 

sql.execute( {  
  query: "SELECT distinct pilediameter FROM bplpile where Project_Id = "+id+" order by PileDiameter asc " 
} ).then( function( result ) {
  //console.log(result)
  totalsheet = result.length;
  var x = 0;
  for (var i = 0; i < totalsheet; i++){
  pilediameter.push(result[i].pilediameter)
  }
 
  var i = 0;
  var loopSheet = function(result){
    
    getExcel(id,result[i].pilediameter,workcode,workname,c,d,function(){
      i++
      if (i < result.length){
        loopSheet(result);
         completed = 0
      }
      else{
        console.log("Completed")
        completed = 1
         workbook1.save(function (err) {
        if (err)
        throw err;
      else
      console.log('congratulations, your workbook created');
 
    
  });  
        callback1();
      
       
      }
    })
    
  }
  // start loopSheet
  
  loopSheet(result)

  //callback();




}, function( err ) {
        console.log( "Something bad happened data:", err );
    } );

}



// execute strdproc 
function getExcel(id,pileno,projCode,projName,c,d,callback){

newRCArr.length = 0;
uniqueRCArr.length = 0;
newHLArr.length = 0;
uniqueHLArr.length = 0;

//newHLArr.length = 0;
//uniqueHLArr.length = 0;

sql.execute( {      
        query: "execute dbo.usp_SummaryList "+id+","+pileno+""
    } ).then( function( result ) {
 
  var totaldata = result.length+5;
  var totaldatatot = result.length+6;


  
    //for (var k=0; k < totalsheet; k++){
  
  var sheet1 = workbook1.createSheet(''+pileno+'', 90, totaldatatot)
  var RCAdditional = 0;
  var HLAdditional = 0;
  var incrementy = 0;
  var incrementyHL = 0;
 
/* start of Cage Breakdown*/
  for (var i = 5; i < totaldata; i++){
  var newRC =  result[i-5].v14;
  var newHL =  result[i-5].v15;

  // Main Bar
  
// 1 must not null
// found (+) skip

if (newRC != null){
if (newRC.indexOf("na") == -1){ //na
if (newRC.indexOf("TEST") == -1){ //TEST
if (newRC.indexOf("+") == -1){ //single HL
if (newRC.indexOf("(") == -1){ //single HL
if (newRC.indexOf(",") == -1){
  //console.log("Single")
newRC = newRC.replace(/\s/g, "") 
newRC =  newRC.split("T");
newRC = newRC[1];
newRCArr.push("T"+newRC)

}
//console.log(newRCArr)
}
else{
newRC = newRC.split(",")
// must get value 14 & 20
for (var q = 0; q < newRC.length; q++){
var str = JSON.stringify(newRC[q]);
str = str.replace(/\s/g, "")
str = str.replace(/['"]+/g, '') 
var MainBar = str.split("T");
var MainBar0 = MainBar[0];
var MainBar1 = MainBar[1];
var MainBar2 = JSON.stringify(MainBar[1]);
MainBar1 = MainBar1.replace(/ *\([^)]*\) */g, "");
newRCArr.push("T"+MainBar1)
}// for loop

}// else multiple
}// no (+)
}//NOT na
}//NOT test
}// not null



 for(var o in newRCArr){
        if(uniqueRCArr.indexOf(newRCArr[o]) === -1){
            uniqueRCArr.push(newRCArr[o]);
        }
    }


RCAdditional = uniqueRCArr.length



// Helical Link

if (newHL != null){
if (newHL.indexOf("TEST") == -1){ //TEST
if (newHL.indexOf("na") == -1){ //na
if (newHL.indexOf("2T") == -1){ //2T
if (newHL.indexOf(",") == -1){ //single HL 
//Single HL
newHL = newHL.replace(/['"]+/g, '');
newHL =  newHL.split("-");
newHL = newHL[0];
newHLArr.push(newHL)

}//single HL
else{
// multiple HL
newHL = newHL.split(",")
// must get value 14 & 20
for (var j = 0; j < newHL.length; j++){
var strHL = JSON.stringify(newHL[j]);
strHL = strHL.replace(/['"]+/g, '');
var BDHL = strHL.split("-");
var BDHL0 = BDHL[0];
BDHL0 = BDHL0.replace(/ *\([^)]*\) */g, "");
BDHL0 = BDHL0.replace(/ /g,'')
newHLArr.push(BDHL0)
}// else 
}// multiple HL
}// not 2T
}// not na
}// not test
}// null

for(var w in newHLArr){
        if(uniqueHLArr.indexOf(newHLArr[w]) === -1){
            uniqueHLArr.push(newHLArr[w]);
        }
    }

HLAdditional = uniqueHLArr.length

}// end loop i=5


if (RCAdditional > 0){
for (var u = 0; u < RCAdditional;u++){
incrementy++
sheet1.set(22+incrementy, 4, ''+uniqueRCArr[u]+'\n Weight (kg)');
sheet1.wrap(22+incrementy, 4, 'true');
sheet1.border(22+incrementy, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
sheet1.border(22+incrementy, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
sheet1.wrap(22+incrementy, 4, 'true');
sheet1.align(22+incrementy, 4, 'center');
sheet1.border(22+incrementy, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.width(22+incrementy, '10');


}

}

if (HLAdditional > 0){
for (var z = 0; z < HLAdditional;z++){
incrementyHL++
sheet1.set(22+incrementy+incrementyHL+2, 4, ''+uniqueHLArr[z]+'\n Weight (kg)');
sheet1.wrap(22+incrementy+incrementyHL+2, 4, 'true');
sheet1.border(22+incrementy+incrementyHL+2, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
sheet1.border(22+incrementy+incrementyHL+2, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
sheet1.wrap(22+incrementy+incrementyHL+2, 4, 'true');
sheet1.align(22+incrementy+incrementyHL+2, 4, 'center');
sheet1.border(22+incrementy+incrementyHL+2, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.width(22+incrementy+incrementyHL+2, '10');


}

}
/* End of Cage Breakdown*/



//console.log("RowPlus"+incrementy)
  //}
    // header
  sheet1.set(2, 1, 'Summary List for '+projCode+'-'+projName+'' );
  //type
  sheet1.set(1, 3, 'Pile Details');
  sheet1.set(8, 3, 'General Details');
  sheet1.set(14, 3, 'Pile Details');
  sheet1.set(21 , 3, 'Steel Cage');
  sheet1.set(28+RCAdditional+HLAdditional, 3, 'Concrete');
  // type center
  sheet1.align(1, 3, 'center');
  sheet1.align(8, 3, 'center');
  sheet1.align(14, 3, 'center');
  sheet1.align(21, 3, 'center');
  sheet1.align(28+RCAdditional+HLAdditional, 3, 'center');
  //merge table
  sheet1.merge({col:1,row:3},{col:6,row:3});
  sheet1.merge({col:8,row:3},{col:12,row:3});
  sheet1.merge({col:14,row:3},{col:19,row:3});
  sheet1.merge({col:21,row:3},{col:26+RCAdditional+HLAdditional,row:3});
  sheet1.merge({col:28+RCAdditional+HLAdditional,row:3},{col:33+RCAdditional+HLAdditional,row:3});
  // type border
  sheet1.border(1, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(2, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(3, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(4, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(5, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(6, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(1, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(2, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(3, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(4, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(5, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(6, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(8, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(9, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(10, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(11, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(12, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(8, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(9, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(10, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(11, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(12, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
  
  
  sheet1.border(14, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(15, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(16, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(17, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
   sheet1.border(18, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(19, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(14, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(15, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(16, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(17, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(18, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(19, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(21, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(22, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(23+RCAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(24+RCAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(25+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(26+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
 
  sheet1.border(21, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(22, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(23+RCAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(24+RCAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(25+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(26+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(28+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(29+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(30+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(31+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(32+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(33+RCAdditional+HLAdditional, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  
  sheet1.border(28+RCAdditional+HLAdditional, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(29+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(30+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(31+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(32+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(33+RCAdditional+HLAdditional, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
 

  // Fill some data
  sheet1.set(1, 4, 'Pile No.');
  sheet1.set(2, 4, 'Rig');
  sheet1.set(3, 4, 'Boring Start Date');
  sheet1.set(4, 4, 'Boring End Date');
  sheet1.set(5, 4, 'Concrete Start Date');
  sheet1.set(6, 4, 'Concrete End Date');
  //gap
  sheet1.set(7, 4, '');
  //end gap
  sheet1.set(8, 4, 'Piling Platform Level (mRL)');
  sheet1.set(9, 4, 'Cut-off Level (mRL)');
  sheet1.set(10, 4, 'Hit Rock Level (mRL)');
  sheet1.set(11, 4, 'Rock Layer Thickness (m)'); // new
  sheet1.set(12, 4, 'Toe Level (mRL)');
   //gap
  sheet1.set(13, 4, '');
  //end gap
  sheet1.set(14, 4, 'Bored Depth - PPL (m)');
  sheet1.set(15, 4, 'Pile Length (m)');
  sheet1.set(16, 4, 'Cavity (m)');
  sheet1.set(17, 4, 'Total Rock Coring (m)');
  sheet1.set(18, 4, 'Paid Rock Coring (m)');
  sheet1.set(19, 4, 'Rock Socket (m)');
    //gap
  sheet1.set(20, 4, '');
  //end gap
  sheet1.set(21, 4, 'Main Reinforcement Content');
  sheet1.set(22, 4, 'Main Bar Weight (kg)'); // new requirement
  sheet1.set(23+RCAdditional, 4, 'Helical/\nSpiral Link');
  sheet1.set(24+RCAdditional, 4, 'Helical/Spiral Weight (kg)'); // new
  sheet1.set(25+RCAdditional+HLAdditional, 4, 'Cage Length');
  sheet1.set(26+RCAdditional+HLAdditional, 4, 'Computed Cage Length');
     //gap
  sheet1.set(27+RCAdditional+HLAdditional, 4, '');
  //end gap
  sheet1.set(28+RCAdditional+HLAdditional, 4, 'Theoretical (m3)');
  sheet1.set(29+RCAdditional+HLAdditional, 4, 'Actual (m3)');
  sheet1.set(30+RCAdditional+HLAdditional, 4, 'Wastage (%)');
  sheet1.set(31+RCAdditional+HLAdditional, 4, 'Grade');
  sheet1.set(32+RCAdditional+HLAdditional, 4, 'DO Number');
  sheet1.set(33+RCAdditional+HLAdditional, 4, 'Concrete Volume (m3)');
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
sheet1.wrap(23+RCAdditional, 4, 'true');
sheet1.wrap(24+RCAdditional, 4, 'true');
sheet1.wrap(25+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(26+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(27+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(28+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(29+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(30+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(31+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(32+RCAdditional+HLAdditional, 4, 'true');
sheet1.wrap(33+RCAdditional+HLAdditional, 4, 'true');

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
sheet1.align(23+RCAdditional, 4, 'center');
sheet1.align(24+RCAdditional, 4, 'center');
sheet1.align(25+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(26+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(27+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(28+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(29+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(30+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(31+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(32+RCAdditional+HLAdditional, 4, 'center');
sheet1.align(33+RCAdditional+HLAdditional, 4, 'center');

  var boredarr = [];
  var pilearr = [];
  var cavityarr = [];
  var rockcoringarr = [];
  var paidrockarr = [];
  var rocksocketarr = [];

  var MainSteelBarArr = [];
  var MainSteelBarBDArr = [];

  var FinalTotalLofHLArr = [];
  var FinalTotalLofHLBDArr = [];

  var TheoreticalArr = [];
  var ActualArr = [];
  var totFinalMB = [];
  var totFinalHL = [];
  

for (var i = 5; i < totaldata; i++){
//parseFloat
var Platform = result[i-5].v5;
if (typeof Platform  != "undefined" || Platform != null){
var PlatformparseFloat = parseFloat(Platform).toFixed(3)
//console.log(PlatformparseFloat)
sheet1.set(8, i, PlatformparseFloat);
}


var numlistNewLinedo = JSON.stringify(result[i-5].v21);
if (numlistNewLinedo != "null"){
numlistNewLinedo = numlistNewLinedo.replace(/['"]+/g, '');
numlistNewLinedo = numlistNewLinedo.replace(/,/g, '\n');
numlistNewLinedo = numlistNewLinedo.replace(/ /g,'');
}
else{
numlistNewLinedo = ""
}


var numlistNewLinecv = JSON.stringify(result[i-5].v22);
if (numlistNewLinecv != "null"){
numlistNewLinecv = numlistNewLinecv.replace(/['"]+/g, '');
numlistNewLinecv = numlistNewLinecv.replace(/,/g, '\n');
numlistNewLinecv = numlistNewLinecv.replace(/ /g,'');
}
else{
numlistNewLinecv ="";
}



var numlistNewLineHRL = JSON.stringify(result[i-5].v7);
if (numlistNewLineHRL != "null"){
numlistNewLineHRL= numlistNewLineHRL.replace(/['"]+/g, '');
numlistNewLineHRL = numlistNewLineHRL.replace(/,/g, '\n');
numlistNewLineHRL = numlistNewLineHRL.replace(/ /g,'');
}
else{
numlistNewLineHRL ="";
}

var numlistNewLineRLT = JSON.stringify(result[i-5].v23);
if (numlistNewLineRLT!= "null"){
numlistNewLineRLT= numlistNewLineRLT.replace(/['"]+/g, '');
numlistNewLineRLT = numlistNewLineRLT.replace(/,/g, '\n');
numlistNewLineRLT = numlistNewLineRLT.replace(/ /g,'');
}
else{
numlistNewLineHRL ="";
}

/* Fix Table
High Tensile */




for (var dash = 0; dash < RCAdditional; dash++){
sheet1.set(23+dash, i, "-");//newreq 
sheet1.wrap(23+dash, i, 'true');
sheet1.border(23+dash, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});  
}


//console.log( numlistNewLinedo )

var HighTenstile = 0;


// Main Bar
// split Reinforcement Content
var MainSteelBar = 0;
var FinalTotalLofHL  = 0;
var computedcagelength = 0;
var Totalboredarr = 0;
var MultipleCageMainSteelBar = []



if (result[i-5].v14 != null){

var checkarray = JSON.stringify(result[i-5].v14);
var checkcagelength = JSON.stringify(result[i-5].v16);

if (checkarray.indexOf("(") == -1){ 
if (result[i-5].v16 != null){
if (checkcagelength.indexOf ("+") == -1){

//("1 only")
var str = result[i-5].v14;
str = str.replace(/ /g,'')
var MainBar = str.split("T");
var MainBar0 = MainBar[0];
var MainBar1 = MainBar[1];
var MainBar1Div1000 = parseFloat(MainBar1) / 1000;
var piledimeter = parseFloat(pileno);
var BeforexNosDiv = parseFloat(piledimeter*MainBar1Div1000)
var PileLength = parseFloat(result[i-5].v10).toFixed(3)
var BeforexNosPlus = parseFloat(PileLength) + parseFloat(BeforexNosDiv);
var BeforexNosTimes = parseFloat(MainBar0) * parseFloat(BeforexNosPlus);
var cagelength = result[i-5].v16;




if (MainBar1  == "10"){
HighTenstile = 0.617;

}
else if (MainBar1  == "12"){
HighTenstile = 0.888;

}
else if (MainBar1  == "16"){
  HighTenstile = 1.579;
}
else if (MainBar1 == "20"){
    HighTenstile = 2.466;
}
else if (MainBar1 == "25"){
    HighTenstile = 3.854;
}
else if (MainBar1  == "32"){
    HighTenstile = 6.313;
}
else if (MainBar1  == "40"){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}

/*
if (HighTenstile == 0){
  if (MainBar1  == 10){
HighTenstile = 0.617;

}
else if (MainBar1  == 12){
HighTenstile = 0.888;

}
else if (MainBar1  == 16){
  HighTenstile = 1.579;
}
else if (MainBar1 == 20){
    HighTenstile = 2.466;
}
else if (MainBar1 == 25){
    HighTenstile = 3.854;
}
else if (MainBar1  == 32){
    HighTenstile = 6.313;
}
else if (MainBar1  == 40){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}
}
*/

var finalCal = (parseFloat(MainBar0)  * parseFloat(HighTenstile) *parseFloat(cagelength));

console.log(parseFloat(HighTenstile) )

var curMainbar = "T"+MainBar1

if (finalCal <= 0 || isNaN(parseFloat(finalCal))){
MainSteelBar = "-"
MainSteelBarArr.push(0)

//MainSteelBarBDArr.push(curMainbar+"-"+0)
}else{
MainSteelBar = parseFloat(finalCal).toFixed(3);
sheet1.set(22, i, MainSteelBar);//newreq 


// re create breakdown


if (result[i-5].v14 != null){
if (result[i-5].v14.indexOf("X") == -1){ //cannot have X
if (result[i-5].v14.indexOf("na") == -1){ //na
if (result[i-5].v14.indexOf("TEST") == -1){ //TEST
if (result[i-5].v14.indexOf("+") == -1){ //dont have + sign
var checkarray = JSON.stringify(result[i-5].v14);
var checkcagelength = JSON.stringify(result[i-5].v16);

if (result[i-5].v14.indexOf("(") == -1){ //single MB
    if (result[i-5].v14.indexOf(",") == -1){// single cannot have ,
   
var BreakDownMB = (uniqueRCArr.length)

if (BreakDownMB > 0){



for (var BDSB = 0 ; BDSB < BreakDownMB;  BDSB++ ){

if ("T"+MainBar1 == uniqueRCArr[BDSB] ){
MainSteelBarBDArr.push("T"+MainBar1+"-"+MainSteelBar)
sheet1.set(23+BDSB, i, MainSteelBar);//newreq 
sheet1.wrap(23+BDSB, i, 'true');
sheet1.border(23+BDSB, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
// console.log("Got Something HEre")  


}else{
//MainSteelBarBDArr.push(curMainbar+"-"+0)
 //console.log("Got Something HEre")  
sheet1.set(23+BDSB, i, "-");//newreq 
sheet1.wrap(23+BDSB, i, 'true');
sheet1.border(23+BDSB, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});  

}
}

// end of breakdown
    
  
    }// single cannot have ,
}//single MB
else{// multiple MB

}// multiple MB

}// dont have + sign
}// not test
}// not na
}// cannot have X
}// cannot null




//end re create breakdown


MainSteelBarArr.push(parseFloat(finalCal).toFixed(3))



}

}
}else{// got plus sign
MainSteelBar = "-"
MainSteelBarArr.push(0)
    
//MainSteelBarBDArr.push(curMainbar+"-"+0)
}
}else{// empty cage length
MainSteelBar = "-"
MainSteelBarArr.push(0)
    
//MainSteelBarBDArr.push(curMainbar+"-"+0)
}
//console.log(MainSteelBarBDArr)

}
else{
  //("Here More than 2")

var HighTenstile = 0;
var checkarray = result[i-5].v14;
checkarray = checkarray.split(",");


// must get value 16 & 20
var mainbar2arr = [];
var mainbar0arr = [];
var mainbar1arr = [];
var mainbar3arr = [];
var newcagelength =  result[i-5].v16;
if (newcagelength != null){
if (newcagelength.indexOf ("+") != -1){ 
newcagelength = newcagelength.split("+")



// must get value 14 & 20
for (var q = 0; q < checkarray.length; q++){
var str = JSON.stringify(checkarray[q]);
str = str.replace(/['"]+/g, '');
var MainBar = str.split("T");
var MainBar0 = MainBar[0];
var MainBar1 = MainBar[1];
var MainBar2 = JSON.stringify(MainBar[1]);
MainBar1 = MainBar1.replace(/ *\([^)]*\) */g, "");
mainbar0arr.push(MainBar0)
mainbar3arr.push("T"+MainBar1)

//check MainBar1
if (MainBar1 == 10){
HighTenstile = 0.617;

}
else if (MainBar1 == 12){
HighTenstile = 0.888;

}
else if (MainBar1 == 16){
  HighTenstile = 1.579;
}
else if (MainBar1 == 20){
    HighTenstile = 2.466;
}
else if (MainBar1 == 25){
    HighTenstile = 3.854;
}
else if (MainBar1 == 32){
    HighTenstile = 6.313;
}
else if (MainBar1 == 40){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}

mainbar1arr.push(HighTenstile)


// get 2 & 1
MainBar2 = MainBar2.match(/\((.*?)\)/)[1]
mainbar2arr.push(MainBar2)

}


var a = newcagelength;
var b = mainbar2arr;

var index = 0, sums = [];
if (newcagelength != null)
for (var p in b){
	sums.push(a.slice(parseInt(index), parseInt(index)+parseInt(b[p])).reduce(function(init, cur){
  	return parseInt(init) + parseInt(cur);
  }, 0));  
  index += b[p];
  index = parseInt(index)
}

//print result:
//console.log(JSON.stringify( sums ));
// = [12,12] = [22,12]

// must get 14 & 20 - Reinforcement Content




var output = [];
var currentArr = [];

for (var f = 0; f < mainbar0arr.length; f++ ){
var y = parseInt(mainbar0arr[f]) * parseInt(JSON.stringify( sums[f] ))
var current = parseFloat(y) * parseFloat(mainbar1arr[f])
output.push(current)

}


var sumt = output.reduce((a, b) => a + b, 0);
if (sumt <= 0 || isNaN(parseFloat(sumt))){
MainSteelBar = "-"
MainSteelBarArr.push(0)
MainSteelBarBDArr.push("T"+MainBar1+"-"+0)
}else{


//do here
// Breakdown 
// re create breakdown

if (result[i-5].v14 != null){

var a4 = mainbar3arr;
var b4 = output;
var obj4 = {};




a4.forEach(function(el4, i4){
  if (el4 in obj4){
    obj4[el4] = obj4[el4] + b4[i4];
  } else {
    obj4[el4] = b4[i4];
  }
});



var finalMBArr = Object.keys(obj4).map( (key, index) => obj4[key]);


var BreakDownMB = (uniqueRCArr.length)
var BDSB, BDMB
var samefactorMB = [];


 for(var io in mainbar3arr){
        if(samefactorMB.indexOf(mainbar3arr[io]) === -1){
            samefactorMB.push(mainbar3arr[io]);
        }
    }



if (BreakDownMB > 0){

for (BDSB = 0 ; BDSB < samefactorMB.length;  BDSB++ ){

MainSteelBarBDArr.push(samefactorMB[BDSB]+"-"+finalMBArr[BDSB])
  //MainSteelBarBDArr.push(samefactorMB[BDSB]+"-"+finalMBArr[BDSB])
  for (BDMB = 0 ; BDMB < BreakDownMB ;  BDMB++ ){

    if (samefactorMB[BDSB] == uniqueRCArr[BDMB]){
    //if (typeof finalMBArr[BDMB] != "undefined"){
   
    sheet1.set(23+BDMB, i, finalMBArr[BDSB].toFixed(1));//newreq   
    //}
    sheet1.wrap(23+BDMB, i, 'true');
    sheet1.border(23+BDMB, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
    }else{
    //MainSteelBarBDArr.push(samefactorMB[BDSB]+"-"+0)
    //sheet1.set(23+BDMB, i, "-");//newreq 
    sheet1.wrap(23+BDMB, i, 'true');
    sheet1.border(23+BDMB, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
}
  }
}
}

}// not null

if (typeof finalMBArr != "undefined"){
totFinalMB.push(finalMBArr)
}
/*
for (var gi = 0; gi < totFinalMB.length; gi++){
for (var hi = 0; hi <uniqueRCArr.length; hi++){

newtotFinalMB.push(totFinalMB[gi][hi])  
}
}


console.log(newtotFinalMB)
*/


// end of breakdown
MainSteelBar = parseFloat(sumt).toFixed(4); // 6 // end result MainSteelBar
MainSteelBarArr.push(parseFloat(sumt).toFixed(4))

}

}else{
  MainSteelBar = "-"
  MainSteelBarArr.push(0)

}//end more than 2

}else {
  MainSteelBar = "-"
  MainSteelBarArr.push(0)
       
}
}
}else {
  MainSteelBar = "-"
  MainSteelBarArr.push(0)
      
}
// end MainBar

/* End of Fix Table */

/* Helical Link */
// start single HL



var BDRC1 = (uniqueRCArr.length)

for (var dash = 0; dash < HLAdditional; dash++){
sheet1.set(25+dash+BDRC1, i, "-");//newreq 
sheet1.wrap(25+dash+BDRC1, i, 'true');
sheet1.border(25+dash+BDRC1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});  
}




var checkarrayHL = result[i-5].v15;

var checkcagelength = JSON.stringify(result[i-5].v16);
var checkplussign = JSON.stringify(result[i-5].v14);
//checkarrayHL = checkarrayHL.split(",");
if (result[i-5].v15 != null){
if (checkarrayHL.indexOf("(") == -1){ 
if (checkplussign.indexOf("(") == -1){ 
if (result[i-5].v16 != null){
if (checkcagelength.indexOf ("+") == -1){
if (checkplussign.indexOf ("+") == -1){

var checkarray = result[i-5].v14;
var str = JSON.stringify(checkarray);
str = str.replace(/['"]+/g, '');
var MainBar = str.split("T");
var MainBar1 = MainBar[1];

var cagelength = result[i-5].v16;
var HelicalLink = result[i-5].v15;
HelicalLink  = HelicalLink.replace(/ /g,'')// remove whitespace
HelicalLink = HelicalLink.replace(/['"]+/g, '');
//HelicalLink = HelicalLink.replace('T', "");
HelicalLink  = HelicalLink.split("-");
var HelicalLink1 = HelicalLink[1]




/*
var PileLength = parseFloat(result[i-5].v10).toFixed(3);
var HelicalLink = result[i-5].v15;
HelicalLink  = HelicalLink.replace(/ /g,'')// remove whitespace
HelicalLink = HelicalLink.replace(/['"]+/g, '');
//HelicalLink = HelicalLink.replace('T', "");
HelicalLink  = HelicalLink.split("-");

//console.log(HelicalLink[1]) 

/* Helical Chart */
var HelicalChart = 0

if (HelicalLink[0] == "R6"){
HelicalChart = 0.222;
}
else if (HelicalLink[0] == "R8"){
HelicalChart = 0.394;
}
else if (HelicalLink[0] == "R10"){
HelicalChart = 0.617;
}
else if (HelicalLink[0] == "T10"){
HelicalChart = 0.617;
}
else if (HelicalLink[0] == "T12"){
HelicalChart = 0.888;
}
else if (HelicalLink[0] == "T16"){
HelicalChart = 1.579;
}
else if (HelicalLink[0] == "T20"){
HelicalChart = 2.466;
}
else if (HelicalLink[0] == "T25"){
HelicalChart = 3.854;
}
else if (HelicalLink[0] == "T32"){
HelicalChart = 0.6313;
}
else if (HelicalLink[0] == "T40"){
HelicalChart = 9.870;
}
else{
HelicalChart = 0;
}

var spacing = parseFloat(HelicalLink[1]) / 1000;
var D = parseFloat(pileno) / 1000;
 /*Start Formula */
var pileLDivSpacingPlusOne = PileLength / spacing;
var pieDimbracket = (D - parseFloat(c) - parseFloat(c)).toFixed(2)
var pileXpie = ((3.142 * pieDimbracket)*(3.142 * pieDimbracket))
var pilePlusrow = (spacing  * spacing )
var pilePlusBoth =  pileXpie + pilePlusrow 
var rightCal = (PileLength /  spacing ) +1
var squarerootpieDim  = rightCal * Math.sqrt(pilePlusBoth);
var finallengthofHL = squarerootpieDim.toFixed(4)

var lapping = (finallengthofHL / PileLength) * parseFloat(d)
var TotalLofHL = (parseFloat(finallengthofHL) + parseFloat(lapping)) * parseFloat(HelicalChart) 


var step1 = (d * parseFloat(MainBar1)/1000);
var step2 = (cagelength - step1);
var step3 = ((step2/(HelicalLink1/1000))+1) *Math.sqrt((Math.pow(3.14159 * (pileno/1000 - c - c).toFixed(1), 2) +parseFloat( Math.pow((HelicalLink1/1000),2)))) //186.7247


if (step3 <= 0 || isNaN(parseFloat(step3))){
FinalTotalLofHL = "-"
FinalTotalLofHLArr.push(0)
//FinalTotalLofHLBDArr.push(uniqueHLArr[BDHL]+"-"+0)
}else{
FinalTotalLofHL = step3.toFixed(4)
FinalTotalLofHLArr.push(step3.toFixed(4))

// Breakdown single HL


var BreakDownHL = (uniqueHLArr.length)
var BDRC1 = (uniqueRCArr.length)
var curHL = HelicalLink[0]


if (BreakDownHL > 0){
for (var BDHL = 0 ; BDHL < BreakDownHL;  BDHL++ ){
if (curHL == uniqueHLArr[BDHL] ){
FinalTotalLofHLBDArr.push(curHL+"-"+FinalTotalLofHL)
sheet1.set(25+BDHL+BDRC1, i, FinalTotalLofHL);//newreq 
sheet1.wrap(25+BDHL+BDRC1, i, 'true');
sheet1.border(25+BDHL+BDRC1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
}else{
sheet1.set(25+BDHL+BDRC1, i, "-");//newreq 
sheet1.wrap(25+BDHL+BDRC1, i, 'true');
sheet1.border(25+BDHL+BDRC1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});  
}
}

}


// end of breakdown single HL


}
// end of single HL
}else{// got plus sign
FinalTotalLofHL = "-"
FinalTotalLofHLArr.push(0)
//FinalTotalLofHLBDArr.push(uniqueHLArr[BDHL]+"-"+0)
}
}else{// got plus sign
FinalTotalLofHL = "-"
FinalTotalLofHLArr.push(0)
//FinalTotalLofHLBDArr.push(uniqueHLArr[BDHL]+"-"+0)
}
}else{// empty check both
FinalTotalLofHL = "-"
FinalTotalLofHLArr.push(0)
//FinalTotalLofHLBDArr.push(uniqueHLArr[BDHL]+"-"+0)
}
}else{// empty cage length
FinalTotalLofHL = "-"
FinalTotalLofHLArr.push(0)
//FinalTotalLofHLBDArr.push(uniqueHLArr[BDHL]+"-"+0)
}


}
else{
//console.log("Multiple HL")
var basiccalarr = [];
var mainbar2arr = [];
var mainbar1arr = [];
var mainbar0arr = [];
var HL1arr = [];
var HL2arr = [];
var HL3arr = [];
var HL3Rarr = [];
var HL4Rarr = [];
var HelicalChart = 0
// getcagelength
var newcagelength =  result[i-5].v16;
if (newcagelength != null){
if (newcagelength.indexOf ("+") != -1){
//console.log("Found")
newcagelength = newcagelength.split("+")


//get 16,20,16,20
var checkarray = result[i-5].v14;
checkarray = checkarray.split(",");


// helical Spiral
var hLspiral = result[i-5].v15;
hLspiral  = hLspiral.replace(/ /g,'')// remove whitespace
hLspiral = hLspiral.split(",");



// must get value 14 & 20
for (var l = 0; l < checkarray.length; l++){
var str = JSON.stringify(checkarray[l]);
str = str.replace(/['"]+/g, '');
var MainBar = str.split("T");
var MainBar0 = MainBar[0];
var MainBar1 = MainBar[1];
var MainBar2 = JSON.stringify(MainBar[1]);
mainbar0arr.push(MainBar0)
MainBar1 = MainBar1.replace(/ *\([^)]*\) */g, "");
mainbar1arr.push(MainBar1)
MainBar2 = MainBar2.match(/\((.*?)\)/)[1]
mainbar2arr.push(MainBar2)


// basic add
//basiccalarr.push(40*MainBar1/1000)


}
if (newcagelength != null)
for (var e1=0; e1 < newcagelength.length; e1++ ){
var HLstr = JSON.stringify(hLspiral[e1]);
if (typeof HLstr != "undefined")
HLstr = HLstr.replace(/['"]+/g, '');
if (typeof HLstr != "undefined")
var HL = HLstr.split("-");
var HL0 = HL[1];
var HL1 = HL0.replace(/ *\([^)]*\) */g, "");
HL1arr.push(HL1)
var HL2 = HL0.match(/\((.*?)\)/)[1] 
HL2arr.push(HL2)

var HL3 = HL[0];
HL3 = HL3.trim()
HL3 = HL3.replace(/^R/, "");
HL3 = HL3.replace(/^T/, "");
HL3arr.push(HL3)

var HL4 = HL[0];
HL4 = HL4.trim()



if (HL4 == "R6"){
HelicalChart = 0.222;
}
else if (HL4 == "R8"){
HelicalChart = 0.394;
}
else if (HL4 == "R10"){
HelicalChart = 0.617;
}
else if (HL4 == "T10"){
HelicalChart = 0.617;
}
else if (HL4 == "T12"){
HelicalChart = 0.888;
}
else if (HL4 == "T16"){
HelicalChart = 1.579;
}
else if (HL4 == "T20"){
HelicalChart = 2.466;
}
else if (HL4 == "T25"){
HelicalChart = 3.854;
}
else if (HL4 == "T32"){
HelicalChart = 0.6313;
}
else if (HL4 == "T40"){
HelicalChart = 9.870;
}
else{
HelicalChart = 0;
}

HL3Rarr.push(HelicalChart)
HL4Rarr.push(HL4)

}


//console.log(q)


var finalHL = [];

var a = mainbar0arr
var b = mainbar2arr
var c1 = mainbar1arr

var k = HL1arr
var m = HL2arr
var p = HL3arr



var d1 = repeatValuesByAmounts(c1, b); // [5,5,10,15,15,15]
var n = repeatValuesByAmounts(k, m); // [5,5,10,15,15,15]
var q = repeatValuesByAmounts(p, m); // [5,5,10,15,15,15]

//* Math.sqrt((3.14159*((parseFloat(pileno)/1000)- c - c))^2 + (n[e]/1000)^2))
if (newcagelength != null)
for (var e=0; e < newcagelength.length; e++ ){
var f = d1[e] * d/1000 //0.64



var h = parseFloat(newcagelength[e]) - f // 11.362

var I = ((h/(n[e]/1000))+1) *Math.sqrt((Math.pow(3.14159 * (pileno/1000 - c - c).toFixed(1), 2) +parseFloat( Math.pow((n[e]/1000),2)))) //186.7247
var J = ((I/h) * parseInt(d) * (q[e]/1000))
var H = I+J

//parseFloat(J).toFixed(4)

finalHL.push( H * parseFloat(HL3Rarr[e]))


}
// combine


var a5 = HL4Rarr;
var b5 = finalHL;
var obj = {};



a5.forEach(function(el, i5){
  if (el in obj){
    obj[el] = obj[el] + b5[i5];
  } else {
    obj[el] = b5[i5];
  }
});

var finalArr = Object.keys(obj).map( (key, index) => obj[key]);



// Breakdown  Multiple HL

var BreakDownMHL = (uniqueHLArr.length)
var BDBL = (uniqueRCArr.length)
var BDMHL, BDM
var samefactor = []
var count = 0;



 for(var ioi in HL4Rarr){
        if(samefactor.indexOf(HL4Rarr[ioi]) === -1){
            samefactor.push(HL4Rarr[ioi]);
        }
    }


if (BreakDownMHL > 0){



for (BDM = 0 ; BDM < samefactor.length;  BDM++ ){

  FinalTotalLofHLBDArr.push(samefactor[BDM]+"-"+finalArr[BDM])
  for (BDMHL = 0 ; BDMHL < BreakDownMHL ;  BDMHL++ ){

  
  if (samefactor[BDM] = uniqueHLArr[BDMHL]){
      sheet1.set(25+BDBL+BDM , i, finalArr[BDM].toFixed(1));//newreq 
      sheet1.wrap(25+BDM+BDBL, i, 'true');
      sheet1.border(25+BDM+BDBL, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
    
  }
  else{
    sheet1.set(25+BDM+BDBL, i, "-");//newreq 
    sheet1.wrap(25+BDM+BDBL, i, 'true');
    sheet1.border(25+BDM+BDBL, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
   
} 
 
  }



}

}



//totFinalHL.push(finalArr)

// end of Breakdown HL Multiple


FinalTotalLofHL  = finalHL.reduce((a, b) => a + b, 0);
if(FinalTotalLofHL <= 0 || isNaN(parseFloat(FinalTotalLofHL))){
   FinalTotalLofHL = "-" 
   FinalTotalLofHLArr.push(0)
   //FinalTotalLofHLBDArr.push(samefactor[BDSB]+"-"+0)

}else{
 FinalTotalLofHL = FinalTotalLofHL.toFixed(4);
 FinalTotalLofHLArr.push(FinalTotalLofHL)
 //FinalTotalLofHLBDArr.push(samefactor[BDSB]+"-"+FinalTotalLofHL)
}

//if no plus sign
}else{
 FinalTotalLofHL = "-" 
 FinalTotalLofHLArr.push(0)
 //FinalTotalLofHLBDArr.push(0)
}

//if null
}else{
 FinalTotalLofHL = "-" 
 FinalTotalLofHLArr.push(0)
 //FinalTotalLofHLBDArr.push(samefactor[BDSB]+"-"+0)
}
}// end of else

// end if not null
}else{
 FinalTotalLofHL = "-"
 FinalTotalLofHLArr.push(0) 
 //FinalTotalLofHLBDArr.push(0)
}


/* End of Helical Link */


/* computed cage length
single cage
PileLength + StarterBar = Answer

special condition
if PL = 11.8 +0.64 = 12.44
output will be 12+0.44
because after plus will SB got more than 12

multiple cage 2 cage example:
PL - main bar1 length for helical calc  = X 
X+SB = y
12+Y

triple cage 3 cage example:
PL - main bar1 length for helical calc - main bar2 length for helical calc = X 
X+SB = y
12+12+Y

single cage*/

//1. get PL
 
var ccPL = parseFloat(result[i-5].v10).toFixed(3);


//2. get starterbar
//single Cage
var cccheckarrayHL = result[i-5].v15; // helical spiral link (T10-150)
var cccheckcagelength = JSON.stringify(result[i-5].v16); // cage length  (12+12)
var cccheckplussign = JSON.stringify(result[i-5].v14);// main reinforcement content 10T25
var cccagearr = [];

if (ccPL != null){
if (cccheckarrayHL != null){
if (result[i-5].v14 != "na"){
if (cccheckarrayHL.indexOf("(") == -1){ //single HL

var ccstr = JSON.stringify(result[i-5].v14);
ccstr = ccstr.replace(/['"]+/g, '');
var ccMainBar = ccstr.split("T");
var ccMainBar0 = ccMainBar[0];
var ccMainBar1 = ccMainBar[1];
ccMainBar1 = ccMainBar1.replace(/ *\([^)]*\) */g, "");
var cca = d*ccMainBar1/1000
var ccb = parseFloat(ccPL)+cca
var ccc = ccb/12
var ccf = parseInt(ccc)


for (var cce = 0; cce < ccf; cce++){
cccagearr.push(12)
}



cctotalcagearr  = cccagearr.reduce((ccg, cch) => ccg + cch, 0);

var ccj = ccb - cctotalcagearr
var cck = parseFloat(ccj).toFixed(3)


if (cccagearr.length < 1){
computedcagelength = ccb.toFixed(3)
}
else{
cccagearr.unshift(cck)
computedcagelength = cccagearr.join("+")

}

//console.log(computedcagelength)

if (cccheckplussign.indexOf("(") == -1){ //single MainBar
if (cccheckplussign.indexOf ("+") == -1){ //error

}else{//error 
computedcagelength = "-"
}
}else{//multiple Main Bar

}

}else{//Multiple HL

//start here 
//1. get PL
 
var ccmPL = parseFloat(result[i-5].v10).toFixed(3);

var ccmcageL = result[i-5].v16;
var ccmmainbar0arr = [];
var ccmmainbar1arr = [];
var ccmmainbar2arr = [];
var MBLarr = [];
var SBarr = [];
var push12 = [];
var reverseArr = [];
var ccmcheckArr = [];



if (ccmcageL != null){
if (ccmcageL.indexOf ("+") != -1){
ccmcageL = ccmcageL.split("+")


//get 16,20,16,20
var ccmcheckarray = result[i-5].v14;
ccmcheckarray = ccmcheckarray.split(",");

// helical Spiral
var ccmhLspiral = result[i-5].v15;
ccmhLspiral  = ccmhLspiral.replace(/ /g,'')// remove whitespace
ccmhLspiral = ccmhLspiral.split(",");


// must get value 14 & 20
for (var l = 0; l < ccmcheckarray.length; l++){
var ccmstr = JSON.stringify(ccmcheckarray[l]);
ccmstr = ccmstr.replace(/['"]+/g, '');
var ccmMainBar = ccmstr.split("T");
var ccmMainBar0 = ccmMainBar[0];
var ccmMainBar1 = ccmMainBar[1];
var ccmMainBar2 = JSON.stringify(ccmMainBar[1]);
ccmmainbar0arr.push(ccmMainBar0)
ccmMainBar1 = ccmMainBar1.replace(/ *\([^)]*\) */g, "");
ccmmainbar1arr.push(ccmMainBar1)
ccmMainBar2 = ccmMainBar2.match(/\((.*?)\)/)[1]
ccmmainbar2arr.push(ccmMainBar2)
}

var ccmd1 = repeatValuesByAmounts(ccmmainbar1arr, ccmmainbar2arr); // [5,5,10,15,15,15]
var newrevArr = [];
for (var ccmi=0; ccmi <ccmd1.length; ccmi++){
  reverseArr.unshift(ccmd1[ccmi])
  var calIndex = reverseArr[ccmi] * d/1000
  newrevArr.push(calIndex)
}

// 12 is constant
var constant = 12

var a = ccmPL
var b = newrevArr
var index = 0
var step1 = 0
var step2 = 0
var step4 = 0
var step5 = 0
var step6 = 0
var final = 0
var TotArr = [];
var step12cnt  = 0;
var cageFinalArr = [];
var finalAns = 0;




// Step 1: 12 - 0.8 (First array) = 11.2
step1 = constant - b[index++]
step12cnt++
// Step 2: a - 11.2 (Step 1 answer) = 14.8
step2 = a - step1
TotArr.push(step1)
// Step 3: check, if step 2 answer is greater than 12 then do step 4 else do step 7
if (step2 > constant) do{
  step12cnt++
  // Step 4: 12 - 1.6 (Second array) = 10.4
  step4 = constant - b[index++]

  // Step 5 : a - (14.8 (S1 [not step 2?] answer) + 10.4 (S4 answer) ) = 4.4
  step5 = a - (step1 + step4)
  TotArr.push(step4)

  // Step 6: check again if step 5 answer is greater than 12 then do step 4 but take (Third array), else do step 7.
}while (step5 > constant && index < b.length){


TotArr = TotArr.reduce((a1, b1) => a1 + b1, 0);
step6 = a - TotArr

// Step 7: 4.4 (Final Answer) + 0.8 (Last array)
final = (step6) + (parseFloat(b[b.length - 1]))
finalAns = (final.toFixed(3))

}

for (var i1=0; i1 < step12cnt; i1++){
  cageFinalArr.push(12)
}

cageFinalArr.push(finalAns)

// shifted
var ShiftCageFinalArr = [];
for (var i2=0; i2 < cageFinalArr.length; i2++){
ShiftCageFinalArr.unshift(cageFinalArr[i2])
}

if (step6 < 0){
computedcagelength = "-"
}
else{
computedcagelength = ShiftCageFinalArr.join("+")
}



}// end of index +
}// not null


}
}else{ // HL Null
computedcagelength = "-"
}
}else{ // HL Null
computedcagelength = "-"
}
}else{ // PL Null
computedcagelength = "-"
}
//multiple Cage


/* end of computed cage length*/

   
sheet1.set(1, i, result[i-5].v1);
sheet1.set(2, i, result[i-5].v2);
sheet1.set(3, i, result[i-5].v3);
sheet1.set(4, i, result[i-5].v25);
sheet1.set(5, i, result[i-5].v4);
sheet1.set(6, i, result[i-5].v26);

sheet1.set(7, i, "");

sheet1.set(8, i, result[i-5].v5);
sheet1.set(9, i, result[i-5].v6);
sheet1.set(10, i, numlistNewLineHRL);
sheet1.set(11, i, numlistNewLineRLT);//new
sheet1.set(12, i, parseFloat(result[i-5].v8).toFixed(3));


//target
sheet1.set(13, i, "");
sheet1.set(14, i, parseFloat(result[i-5].v9).toFixed(3));
sheet1.set(15, i, parseFloat(result[i-5].v10).toFixed(3));
sheet1.set(16, i, result[i-5].v11);
sheet1.set(17, i, parseFloat(result[i-5].v12).toFixed(1));
sheet1.set(18, i, parseFloat(result[i-5].v24).toFixed(1)); //new paid rock coring
sheet1.set(19, i, parseFloat(result[i-5].v13).toFixed(1));

//new requirement

  if (parseFloat(result[i-5].v9).toFixed(3) != "-" || parseFloat(result[i-5].v9).toFixed(3) != null || parseFloat(result[i-5].v9).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v9).toFixed(3)) ){
  boredarr.push(parseFloat(result[i-5].v9).toFixed(3))
  }
  else{
  boredarr.push(0)
  }

  if (parseFloat(result[i-5].v10).toFixed(3) != "-" || parseFloat(result[i-5].v10).toFixed(3) != null || parseFloat(result[i-5].v10).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v10).toFixed(3))){
  pilearr.push(parseFloat(result[i-5].v10).toFixed(3))
  }
  else{
  pilearr.push(0)
  }

  if (result[i-5].v11 != "" ){
  cavityarr.push(parseFloat(result[i-5].v11).toFixed(3))
    }
  else if (isNaN(result[i-5].v11)){
  cavityarr.push(0)
  }
  else{
  cavityarr.push(0)
  }

  if (parseFloat(result[i-5].v12).toFixed(3) != "-" || parseFloat(result[i-5].v12).toFixed(3) != null || parseFloat(result[i-5].v12).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v12).toFixed(3))){
  rockcoringarr.push(parseFloat(result[i-5].v12).toFixed(3))
  }
  else{
  rockcoringarr.push(0)
  }
  if (parseFloat(result[i-5].v24).toFixed(3) != "-" || parseFloat(result[i-5].v24).toFixed(3) != null || parseFloat(result[i-5].v24).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v24).toFixed(3))){
  paidrockarr.push(parseFloat(result[i-5].v24).toFixed(3))
  }
  else{
  paidrockarr.push(0)
  }
  if (parseFloat(result[i-5].v13).toFixed(3) != "-" || parseFloat(result[i-5].v13).toFixed(3) != null || parseFloat(result[i-5].v13).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v13).toFixed(3))){
  rocksocketarr.push(parseFloat(result[i-5].v13).toFixed(3))
  }
  else{
  rocksocketarr.push(0)
  }

  // Theoretical & Actual

  if (parseFloat(result[i-5].v17).toFixed(3) != "-" || parseFloat(result[i-5].v17).toFixed(3) != null || parseFloat(result[i-5].v17).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v17).toFixed(3))){
  TheoreticalArr.push(parseFloat(result[i-5].v17).toFixed(3))
  }
  else{
  TheoreticalArr.push(0)
  }

  if (parseFloat(result[i-5].v18).toFixed(3) != "-" || parseFloat(result[i-5].v18).toFixed(3) != null || parseFloat(result[i-5].v18).toFixed(3) != "" || isNaN(parseFloat(result[i-5].v18).toFixed(3))){
  ActualArr.push(parseFloat(result[i-5].v18).toFixed(3))
  }
  else{
  ActualArr.push(0)
  }
  

sheet1.set(20, i, "");

var resultwhitespace = result[i-5].v15;
if (resultwhitespace != null){
resultwhitespace = resultwhitespace.replace(/ /g,'')
}

sheet1.set(21, i, result[i-5].v14);
sheet1.set(22, i, MainSteelBar);//new
sheet1.set(23+RCAdditional, i, resultwhitespace);
sheet1.set(24+RCAdditional, i, FinalTotalLofHL);//new
sheet1.set(25+RCAdditional+HLAdditional, i, result[i-5].v16);

if (isNaN(computedcagelength)){
sheet1.set(26+RCAdditional+HLAdditional, i, "-");// new computed cage
}
else{
sheet1.set(26+RCAdditional+HLAdditional, i, computedcagelength);// new computed cage
}

sheet1.set(27+RCAdditional+HLAdditional, i, "");

sheet1.set(28+RCAdditional+HLAdditional, i, parseFloat(result[i-5].v17).toFixed(1));
sheet1.set(29+RCAdditional+HLAdditional, i, parseFloat(result[i-5].v18).toFixed(1));
sheet1.set(30+RCAdditional+HLAdditional, i, parseFloat(result[i-5].v19).toFixed(1));
sheet1.set(31+RCAdditional+HLAdditional, i, result[i-5].v20);
sheet1.set(32+RCAdditional+HLAdditional, i, numlistNewLinedo);
sheet1.set(33+RCAdditional+HLAdditional, i, numlistNewLinecv);


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
sheet1.wrap(23+RCAdditional, i, 'true');
sheet1.wrap(24+RCAdditional, i, 'true');
sheet1.wrap(25+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(26+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(27+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(28+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(29+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(30+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(31+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(32+RCAdditional+HLAdditional, i, 'true');
sheet1.wrap(33+RCAdditional+HLAdditional, i, 'true');

// border
sheet1.border(1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(2, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(3, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(4, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(5, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(6, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(5, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(8, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(9, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(10, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(11, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(12, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(11, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(14, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(15, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(16, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(17, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(18, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(19, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(18, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(21, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(22, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(23+RCAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(24+RCAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(25+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(26+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(24, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(28+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(29+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(30+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(31+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(32+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(33+RCAdditional+HLAdditional, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});


//sheet1.numberFormat(2,1, 10); // equivalent to above
    
}


var Totalboredarr1  = boredarr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalboredarr = Totalboredarr1.toFixed('3')

var Totalpilearr1  = pilearr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalpilearr = Totalpilearr1.toFixed('3')

var Totalcavityarr1  = cavityarr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalcavityarr = Totalcavityarr1.toFixed('1')

var Totalrockcoringarr1  = rockcoringarr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalrockcoringarr = Totalrockcoringarr1.toFixed('1')

var Totalpaidrockarr1  = paidrockarr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalpaidrockarr = Totalpaidrockarr1.toFixed('1')

var Totalrocksocketarr1  = rocksocketarr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var Totalrocksocketarr = Totalrocksocketarr1.toFixed('1')



//cal steel & helical
var TotalMainSteelBarArr1  = MainSteelBarArr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var TotalMainSteelBarArr  = TotalMainSteelBarArr1.toFixed('1')

var TotalFinalTotalLofHLArr1  = FinalTotalLofHLArr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var TotalFinalTotalLofHLArr  = TotalFinalTotalLofHLArr1.toFixed('1')

//Theoretical & Actual
var TotalTheoreticalArr1  = TheoreticalArr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var TotalTheoreticalArr  = TotalTheoreticalArr1.toFixed('1')

var TotalActualArr1  = ActualArr.reduce((r,t) => parseFloat(r) + parseFloat(t), 0);
var TotalActualArr  = TotalActualArr1.toFixed('1')


// Main Bar Breakdown
// re sturcute breakdown total
var calmainbartot = MainSteelBarBDArr.length;
var countRC = (uniqueRCArr.length)
var newcal = [];
var newTArr = [];
var minT1Arr = [];
var minimizeTarr = []


 

// get value of T


if (MainSteelBarBDArr.length != 0){
for (var tt = 0; tt < calmainbartot; tt++ ){
// loop to get T
var minT = MainSteelBarBDArr[tt].split("-")
minT0 = minT[0];
newTArr.push(minT0)
minT1 = minT[1];
minT1Arr.push(minT1)
}

// minimize T



for(var rt in newTArr){
        if(minimizeTarr.indexOf(newTArr[rt]) === -1){
            minimizeTarr.push(newTArr[rt]);
        }
    }

// confirm T array
//console.log(minimizeTarr)

if(uniqueRCArr.length < 2){
//console.log("Single")

var TotalBDRCArr1  = minT1Arr.reduce((h7,t7) => parseFloat(h7) + parseFloat(t7), 0);
var TotalBDRCArr  = TotalBDRCArr1.toFixed('1')
sheet1.set(22+RCAdditional, totaldatatot-1, TotalBDRCArr);//new
}
else{

  // creating Object
  for (var tm = 0; tm < calmainbartot; tm++){
  var newsplittotMB = MainSteelBarBDArr[tm].split("-");
  newcal.push({'name':newsplittotMB[0],'value':newsplittotMB[1]})
  }

  var obj4 = newcal
  var holder = {};

    obj4.forEach(function (d) {
    if(holder.hasOwnProperty(d.name)) {
       holder[d.name] = holder[d.name] + parseFloat(d.value)
    } else {       
       holder[d.name] = parseFloat(d.value)
    }
    });


var obj2 = [];
var cnt = 0
var keyObj = Object.keys(holder)

//console.log(minimizeTarr)

for (var keyi = 0; keyi <keyObj.length; keyi++ ){
  for (var titlei = 0; titlei < uniqueRCArr.length; titlei++){


if (keyObj[keyi] == uniqueRCArr[titlei]){
  var valueRC = parseFloat(holder[keyObj[keyi]]).toFixed(1)
 sheet1.set(23+titlei, totaldatatot-1,valueRC);//new
}
  }
}
}

}    

// Helical link total


var calHLtot = FinalTotalLofHLBDArr.length
var countHL = (uniqueHLArr.length)
var newHLcal = [];
var newTHLarr = [];
var minT1HLarr = [];
var minimizeHLarr = [];

// get T value HL

if (calHLtot != 0){
for (var ttHL = 0; ttHL < calmainbartot; ttHL++ ){
// loop to get T
var minTHL = FinalTotalLofHLBDArr[ttHL].split("-")
minT0HL = minTHL[0];
minT0HL = minT0HL.replace(/ /g,'') 
newTHLarr.push(minT0HL)
minT1HL = minTHL[1];
minT1HLarr.push(minT1HL)
}

}

// minimize T

for(var rtHL in newTHLarr){
        if(minimizeHLarr.indexOf(newTHLarr[rtHL]) === -1){
            minimizeHLarr.push(newTHLarr[rtHL]);
        }
    }

if(uniqueHLArr.length < 2){
var TotalBDHLArr1  = minT1HLarr.reduce((h8,t8) => parseFloat(h8) + parseFloat(t8), 0);
var TotalBDHLArr  = TotalBDHLArr1.toFixed('1')
sheet1.set(24+RCAdditional+HLAdditional, totaldatatot-1, TotalBDHLArr);//new
}else{

 // creating Object
  for (var tmHL = 0; tmHL < calHLtot; tmHL++){
  var newsplittotHL = FinalTotalLofHLBDArr[tmHL].split("-");
  newHLcal.push({'name':newsplittotHL[0],'value':newsplittotHL[1]})
  }



  var obj5 = newHLcal
  var holderHL = {};

    obj5.forEach(function (dHL) {
    if(holderHL.hasOwnProperty(dHL.name)) {
       holderHL[dHL.name] = holderHL[dHL.name] + parseFloat(dHL.value)
    } else {       
       holderHL[dHL.name] = parseFloat(dHL.value)
    }
    });


var obj2HL = [];
var keyObjHL = Object.keys(holderHL)



for (var keyHLi = 0; keyHLi <keyObjHL.length; keyHLi++ ){
  for (var titleHLi = 0; titleHLi < uniqueHLArr.length; titleHLi++){


if (keyObjHL[keyHLi] == uniqueHLArr[titleHLi]){

var valueHL = parseFloat(holderHL[keyObjHL[keyHLi]]).toFixed(1)
 sheet1.set(25+titleHLi+RCAdditional, totaldatatot-1, valueHL );//new
}
  }
}


}






/*
var countHL = (uniqueHLArr.length)

if(countHL < 2){
var TotalBDRCArr1  = FinalTotalLofHLBDArr.reduce((h8,t8) => parseFloat(h8) + parseFloat(t8), 0);
var TotalBDHLArr  = TotalBDRCArr1.toFixed('1')
sheet1.set(24+RCAdditional+HLAdditional, totaldatatot-1, TotalBDHLArr);//new
}
else if(countHL > 1){

if (typeof totFinalHL[0] != "undefined"){
newArrayHL = totFinalHL[0].map(function(col, i4) { 
  return totFinalHL.map(function(row) { 
     return row[i4]; 
  })
});



for (var th = 0; th < newArrayHL.length; th++){


var TotalBDHLArr1  = newArrayHL[th].reduce((h9,t9) => parseFloat(h9) + parseFloat(t9), 0);
var TotalBDHLArr  = TotalBDHLArr1.toFixed('1')
sheet1.set(25+th+RCAdditional, totaldatatot-1, TotalBDHLArr);//new

}
}// not undefined
}

// end of Helical Link Total


//sheet1.set(23+RCAdditional, totaldatatot-1, totalofBDMB[0]);//new
*/

sheet1.set(1,totaldatatot-1,"Total");
sheet1.set(14,totaldatatot-1,Totalboredarr);
sheet1.set(15,totaldatatot-1,Totalpilearr);
sheet1.set(16,totaldatatot-1,Totalcavityarr);
sheet1.set(17,totaldatatot-1,Totalrockcoringarr);
sheet1.set(18,totaldatatot-1,Totalpaidrockarr);
sheet1.set(19,totaldatatot-1,Totalrocksocketarr); //new paid rock coring


sheet1.set(22, totaldatatot-1, TotalMainSteelBarArr);//new
sheet1.set(24+RCAdditional, totaldatatot-1, TotalFinalTotalLofHLArr);//new

sheet1.set(28+RCAdditional+HLAdditional, totaldatatot-1, TotalTheoreticalArr);
sheet1.set(29+RCAdditional+HLAdditional, totaldatatot-1, TotalActualArr);

sheet1.font(1, totaldatatot-1, {bold:'true'});

//bold total
sheet1.font(1, totaldatatot-1, {bold:'true'});
sheet1.font(14, totaldatatot-1, {bold:'true'});
sheet1.font(15, totaldatatot-1, {bold:'true'});
sheet1.font(16, totaldatatot-1, {bold:'true'});
sheet1.font(17, totaldatatot-1, {bold:'true'});
sheet1.font(18, totaldatatot-1, {bold:'true'});
sheet1.font(19, totaldatatot-1, {bold:'true'});
sheet1.font(22, totaldatatot-1, {bold:'true'});
sheet1.font(24+RCAdditional, totaldatatot-1, {bold:'true'});
sheet1.font(28+RCAdditional+HLAdditional, totaldatatot-1, {bold:'true'});
sheet1.font(29+RCAdditional+HLAdditional, totaldatatot-1, {bold:'true'});


sheet1.width(1, '10');
sheet1.width(2, '10');
sheet1.width(3, '20');
sheet1.width(4, '20');
sheet1.width(5, '20');
sheet1.width(6, '20');
sheet1.width(7, '5');
sheet1.width(8, '10');
sheet1.width(9, '10');
sheet1.width(10, '10');
sheet1.width(11, '10');
sheet1.width(12, '10');
sheet1.width(13, '5');
sheet1.width(14, '10');
sheet1.width(15, '10');
sheet1.width(16, '10');
sheet1.width(17, '10');
sheet1.width(18, '10');
sheet1.width(19, '10');
sheet1.width(20, '5');
sheet1.width(21, '15');
sheet1.width(22, '10');
sheet1.width(23+RCAdditional, '10');
sheet1.width(24+RCAdditional, '15');
sheet1.width(25+RCAdditional+HLAdditional, '10');
sheet1.width(26+RCAdditional+HLAdditional, '10');
sheet1.width(27+RCAdditional+HLAdditional, '5');
sheet1.width(28+RCAdditional+HLAdditional, '12');
sheet1.width(29+RCAdditional+HLAdditional, '10');
sheet1.width(30+RCAdditional+HLAdditional, '10');
sheet1.width(31+RCAdditional+HLAdditional, '15');
sheet1.width(32+RCAdditional+HLAdditional, '15');
sheet1.width(33+RCAdditional+HLAdditional, '10');


// border
sheet1.border(1, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(2, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(3, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(4, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(5, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(6, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(5, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(8, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(9, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(10, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(11, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(12, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(11, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(14, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(15, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(16, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(17, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(18, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(19, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(18, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(21, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(22, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(23+RCAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(24+RCAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(25+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(26+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(24, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(28+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(29+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(30+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(31+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(32+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(33+RCAdditional+HLAdditional, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});


//buat sini
// Breakdown Style

var styling = (uniqueRCArr.length)
if (styling > 0){

for (var RC1 = 1 ; RC1 < styling+1;  RC1++ ){
sheet1.font(22+RC1, totaldatatot-1, {bold:'true'});
sheet1.width(22+RC1, '10');
sheet1.border(22+RC1, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
}
}

var stylingHL = (uniqueHLArr.length)
if (stylingHL > 0){

for (var RC2 = 0 ; RC2 < stylingHL;  RC2++ ){
sheet1.font(25+RCAdditional+RC2, totaldatatot-1, {bold:'true'});
sheet1.width(25+RCAdditional+RC2, '10');
sheet1.border(25+RCAdditional+RC2, totaldatatot-1, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
}
}

// end of style




sheet1.merge({col:2,row:1},{col:33+RCAdditional+HLAdditional,row:1});
callback();




    
    }, function( err ) {
        console.log( "Something bad happened Dem:", err );
    } );
//    
}//get excelbuilder
// end function 

function repeatValuesByAmounts(values, amounts) {
  return values.reduce(function(arr, next, index) {
    var i = amounts[index];
    while (i > 0) {
      arr.push(parseFloat(next));
      i -= 1;
    }
    return arr;
  }, []);
}

function CallToDownload(){
var totaldoc = workbookname.length
var fileService = azure.createFileService('summarylist','LAwgoa2sUctdxV+x4ar2yZpfnbfH9H4dRujhmyLlKJ5Rel2fO2E72syl+FITMAoyXgCtrSO0lNadACZzJGmpow=='); 
for (var i = 0; i < totaldoc; i++){
 // https://sungeo.scm.azurewebsites.net/dev/api/files/wwwroot/Summary_List/Summary%20List%20for%201891642800-SVMC.xlsx
      excellist += '<a  href="https://sungeo.scm.azurewebsites.net/dev/api/files/wwwroot/Summary_List/'+workbookname[i]+'" download>'+workbookname[i]+'</a><br>'
 
fileService.createShareIfNotExists('summarylistfile', function(error, result, response) {
  if (!error) {
    // if result = true, share was created.
    // if result = false, share already existed.
  }
});

fileService.createFileFromLocalFile('summarylistfile', '',workbookname[i] , '../Summary_List/'+workbookname[i], function(error, result, response) {
  if (!error) {
    // file uploaded
  }

  if (!error) {
   console.log("Completed Uploaded")
  }
  else{
    console.log(error);
  }
});

 
   
  } 
http.createServer(function(req, res){

    //res.writeHead(200, {'Content-Type':'text/plain'});
    //res.end('Hello World\n');
    res.writeHead(200, {'Content-Type':'text/html'});
    res.end('<!DOCTYPE html>\
<html>\
<body>\
'+excellist+'\
</body>\
</html>');

}).listen(port);    
 
}




/*
 router.get('/', function(req, res) {
  res.render('index', { title: 'Express',test: 'Node JS'});
 
});

*/


module.exports = router;
