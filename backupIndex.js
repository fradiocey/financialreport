var express = require('express');
var router = express.Router();
var async = require("async");
var SMB2 = require('smb2');

var fs = require('fs-extra')
const MFS = require("mfs");

/*// create an SMB2 instance 
var smb2Client = new SMB2({
  share:'\\\\mymswfs01\\'
, domain:'sunway.com'
, username:'mazim'
, password:'Fradiocey10'
, debug: false,
  autoCloseTimeout: 0
});

smb2Client.mkdir('180-Public\\BPMS\\Test', function (err) {
    if (err) throw err;
    console.log('Folder created!');
});
*/

 
/*fs.copy('smb:\\\\mymswfs01\\180-Public\\BPMS\\Book1.xlsx', '//mymswfs01/180-Public/Test/BPMS/sample.xlsx', function (err) {
  if (err) return console.error(err)
  console.log("success!")
}) // copies file */


/*function copyFile(source, target, cb) {
  var cbCalled = false;

  var rd = fs.createReadStream(source);
  rd.on("error", function(err) {
    done(err);
  });
  var wr = fs.createWriteStream(target);
  wr.on("error", function(err) {
    done(err);
  });
  wr.on("close", function(ex) {
    done();
  });
  rd.pipe(wr);

  function done(err) {
    if (!cbCalled) {
      cb(err);
      cbCalled = true;
    }
  }

 
/*fs.copy('/tmp/mydir', '/tmp/mynewdir', function (err) {
  if (err) return console.error(err)
  console.log('success!')
}) // copies directory, even if it has subdirectories or files*/

var sql = require("seriate");
var prodID = 5;
var PileDim = 880;
var operation = 0;

var excelbuilder = require('msexcel-builder-colorfix');
//var excelbuilder = require('msexcel-builder-colorfix-intfix');
//var excelbuilder = require('msexcel-builder');
//var excelbuilder = require('msexcel-builder-colorfix');
var workbookName = [];
var workbookID = [];
var pilediameter = [];

var totalexcel = 0;
var totalsheet = 0;
var completed = 0;

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
    "options" : {encrypt: true,
                "requestTimeout": 90000}  
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

 var workbook1 = "";
//get file name
sql.execute( { 
  query: "SELECT id, ProjectCode, ProjectName, ConcreteCover, SteelDiameter FROM BplProject " 
} ).then( function( result ) {
  totalexcel = result.length;
  //for (var i=0; i < 4; i++){
  //workbookID.push(result[i].ProjectCode);
  //workbookName.push(result[i].ProjectName);
  //getPile(result[i].id,result[i].ProjectCode,result[i].ProjectName);
  //}
  
  var y = 0;
  var loopWorkbook = function(result){
    var projCode = result[y].ProjectCode
     console.log ("RUN WORKBOOK"+y) 
    getPile(result[y].id,result[y].ProjectCode,result[y].ProjectName,result[y].ConcreteCover,result[y].SteelDiameter,function(){
    
      y++
      if (y < result.length){
      loopWorkbook(result);
      }
      
    })
  }
  //start loopWorkbook
  loopWorkbook(result) 

  
}, function( err ) {
        console.log( "Something bad happened:", err );
    } );

function getPile(id,workcode,workname,c,d,callback1){
workbook1 = excelbuilder.createWorkbook('./', 'Summary List for '+workcode+'-'+workname+'.xlsx');

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
        //console.log("Completed")
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
        console.log( "Something bad happened:", err );
    } );

}



// execute strdproc 
function getExcel(id,pileno,projCode,projName,c,d,callback){

sql.execute( {      
        query: "execute dbo.usp_SummaryList "+id+","+pileno+""
    } ).then( function( result ) {
 
  var totaldata = result.length+5;
  //for (var k=0; k < totalsheet; k++){
  
  var sheet1 = workbook1.createSheet(''+pileno+'', 30, totaldata)
  //}
    // header
  sheet1.set(2, 1, 'Summary List for '+projCode+'-'+projName+'' );
  //type
  sheet1.set(1, 3, 'Pile Details');
  sheet1.set(6, 3, 'General Details');
  sheet1.set(12, 3, 'Pile Details');
  sheet1.set(19, 3, 'Steel Cage');
  sheet1.set(25, 3, 'Concrete');
  // type center
  sheet1.align(1, 3, 'center');
  sheet1.align(6, 3, 'center');
  sheet1.align(12, 3, 'center');
  sheet1.align(19, 3, 'center');
  sheet1.align(25, 3, 'center');
  //merge table
  sheet1.merge({col:1,row:3},{col:4,row:3});
  sheet1.merge({col:6,row:3},{col:10,row:3});
  sheet1.merge({col:12,row:3},{col:17,row:3});
  sheet1.merge({col:19,row:3},{col:23,row:3});
  sheet1.merge({col:25,row:3},{col:30,row:3});
  // type border
  sheet1.border(1, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(2, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(3, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(4, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(1, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(2, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(3, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(4, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(6, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(7, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(8, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(9, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(10, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(6, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(7, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(8, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(9, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(10, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
  
  
  sheet1.border(12, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(13, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(14, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(15, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
   sheet1.border(16, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(17, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(12, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(13, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(14, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(15, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(16, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(17, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(19, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(20, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(21, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(22, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(23, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
 
  sheet1.border(19, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(20, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(21, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(22, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(23, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(25, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(26, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(27, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(28, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(29, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(30, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  
  sheet1.border(25, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(26, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(27, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(28, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(29, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(30, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
 

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
  sheet1.set(9, 4, 'Thickness'); // new
  sheet1.set(10, 4, 'Toe Level (mRL)');
   //gap
  sheet1.set(11, 4, '');
  //end gap
  sheet1.set(12, 4, 'Bored Depth PPL (m)');
  sheet1.set(13, 4, 'Pile Length (m)');
  sheet1.set(14, 4, 'Cavity (m)');
  sheet1.set(15, 4, 'Total Rock Coring (m)');
  sheet1.set(16, 4, 'Paid Rock Coring (m)');
  sheet1.set(17, 4, 'Rock Socket (m)');
    //gap
  sheet1.set(18, 4, '');
  //end gap
  sheet1.set(19, 4, 'Reinforcement Content');
  sheet1.set(20, 4, 'Main Steel Bar Weight'); // new requirement
  sheet1.set(21, 4, 'Helical/Spiral');
  sheet1.set(22, 4, 'Helical Link (KG)'); // new
  sheet1.set(23, 4, 'Cage Length');
     //gap
  sheet1.set(24, 4, '');
  //end gap
  sheet1.set(25, 4, 'Theoretical');
  sheet1.set(26, 4, 'Actual');
  sheet1.set(27, 4, 'Wastage (%)');
  sheet1.set(28, 4, 'Grade');
  sheet1.set(29, 4, 'DO Number');
  sheet1.set(30, 4, 'Concrete Volume');
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
sheet1.wrap(26, 4, 'true');
sheet1.wrap(27, 4, 'true');
sheet1.wrap(28, 4, 'true');
sheet1.wrap(29, 4, 'true');
sheet1.wrap(30, 4, 'true');

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
sheet1.align(26, 4, 'center');
sheet1.align(27, 4, 'center');
sheet1.align(28, 4, 'center');
sheet1.align(29, 4, 'center');
sheet1.align(30, 4, 'center');

for (var i = 5; i < totaldata; i++){
//parseFloat
var Platform = result[i-5].v5;
if (typeof Platform  != "undefined" || Platform != null){
var PlatformparseFloat = parseFloat(Platform).toFixed(3)
//console.log(PlatformparseFloat)
sheet1.set(6, i, PlatformparseFloat);
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
/* Fix Table
High Tensile */

//console.log( numlistNewLinedo )

var HighTenstile = 0;


// Main Bar
// split Reinforcement Content
var MainSteelBar = 0;
var FinalTotalLofHL  = 0;
var MultipleCageMainSteelBar = []

if (result[i-5].v14 != null){

var checkarray = JSON.stringify(result[i-5].v14);

if (checkarray.indexOf("(") == -1){ 

//("1 only")
var str = JSON.stringify(result[i-5].v14);
str = str.replace(/['"]+/g, '');
var MainBar = str.split("T");
var MainBar0 = MainBar[0];
var MainBar1 = MainBar[1];
var MainBar1Div1000 = parseFloat(MainBar1) / 1000;
var piledimeter = parseFloat(pileno);
var BeforexNosDiv = parseFloat(piledimeter*MainBar1Div1000)
var PileLength = parseFloat(result[i-5].v10).toFixed(3)
var BeforexNosPlus = parseFloat(PileLength) + parseFloat(BeforexNosDiv);
var BeforexNosTimes = parseFloat(MainBar0) * parseFloat(BeforexNosPlus);

if (MainBar[1] == "10"){
HighTenstile = 0.617;

}
else if (MainBar[1] == "12"){
HighTenstile = 0.888;

}
else if (MainBar[1] == "16"){
  HighTenstile = 1.579;
}
else if (MainBar[1] == "20"){
    HighTenstile = 2.466;
}
else if (MainBar[1]== "25"){
    HighTenstile = 3.854;
}
else if (MainBar[1] == "32"){
    HighTenstile = 6.313;
}
else if (MainBar[1] == "40"){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}


var finalCal = parseFloat(BeforexNosTimes) * parseFloat(HighTenstile);

MainSteelBar = parseFloat(finalCal).toFixed(3);

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
var newcagelength =  result[i-5].v16;
if (newcagelength != null)
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

for (var f = 0; f < mainbar0arr.length; f++ ){
var y = parseInt(mainbar0arr[f]) * parseInt(JSON.stringify( sums[f] ))
var current = parseFloat(y) * parseFloat(mainbar1arr[f])
output.push(current)

}



var sumt = output.reduce((a, b) => a + b, 0);
MainSteelBar = parseFloat(sumt).toFixed(4); // 6 // end result MainSteelBar



}//end more than 2

}// end MainBar



/* End of Fix Table */

/* Helical Link */
// start single HL

var checkarrayHL = result[i-5].v15;
//checkarrayHL = checkarrayHL.split(",");
if (result[i-5].v15 != null)
if (checkarrayHL.indexOf("(") == -1){ 

var PileLength = parseFloat(result[i-5].v10).toFixed(3);
var HelicalLink = JSON.stringify(result[i-5].v15);
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
/* Start Formula */
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
FinalTotalLofHL = TotalLofHL.toFixed(4)
// end of single HL

}
else{
console.log("Multiple HL")
var basiccalarr = [];
var mainbar2arr = [];
var mainbar1arr = [];
var mainbar0arr = [];
var HL1arr = [];
var HL2arr = [];
var HL3arr = [];
var HL3Rarr = [];
var HelicalChart = 0
// getcagelength
var newcagelength =  result[i-5].v16;
if (newcagelength != null)
newcagelength = newcagelength.split("+")


//get 16,20,16,20
var checkarray = result[i-5].v14;
checkarray = checkarray.split(",");

// helical Spiral
var hLspiral = result[i-5].v15;
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
console.log(MainBar2)
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

var h = parseFloat(newcagelength[e]) - f // 11.36

var I = ((h/(n[e]/1000))+1) *Math.sqrt((Math.pow(3.14159 * (pileno/1000 - c - c).toFixed(1), 2) +parseFloat( Math.pow((n[e]/1000),2)))) //186.7247
var J = ((I/h) * parseInt(d) * (q[e]/1000))
var H = I+J

//parseFloat(J).toFixed(4)



console.log(parseFloat(HL3Rarr[e]))

finalHL.push( H * parseFloat(HL3Rarr[e]))

}


console.log(finalHL)
var sumHL = finalHL.reduce((U, P) => U + P, 0);
var FinalTotalLofHL1 = parseFloat(sumHL).toFixed(3); // 6 // end result MainSteelBar
FinalTotalLofHL  = finalHL.reduce((a, b) => a + b, 0);
FinalTotalLofHL = FinalTotalLofHL.toFixed(4);

}




/* End of Helical Link */

   
sheet1.set(1, i, result[i-5].v1);
sheet1.set(2, i, result[i-5].v2);
sheet1.set(3, i, result[i-5].v3);
sheet1.set(4, i, result[i-5].v4);
sheet1.set(5, i, "");

sheet1.set(7, i, result[i-5].v6);
sheet1.set(8, i, result[i-5].v7);
sheet1.set(9, i, result[i-5].v23);//new
sheet1.set(10, i, parseFloat(result[i-5].v8).toFixed(3));

sheet1.set(11, i, "");
sheet1.set(12, i, parseFloat(result[i-5].v9).toFixed(3));
sheet1.set(13, i, parseFloat(result[i-5].v10).toFixed(3));
sheet1.set(14, i, result[i-5].v11);
sheet1.set(15, i, parseFloat(result[i-5].v12).toFixed(1));
sheet1.set(16, i, parseFloat(result[i-5].v24).toFixed(1)); //new paid rock coring
sheet1.set(17, i, parseFloat(result[i-5].v13).toFixed(1));

sheet1.set(18, i, "");

sheet1.set(19, i, result[i-5].v14);
sheet1.set(20, i, MainSteelBar);//new
sheet1.set(21, i, result[i-5].v15);
sheet1.set(22, i, FinalTotalLofHL);//new
sheet1.set(23, i, result[i-5].v16);
sheet1.set(24, i, "");

sheet1.set(25, i, parseFloat(result[i-5].v17).toFixed(1));
sheet1.set(26, i, parseFloat(result[i-5].v18).toFixed(1));
sheet1.set(27, i, parseFloat(result[i-5].v19).toFixed(1));
sheet1.set(28, i, result[i-5].v20);
sheet1.set(29, i, numlistNewLinedo);
sheet1.set(30, i, numlistNewLinecv);


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
sheet1.wrap(26, i, 'true');
sheet1.wrap(27, i, 'true');
sheet1.wrap(28, i, 'true');
sheet1.wrap(29, i, 'true');
sheet1.wrap(30, i, 'true');

// border
sheet1.border(1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(2, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(3, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(4, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(5, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(6, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(7, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(8, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(9, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(10, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(11, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(12, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(13, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(14, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(15, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(16, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(17, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(18, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(19, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(20, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(21, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(22, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(23, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(24, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(25, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(26, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(27, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(28, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(29, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(30, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});


//sheet1.numberFormat(2,1, 10); // equivalent to above
    
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
sheet1.width(10, '10');
sheet1.width(11, '5');
sheet1.width(12, '10');
sheet1.width(13, '10');
sheet1.width(14, '10');
sheet1.width(15, '10');
sheet1.width(16, '10');
sheet1.width(17, '10');
sheet1.width(18, '5');
sheet1.width(19, '10');
sheet1.width(20, '10');
sheet1.width(21, '10');
sheet1.width(22, '10');
sheet1.width(23, '10');
sheet1.width(24, '5');
sheet1.width(25, '10');
sheet1.width(26, '10');
sheet1.width(27, '10');
sheet1.width(28, '10');
sheet1.width(29, '15');
sheet1.width(30, '10');


sheet1.merge({col:1,row:1},{col:30,row:1});
callback();

 
    
    }, function( err ) {
        console.log( "Something bad happened:", err );
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



router.get('/', function(req, res) {
  res.render('index', { title: 'Express',test: 'Node JS'});
 
})

module.exports = router;
