
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
var newRCArr = [];
var newHLArr = [];
var uniqueRCArr = [];
var uniqueHLArr = [];
var newArrayMB = [];

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
var workbook1 = excelbuilder.createWorkbook('./Summary_List', 'Financial Report.xlsx');
var curr = 0;
function createWorkSheet(id,totalexcel,workcode,workname,c,d,callback1){
 
var sysdb = [];

var conarr = [];
var revarr = [];
var revJArr = [];
var steelArr = [];
var totalSteelArr = [];


    var sheetname = workcode +"-"+ workname;
    var sheet1 = workbook1.createSheet(sheetname, 100, 120);
     
    // Fill some data
    sheet1.set(1+1, 1+2, 'Date');
    //sheet1.set(2, 1, 'Date DB');
    sheet1.set(2+1, 1+2, 'Soil Value \n (RM)');
    sheet1.set(3+1, 1+2, 'Rock Value \n (RM)');
    sheet1.set(4+1, 1+2, 'Steel \n (RM)');
    sheet1.set(5+1, 1+2, 'Concrete \n (RM)');
    sheet1.set(6+1, 1+2, 'Movement \n (RM)');
    sheet1.set(7+1, 1+2, 'Total \n (RM)');


    // type center
  sheet1.align(1+1, 1+2, 'center');
  sheet1.align(2+1, 1+2, 'center');
  sheet1.align(3+1, 1+2, 'center');
  sheet1.align(4+1, 1+2, 'center');
  sheet1.align(5+1, 1+2, 'center');
  sheet1.align(6+1, 1+2, 'center');
  sheet1.align(7+1, 1+2, 'center');

  // type border
  sheet1.border(1+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(2+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(3+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(4+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(5+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(6+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});
  sheet1.border(7+1, 1+2, {left:'medium',top:'medium',right:'medium',bottom:'medium'});

  sheet1.width(1+1, '15');
  sheet1.width(2+1, '15');
  sheet1.width(3+1, '15');
  sheet1.width(4+1, '15');
  sheet1.width(5+1, '15');
  sheet1.width(6+1, '15');
  sheet1.width(7+1, '15');

    var today = new Date();
    var year = today.getFullYear();
    var month = today.getMonth();
    var date = today.getDate();
    var dayslists = [];


for(var i=0; i<30; i++){
var day=new Date(year, month - 1, date + i+1);    
var datefield = (moment(day).format('MMM DD YYYY'))
var dateconrev = (moment(day).format('YYYY-MM-DD'))
dayslists.push(datefield)

var b =0


 //get concrete & Revenue
sql.execute( { 
 query: "execute dbo.usp_CountRevenue "+id+",'"+dateconrev+"' "
} ).then( function( result ) {

b++

var newdate = result[0].DailyDate;
var newdatefilter = (moment(newdate).format('MMM DD YYYY'))

revJArr.push({"Date":newdatefilter,"Rev":result[0].SoilValue,"Rock":result[0].RockValue,"Con":result[0].ConcreteValue,"Mov":result[0].MovementPiles})


//revarr.push(result[0].RevenueValue)
//conarr.push(result[0].ConcreteValue)


if (b == 30){
// Save it 

//console.log(revJArr)

for (var g = 0; g < dayslists.length; g++){  


    sheet1.set(2+1, g+2+2, 0);
    sheet1.set(3+1, g+2+2, 0);
   
    sheet1.set(5+1, g+2+2, 0);
    sheet1.set(6+1, g+2+2, 0);
   



    sheet1.set(1+1, g+2+2, dayslists[g]);
for (var q = 0; q < 30; q++){ 

  if (revJArr[q].Date == dayslists[g] ){
    sheet1.set(2+1, g+2+2, parseFloat(revJArr[q].Rev).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
    sheet1.set(6, g+2+2, parseFloat(revJArr[q].Con).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
sheet1.set(4, g+2+2, parseFloat(revJArr[q].Rock).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
sheet1.set(7, g+2+2, parseFloat(revJArr[q].Mov).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));


sheet1.align(3, g+2+2, 'right');
sheet1.align(6, g+2+2, 'right');
sheet1.align(4, g+2+2, 'right');
sheet1.align(7, g+2+2, 'right');


  }
    //sheet1.set(2, g+2, datedb[h]);
 
  }
 
}



// steelArr
for (var u = 0; u < dayslists.length; u++){
     // type border
  sheet1.border(1+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(2+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(3+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(4+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(5+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(6+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sheet1.border(7+1, u+2+2, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
  sysdb.push(dayslists[u])
  }



// Steel

var startdate = moment( moment().subtract(29, 'days') ).format("YYYY-MM-DD")
var enddate = moment().format("YYYY-MM-DD")

var steelratecal = 0;
//get steelrate
sql.execute( { 
 query: "select DISTINCT steelrate from BplPjtFinancialRate where Project_Id = "+id+""
} ).then( function( result ) {

steelratecal = result[0].steelrate;

})


// get date 
sql.execute( { 
 //query: "select b.id, b.pileno, b.CageInstallStartTimeStamp, b.TotalCageLength,b.PileLength, b.PileDiameter, b.mainreinforcementcontent, b.helicalreinforcementcontent from bplpile as b where project_id = 5 and CAST(b.CageInstallStartTimeStamp as date) BETWEEN  '2016-11-16' AND '2016-12-15' and (MainReinforcementContent is not null or helicalreinforcementcontent is not null) order by b.CageInstallStartTimeStamp"

  query: "select b.id, b.pileno, b.CageInstallStartTimeStamp, b.TotalCageLength, b.PileLength, b.PileDiameter,b.mainreinforcementcontent, b.helicalreinforcementcontent, CAST(b.CageInstallStartTimeStamp as date) as cdate from bplpile as b where project_id = "+id+" and CAST(b.CageInstallStartTimeStamp as date) BETWEEN  '"+startdate+"' AND '"+enddate+"' and (MainReinforcementContent is not null or helicalreinforcementcontent is not null) order by b.CageInstallStartTimeStamp"
} ).then( function( result ) {


// get result
// date

for (var t = 0; t < result.length; t++){


var str1 = result[t].cdate;

var str = str1.toString()
var datares = str.slice(4,15);

//pile total cagelength 12+12+8.39
var TotalCageLength = result[t].TotalCageLength

//pile length
var PileLength = result[t].PileLength

//mainbar
var mainreinforcementcontent = result[t].mainreinforcementcontent


//helicallink
var helicalreinforcementcontent = result[t].helicalreinforcementcontent

//PileDiameter
var PileDiameter = result[t].PileDiameter

//PileNo
var pileno = result[t].pileno


var newRC =  mainreinforcementcontent ;
var newHL =  helicalreinforcementcontent

steelArr.push({"PileNo":pileno,"Date":datares,"MB":newRC,"HL":newHL,"PileD":PileDiameter,"cageL":TotalCageLength,"PileL":PileLength})


if (t+1 ==result.length){

for (var i = 0; i < result.length; i++){
var mb = steelArr[i].MB
var hl = steelArr[i].HL
var cageL  = steelArr[i].cageL
var date =  steelArr[i].Date
var pileno = steelArr[i].PileNo
var pileD = steelArr[i].PileD


if (mb != null){
if (mb.indexOf("na") == -1){ //na
if (mb.indexOf("TEST") == -1){ //TEST
if (mb.indexOf("+") == -1){ //single HL
if (cageL != null){ //single HL  


if (mb.indexOf("(") == -1){//single

var newRC = mb;
newRC = newRC.replace(/\s/g, "") 
newRC =  newRC.split("T");
var MainBar0 = newRC[0];
newRC = newRC[1];
newRCArr.push("T"+newRC)


// MB Table
if (newRC  == "10"){
HighTenstile = 0.617;

}
else if (newRC  == "12"){
HighTenstile = 0.888;

}
else if (newRC  == "16"){
  HighTenstile = 1.579;
}
else if (newRC == "20"){
    HighTenstile = 2.466;
}
else if (newRC == "25"){
    HighTenstile = 3.854;
}
else if (newRC  == "32"){
    HighTenstile = 6.313;
}
else if (newRC  == "40"){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}

if (HighTenstile == 0){
  if (newRC  == 10){
HighTenstile = 0.617;

}
else if (newRC  == 12){
HighTenstile = 0.888;

}
else if (newRC  == 16){
  HighTenstile = 1.579;
}
else if (newRC == 20){
    HighTenstile = 2.466;
}
else if (newRC == 25){
    HighTenstile = 3.854;
}
else if (newRC  == 32){
    HighTenstile = 6.313;
}
else if (newRC  == 40){
    HighTenstile = 9.870;
}
else{
    HighTenstile = 0;
}
}


var MBSstep1 = parseFloat(MainBar0) * parseFloat(cageL)
var MBSstep2 = (parseFloat(d)* parseFloat(newRC)) / 1000 //starter bar
var MBSstep3 = parseFloat(MBSstep1) * parseFloat(HighTenstile)

//totalSteelArr.push({"pileno":pileno,"date":date,"val":MBSstep3})


// HL Single



var checkarrayHL = hl;

var checkcagelength = JSON.stringify(cageL);

//checkarrayHL = checkarrayHL.split(",");
if (hl!= null){
if (checkarrayHL.indexOf("(") == -1){ 
if (cageL != null){
if (checkcagelength.indexOf ("+") == -1){



var cagelength = cageL;
var HelicalLink = hl;
HelicalLink  = HelicalLink.replace(/ /g,'')// remove whitespace
HelicalLink = HelicalLink.replace(/['"]+/g, '');
//HelicalLink = HelicalLink.replace('T', "");
HelicalLink  = HelicalLink.split("-");
var HelicalLink1 = HelicalLink[1]



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
var D = parseFloat(pileD) / 1000;
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


var step1 = (d * parseFloat(newRC)/1000);
var step2 = (cagelength - step1);
var step3 = ((step2/(HelicalLink1/1000))+1) *Math.sqrt((Math.pow(3.14159 * (pileD/1000 - c - c).toFixed(1), 2) +parseFloat( Math.pow((HelicalLink1/1000),2)))) //186.7247

var totalmbhl = parseFloat(MBSstep3) + parseFloat (step3)

totalSteelArr.push({"pileno":pileno,"date":date,"val":totalmbhl})




}
}
}
}





}
else{//multiple
//console.log(steelArr[i].MB +"-"+ steelArr[i].Date)
//console.log("Multiple:"+steelArr[i].Date +"-"+ steelArr[i].MB)


var HighTenstile = 0;

var mb = steelArr[i].MB
var cageL  = steelArr[i].cageL
var date =  steelArr[i].Date
var pileno = steelArr[i].PileNo


var checkarray = mb;
checkarray = checkarray.split(",");


// must get value 16 & 20
var mainbar2arr = [];
var mainbar0arr = [];
var mainbar1arr = [];
var mainbar3arr = [];
var newcagelength =  cageL;
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

var output = [];
var currentArr = [];

for (var f = 0; f < mainbar0arr.length; f++ ){
var y = parseInt(mainbar0arr[f]) * parseInt(JSON.stringify( sums[f] ))
var current = parseFloat(y) * parseFloat(mainbar1arr[f])
output.push(current)

}


var sumt = output.reduce((a, b) => a + b, 0);





// Multiple HL



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
var newcagelength =  cageL;
if (newcagelength != null){
if (newcagelength.indexOf ("+") != -1){
//console.log("Found")
newcagelength = newcagelength.split("+")


//get 16,20,16,20
var checkarray = mb;
checkarray = checkarray.split(",");


// helical Spiral
var hLspiral = hl;
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
//console.log("HL ERROR"+HLstr)
var HL1 = HL0.replace(/ *\([^)]*\) */g, ""); 
//var HL1 = HL0
HL1arr.push(HL1)
var HL2 = HL0.match(/\((.*?)\)/)[1] 
//var HL2 = HL0
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

var I = ((h/(n[e]/1000))+1) *Math.sqrt((Math.pow(3.14159 * (pileD/1000 - c - c).toFixed(1), 2) +parseFloat( Math.pow((n[e]/1000),2)))) //186.7247
var J = ((I/h) * parseInt(d) * (q[e]/1000))
var H = I+J



//parseFloat(J).toFixed(4)

finalHL.push((I/h))


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
var totaldoublembhl =  parseFloat(sumt) +parseFloat (finalArr)

totalSteelArr.push({"pileno":pileno,"date":date,"val":totaldoublembhl})




//totalSteelArr.push({"pileno":pileno,"date":date,"val":sumt})

}
}
}
}




}



}//cageL not null
else{
  //console.log(steelArr[i]) // you cannot get this val because of cagel = null
}
}// not +
}// not test
}// not na
}// not null



}



var temp = {};
var obj = null;
for(var i=0; i < totalSteelArr.length; i++) {
   obj=totalSteelArr[i];

   if(!temp[obj.date]) {
       temp[obj.date] = obj;
   } else {
       temp[obj.date].val += obj.val;
   }
}
var result1 = [];
for (var prop in temp){
    result1.push(temp[prop]);
}



var totalall = [];
for (var g = 0; g < dayslists.length; g++){    
  sheet1.set(4+1, g+2+2, 0);
   
for(var b = 0; b < result1.length; b++){



if (dayslists[g] == result1[b].date){


var convert = parseFloat(result1[b].val).toFixed(2) * steelratecal
sheet1.set(5, g+2+2, convert.toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
totalall.push({date:result1[b].date,steel:convert})
sheet1.align(5, g+2+2, 'right');




}



}
}


}
// end of main bar
// helical link



}// end loop


console.log(totalall)
var totalsum = [];



for (var f = 0; f <revJArr.length; f++){
    for (var u = 0; u < totalall.length; u++){

        
        if (revJArr[f].Date == totalall[u].date){
        
         var sumall = parseFloat(revJArr[f].Rev) + parseFloat(revJArr[f].Rock) + parseFloat(totalall[u].steel) + parseFloat(revJArr[f].Con) + parseFloat(revJArr[f].Mov)
         totalsum.push({date:revJArr[f].Date,val:sumall})
        }
        else{
             //console.log(revJArr[f].Date == totalall[u].date)
        //totalsum.push({date:revJArr[f].Date,rev:revJArr[f].Rev,steel:0,con:revJArr[f].Con})
        }
    }
}






var totalsumnosteel = [];
for (var u = 0; u < dayslists.length; u++){
     for (var f = 0; f <revJArr.length; f++){
 
        if (dayslists[u] == revJArr[f].Date){
            
         var sumallnosteel = parseFloat(revJArr[f].Rev) + parseFloat(revJArr[f].Rock) + parseFloat(revJArr[f].Con) + parseFloat(revJArr[f].Mov)
         totalsumnosteel.push({date:revJArr[f].Date,val:sumallnosteel})
        
    
        }
        else{
       
             //console.log(revJArr[f].Date == totalall[u].date)
        //totalsum.push({date:revJArr[f].Date,rev:revJArr[f].Rev,steel:0,con:revJArr[f].Con})
        }
    }
}





for (var g = 0; g < dayslists.length; g++){   
     sheet1.set(7+1, g+2+2, 0);
for (var w = 0; w < totalsumnosteel.length; w++){

if (dayslists[g] == totalsumnosteel[w].date){
sheet1.set(8, g+2+2, (totalsumnosteel[w].val).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
sheet1.align(8, g+2+2, 'right');
}



}
}







for (var g = 0; g < dayslists.length; g++){ 
for (var w = 0; w < totalsum.length; w++){

if (dayslists[g] == totalsum[w].date){
sheet1.set(8, g+2+2, (totalsum[w].val).toFixed(0).replace(/./g, function(c, i, a) {
    return i && c !== "." && ((a.length - i) % 3 === 0) ? ',' + c : c;
}));
sheet1.align(8, g+2+2, 'right');

workbook1.save(function (err) {
        if (err)
        throw err;
      else
      console.log('congratulations, your workbook created');

  }); 
}


}
}
})

// end of steel









}// end for loop if complete 30

});




}


callback1();


}


// main
sql.execute( { 
 query: "SELECT id, ProjectCode, ProjectName, ConcreteCover, SteelDiameter FROM BplProject where id in (5,7,8)" 
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

module.exports = router;
