//This is the most updated version as of April 5th, 2022

//These variables activate the selected cell.
var ss = SpreadsheetApp.getActiveSheet();
var activeCell = ss.getActiveCell()
var valueActiveCell = activeCell.getValue();

//These variables allow the purchase to be shown in the cell to the right of the selected cell.
//The only time the range and offset change are for updating points, whole group rewards, and applying interest functions.  
var newRange = activeCell.offset(0, 1, 1);
var nextVal = newRange.getValue();

//These functions correspond with the main buttons for subtracting points
function rewardA() {
  var A = "A ";
 
  activeCell.setValue(valueActiveCell - 30);
  
  newRange.setValue(nextVal + A);
}

function rewardB() {
  var B = "B ";
 
  activeCell.setValue(valueActiveCell - 40);
  
  newRange.setValue(nextVal + B);
}

function rewardC() {
  var C = "C ";
 
  activeCell.setValue(valueActiveCell - 50);
  
  newRange.setValue(nextVal + C);
}

function rewardD() {
  var D = "D ";
 
  activeCell.setValue(valueActiveCell - 60);
  
  newRange.setValue(nextVal + D);
}

function rewardE() {
 var nextVal = newRange.getValue();
  var E = "E ";
 
  activeCell.setValue(valueActiveCell - 70);
  
  newRange.setValue(nextVal + E);
}

function rewardF() {
  var F = "F ";
 
  activeCell.setValue(valueActiveCell - 80);
  
  newRange.setValue(nextVal + F);
}

function rewardG() {
  var nextVal = newRange.getValue();
  var G = "G ";
 
  activeCell.setValue(valueActiveCell - 100);
  
  newRange.setValue(nextVal + G);
}

function rewardH() {
  var H = "H ";
 
  activeCell.setValue(valueActiveCell - 120);
  
  newRange.setValue(nextVal + H);
}

function rewardI() {
  var I = "I ";
 
  activeCell.setValue(valueActiveCell - 130);
  
  newRange.setValue(nextVal + I);
}

function rewardJ() {
  var J = "J ";
 
  activeCell.setValue(valueActiveCell - 150);
  
  newRange.setValue(nextVal + J);
}

function rewardK() {
  var K = "K ";
 
  activeCell.setValue(valueActiveCell - 200);
  
  newRange.setValue(nextVal + K);
}

function rewardL() {
  var L = "L ";
 
  activeCell.setValue(valueActiveCell - 300);
  
  newRange.setValue(nextVal + L);
}

function rewardM() {
  var M = "M ";
 
  activeCell.setValue(valueActiveCell - 500);
  
  newRange.setValue(nextVal + M);
}

//WGR stands for Whole Group Rewards
function WGR25() {
  var newRange = ss.getRange('F20');
  var nextVal = newRange.getValue();
 
  activeCell.setValue(valueActiveCell - 25);
  
  newRange.setValue(nextVal + 25);
}

function WGR50() {
  var newRange = ss.getRange('F20');
  var nextVal = newRange.getValue();
 
  activeCell.setValue(valueActiveCell - 50);
  
  newRange.setValue(nextVal + 50);
}

function WGR100() {
  var newRange = ss.getRange('F20');
  var nextVal = newRange.getValue();
 
  activeCell.setValue(valueActiveCell - 100);
  
  newRange.setValue(nextVal + 100);
}

function WGR500() {
  var newRange = ss.getRange('F20');
  var nextVal = newRange.getValue();
 
  activeCell.setValue(valueActiveCell - 500);
  
  newRange.setValue(nextVal + 500);
}

//this part just sets up the updating points function
function uploadPoints() {
  var ss = SpreadsheetApp.getActiveSheet();
  var activeCell = ss.getActiveCell()
  var valueActiveCell = activeCell.getValue();
  
  var positiveNew = activeCell.offset(0, 7, 1);
  var pVal = positiveNew.getValue();
  
  var negativeNew = activeCell.offset(0, 8, 1);
  var nVal = negativeNew.getValue();
  
  var newRange = activeCell.offset(1, 0, 1);
  var next = newRange.activateAsCurrentCell();
  var nextVal = next.getValue();
  
 activeCell.setValue(valueActiveCell + pVal - nVal).protect();
 next.setValue(nextVal);
  
}

//assign this script for updating points
function runUploadPoints() {
  var spreadsheet = SpreadsheetApp.getActive();
for (var i = 0; i<25; i++) {
  uploadPoints();
  spreadsheet.getCurrentCell().activate();
}
}

//this part just sets up the interest value for 10%
function newInterest() {
  var ss = SpreadsheetApp.getActiveSheet();
  var activeCell = ss.getActiveCell()
  var valueActiveCell = activeCell.getValue();
  var interest = valueActiveCell * (1/10);
 
  var newRange = activeCell.offset(1, 0, 1);
  var next = newRange.activateAsCurrentCell();
  var nextVal = next.getValue();

    activeCell.setValue(valueActiveCell + interest);
    next.setValue(nextVal);

}

//assign this script for applying interest
function runNewInterest() {
  var spreadsheet = SpreadsheetApp.getActive();
for (var i = 0; i<25; i++) {
  newInterest();
  spreadsheet.getCurrentCell().activate();
}
}

//these are for a possible update in the future. I currently charge 5 Dojo Points to use the restroom and have a linked Google Form that reflects the responses in my personal Dojo Bank in real time. I run the restroom script once a week and it subtracts points from the students' totals.
function restRoom() {
  var ss = SpreadsheetApp.getActiveSheet();
  var activeCell = ss.getActiveCell()
  var valueActiveCell = activeCell.getValue();
  
  var restroomNew = activeCell.offset(0, 9, 1);
  var GetrVal = restroomNew.getValue();
  var rVal = GetrVal * 5;
  
  var newRange = activeCell.offset(1, 0, 1);
  var next = newRange.activateAsCurrentCell();
  var nextVal = next.getValue();
  
 activeCell.setValue(valueActiveCell - rVal).protect();
 next.setValue(nextVal);
  
}

function runRestRoom() {
  var spreadsheet = SpreadsheetApp.getActive();
for (var i = 0; i<25; i++) {
  restRoom();
  spreadsheet.getCurrentCell().activate();
}
}