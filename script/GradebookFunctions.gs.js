/**
 * A set of functions for Google Apps Script that automates the tedious parts of setting up a 
 * 
 * My hope is that this set of gradebook tools is useful to as many as possible. Please cite as XXX in any publication.
 * 
 * BSD Zero Clause License:
 * Copyright (C) 2024 by Kyle Broaders (broaders@mtholyoke.edu)
 * Permission to use, copy, modify, and/or distribute this software for any purpose with or without fee is hereby granted.
 * 
 * THE SOFTWARE IS PROVIDED ""AS IS"" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS SOFTWARE INCLUDING
 * ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL,
 * DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
 * WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE
 * OR PERFORMANCE OF THIS SOFTWARE.
 * 
 */

//--------------------------------------------------------------------------------
// Global Variables that lazy load via: 
// https://stackoverflow.com/questions/70056310/avoid-repeating-use-global-variables-in-google-apps-script-or-not
const g = {};//global object
const addGetter_ = (name, value, obj = g) => {
  Object.defineProperty(obj, name, {
    enumerable: true,
    configurable: true,
    get() {
      delete this[name];
      return (this[name] = value());
    },
  });
  return obj;
};

//MY GLOBAL VARIABLES in g
[
  ['ss', () => SpreadsheetApp.getActive()],
  ['rosterSheet', () => g.ss.getSheetByName("Roster")],
  ['AllGrades', () => g.ss.getSheetByName("AllGrades")],
  ['settings', () => g.ss.getSheetByName("settings")],
  ['studentNames', () => g.rosterSheet.getRange(g.settings.getRange("C2").getValue()).getValues().filter(String).flat()],
  ['studentEmails', () => g.rosterSheet.getRange(g.settings.getRange("C3").getValue()).getValues().filter(String).flat()],
  ['numStudents', () => g.studentNames.length],
  ['gradeRange', () => g.settings.getRange("C4").getValue()],
  ['studentNameCell', () => g.settings.getRange("C5").getValue()],
  ['linkCol', () => g.settings.getRange("C6").getValue()]
].forEach(([n, v]) => addGetter_(n, v));

// -----------------------------------------------------------------------------
// SET UP SHEETS

/**
 * Duplicates the `Template` sheet for each student listed in the `Roster` sheet
 * Formulas within each sheet will pull the appropriate data from 'AllGrades'
 * These formulas require the student name, so it is pasted into studentNameCell
 * 
 * Makes use of information in global variable g:
 * - ss
 * - studentNames
 * - studentNameCell
 */

function makeStudentSheets(){
  ss = g.ss;
  names = g.studentNames;

  ss.setActiveSheet(ss.getSheetByName('Template'), true);

  for (let thisStudent of g.studentNames){
    // Checks if sheet already exists before making a new one
    if(!ss.getSheetByName(thisStudent)){
      ss.duplicateActiveSheet();
      ss.getActiveSheet().setName(thisStudent);
      ss.getActiveSheet().getRange(g.studentNameCell).setValue(thisStudent);
      Logger.log("Made sheet for "+thisStudent);
    }
  }
}

/**
 * Makes a copy for each student of a formatted but blank template spreadsheet in the destnation folder.
 * Each sheet imports values from the main gradebook using importRange.
 * 
 * A link to the new spreadsheet is added to linkCol in the `Roster` sheet
 * 
 * Makes use of information in global variable g:
 * - ss
 * - rosterSheet
 * - templateID
 * - destinationID
 * - studentNames
 * - linkCol
 * - studentEmails
 */
function makeIndividualSpreadsheets(){

  try{
    var rosterID  = g.ss.getId();
    var template = DriveApp.getFileById(g.templateID);
    var destFolder = DriveApp.getFolderById(g.destinationID);
  }catch (error) {
    Logger.log(`ID values not set correctly in settings tab: ${error.toString()}`);
  }

  let names = g.studentNames;

  for (var i=0; i<g.numStudents; i++) {
    //Only make a new sheet if it hasn't been done before
    if(isLinked(names[i])){
      Logger.log("Continuing on "+names[i]);
      continue;
    }

    var copyFile = template.makeCopy(names[i],destFolder);

    // Put a link to the individual sheet in the `Roster` sheet
    var studentRow = i+2;
    var linkCell = g.linkCol+studentRow
    g.rosterSheet.getRange(linkCell).setValue(copyFile.getUrl());
    Logger.log("Linked "+names[i]+" in "+linkCell);

    // Paste import formula into the sheet with the appropriate student info
    var copiedSS = SpreadsheetApp.open(copyFile);
    copiedSS.getSheets()[0].getRange("A1").setValue("=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/"+rosterID+"\",\"'"+names[i]+"'!A1:J50\")");
    Logger.log("New SS for "+names[i]);

    // Adjust permissions so that the import works and only the right student can view the sheet
    var studentSheetID = copiedSS.getId();
    addImportrangePermission(studentSheetID,rosterID);
    addViewPermission(g.studentEmails[i],studentSheetID)
  }
}

// -----------------------------------------------------------------------------
// ADJUST EXISTING SHEETS

/**
 * Helper function to run writeToCells from within Apps Script interface
 * 
 * Replaces the content in 'cell' with 'value' in every student sheet
 */
function writeValueToAllSheets(){
  var cell = 'CELL ADDRESS';
  var value = "REPLACE WITH NEW CONTENT";
  writeToCells(cell, value);
}

/**
 * Helper function to run copyRangeToStudents from within Apps Script interface
 * 
 * Copies contents of the template sheet in `theRange` to every student sheet
 * Used to update or fix a section of the student sheets without having to remake them
 */

function copyFromTemplateToStudents(){
  var theRange = "D3:D34";
  copyRangeToStudents(theRange);
}

// -----------------------------------------------------------------------------
// UTILITY FUNCTIONS — shouldn't need to execute directly

/**
 * Replaces the contents of theCell with theValue in every student sheet
 * Makes use of information in global variable g:
 * - ss
 * - studentNames
 */
function writeToCells(theCell, theValue){
  var ss = g.ss;
  for (let thisStudent of g.studentNames){
    var thisSheet = ss.getSheetByName(thisStudent);
    thisSheet.getRange(theCell).setValue(theValue);
  }
}

/**
 * Copies the contents of copyRange in the `Templates` sheet and pastes it to each student sheet
 * 
 * Makes use of information in global variable g:
 * - ss
 * - studentNames
 */
function copyRangeToStudents(copyRange){
  var ss = g.ss;
  var valuesRange = ss.getSheetByName('Template').getRange(copyRange);

  for (let thisStudent of g.studentNames){
    var thisSheet = ss.getSheetByName(thisStudent);
    valuesRange.copyTo(thisSheet.getRange(copyRange));
  }
  
}

/**
 * Checks if studentName already has a link in the linkColumn of the `Roster` sheet
 * 
 * Makes use of information in global variable g:
 * - rosterSheet
 * - studentNames
 * - linkCol
 */
function isLinked(studentName){
  const studentRow = g.studentNames.indexOf(studentName)+2
  if (studentRow < 0){
    throw new Error("Can't check if student is linked because they were not found in the Roster");
  }
  return !g.rosterSheet.getRange(g.linkCol+studentRow).isBlank();
}

/**
 * Gives permission to importRange in an individual student sheet without having to open the sheet or click "Allow Access"
 * 
 * Taken from https://stackoverflow.com/questions/28038768/how-to-allow-access-for-importrange-function-via-apps-script
 * 
 * Makes use of undocumented functionality but has worked 2020-2024.
 */

function addImportrangePermission(importSheetID,donorSheetID) {
  
  // add permission by fetching this url
  const url ="https://docs.google.com/spreadsheets/d/"+importSheetID+"/externaldata/addimportrangepermissions?donorDocId="+donorSheetID;

  const token = ScriptApp.getOAuthToken();

  const params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, params);
}

/**
 * Updates student spreadsheet to give them read access without emailing them
 * 
 * Found resource at https://www.pbainbridge.co.uk/2020/04/drive-api-share-file-without-email.html
 * 
 * Requires the Drive API v2 to be activated. Click "Services" in the left sidebar, then add Drive API version v2 with the name "Drive"
 */
function addViewPermission(thisUser,thisSheetID) {

  var resource = {
    // user email address or domain name
    value: thisUser, 
    
    // Options are "user", "group", "domain", "anyone" or "default"
    type: 'user',                
    
    // Options are: "owner", "organizer", "fileOrganizer", "writer" or "reader"
    role: 'reader'               
  };
  
  var optionalArgs = {sendNotificationEmails: false};
  
  Drive.Permissions.insert(resource, thisSheetID, optionalArgs);
}

/**
 * Takes a column number and returns a string with a spreadsheet-style letter (zero indexed).
 * 
 * Example: 0 -> "A", 25 -> "Z", 26 -> "AA", 29 -> "AD"
 */

function getColAlpha(colNum){
	var letters='ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');  // Makes an array of letters
  var lastLetter = letters[colNum % letters.length];   // Gets the base26 final digit as a letter

  // Recursively calls getColAlpha until it can return just 1 letter then combines outcomes for each previous call 
  if (Math.trunc(colNum / letters.length) >= 1) {
    var otherLetters = getColAlpha(Math.trunc(colNum / letters.length)-1);
    return otherLetters+lastLetter; 
	}

  return lastLetter;
}

/**
 * Returns the name of the current sheet. For use as a spreadsheet formula.
 */
function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}


/**
 * Returns the ID of the current sheet. For use as a spreadsheet formula.
 */
function thisSheetID() {
  return g.ss.getId();
}