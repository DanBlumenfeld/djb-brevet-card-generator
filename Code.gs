/*****************************************************************************/
/* Original concept, and bulk of the field-based merge code, taken from here: https://github.com/hadaf/SheetsToDocsMerge */

/* 
 *  Make sure the first row of the source data spreadsheet is column headers.
 *
 *  Reference the column headers in the template by enclosing the header in square brackets.
 *  Example: "RUSA Member #:  [RUSA Number]"
 */
/*****************************************************************************/


function generateCardFronts() {
  
  /*****************************************************************************/
  /*Stuff to edit only if the template or source sheet changes                 */
  /*****************************************************************************/
  var selectedTemplateId = "1qfR-l2BCHVaOm6KSihYEFevtd7Wmc67r8ZcRGeypKXQ";//Copy and paste the ID of the template document here (you can find this in the document's URL)
  var sheetName = "Submissions";
  //In the tokens, make sure to escape square braces with two \\
  var tokenOrg = "\\[Org\\]";
  var tokenRideType = "\\[Ride Type\\]";
  var tokenDistance = "\\[Distance\\]";
  var tokenDate = "\\[Date\\]";
  //Do we need a page break after each copy?
  var addPageBreak = false;    
  
  /*****************************************************************************/
  /*Stuff to edit for each event                                               */
  /*****************************************************************************/
  var org = "RUSA"; //Should be RUSA or ACP...the sponsoring body
  var rideType = "Populaire"; //should be Populaire or Brevet or (maybe) Grand Randonnée
  var distance = "100"; //Should be distance in KM
  var brevetDate = "08 March 2015"; //Should be in the format DD Month YYYY, such as 08 March 2015
  var outputFileName = "NorthHillsPopulaire_BrevetCardFront"; //What do you want to name the newly created brevet card fronts doc?  
    
  
  /*****************************************************************************/
  /*Don't edit anything below unless you wish to change the script logic.      */
  /*****************************************************************************/
  var templateFile = DriveApp.getFileById(selectedTemplateId);
  var mergedFile = templateFile.makeCopy();//make a copy of the template file to use for the merged File. Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file, otherwise, if the destination file and the template file are separate, a Google bug will prevent copying of images from the template to the destination. See the description of the bug here: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14
  mergedFile.setName(outputFileName);
  var mergedDoc = DocumentApp.openById(mergedFile.getId());
  var bodyElement = mergedDoc.getBody();//the body of the merged document, which is at this point the same as the template doc.
  var bodyCopy = bodyElement.copy();//make a copy of the body
  
  bodyElement.clear();//clear the body of the mergedDoc so that we can write the new data in it.
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];//First row of the sheet must be the the field names

  for (var i = 1; i < numRows; i++) {//data values start from the second row of the sheet 
    var row = values[i];
    var body = bodyCopy.copy();   
    
    //Populate from spreadsheet data
    for (var f = 0; f < fieldNames.length; f++) {
      var currFieldName = fieldNames[f];
      currFieldName = currFieldName.replace("(", "\\(").replace(")", "\\)"); //Escape parentheses if necessary
      body.replaceText("\\[" + currFieldName + "\\]", row[f]);//replace [fieldName] with the respective data value. 
    }
    
     //Populate from ride-specific values
    body.replaceText(tokenOrg, org);
    body.replaceText(tokenRideType, rideType);
    body.replaceText(tokenDistance, distance);
    body.replaceText(tokenDate, brevetDate);
    
    var numChildren = body.getNumChildren();//number of the contents in the template doc
   
    for (var c = 0; c < numChildren; c++) {//Go over all the content of the template doc, and replicate it for each row of the data.
      var child = body.getChild(c);
      child = child.copy();
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child.getBlob());
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        mergedDoc.appendTable(child);
      } else {
        Logger.log("Unknown element type: " + child);
      }
   }
        
   if(addPageBreak == true)
     mergedDoc.appendPageBreak();

  }
}

function generateCardBacks() {
  /*****************************************************************************/
  /*Stuff to edit only if the template changes                                 */
  /*****************************************************************************/
  var selectedTemplateId = "1oQnFMXm9GGdljgPWFnS11GHF0kHFHVcUohezj60W234";//Copy and paste the ID of the template document here (you can find this in the document's URL)
  var sheetName = "Controles";
  //In the tokens, make sure to escape square braces with two \\, and use # for the number
  var tokenName = "\\[Controle Name #\\]";
  var tokenAddress = "\\[Controle Address #\\]";
  var tokenTimes = "\\[Controle Times #\\]";
  var tokenDistance = "\\[Controle Distance #\\]";
  var tokenStartTime = "\\[start\\]";
  //Do we need a page break after each copy?
  var addPageBreak = false;    
  
  /*****************************************************************************/
  /*Stuff to edit for each event                                               */
  /*****************************************************************************/
  var startTime = "09:00"; //Should be 24-hour time
  var outputFileName = "NorthHillsPopulaire_BrevetCardBack"; //What do you want to name the newly created brevet card backs doc?
  
  
  /*****************************************************************************/
  /*Don't edit anything below unless you wish to change the script logic.      */
  /*****************************************************************************/
  var templateFile = DriveApp.getFileById(selectedTemplateId);
  var mergedFile = templateFile.makeCopy();//make a copy of the template file to use for the merged File. Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file, otherwise, if the destination file and the template file are separate, a Google bug will prevent copying of images from the template to the destination. See the description of the bug here: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14
  mergedFile.setName(outputFileName);
  var mergedDoc = DocumentApp.openById(mergedFile.getId());
  var bodyElement = mergedDoc.getBody();//the body of the merged document, which is at this point the same as the template doc.
  var bodyCopy = bodyElement.copy();//make a copy of the body
  
  bodyElement.clear();//clear the body of the mergedDoc so that we can write the new data in it.
  
  //TODO: get the number of rows in Submissions sheet, use it to determine how many backs are needed
  var numBacks = 16;
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);  
  var rows = sheet.getDataRange();
  var values = rows.getValues();  
  
  for(var j=0; j < numBacks; j++) {
    var body = bodyCopy.copy();
  
    
    //Write the start time
    body.replaceText(tokenStartTime, startTime);
  
    for (var i = 1; i <= 10; i++) {//data values start from the second row of the sheet, there are ten data rows
     var row = values[i];
     var tokenToReplace;
      
     //Name
     tokenToReplace = tokenName.replace("#", i);
     body.replaceText(tokenToReplace, row[0]);
     //Address
     tokenToReplace = tokenAddress.replace("#", i);
     body.replaceText(tokenToReplace, row[1]);     
     //Times
     tokenToReplace = tokenTimes.replace("#", i);
     body.replaceText(tokenToReplace, row[2]);
     //Distance
     tokenToReplace = tokenDistance.replace("#", i);
     body.replaceText(tokenToReplace, row[3]);
   }
  
  var numChildren = body.getNumChildren();//number of the contents in the template doc
   
    for (var c = 0; c < numChildren; c++) {//Go over all the content of the template doc, and replicate it for each row of the data.
      var child = body.getChild(c);
      child = child.copy();
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child.getBlob());
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        mergedDoc.appendTable(child);
      } else {
        Logger.log("Unknown element type: " + child);
      }
  }
  
  }  
   
}


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the doGenerate() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Generate brevet card fronts",
    functionName : "generateCardFronts"},
   {name : "Generate brevet card backs",
    functionName : "generateCardBacks"
  }];
  spreadsheet.addMenu("Brevet Cards", entries);
};