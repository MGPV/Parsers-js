//Marcelo García Pablos Vélez
//dont copy this code without permission

const FOLDER_NAME_PDFS = '' //folder name hidden for security reasons

//pdf to google doc

function convertPdftoDoc(id, ocrLanguage="en") { 
  const pdf = DriveApp.getFileById(id)
  const resource = {
    title: pdf.getName(),
    mimeType: pdf.getMimeType()
  }
  const mediaData = pdf.getBlob()
  const options = {
    convert: true,
    ocr: true,
    ocrLanguage
  }
  const newFile = Drive.Files.insert(resource, mediaData, options)
  return DocumentApp.openById(newFile.id)
}

//read text from google docs

function readTextFromPdf(id, trashDocFile=true){
  const doc = convertPdftoDoc(id)
  const text = doc.getBody().getText()
  DriveApp.getFileById(doc.getId()).setTrashed(trashDocFile)
  return text
}

function getPdfFiles(folderName = FOLDER_NAME_PDFS){
  const ss = SpreadsheetApp.getActive()
  const currentFolder = DriveApp.getFileById(ss.getId()).getParents().next()
  const folders = currentFolder.getFoldersByName(folderName)
  if(!folders.hasNext()) return []
  const folder = folders.next()
  const files = folder.getFilesByType(MimeType.PDF)
  const ids = []
  while(files.hasNext()){
    ids.push(files.next().getId())
  }
  return ids
}

function getInvoiceDataFromText(text){
  const lines = text.split("\n")
  return {
    val1: lines[13].slice(13),
    val2: lines[19].slice(44),
    val3: lines[58],
    val4: lines[50],
  }
}

//export data to google sheets

//overwrite data and include headers
function exportToSheetWithHeaders(){
  const values = [
    ["val1", "val2", "val3", "val4"]
  ]
  const ids = getPdfFiles()
  ids.forEach(id =>{
    const text = readTextFromPdf(id)
    const {val1, val2, val3, val4} = getInvoiceDataFromText(text)
    values.push([
      val1,
      val2,
      val3,
      val4
    ])
  })
  const outputSheet = SpreadsheetApp.getActive().getActiveSheet()
  outputSheet.getRange(1,1, values.length, values[0].length).setValues(values)
}

//same concept but gets the end of the data in the sheet and appends new data
function exportToSheetWithoutHeaders(){
  const values = []
  const ids = getPdfFiles()
  ids.forEach(id =>{
    const text = readTextFromPdf(id)
    const {val1, val2, val3, val4} = getInvoiceDataFromText(text)
    values.push([
      val1,
      val2,
      val3,
      val4
    ])
  })
  
  const outputSheet = SpreadsheetApp.getActive().getActiveSheet()
  const lastRow = outputSheet.getLastRow()
  outputSheet.getRange(lastRow + 1, 1, values.length, values[0].length).setValues(values)
}

function colorMatchingCells() { //function to color matching cells
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeToCheck = sheet.getRange("A2:A111");
  var searchRange = sheet.getRange("L2:L67");
  var searchValues = [].concat.apply([], searchRange.getValues());
  var expandedSearchValues = [];
  
  for (var i = 0; i < searchValues.length; i++) {
    var stringValue = String(searchValues[i]);
    var splitValues = stringValue.split("-"); //if the value contains two numbers separated by "-"
    expandedSearchValues.push(...splitValues);
  }
  
  //create a map to keep track of the count of each unique value
  var searchMap = expandedSearchValues.reduce(function(map, val) {
    map[val] = (map[val] || 0) + 1;
    return map;
  }, {});
  
  //loop through each cell in the range to check for matches
  for (var i = 1; i <= rangeToCheck.getLastRow() - 1; i++) {
    var cellValue = rangeToCheck.getCell(i, 1).getValue();
    
    if (searchMap[cellValue] > 0) {
      var cellToColor = sheet.getRange(i + 1, 1);
      cellToColor.setBackground("#94DF8B");
      searchMap[cellValue]--; //decrement the count in the searchMap
    }
  }
}

function colorMatchingCellsReversed() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var rangeToCheck = sheet.getRange("L2:L154");
    var searchRange = sheet.getRange("A2:A132");
    var searchValues = [].concat.apply([], searchRange.getValues());
    var searchMap = searchValues.reduce(function(map, val) { //create a map to keep track of the count of each unique value
    map[val] = (map[val] || 0) + 1;
    return map;
  }, {});

  //loop through each cell in the range to check for matches
  for (var i = 1; i <= rangeToCheck.getLastRow() - 1; i++) {
    var cellValue = rangeToCheck.getCell(i, 1).getValue();
    var stringValue = String(cellValue); //convert the value to a string
    var splitValues = stringValue.split("-"); //if the value contains two numbers separated by "-"
    var foundAll = true; //boolean to keep track of whether all parts of the value were found in the searchMap
    
    for (var j = 0; j < splitValues.length; j++) {
      if (!(searchMap[splitValues[j]] > 0)) {
        foundAll = false;
        break;
      }
    }
    
    if (foundAll) { //if the values match, color the column green
      var cellToColor = sheet.getRange(i + 1, 12);
      cellToColor.setBackground("#94DF8B");
      
      for (var j = 0; j < splitValues.length; j++) { //decrease count in the map for each part of the value
        searchMap[splitValues[j]]--;
      }
    }
  }
}



function capitalizeColumnU() { //small function to capitalize column U
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("U2:U102");
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    values[i][0] = ('' + values[i][0]).toUpperCase();
  }
  
  range.setValues(values);
}