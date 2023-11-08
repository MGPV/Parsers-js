function parseXMLAndCopyToSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    var folderId = ''; //replace with google drive folder ID (removed for security purposes)
    var files = DriveApp.getFolderById(folderId).getFiles();
  
    while (files.hasNext()) { //until no more xmls are in the folder
      var file = files.next();
      var xmlData = file.getBlob().getDataAsString();
      var document = XmlService.parse(xmlData); //parse xml file
      var root = document.getRootElement();
      
      // get attributes
      var fechaEmision = root.getAttribute('Fecha').getValue().split("T")[0];
      var total = root.getAttribute('Total').getValue();
      var folio = root.getAttribute('Folio').getValue();
      var serie = root.getAttribute('Serie').getValue();
      var condiciones = root.getAttribute('CondicionesDePago').getValue();
  
      //write the variables into the row
      var row = [serie, folio, fechaEmision, condiciones, total];

      sheet.appendRow(row);
    }
  }
  
  function onOpen() { //custom button
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Código')
      .addItem('Leer XMLs', 'parseXMLAndCopyToSheet')
      .addToUi();
  }
  