function getAttributeValue(attribute) { //function to get attribute value and check if it exists because some xmls dont have all values
    return attribute ? attribute.getValue() : '';
  }
  
  function parseXMLAndCopyToSheet() {
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    var folderId = ''; // google drive folder ID (hidden for security purposes)
    var files = DriveApp.getFolderById(folderId).getFiles();
  
    while(files.hasNext()) { //read xmls until there are no more in the folder
      var file = files.next();
  
        var xmlData = file.getBlob().getDataAsString('UTF-8');
        xmlData = xmlData.replace(/^\uFEFF/, '');
  
            //parse XMLs
            var document = XmlService.parse(xmlData);
            var root = document.getRootElement();
  
            var tfdNamespace = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');
            var cfdiNamespace = XmlService.getNamespace('cfdi', 'http://www.sat.gob.mx/cfd/4');
            
            var comprobante = root.getChild('Complemento', cfdiNamespace);
            var impuestos = root.getChild('Impuestos', cfdiNamespace)
            var timbreFiscalDigital = comprobante.getChild('TimbreFiscalDigital', tfdNamespace);
  
            var fecha = getAttributeValue(root.getAttribute('Fecha')).split("T")[0];
            var uuid = getAttributeValue(timbreFiscalDigital.getAttribute('UUID'));
            var subTotal = getAttributeValue(root.getAttribute('SubTotal'));
            var total = getAttributeValue(root.getAttribute('Total'));
            var folio = getAttributeValue(root.getAttribute('Folio'));
            var serie = getAttributeValue(root.getAttribute('Serie'));
            var impuestosTrasladados = getAttributeValue(impuestos.getAttribute('TotalImpuestosTrasladados'));
  
            var emisor = root.getChild('Emisor', cfdiNamespace);
            var rfc = getAttributeValue(emisor.getAttribute('Rfc'));
            var nombre = getAttributeValue(emisor.getAttribute('Nombre'));
          
            
            //extract concepts
            var conceptos = root.getChild('Conceptos', root.getNamespace('cfdi')).getChildren('Concepto', root.getNamespace('cfdi'));
            
            // loop through each concept
            for (var i = 0; i < conceptos.length; i++) {
              var concepto = conceptos[i];
              
              // extract everything from each concept
              var descripcion = getAttributeValue(concepto.getAttribute('Descripcion'));
              var clave = getAttributeValue(concepto.getAttribute('ClaveUnidad'));
              var cantidad = getAttributeValue(concepto.getAttribute('Cantidad'));
              var unidad = getAttributeValue(concepto.getAttribute('Unidad'));
              var valorUnitario = getAttributeValue(concepto.getAttribute('ValorUnitario'));
              
              //write everything in the sheet
              var row = (i == 0) ? [fecha, nombre, rfc, serie, folio, uuid, descripcion, clave, cantidad, unidad, valorUnitario, subTotal, impuestosTrasladados, total] : ['', '', '', '', '', '', descripcion, clave, cantidad, unidad, valorUnitario, '', '', ''];
  
  
              //append row to sheet
              sheet.appendRow(row);
              }
      }
      
    }
  
  
  function onOpen() { //custom button
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Código')
      .addItem('Leer XMLs', 'parseXMLAndCopyToSheet')
      .addToUi();
  }
  