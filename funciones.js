let estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {

  SpreadsheetApp.getUi().createMenu('Menu CBLUNA')
     .addItem('Mostrar barra lateral', 'mostrarBarraLateral')
     .addItem('Mostrar Propiedaes de una celda', 'obtenerAtributosCelda')
     .addToUi();

}

function mostrarBarraLateral() {
   let ui =HtmlService.createHtmlOutputFromFile('BarraLateral').setTitle('Barra lateral CBLUNA');
   SpreadsheetApp.getUi().showSidebar(ui)
}

function aplicarEstilo1(){
  let hojaActual =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let celdas = hojaActual.getActiveRange();
  celdas.setBackground('blue')
        .setFontColor('white')
        .setHorizontalAlignment('center')
        .setValue('Estilo 1')
}
function aplicarEstilo2(){
   let hojaActual =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   let celdas = hojaActual.getActiveRange();
   celdas.setBackground('green')
        .setFontColor('white')
        .setFontWeight('bold')
        .setValue('Estilo 2')
}
function copiarFormatoCelda() {
  let celda =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(celdaOrigen).getE;

  // Obtiene el formato de la celda de origen
  var formatoOrigen = celdaOrigen.getActiveSheet().getRange(celdaOrigen).getEffectiveFormat();

  // Aplica el formato a las celdas del rango de destino
  rangoDestino.getActiveSheet().getRange(rangoDestino).applyFormat(formatoOrigen);
}

function guardarEstilo(){
   let celda =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
   
    estilo_sheet.setProperty('color', celda.getTextStyle())
                .setProperty('colorFondo', celda.getBackground())
                .setProperty('size', celda.getFontSize());

  
 
}
function obtenerAtributosCelda() {
  // Obtiene la hoja de cálculo activa
  var hoja = SpreadsheetApp.getActiveSheet();

  // Obtiene la celda seleccionada
  var celda = hoja.getActiveCell();

  // Obtiene el color de fondo de la celda
  var colorFondo = celda.getBackground();

  // Obtiene el tamaño del texto de la celda
  var tamanoTexto = celda.getFontSize();

  // Obtiene el color del texto de la celda
  var colorTexto = celda.getTextStyle().textColor;

  // Muestra los valores obtenidos en un cuadro de diálogo
  var mensaje = "Color de fondo: " + colorFondo + "\n" +
              "Tamaño del texto: " + tamanoTexto + "\n" +
              "Color del texto: " + colorTexto;
var ui = SpreadsheetApp.getUi();
var response = ui.alert(mensaje, ui.ButtonSet.YES_NO);

// Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
  } else {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  }
  
}

function copiarEstilo(){
  let hojaActual =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let celdas = hojaActual.getActiveRange();

  
  
  celdas.setTextStyle(estilos_sheet.getProperty('color'))
        .setBackground(estilos_sheet.getProperty('colorFondo'))
        .setFontSize(estilos_sheet.getProperty('size'));
        
}



function borrarEstilos(){
   let hojaActual =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    hojaActual.getActiveRange().clear({formatOnly: true});
}


function borrarTodo(){
   let hojaActual =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    hojaActual.getActiveRange().clear();
}
   
  