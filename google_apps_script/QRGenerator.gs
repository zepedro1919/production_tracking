/**
 * ============================================================
 * QR CODE GENERATOR
 * ============================================================
 * 
 * The unique ID for each row is built on-the-fly from:
 *   enc + "_" + codigo + "_" + descricao + "_" + quantidadeProducao
 * 
 * This combination is naturally unique per row.
 * No extra "id" column needed — it's computed when generating QR codes
 * and again when scanning (in findRowById).
 *
 * The QR code encodes:
 *   https://script.google.com/.../exec?id=<that_id>
 *
 * Run generateAllQrCodes() from the menu after deploying.
 */


/**
 * Add a custom menu to the spreadsheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🏭 Produção")
    .addItem("Gerar QR Codes", "generateAllQrCodes")
    .addItem("Limpar QR Codes", "clearAllQrCodes")
    .addSeparator()
    .addItem("Reset Posição Etiqueta", "resetLabelPosition")
    .addItem("Reset Todas Picagens", "resetAllPickings")
    .addSeparator()
    .addItem("Testar Ligação Impressora", "testPrintAgentConnection")
    .addItem("Testar Impressão", "testLabelGeneration")
    .addToUi();
}


/**
 * Generate unique IDs + QR code IMAGE formulas for all data rows.
 */
function generateAllQrCodes() {
  var deployUrl = getWebAppUrl();
  
  if (!deployUrl || deployUrl === "PASTE_YOUR_DEPLOY_URL_HERE") {
    SpreadsheetApp.getUi().alert(
      "⚠️ URL não configurada!\n\n" +
      "Execute primeiro: setWebAppUrl() e cole o URL do web app deployado."
    );
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Nenhum dado encontrado.");
    return;
  }
  
  // Ensure headers exist
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Nenhum dado encontrado.");
    return;
  }
  
  // Read columns A-D (enc, codigo, descricao, quantidadeProducao) to build IDs
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A-D
  
  var qrColIndex = COL.QR_CODE + 1;  // column F (1-based)
  
  var count = 0;
  
  for (var i = 0; i < dataRange.length; i++) {
    var row = i + 2; // 1-based sheet row
    var enc       = String(dataRange[i][0]).trim();
    var codigo    = String(dataRange[i][1]).trim();
    var descricao = String(dataRange[i][2]).trim();
    var qtdProd   = String(dataRange[i][3]).trim();
    
    // Skip empty rows
    if (!enc && !codigo) continue;
    
    // Build ID: enc_codigo_descricao_quantidadeProducao
    var id = enc + "_" + codigo + "_" + descricao + "_" + qtdProd;
    
    // Build the scan URL
    var scanUrl = deployUrl + "?id=" + encodeURIComponent(id);
    
    // QR code via qrserver.com with quiet zone margin (margin=4)
    var qrImageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=4&data=" + 
                     encodeURIComponent(scanUrl);
    
    // Set IMAGE formula - mode 1 = stretch to fit cell
    var formula = '=IMAGE("' + qrImageUrl + '")';
    sheet.getRange(row, qrColIndex).setFormula(formula);
    
    count++;
  }
  
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    "✅ QR Codes gerados!\n\n" +
    "Total: " + count + " códigos QR.\n" +
    "Cada QR contém um ID único baseado em:\n" +
    "  encomenda + código + descrição + quantidade\n\n" +
    "Pode agora imprimir a coluna F (QR Codes)."
  );
}


/**
 * Clear all QR codes (keeps the IDs intact so they can be regenerated)
 */
function clearAllQrCodes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    sheet.getRange(2, COL.QR_CODE + 1, lastRow - 1, 1).clearContent();
    SpreadsheetApp.flush();
  }
  
  SpreadsheetApp.getUi().alert("QR Codes removidos (IDs mantidos).");
}


/**
 * Clear all QR codes (full reset — will need to regenerate)
 */
function clearAllQrCodesAndIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    sheet.getRange(2, COL.QR_CODE + 1, lastRow - 1, 1).clearContent();
    SpreadsheetApp.flush();
  }
  
  SpreadsheetApp.getUi().alert("QR Codes removidos.");
}


/**
 * Reset the label position counter back to 1
 */
function resetLabelPosition() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var posSheet = ss.getSheetByName(CONFIG.SHEET_POSITION);
  posSheet.getRange(CONFIG.POSITION_CELL).setValue(1);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("Posição da etiqueta reiniciada para 1.");
}


/**
 * Reset all quantidadePicada values to 0
 */
function resetAllPickings() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "⚠️ Confirmar Reset",
    "Tem a certeza que quer reiniciar TODAS as picagens para 0?",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  var lastRow = sheet.getLastRow();
  
  if (lastRow >= 2) {
    // Clear quantidadePicada column (F = column 6)
    sheet.getRange(2, COL.QTD_PIC + 1, lastRow - 1, 1).setValue(0);
    SpreadsheetApp.flush();
  }
  
  ui.alert("Todas as picagens reiniciadas para 0.");
}


// ---- Web App URL management ----

var PROPERTY_KEY_URL = "WEB_APP_URL";

/**
 * Store the deployed web app URL in Script Properties
 */
function setWebAppUrl(url) {
  if (!url) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "URL do Web App",
      "Cole aqui o URL do web app deployado\n(ex: https://script.google.com/macros/s/.../exec):",
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return;
    url = response.getResponseText().trim();
  }
  
  PropertiesService.getScriptProperties().setProperty(PROPERTY_KEY_URL, url);
  SpreadsheetApp.getUi().alert("✅ URL guardado: " + url);
}

/**
 * Get the stored web app URL
 */
function getWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty(PROPERTY_KEY_URL) || "PASTE_YOUR_DEPLOY_URL_HERE";
}
