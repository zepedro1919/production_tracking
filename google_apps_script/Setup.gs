/**
 * ============================================================
 * SETUP & UTILITIES
 * ============================================================
 * 
 * Run setup() once after creating the Apps Script project
 * to initialize the spreadsheet structure.
 */


/**
 * Initial setup — run this once from the Apps Script editor.
 * Creates/verifies the required sheets and sets initial values.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Verify "encomendas" sheet exists
  var encSheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  if (!encSheet) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Folha '" + CONFIG.SHEET_ENCOMENDAS + "' não encontrada!\n\n" +
      "Renomeie a folha com os dados das encomendas para: " + CONFIG.SHEET_ENCOMENDAS
    );
    return;
  }
  
  // 2. Verify/create "current_position" sheet
  var posSheet = ss.getSheetByName(CONFIG.SHEET_POSITION);
  if (!posSheet) {
    posSheet = ss.insertSheet(CONFIG.SHEET_POSITION);
  }
  
  // Set headers and initial position
  posSheet.getRange("A1").setValue("Campo");
  posSheet.getRange("B1").setValue("Valor");
  posSheet.getRange("A2").setValue("Posição Atual");
  
  // Only set position to 1 if it's empty
  var currentVal = posSheet.getRange(CONFIG.POSITION_CELL).getValue();
  if (!currentVal) {
    posSheet.getRange(CONFIG.POSITION_CELL).setValue(1);
  }
  
  // Draw the label position reference grid
  posSheet.getRange("G2").setValue("Posição");
  posSheet.getRange("G3").setValue("1 (Topo)");
  posSheet.getRange("G4").setValue("2 (Baixo)");
  
  // 3. Verify encomendas headers (6 columns)
  var headers = encSheet.getRange(1, 1, 1, 6).getValues()[0];
  var expectedHeaders = ["enc", "codigo", "descricao", "quantidadeProducao", "quantidadePicada", "qrCode"];
  
  var headersOk = true;
  for (var i = 0; i < expectedHeaders.length; i++) {
    if (String(headers[i]).toLowerCase().trim() !== expectedHeaders[i].toLowerCase()) {
      headersOk = false;
      break;
    }
  }
  
  if (!headersOk) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Cabeçalhos da folha '" + CONFIG.SHEET_ENCOMENDAS + "' não correspondem!\n\n" +
      "Esperado: " + expectedHeaders.join(", ") + "\n" +
      "Encontrado: " + headers.join(", ")
    );
    return;
  }
  
  // 4. Initialize quantidadePicada with 0 where empty
  var lastRow = encSheet.getLastRow();
  if (lastRow >= 2) {
    var picRange = encSheet.getRange(2, COL.QTD_PIC + 1, lastRow - 1, 1);
    var picValues = picRange.getValues();
    for (var j = 0; j < picValues.length; j++) {
      if (picValues[j][0] === "" || picValues[j][0] === null || picValues[j][0] === undefined) {
        picValues[j][0] = 0;
      }
    }
    picRange.setValues(picValues);
  }
  
  SpreadsheetApp.flush();
  
  SpreadsheetApp.getUi().alert(
    "✅ Setup completo!\n\n" +
    "Próximos passos:\n" +
    "1. Configurar PRINT_AGENT_URL e PRINT_AGENT_TOKEN em Config.gs\n" +
    "2. Iniciar o print_agent.py no PC da impressora\n" +
    "3. Deploy → New deployment → Web app\n" +
    "4. Execute as: Me | Access: Anyone\n" +
    "5. Copie o URL e execute setWebAppUrl()\n" +
    "6. Execute 'Gerar QR Codes' no menu Produção\n" +
    "7. Imprima os QR Codes e distribua aos trabalhadores"
  );
}


/**
 * Test label generation locally (for debugging)
 */
function testLabelGeneration() {
  var testData = {
    enc: "ENC-A-2251910",
    codigo: "A19AR0000831789A",
    descricao: "Armário Suspenso c/Portas Correr 100x45x80cm - (Me)Branco/R9016",
    picada: 3,
    producao: 10
  };
  
  var pdf = generateLabelPdf(testData, 1);
  Logger.log("PDF generated: " + pdf.getBytes().length + " bytes");
  
  // Try sending to print agent
  var result = sendToPrintAgent(pdf);
  if (result.success) {
    SpreadsheetApp.getUi().alert("✅ Etiqueta de teste enviada para a impressora!");
  } else {
    SpreadsheetApp.getUi().alert("❌ Falha ao enviar para impressora: " + result.error + 
      "\n\nVerifique que o print_agent.py está a correr em " + CONFIG.PRINT_AGENT_URL);
  }
}


/**
 * Test the full scan flow with row 2 (for debugging)
 */
function testScanFlow() {
  var result = processPickagem(2);
  Logger.log(JSON.stringify(result));
  
  if (result.error) {
    SpreadsheetApp.getUi().alert("Erro: " + result.error);
  } else {
    SpreadsheetApp.getUi().alert(result.message);
  }
}


/**
 * Test connectivity to the print agent
 */
function testPrintAgentConnection() {
  var url = CONFIG.PRINT_AGENT_URL + "/health";
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      validateHttpsCertificates: false
    });
    var code = response.getResponseCode();
    var body = response.getContentText();
    SpreadsheetApp.getUi().alert(
      "Resultado: HTTP " + code + "\n\n" + body
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "❌ Não foi possível ligar ao print agent!\n\n" +
      "URL: " + url + "\n" +
      "Erro: " + e.message + "\n\n" +
      "Verifique que:\n" +
      "1. O print_agent.py está a correr\n" +
      "2. O IP/porta em Config.gs está correcto\n" +
      "3. O PC da impressora está na mesma rede"
    );
  }
}
