/**
 * ============================================================
 * MAIN WEB APP — handles QR scan requests (backend only)
 * ============================================================
 * 
 * Deploy as: Web App
 *   - Execute as: Me
 *   - Who has access: Anyone (so workers can scan without login)
 *
 * The unique ID for each row is built on-the-fly from:
 *   enc_codigo_descricao_quantidadeProducao
 *
 * QR codes encode: https://script.google.com/.../exec?id=<that_id>
 *
 * On scan, the script rebuilds IDs from columns A-D and finds
 * the matching row, regardless of row order.
 *
 * The worker scans → the label prints. That's it.
 */

// ---------- Column indices (0-based) in "encomendas" sheet ----------
//   A      B       C          D                   E                F
//   enc | codigo | descricao | quantidadeProducao | quantidadePicada | qrCode
var COL = {
  ENC:       0,  // A - enc
  CODIGO:    1,  // B - codigo
  DESCRICAO: 2,  // C - descricao
  QTD_PROD:  3,  // D - quantidadeProducao
  QTD_PIC:   4,  // E - quantidadePicada
  QR_CODE:   5   // F - qrCode
};

// Total columns to read per row
var TOTAL_COLS = 6;


/**
 * GET handler — called when a worker scans the QR code.
 * Returns plain text so it works on any phone/browser instantly.
 */
function doGet(e) {
  try {
    var id = e.parameter.id;
    if (!id) {
      return textResponse("ERRO: Parâmetro 'id' em falta.");
    }
    
    var rowNum = findRowById(id);
    if (!rowNum) {
      return textResponse("ERRO: Produto não encontrado (id=" + id + ").");
    }
    
    var result = processPickagem(rowNum);
    
    if (result.error) {
      return textResponse("ERRO: " + result.error);
    }
    
    return textResponse(result.message);
    
  } catch (err) {
    return textResponse("ERRO: " + err.message);
  }
}


/**
 * Find the sheet row number for a given unique ID.
 * Rebuilds IDs from columns A-D (enc_codigo_descricao_qtdProd)
 * and compares with the scanned ID.
 * Returns the 1-based row number, or null if not found.
 */
function findRowById(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return null;
  
  // Read columns A-D (enc, codigo, descricao, quantidadeProducao)
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowId = String(data[i][0]).trim() + "_" +
                String(data[i][1]).trim() + "_" +
                String(data[i][2]).trim() + "_" +
                String(data[i][3]).trim();
    if (rowId === String(id)) {
      return i + 2; // convert 0-based array index to 1-based sheet row
    }
  }
  
  return null;
}


/**
 * Return a simple plain-text response.
 */
function textResponse(text) {
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Core logic: validate, increment, generate label, print, advance position.
 */
function processPickagem(rowNum) {
  var lock = LockService.getScriptLock();
  try {
    // Prevent race conditions if two workers scan at the same time
    lock.waitLock(10000); // wait up to 10s
  } catch (e) {
    return { error: "Sistema ocupado, tente novamente." };
  }
  
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_ENCOMENDAS);
    
    if (!sheet) {
      return { error: "Folha '" + CONFIG.SHEET_ENCOMENDAS + "' não encontrada." };
    }
    
    // Read the row (6 columns: A-F)
    var range  = sheet.getRange(rowNum, 1, 1, TOTAL_COLS);
    var values = range.getValues()[0];
    
    var enc       = values[COL.ENC];
    var codigo    = values[COL.CODIGO];
    var descricao = values[COL.DESCRICAO];
    var qtdProd   = parseInt(values[COL.QTD_PROD], 10) || 0;
    var qtdPic    = parseInt(values[COL.QTD_PIC],  10) || 0;
    
    // Validate: can we still do picagem?
    if (qtdPic >= qtdProd) {
      return {
        error: "Picagem completa! " + descricao + " (" + qtdProd + "/" + qtdProd + ")"
      };
    }
    
    // Increment quantidadePicada
    var newQtdPic = qtdPic + 1;
    sheet.getRange(rowNum, COL.QTD_PIC + 1).setValue(newQtdPic);
    SpreadsheetApp.flush();
    
    // Get current label position (1-14)
    var posSheet   = ss.getSheetByName(CONFIG.SHEET_POSITION);
    var currentPos = parseInt(posSheet.getRange(CONFIG.POSITION_CELL).getValue(), 10) || 1;
    
    // Generate the label PDF
    var labelData = {
      enc:       enc,
      codigo:    codigo,
      descricao: descricao,
      picada:    newQtdPic,
      producao:  qtdProd
    };
    
    var pdfBlob = generateLabelPdf(labelData, currentPos);
    
    // Send PDF to print agent → prints automatically
    var printResult = sendToPrintAgent(pdfBlob);
    
    // Advance position: 1→2→...→14→1
    var nextPos = (currentPos % CONFIG.TOTAL_LABELS) + 1;
    posSheet.getRange(CONFIG.POSITION_CELL).setValue(nextPos);
    SpreadsheetApp.flush();
    
    var printStatus = printResult.success 
      ? "Impressão enviada ✓" 
      : "AVISO: Impressão falhou (" + printResult.error + ")";
    
    return {
      error: null,
      message: "✓ PICAGEM REGISTADA\n" +
               "Encomenda: " + enc + "\n" +
               "Código: " + codigo + "\n" +
               "Picagem: " + newQtdPic + "/" + qtdProd + "\n" +
               "Etiqueta posição: " + currentPos + "\n" +
               printStatus
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Send a PDF blob directly to the print agent via HTTP POST.
 */
function sendToPrintAgent(pdfBlob) {
  var url = CONFIG.PRINT_AGENT_URL + "/print";
  
  try {
    var response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/pdf",
      payload: pdfBlob.getBytes(),
      headers: {
        "Authorization": "Bearer " + CONFIG.PRINT_AGENT_TOKEN
      },
      muteHttpExceptions: true,
      validateHttpsCertificates: false
    });
    
    var code = response.getResponseCode();
    if (code === 200) {
      return { success: true };
    } else {
      var body = response.getContentText();
      Logger.log("Print agent error: " + code + " " + body);
      return { success: false, error: "HTTP " + code };
    }
  } catch (e) {
    Logger.log("Print agent unreachable: " + e.message);
    return { success: false, error: e.message };
  }
}
