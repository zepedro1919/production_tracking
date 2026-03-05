/**
 * LABEL PDF GENERATOR - Half-A4 labels (1 col x 2 rows)
 * Position 1 = top half, Position 2 = bottom half
 */
function generateLabelPdf(data, position) {
  var pos0 = position - 1;
  var row = Math.floor(pos0 / CONFIG.LABEL_COLS);
  var col = pos0 % CONFIG.LABEL_COLS;
  var labelX = CONFIG.MARGIN_LEFT_CM + col * (CONFIG.LABEL_WIDTH_CM + CONFIG.GAP_HORIZONTAL_CM);
  var labelY = CONFIG.MARGIN_TOP_CM + row * (CONFIG.LABEL_HEIGHT_CM + CONFIG.GAP_VERTICAL_CM);
  var htmlContent = buildLabelHtml(data, labelX, labelY);
  var blob = HtmlService.createHtmlOutput(htmlContent)
    .getBlob()
    .getAs("application/pdf")
    .setName("label.pdf");
  return blob;
}

function buildLabelHtml(data, x_cm, y_cm) {
  var w = CONFIG.LABEL_WIDTH_CM;
  var h = CONFIG.LABEL_HEIGHT_CM;
  var desc = data.descricao;
  if (desc.length > 120) {
    desc = desc.substring(0, 117) + "...";
  }
  var logoBase64 = "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";
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
  '@page { size: A4; margin: 0; } ' +
  '* { margin: 0; padding: 0; box-sizing: border-box; } ' +
  'body { width: 21cm; height: 29.7cm; position: relative; font-family: Arial, Helvetica, sans-serif; } ' +
  '.label { position: absolute; left: ' + x_cm + 'cm; top: ' + y_cm + 'cm; width: ' + w + 'cm; height: ' + h + 'cm; padding: 0.8cm; overflow: hidden; } ' +
  '.label-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 0.5cm; } ' +
  '.label-info { flex: 1; } ' +
  '.label-logo { width: 4cm; height: auto; margin-left: 0.3cm; } ' +
  '.label-line { font-size: 14pt; line-height: 1.6; } ' +
  '.label-line strong { font-size: 15pt; } ' +
  '.label-line.desc { font-size: 12pt; line-height: 1.4; margin-top: 0.3cm; } ' +
  '.label-counter { font-size: 28pt; font-weight: bold; margin-top: 0.5cm; } ' +
  '</style></head><body>' +
  '<div class="label">' +
  '<div class="label-header">' +
  '<div class="label-info">' +
  '<div class="label-line"><strong>Encomenda:</strong> ' + escapeHtml(data.enc) + '</div>' +
  '<div class="label-line"><strong>Codigo:</strong> ' + escapeHtml(data.codigo) + '</div>' +
  '</div>' +
  '<img class="label-logo" src="data:image/png;base64,' + logoBase64 + '" />' +
  '</div>' +
  '<div class="label-line desc"><strong>Descricao:</strong> ' + escapeHtml(desc) + '</div>' +
  '<div class="label-counter">' + data.picada + ' / ' + data.producao + '</div>' +
  '</div></body></html>';
  return html;
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
