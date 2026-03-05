/**
 * ============================================================
 * CONFIGURATION — Edit these values to match your setup
 * ============================================================
 */

// ---------- Sheet names ----------
var CONFIG = {
  // Name of the sheet with orders data
  SHEET_ENCOMENDAS: "encomendas",
  
  // Name of the sheet tracking label position
  SHEET_POSITION: "current_position",
  
  // Cell in current_position sheet that stores the current label slot (1-2)
  POSITION_CELL: "B2",

  // ---------- Print Agent ----------
  // URL of the print agent (via ngrok tunnel to the local PC)
  PRINT_AGENT_URL: "https://16b9-83-223-225-181.ngrok-free.app",
  
  // Auth token (must match the one in print_agent.py)
  PRINT_AGENT_TOKEN: "producao2026",

  // ---------- A4 Label Layout (in cm) ----------
  PAGE_WIDTH_CM: 21.0,
  PAGE_HEIGHT_CM: 29.7,
  
  MARGIN_TOP_CM: 1.0,
  MARGIN_BOTTOM_CM: 1.0,
  MARGIN_LEFT_CM: 1.0,
  MARGIN_RIGHT_CM: 1.0,
  
  // No horizontal gap (only 1 column)
  GAP_HORIZONTAL_CM: 0.0,
  
  // Gap between the two rows
  GAP_VERTICAL_CM: 0.5,
  
  LABEL_COLS: 1,
  LABEL_ROWS: 2,
  
  // ---------- Computed (will be set in init) ----------
  LABEL_WIDTH_CM: 0,
  LABEL_HEIGHT_CM: 0,
  TOTAL_LABELS: 2
};

// Compute label dimensions
(function() {
  var usableWidth  = CONFIG.PAGE_WIDTH_CM - CONFIG.MARGIN_LEFT_CM - CONFIG.MARGIN_RIGHT_CM - CONFIG.GAP_HORIZONTAL_CM;
  var usableHeight = CONFIG.PAGE_HEIGHT_CM - CONFIG.MARGIN_TOP_CM - CONFIG.MARGIN_BOTTOM_CM;
  
  CONFIG.LABEL_WIDTH_CM  = usableWidth / CONFIG.LABEL_COLS;   // 19 cm
  CONFIG.LABEL_HEIGHT_CM = (usableHeight - CONFIG.GAP_VERTICAL_CM) / CONFIG.LABEL_ROWS; // ~13.35 cm
  CONFIG.TOTAL_LABELS    = CONFIG.LABEL_COLS * CONFIG.LABEL_ROWS; // 2
})();
