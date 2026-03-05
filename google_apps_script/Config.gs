/**
 * ============================================================
 * CONFIGURATION — Edit these values to match your setup
 * ============================================================
 */

// ---------- Sheet names ----------
var CONFIG = {
  // Name of the sheet with orders data
  SHEET_ENCOMENDAS: "encomendas",

  // ---------- Print Agent ----------
  // URL of the print agent (via tunnel to the local PC)
  PRINT_AGENT_URL: "https://ALTERAR-COM-URL-DO-TUNNEL",
  
  // Auth token (must match the one in print_agent.py)
  PRINT_AGENT_TOKEN: "producao2026",

  // ---------- Label Size (in cm) ----------
  // Impressora de etiquetas 62mm x 29mm
  LABEL_WIDTH_CM: 6.2,
  LABEL_HEIGHT_CM: 2.9,
  
  // Layout: 1 etiqueta por pagina (impressora de etiquetas)
  LABEL_COLS: 1,
  LABEL_ROWS: 1,
  TOTAL_LABELS: 1,
  
  // Margens minimas (impressora de etiquetas)
  MARGIN_TOP_CM: 0.0,
  MARGIN_LEFT_CM: 0.0
};
