"""
Print Agent - Producao Nautilus
Servidor local que recebe PDFs e imprime automaticamente.
Versao portatil com auto-deteccao de impressora e SumatraPDF.
"""

import os
import sys
import json
import time
import shutil
import logging
import tempfile
import subprocess
from datetime import datetime
from pathlib import Path

try:
    from flask import Flask, request, jsonify
except ImportError:
    print("[ERRO] Flask nao instalado. Corra: pip install flask requests")
    sys.exit(1)

# ============================================================
# CONFIGURACAO
# ============================================================

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULT_CONFIG = {
    "port": 5555,
    "auth_token": "producao2026",
    "printer_name": "",
    "sumatra_path": "",
    "temp_dir": ""
}


def load_config():
    """Carrega config.json ou cria com valores detectados automaticamente."""
    config = DEFAULT_CONFIG.copy()

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                config.update(saved)
                logging.info(f"Configuracao carregada de {CONFIG_FILE}")
        except Exception as e:
            logging.warning(f"Erro ao ler config: {e}. A usar valores por defeito.")

    # Auto-detectar o que falta
    changed = False

    if not config["printer_name"]:
        detected = auto_detect_printer()
        if detected:
            config["printer_name"] = detected
            changed = True
            logging.info(f"Impressora auto-detectada: {detected}")

    if not config["sumatra_path"] or not os.path.exists(config["sumatra_path"]):
        detected = auto_detect_sumatra()
        if detected:
            config["sumatra_path"] = detected
            changed = True
            logging.info(f"SumatraPDF auto-detectado: {detected}")

    if not config["temp_dir"] or not os.path.isabs(config["temp_dir"]):
        config["temp_dir"] = os.path.join(tempfile.gettempdir(), "print_agent_nautilus")
        changed = True

    # Garantir que pasta temp existe
    try:
        os.makedirs(config["temp_dir"], exist_ok=True)
    except (PermissionError, OSError):
        # Se a pasta temp for de outro PC/utilizador, usar a local
        config["temp_dir"] = os.path.join(tempfile.gettempdir(), "print_agent_nautilus")
        os.makedirs(config["temp_dir"], exist_ok=True)
        changed = True

    if changed:
        save_config(config)

    return config


def save_config(config):
    """Guarda configuracao no ficheiro."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        logging.info(f"Configuracao guardada em {CONFIG_FILE}")
    except Exception as e:
        logging.warning(f"Erro ao guardar config: {e}")


# ============================================================
# AUTO-DETECCAO
# ============================================================

def list_printers():
    """Lista todas as impressoras instaladas no Windows."""
    printers = []
    try:
        result = subprocess.run(
            ["wmic", "printer", "get", "name"],
            capture_output=True, text=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
        )
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            for line in lines[1:]:  # skip header
                name = line.strip()
                if name and name.lower() != "name":
                    printers.append(name)
    except Exception as e:
        logging.warning(f"Erro ao listar impressoras: {e}")

    return printers


def auto_detect_printer():
    """Detecta a impressora mais provavel (prioriza EPSON/RICOH, ignora virtuais)."""
    printers = list_printers()
    if not printers:
        return ""

    # Impressoras a ignorar (virtuais)
    virtual_keywords = [
        "microsoft", "pdf", "xps", "onenote", "fax",
        "send to", "nul", "file"
    ]

    # Filtrar impressoras fisicas
    physical = []
    for p in printers:
        lower = p.lower()
        if not any(kw in lower for kw in virtual_keywords):
            physical.append(p)

    if not physical:
        return printers[0] if printers else ""

    # Priorizar EPSON ou RICOH
    for p in physical:
        lower = p.lower()
        if "epson" in lower or "ricoh" in lower:
            return p

    return physical[0]


def auto_detect_sumatra():
    """Procura SumatraPDF em localizacoes comuns."""
    search_paths = [
        # Pasta local do instalador
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools", "SumatraPDF.exe"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "SumatraPDF.exe"),
        # AppData do utilizador
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        # Program Files
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
        # Raiz
        r"C:\SumatraPDF\SumatraPDF.exe",
    ]

    # Adicionar pastas de todos os utilizadores
    users_dir = r"C:\Users"
    if os.path.exists(users_dir):
        try:
            for user_dir in os.listdir(users_dir):
                try:
                    user_sumatra = os.path.join(users_dir, user_dir, "AppData", "Local", "SumatraPDF", "SumatraPDF.exe")
                    if user_sumatra not in search_paths:
                        search_paths.append(user_sumatra)
                except (PermissionError, OSError):
                    pass
        except (PermissionError, OSError):
            pass

    for path in search_paths:
        if path and os.path.isfile(path):
            return path

    # Tentar encontrar via where
    try:
        result = subprocess.run(
            ["where", "SumatraPDF"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip().split('\n')[0].strip()
    except Exception:
        pass

    return ""


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)
config = {}


def print_pdf_sumatra(pdf_path):
    """Imprime PDF usando SumatraPDF (silencioso, rapido)."""
    sumatra = config.get("sumatra_path", "")
    printer = config.get("printer_name", "")

    if not sumatra or not os.path.exists(sumatra):
        raise FileNotFoundError(f"SumatraPDF nao encontrado em: {sumatra}")

    cmd = [
        sumatra,
        "-print-to", printer if printer else "-print-to-default",
        "-silent",
        "-print-settings", "fit",
        pdf_path
    ]

    # Ajustar comando se usar impressora default
    if not printer:
        cmd = [
            sumatra,
            "-print-to-default",
            "-silent",
            "-print-settings", "fit",
            pdf_path
        ]

    logging.info(f"SumatraPDF cmd: {' '.join(cmd)}")

    creation_flags = 0
    if hasattr(subprocess, 'CREATE_NO_WINDOW'):
        creation_flags = subprocess.CREATE_NO_WINDOW

    result = subprocess.run(
        cmd,
        capture_output=True, text=True, timeout=30,
        creationflags=creation_flags
    )

    if result.returncode != 0:
        logging.warning(f"SumatraPDF stderr: {result.stderr}")

    return result.returncode


def print_pdf_fallback(pdf_path):
    """Impressao via os.startfile (abre dialogo de impressao)."""
    logging.info("A usar metodo alternativo (os.startfile)...")
    try:
        os.startfile(pdf_path, "print")
        return 0
    except Exception as e:
        logging.error(f"Falha no metodo alternativo: {e}")
        return 1


def do_print(pdf_path):
    """Tenta imprimir: SumatraPDF primeiro, fallback depois."""
    try:
        return print_pdf_sumatra(pdf_path)
    except FileNotFoundError:
        logging.warning("SumatraPDF indisponivel, a usar fallback...")
        return print_pdf_fallback(pdf_path)
    except Exception as e:
        logging.error(f"Erro SumatraPDF: {e}")
        return print_pdf_fallback(pdf_path)


@app.route("/health", methods=["GET"])
def health():
    """Endpoint de verificacao de saude."""
    printers = list_printers()
    return jsonify({
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "printer": config.get("printer_name", "N/A"),
        "sumatra": os.path.exists(config.get("sumatra_path", "")),
        "printers_available": printers,
        "version": "2.0-portable"
    })


@app.route("/print", methods=["POST"])
def print_endpoint():
    """Recebe PDF e imprime."""

    # Verificar token
    auth = request.headers.get("Authorization", "")
    expected = f"Bearer {config.get('auth_token', '')}"
    if auth != expected:
        logging.warning(f"Auth falhou: recebido '{auth[:20]}...'")
        return jsonify({"error": "Nao autorizado"}), 401

    # Verificar ficheiro
    if "file" not in request.files:
        return jsonify({"error": "Nenhum ficheiro recebido"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Ficheiro sem nome"}), 400

    # Guardar temporariamente
    temp_dir = config.get("temp_dir", tempfile.gettempdir())
    os.makedirs(temp_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c if c.isalnum() or c in ".-_" else "_" for c in file.filename)
    pdf_path = os.path.join(temp_dir, f"{timestamp}_{safe_name}")

    try:
        file.save(pdf_path)
        file_size = os.path.getsize(pdf_path)
        logging.info(f"PDF recebido: {safe_name} ({file_size} bytes)")

        # Imprimir
        result = do_print(pdf_path)

        if result == 0:
            logging.info(f"Impressao enviada com sucesso: {safe_name}")
            return jsonify({
                "status": "success",
                "message": f"Impresso: {safe_name}",
                "printer": config.get("printer_name", "default"),
                "size_bytes": file_size
            })
        else:
            logging.error(f"Falha na impressao: codigo {result}")
            return jsonify({
                "error": f"Falha na impressao (codigo: {result})",
                "printer": config.get("printer_name", "default")
            }), 500

    except Exception as e:
        logging.error(f"Erro ao processar impressao: {e}")
        return jsonify({"error": str(e)}), 500

    finally:
        # Limpar ficheiro temp (com delay para SumatraPDF acabar)
        try:
            time.sleep(2)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
        except Exception:
            pass


@app.route("/print-url", methods=["POST"])
def print_url_endpoint():
    """Recebe URL de PDF, descarrega e imprime."""
    import requests as req

    # Verificar token
    auth = request.headers.get("Authorization", "")
    expected = f"Bearer {config.get('auth_token', '')}"
    if auth != expected:
        return jsonify({"error": "Nao autorizado"}), 401

    data = request.get_json()
    if not data or "url" not in data:
        return jsonify({"error": "URL em falta"}), 400

    url = data["url"]
    logging.info(f"A descarregar PDF de: {url}")

    try:
        resp = req.get(url, timeout=30)
        resp.raise_for_status()

        temp_dir = config.get("temp_dir", tempfile.gettempdir())
        os.makedirs(temp_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_path = os.path.join(temp_dir, f"{timestamp}_url_download.pdf")

        with open(pdf_path, "wb") as f:
            f.write(resp.content)

        result = do_print(pdf_path)

        try:
            time.sleep(2)
            os.remove(pdf_path)
        except Exception:
            pass

        if result == 0:
            return jsonify({"status": "success", "message": "Impresso com sucesso"})
        else:
            return jsonify({"error": "Falha na impressao"}), 500

    except Exception as e:
        logging.error(f"Erro: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/config", methods=["GET"])
def get_config():
    """Mostra configuracao actual (sem token)."""
    safe = {k: v for k, v in config.items() if k != "auth_token"}
    safe["printers"] = list_printers()
    return jsonify(safe)


@app.route("/config", methods=["POST"])
def update_config():
    """Actualiza configuracao em runtime."""
    global config

    auth = request.headers.get("Authorization", "")
    expected = f"Bearer {config.get('auth_token', '')}"
    if auth != expected:
        return jsonify({"error": "Nao autorizado"}), 401

    data = request.get_json()
    if not data:
        return jsonify({"error": "Dados em falta"}), 400

    allowed_keys = ["printer_name", "sumatra_path"]
    for key in allowed_keys:
        if key in data:
            config[key] = data[key]
            logging.info(f"Config actualizada: {key} = {data[key]}")

    save_config(config)
    return jsonify({"status": "ok", "config": {k: v for k, v in config.items() if k != "auth_token"}})


# ============================================================
# MAIN
# ============================================================

def print_banner():
    """Mostra banner inicial com info de configuracao."""
    print()
    print("=" * 60)
    print("  PRINT AGENT - Producao Nautilus v2.0")
    print("=" * 60)
    print(f"  Porta:       {config.get('port', 5555)}")
    print(f"  Impressora:  {config.get('printer_name', 'N/A')}")
    print(f"  SumatraPDF:  {'OK' if os.path.exists(config.get('sumatra_path', '')) else 'NAO ENCONTRADO'}")
    print(f"  Config:      {CONFIG_FILE}")
    print(f"  Temp:        {config.get('temp_dir', 'N/A')}")
    print("=" * 60)
    print()

    if not config.get("printer_name"):
        print("  [AVISO] Nenhuma impressora configurada!")
        print("  Impressoras disponiveis:")
        for p in list_printers():
            print(f"    - {p}")
        print()

    if not os.path.exists(config.get("sumatra_path", "")):
        print("  [AVISO] SumatraPDF nao encontrado!")
        print("  Vai usar metodo alternativo de impressao.")
        print()


if __name__ == "__main__":
    # Setup logging
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "print_agent.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding="utf-8")
        ]
    )

    # Carregar configuracao
    config = load_config()
    print_banner()

    port = config.get("port", 5555)
    logging.info(f"A iniciar servidor na porta {port}...")

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
        use_reloader=False
    )
