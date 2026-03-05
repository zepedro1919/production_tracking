"""
Imprimir etiquetas de componentes - Brother QL-570 LE (62mm x 29mm)
Lê componentes.txt e imprime uma etiqueta por linha com o nome centrado.

Uso:
    python imprimir_etiquetas.py                  (imprime TODAS)
    python imprimir_etiquetas.py --preview         (so gera PDFs, nao imprime)
    python imprimir_etiquetas.py --filter PARAFUSO (so linhas que contem PARAFUSO)
    python imprimir_etiquetas.py --dry-run         (mostra o que faria)
"""

import os
import sys
import json
import time
import subprocess
import tempfile
import argparse

# ---------- Caminhos ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INSTALLER_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "installer")
COMPONENTES_FILE = os.path.join(SCRIPT_DIR, "componentes.txt")
CONFIG_FILE = os.path.join(INSTALLER_DIR, "config.json")

# Dimensoes da etiqueta em cm
LABEL_W = 6.2
LABEL_H = 2.9


def safe_print(msg=""):
    try:
        print(msg)
    except Exception:
        try:
            print(msg.encode("ascii", errors="replace").decode("ascii"))
        except Exception:
            pass


def load_componentes(filepath, filtro=""):
    """Le o ficheiro e devolve lista de nomes unicos (sem linhas vazias)."""
    if not os.path.exists(filepath):
        safe_print("ERRO: Ficheiro nao encontrado: %s" % filepath)
        sys.exit(1)

    with open(filepath, "r", encoding="utf-8") as f:
        linhas = f.readlines()

    nomes = []
    vistos = set()
    for linha in linhas:
        nome = linha.strip()
        if not nome:
            continue
        if filtro and filtro.upper() not in nome.upper():
            continue
        chave = nome.upper()
        if chave not in vistos:
            vistos.add(chave)
            nomes.append(nome)

    return nomes


def escape_html(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def gerar_html_etiqueta(nome):
    """Gera HTML para uma etiqueta 62x29mm com o nome centrado."""
    w = LABEL_W
    h = LABEL_H
    nome_esc = escape_html(nome)

    # Ajustar tamanho da fonte com base no comprimento
    if len(nome) <= 10:
        font_size = "14pt"
    elif len(nome) <= 18:
        font_size = "11pt"
    elif len(nome) <= 25:
        font_size = "9pt"
    else:
        font_size = "7.5pt"

    html = (
        '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        '@page { size: %(w)scm %(h)scm; margin: 0; } '
        '* { margin: 0; padding: 0; box-sizing: border-box; } '
        'body { width: %(w)scm; height: %(h)scm; display: flex; '
        'align-items: center; justify-content: center; '
        'font-family: Arial, Helvetica, sans-serif; overflow: hidden; } '
        '.nome { font-size: %(fs)s; font-weight: bold; text-align: center; '
        'padding: 0.1cm; word-wrap: break-word; } '
        '</style></head><body>'
        '<div class="nome">%(nome)s</div>'
        '</body></html>'
    ) % {"w": w, "h": h, "fs": font_size, "nome": nome_esc}

    return html


def html_para_pdf_google(html_content):
    """
    Converte HTML para PDF. Tenta usar a mesma abordagem que o print_agent:
    gera um HTML file e converte via ferramenta disponivel.
    Para simplicidade, guarda como HTML e usa SumatraPDF para imprimir.
    """
    # Guardar como ficheiro HTML temporario
    fd, html_path = tempfile.mkstemp(suffix=".html", prefix="etiqueta_")
    os.close(fd)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return html_path


def encontrar_sumatra():
    """Procura SumatraPDF."""
    # 1. config.json do installer
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                path = cfg.get("sumatra_path", "")
                if path and os.path.exists(path):
                    return path
        except Exception:
            pass

    # 2. Pasta tools do installer
    tools_path = os.path.join(INSTALLER_DIR, "tools", "SumatraPDF.exe")
    if os.path.exists(tools_path):
        return tools_path

    # 3. Localizacoes comuns
    locais = [
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "SumatraPDF", "SumatraPDF.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "SumatraPDF", "SumatraPDF.exe"),
        r"C:\SumatraPDF\SumatraPDF.exe",
    ]
    for p in locais:
        if p and os.path.isfile(p):
            return p

    return ""


def encontrar_impressora():
    """Detecta impressora a partir de config.json ou wmic."""
    # 1. config.json
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                name = cfg.get("printer_name", "")
                if name and name != "DEFAULT":
                    return name
        except Exception:
            pass

    # 2. Detectar via wmic (preferir Brother)
    try:
        result = subprocess.run(
            ["wmic", "printer", "get", "name"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            for line in result.stdout.strip().split("\n"):
                name = line.strip()
                if name and name.lower() != "name" and "brother" in name.lower():
                    return name
    except Exception:
        pass

    return ""


def imprimir_html(html_path, sumatra, impressora):
    """Imprime um ficheiro HTML via SumatraPDF."""
    cmd = [sumatra, "-print-to", impressora, "-silent", "-print-settings", "fit", html_path]
    if not impressora:
        cmd = [sumatra, "-print-to-default", "-silent", "-print-settings", "fit", html_path]

    creation_flags = 0
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        creation_flags = subprocess.CREATE_NO_WINDOW

    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=30,
        creationflags=creation_flags
    )
    return result.returncode == 0


def main():
    parser = argparse.ArgumentParser(description="Imprimir etiquetas de componentes")
    parser.add_argument("--preview", action="store_true", help="So gerar ficheiros, nao imprimir")
    parser.add_argument("--dry-run", action="store_true", help="Mostrar o que faria sem executar")
    parser.add_argument("--filter", type=str, default="", help="Filtrar componentes (ex: PARAFUSO)")
    parser.add_argument("--file", type=str, default=COMPONENTES_FILE, help="Ficheiro de componentes")
    parser.add_argument("--delay", type=float, default=1.0, help="Segundos entre impressoes (default: 1)")
    args = parser.parse_args()

    safe_print()
    safe_print("=" * 55)
    safe_print("  IMPRESSAO DE ETIQUETAS - COMPONENTES")
    safe_print("  Etiqueta: %.1f x %.1f cm" % (LABEL_W, LABEL_H))
    safe_print("=" * 55)
    safe_print()

    # Ler componentes
    componentes = load_componentes(args.file, args.filter)
    if not componentes:
        safe_print("  Nenhum componente encontrado.")
        return

    safe_print("  Componentes a imprimir: %d" % len(componentes))
    for i, nome in enumerate(componentes, 1):
        safe_print("    %2d. %s" % (i, nome))
    safe_print()

    if args.dry_run:
        safe_print("  [DRY-RUN] Nenhuma etiqueta foi impressa.")
        return

    # Verificar SumatraPDF e impressora (se nao for preview)
    sumatra = ""
    impressora = ""
    if not args.preview:
        sumatra = encontrar_sumatra()
        if not sumatra:
            safe_print("  [ERRO] SumatraPDF nao encontrado!")
            safe_print("  Use --preview para gerar os ficheiros sem imprimir.")
            return
        safe_print("  SumatraPDF: %s" % sumatra)

        impressora = encontrar_impressora()
        safe_print("  Impressora: %s" % (impressora if impressora else "(predefinida)"))
        safe_print()

        # Confirmar
        resp = input("  Imprimir %d etiquetas? (S/N): " % len(componentes)).strip().upper()
        if resp != "S":
            safe_print("  Cancelado.")
            return

    safe_print()

    # Pasta para preview
    preview_dir = os.path.join(SCRIPT_DIR, "preview")
    if args.preview:
        os.makedirs(preview_dir, exist_ok=True)

    # Gerar e imprimir
    ok = 0
    falhas = 0
    temp_files = []

    for i, nome in enumerate(componentes, 1):
        html = gerar_html_etiqueta(nome)

        if args.preview:
            # Guardar na pasta preview
            safe_name = nome.replace(" ", "_").replace("/", "-")
            html_path = os.path.join(preview_dir, "%02d_%s.html" % (i, safe_name))
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html)
            safe_print("  [%2d/%d] Gerado: %s" % (i, len(componentes), os.path.basename(html_path)))
        else:
            # Imprimir
            html_path = html_para_pdf_google(html)
            temp_files.append(html_path)

            sucesso = imprimir_html(html_path, sumatra, impressora)
            if sucesso:
                ok += 1
                safe_print("  [%2d/%d] OK: %s" % (i, len(componentes), nome))
            else:
                falhas += 1
                safe_print("  [%2d/%d] FALHOU: %s" % (i, len(componentes), nome))

            # Pausa entre impressoes para nao sobrecarregar
            if i < len(componentes):
                time.sleep(args.delay)

    # Limpar temporarios
    for tmp in temp_files:
        try:
            os.remove(tmp)
        except Exception:
            pass

    safe_print()
    safe_print("=" * 55)
    if args.preview:
        safe_print("  CONCLUIDO! Ficheiros em: %s" % preview_dir)
        safe_print("  Abra os .html no browser para pre-visualizar.")
    else:
        safe_print("  CONCLUIDO! %d impressas, %d falharam." % (ok, falhas))
    safe_print("=" * 55)
    safe_print()


if __name__ == "__main__":
    main()
