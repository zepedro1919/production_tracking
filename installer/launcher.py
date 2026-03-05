"""
Launcher - Arranca o Print Agent e o tunel LocalTunnel num so processo.
Usa py-localtunnel (100%% Python, sem binarios externos).
Compativel com Windows 7/8/10/11.
"""

import os
import sys
import json
import time
import threading
import logging
import subprocess as sp

# ==========================================================
# MONKEY-PATCH DEFINITIVO para Win7 (cp850):
# Substituir o metodo emit() do StreamHandler para NUNCA crashar.
# Isto protege TODOS os handlers, incluindo os criados pelo
# Flask/werkzeug DEPOIS do nosso setup.
# ==========================================================
_original_stream_emit = logging.StreamHandler.emit

def _safe_emit(self, record):
    try:
        _original_stream_emit(self, record)
    except (UnicodeEncodeError, UnicodeDecodeError, OSError):
        try:
            record.msg = str(record.msg).encode("ascii", errors="replace").decode("ascii")
            _original_stream_emit(self, record)
        except Exception:
            pass
    except Exception:
        pass

logging.StreamHandler.emit = _safe_emit

# Pasta base = pasta onde este script esta
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_JSON = os.path.join(BASE_DIR, "config.json")
LOG_FILE = os.path.join(BASE_DIR, "launcher.log")


def safe_print(msg=""):
    """Print que nunca crashoa, mesmo com chars estranhos no Win7."""
    try:
        print(msg)
    except Exception:
        try:
            print(msg.encode("ascii", errors="replace").decode("ascii"))
        except Exception:
            pass


def ensure_packages():
    """Instalar py-localtunnel e requests se nao existirem."""
    missing = []
    try:
        import py_localtunnel  # noqa: F401
    except ImportError:
        missing.append("py-localtunnel")
    try:
        import requests  # noqa: F401
    except ImportError:
        missing.append("requests")

    if missing:
        safe_print("  A instalar: %s ..." % ", ".join(missing))
        sp.check_call(
            [sys.executable, "-m", "pip", "install"] + missing,
            stdout=sp.DEVNULL, stderr=sp.DEVNULL
        )
        safe_print("  [OK] Pacotes instalados.")


def start_localtunnel(port, subdomain=""):
    """
    Arranca o localtunnel e retorna o URL publico.
    Usa a API Python do py-localtunnel directamente.
    O tunnel corre em background (threads internas).
    Retorna (tunnel_obj, public_url).
    """
    from py_localtunnel.tunnel import Tunnel

    tunnel = Tunnel()
    url = tunnel.get_url(subdomain)
    safe_print("  URL obtido: %s" % url)

    # create_tunnel() bloqueia - correr em thread
    tunnel_thread = threading.Thread(
        target=tunnel.create_tunnel,
        args=(port, "localhost"),
        daemon=True
    )
    tunnel_thread.start()
    time.sleep(1)

    return tunnel, url


def main():
    # ==========================================================
    # LOGGING VAI APENAS PARA FICHEIRO - nunca para o terminal.
    # O terminal do Win7 (cp850) nao suporta UTF-8 e
    # causa crash em logging/__init__.py stream.write().
    # ==========================================================
    for h in logging.root.handlers[:]:
        logging.root.removeHandler(h)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.root.setLevel(logging.INFO)
    logging.root.addHandler(file_handler)

    for logger_name in ["werkzeug", "urllib3", "requests"]:
        lg = logging.getLogger(logger_name)
        lg.handlers = []
        lg.addHandler(file_handler)
        lg.propagate = False

    safe_print()
    safe_print("=" * 60)
    safe_print("  PRINT AGENT NAUTILUS - LAUNCHER")
    safe_print("  (LocalTunnel - compativel com Windows 7+)")
    safe_print("=" * 60)
    safe_print()

    # -------------------------------------------------------
    # 1. Verificar/instalar dependencias
    # -------------------------------------------------------
    safe_print("  [1/4] A verificar dependencias...")
    ensure_packages()
    safe_print("  [OK] py-localtunnel e requests disponiveis.")
    safe_print()

    # -------------------------------------------------------
    # 2. Carregar config para obter a porta
    # -------------------------------------------------------
    port = 5555
    lt_subdomain = ""
    if os.path.exists(CONFIG_JSON):
        try:
            with open(CONFIG_JSON, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                port = cfg.get("port", 5555)
                lt_subdomain = cfg.get("localtunnel_subdomain", "")
        except Exception:
            pass
    safe_print("  Porta: %d" % port)
    if lt_subdomain:
        safe_print("  Subdominio pedido: %s" % lt_subdomain)
    safe_print()

    # -------------------------------------------------------
    # 3. Arrancar Print Agent em thread separada
    # -------------------------------------------------------
    safe_print("  [2/4] A arrancar Print Agent na porta %d ..." % port)
    safe_print()

    def run_print_agent():
        try:
            if BASE_DIR not in sys.path:
                sys.path.insert(0, BASE_DIR)

            import print_agent
            print_agent.config = print_agent.load_config()
            print_agent.print_banner()

            wlog = logging.getLogger("werkzeug")
            wlog.handlers = []
            wlog.addHandler(file_handler)
            wlog.propagate = False

            print_agent.app.run(
                host="0.0.0.0",
                port=port,
                debug=False,
                use_reloader=False
            )
        except Exception as e:
            logging.error("Erro Print Agent: %s", e)
            safe_print("  [ERRO] Print Agent: %s" % str(e))

    agent_thread = threading.Thread(target=run_print_agent, daemon=True)
    agent_thread.start()
    time.sleep(3)

    # Verificar se arrancou
    import urllib.request
    try:
        resp = urllib.request.urlopen("http://localhost:%d/health" % port, timeout=5)
        if resp.getcode() == 200:
            safe_print("  [OK] Print Agent a correr!")
        else:
            safe_print("  [AVISO] Print Agent pode nao estar a responder...")
    except Exception:
        safe_print("  [AVISO] Print Agent ainda a arrancar...")
    safe_print()

    # -------------------------------------------------------
    # 4. Arrancar tunel LocalTunnel
    # -------------------------------------------------------
    safe_print("  [3/4] A criar tunel LocalTunnel...")
    safe_print()

    tunnel_obj = None
    MAX_RETRIES = 3

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            tunnel_obj, public_url = start_localtunnel(port, lt_subdomain)

            # Garantir HTTPS
            if public_url.startswith("http://"):
                public_url = public_url.replace("http://", "https://", 1)

            safe_print()
            safe_print("=" * 60)
            safe_print()
            safe_print("  TUDO A FUNCIONAR!")
            safe_print()
            safe_print("  URL DO TUNNEL (copie este URL):")
            safe_print()
            safe_print("  >>> %s <<<" % public_url)
            safe_print()
            safe_print("=" * 60)
            safe_print()
            safe_print("  PROXIMO PASSO:")
            safe_print("    1. Copie o URL acima")
            safe_print("    2. Abra o Google Apps Script")
            safe_print("    3. No Config.gs, altere:")
            safe_print('       PRINT_AGENT_URL: "%s"' % public_url)
            safe_print("    4. Guarde e EDITE a implementacao existente")
            safe_print("       (NAO crie uma nova implementacao!)")
            safe_print()
            safe_print("  NOTAS:")
            safe_print("    - LocalTunnel e gratuito e sem conta")
            safe_print("    - O URL muda se reiniciar este script")
            safe_print("    - NAO feche esta janela!")
            safe_print()
            safe_print("=" * 60)

            # Copiar para clipboard
            try:
                process = sp.Popen(["clip"], stdin=sp.PIPE)
                process.communicate(public_url.encode("utf-8"))
                safe_print()
                safe_print("  [INFO] URL copiado para a area de transferencia!")
            except Exception:
                pass

            # Guardar URL em ficheiro
            url_file = os.path.join(BASE_DIR, "tunnel_url.txt")
            with open(url_file, "w") as f:
                f.write(public_url)
            safe_print("  [INFO] URL guardado em: %s" % url_file)

            safe_print()
            safe_print("  [4/4] Sistema completo. Prima Ctrl+C para parar.")
            safe_print()
            break  # Sucesso, sair do loop

        except Exception as e:
            safe_print("  [AVISO] Tentativa %d/%d falhou: %s" % (attempt, MAX_RETRIES, str(e)))
            if attempt < MAX_RETRIES:
                safe_print("  A tentar de novo em 5 segundos...")
                time.sleep(5)
            else:
                safe_print()
                safe_print("  [ERRO] Falha ao criar tunel apos %d tentativas." % MAX_RETRIES)
                safe_print()
                safe_print("  Possiveis causas:")
                safe_print("    - Sem ligacao a internet")
                safe_print("    - localtunnel.me indisponivel")
                safe_print()
                safe_print("  Print Agent continua em localhost:%d" % port)
                safe_print("  Log detalhado em: %s" % LOG_FILE)
                safe_print()

    # Manter vivo
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        safe_print()
        safe_print("  A encerrar...")
        if tunnel_obj:
            try:
                tunnel_obj.stop_tunnel()
            except Exception:
                pass
        safe_print("  Sistema parado.")


if __name__ == "__main__":
    main()
