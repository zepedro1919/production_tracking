"""
Launcher - Arranca o Print Agent e um tunel publico num so processo.
Tenta multiplos servicos de tunnel (Pinggy, LocalTunnel) com auto-reconnect.
100% Python, sem binarios externos. Compativel com Windows 7/8/10/11.
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
    """Instalar pacotes de tunnel se nao existirem."""
    needed = {
        "pinggy": "pinggy",
        "requests": "requests",
    }
    missing = []
    for mod, pkg in needed.items():
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)

    if missing:
        safe_print("  A instalar: %s ..." % ", ".join(missing))
        sp.check_call(
            [sys.executable, "-m", "pip", "install"] + missing,
            stdout=sp.DEVNULL, stderr=sp.DEVNULL
        )
        safe_print("  [OK] Pacotes instalados.")


# ===========================================================
# TUNNEL PROVIDERS
# ===========================================================

def try_pinggy(port):
    """
    Tenta criar tunnel via Pinggy Python SDK.
    Free tier: 60 min timeout, URL aleatorio, sem conta necessaria.
    Retorna (tunnel_obj, public_url) ou levanta excepcao.
    """
    import pinggy

    safe_print("  [Pinggy] A criar tunnel...")
    tunnel = pinggy.start_tunnel(forwardto="localhost:%d" % port)

    # Obter URL publico
    urls = tunnel.urls
    if not urls:
        raise Exception("Pinggy nao devolveu URLs")

    # Preferir HTTPS
    public_url = None
    for u in urls:
        if u.startswith("https://"):
            public_url = u
            break
    if not public_url:
        public_url = urls[0]

    safe_print("  [Pinggy] Tunnel criado: %s" % public_url)
    return tunnel, public_url


def try_localtunnel(port, subdomain=""):
    """
    Tenta criar tunnel via py-localtunnel (localtunnel.me).
    Free: sem conta, sem timeout, mas menos estavel.
    Retorna (tunnel_obj, public_url) ou levanta excepcao.
    """
    # Instalar py-localtunnel se necessario
    try:
        from py_localtunnel.tunnel import Tunnel
    except ImportError:
        sp.check_call(
            [sys.executable, "-m", "pip", "install", "py-localtunnel"],
            stdout=sp.DEVNULL, stderr=sp.DEVNULL
        )
        from py_localtunnel.tunnel import Tunnel

    safe_print("  [LocalTunnel] A criar tunnel...")
    tunnel = Tunnel()
    url = tunnel.get_url(subdomain)

    tunnel_thread = threading.Thread(
        target=tunnel.create_tunnel,
        args=(port, "localhost"),
        daemon=True
    )
    tunnel_thread.start()
    time.sleep(2)

    if url.startswith("http://"):
        url = url.replace("http://", "https://", 1)

    safe_print("  [LocalTunnel] Tunnel criado: %s" % url)
    return tunnel, url


def create_tunnel(port, subdomain=""):
    """
    Tenta criar um tunnel usando os providers disponiveis.
    Ordem: Pinggy (mais robusto) -> LocalTunnel (fallback).
    Retorna (provider_name, tunnel_obj, public_url).
    """
    # --- Tentar Pinggy primeiro ---
    try:
        tunnel, url = try_pinggy(port)
        return "Pinggy", tunnel, url
    except Exception as e:
        safe_print("  [Pinggy] Falhou: %s" % str(e))
        logging.warning("Pinggy falhou: %s", e)

    # --- Fallback: LocalTunnel ---
    try:
        tunnel, url = try_localtunnel(port, subdomain)
        return "LocalTunnel", tunnel, url
    except Exception as e:
        safe_print("  [LocalTunnel] Falhou: %s" % str(e))
        logging.warning("LocalTunnel falhou: %s", e)

    raise Exception("Todos os servicos de tunnel falharam")


def check_tunnel_health(url, timeout=10):
    """Verificar se o tunnel responde."""
    import urllib.request
    try:
        req = urllib.request.Request(url + "/health", method="GET")
        req.add_header("Bypass-Tunnel-Reminder", "true")
        resp = urllib.request.urlopen(req, timeout=timeout)
        return resp.getcode() == 200
    except Exception:
        return False


def main():
    # ==========================================================
    # LOGGING PARA FICHEIRO
    # ==========================================================
    for h in logging.root.handlers[:]:
        logging.root.removeHandler(h)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.root.setLevel(logging.INFO)
    logging.root.addHandler(file_handler)

    for logger_name in ["werkzeug", "urllib3", "requests", "pinggy"]:
        lg = logging.getLogger(logger_name)
        lg.handlers = []
        lg.addHandler(file_handler)
        lg.propagate = False

    safe_print()
    safe_print("=" * 60)
    safe_print("  PRINT AGENT NAUTILUS - LAUNCHER")
    safe_print("  (Tunnel automatico - Windows 7+)")
    safe_print("=" * 60)
    safe_print()

    # -------------------------------------------------------
    # 1. Verificar/instalar dependencias
    # -------------------------------------------------------
    safe_print("  [1/4] A verificar dependencias...")
    ensure_packages()
    safe_print("  [OK] Dependencias disponiveis.")
    safe_print()

    # -------------------------------------------------------
    # 2. Carregar config
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
    safe_print()

    # -------------------------------------------------------
    # 3. Arrancar Print Agent
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
    # 4. Arrancar tunnel com auto-reconnect
    # -------------------------------------------------------
    safe_print("  [3/4] A criar tunnel...")
    safe_print()

    tunnel_obj = None
    public_url = None
    provider = None
    MAX_CONNECT_RETRIES = 5

    for attempt in range(1, MAX_CONNECT_RETRIES + 1):
        try:
            provider, tunnel_obj, public_url = create_tunnel(port, lt_subdomain)
            break
        except Exception as e:
            safe_print("  [AVISO] Tentativa %d/%d falhou: %s" % (attempt, MAX_CONNECT_RETRIES, str(e)))
            if attempt < MAX_CONNECT_RETRIES:
                wait = min(attempt * 5, 30)
                safe_print("  A tentar de novo em %d segundos..." % wait)
                time.sleep(wait)

    if not public_url:
        safe_print()
        safe_print("  [ERRO] Nao foi possivel criar tunnel.")
        safe_print("  Print Agent continua em localhost:%d" % port)
        safe_print("  Log: %s" % LOG_FILE)
        safe_print()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            safe_print("  Sistema parado.")
        return

    # --- Mostrar resultado ---
    safe_print()
    safe_print("=" * 60)
    safe_print()
    safe_print("  TUDO A FUNCIONAR!  (via %s)" % provider)
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
    safe_print("    - NAO feche esta janela!")
    if provider == "Pinggy":
        safe_print("    - Pinggy free: tunnel reconecta automaticamente")
    safe_print("    - O URL muda se reiniciar este script")
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

    # Guardar URL
    url_file = os.path.join(BASE_DIR, "tunnel_url.txt")
    with open(url_file, "w") as f:
        f.write(public_url)
    safe_print("  [INFO] URL guardado em: %s" % url_file)

    safe_print()
    safe_print("  [4/4] Sistema completo. Prima Ctrl+C para parar.")
    safe_print()

    # --- Loop principal com health-check e auto-reconnect ---
    health_interval = 60  # verificar a cada 60 segundos
    last_check = time.time()

    try:
        while True:
            time.sleep(5)
            now = time.time()
            if now - last_check >= health_interval:
                last_check = now
                if not check_tunnel_health(public_url):
                    safe_print()
                    safe_print("  [AVISO] Tunnel parece em baixo. A reconectar...")
                    logging.warning("Tunnel down, reconnecting...")

                    # Tentar parar o antigo
                    if tunnel_obj:
                        try:
                            if hasattr(tunnel_obj, 'stop'):
                                tunnel_obj.stop()
                            elif hasattr(tunnel_obj, 'stop_tunnel'):
                                tunnel_obj.stop_tunnel()
                        except Exception:
                            pass

                    # Reconectar
                    reconnected = False
                    for retry in range(1, 4):
                        try:
                            provider, tunnel_obj, new_url = create_tunnel(port, lt_subdomain)
                            if new_url.startswith("http://"):
                                new_url = new_url.replace("http://", "https://", 1)
                            public_url = new_url
                            reconnected = True

                            safe_print("  [OK] Reconectado via %s!" % provider)
                            safe_print("  >>> %s <<<" % public_url)

                            # Se URL mudou, avisar
                            safe_print()
                            safe_print("  [!] ATENCAO: O URL pode ter mudado!")
                            safe_print("      Verifique se precisa de atualizar o Config.gs")

                            with open(url_file, "w") as f:
                                f.write(public_url)

                            try:
                                process = sp.Popen(["clip"], stdin=sp.PIPE)
                                process.communicate(public_url.encode("utf-8"))
                            except Exception:
                                pass

                            break
                        except Exception as e:
                            safe_print("  [AVISO] Reconnect %d/3 falhou: %s" % (retry, str(e)))
                            time.sleep(10)

                    if not reconnected:
                        safe_print("  [ERRO] Nao foi possivel reconectar o tunnel.")
                        safe_print("  Print Agent continua em localhost:%d" % port)

    except KeyboardInterrupt:
        safe_print()
        safe_print("  A encerrar...")
        if tunnel_obj:
            try:
                if hasattr(tunnel_obj, 'stop'):
                    tunnel_obj.stop()
                elif hasattr(tunnel_obj, 'stop_tunnel'):
                    tunnel_obj.stop_tunnel()
            except Exception:
                pass
        safe_print("  Sistema parado.")


if __name__ == "__main__":
    main()
