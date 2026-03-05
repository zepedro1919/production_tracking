"""
Launcher - Arranca o Print Agent e o tunel ngrok num so processo.
Usa pyngrok para gerir o ngrok automaticamente (compativel com Windows 7).
"""

import os
import sys
import json
import time
import threading
import logging

# ==========================================================
# MONKEY-PATCH DEFINITIVO para Win7 (cp850):
# Substituir o metodo emit() do StreamHandler para NUNCA crashar.
# Isto protege TODOS os handlers, incluindo os criados pelo
# Flask/werkzeug/pyngrok DEPOIS do nosso setup.
# ==========================================================
_original_stream_emit = logging.StreamHandler.emit

def _safe_emit(self, record):
    try:
        _original_stream_emit(self, record)
    except (UnicodeEncodeError, UnicodeDecodeError, OSError):
        # Fallback: converter para ASCII e tentar de novo
        try:
            record.msg = str(record.msg).encode("ascii", errors="replace").decode("ascii")
            _original_stream_emit(self, record)
        except Exception:
            pass  # Desistir silenciosamente - nao crashar
    except Exception:
        pass  # Nunca crashar por causa de logging

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


def main():
    # ==========================================================
    # LOGGING VAI APENAS PARA FICHEIRO - nunca para o terminal.
    # O terminal do Win7 (cp850) nao suporta UTF-8 do ngrok e
    # causa crash em logging/__init__.py stream.write().
    # ==========================================================
    for h in logging.root.handlers[:]:
        logging.root.removeHandler(h)

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w")
    file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.root.setLevel(logging.INFO)
    logging.root.addHandler(file_handler)

    # Tambem redirigir loggers conhecidos que possam criar handlers proprios
    for logger_name in ["werkzeug", "pyngrok", "pyngrok.ngrok", "pyngrok.process"]:
        lg = logging.getLogger(logger_name)
        lg.handlers = []
        lg.addHandler(file_handler)
        lg.propagate = False

    safe_print()
    safe_print("=" * 60)
    safe_print("  PRINT AGENT NAUTILUS - LAUNCHER")
    safe_print("=" * 60)
    safe_print()

    # -------------------------------------------------------
    # 1. Verificar/instalar pyngrok
    # -------------------------------------------------------
    safe_print("  [1/4] A verificar pyngrok...")
    try:
        from pyngrok import ngrok, conf
    except ImportError:
        safe_print("  pyngrok nao encontrado. A instalar...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyngrok"])
        from pyngrok import ngrok, conf

    safe_print("  [OK] pyngrok disponivel.")
    safe_print()

    # -------------------------------------------------------
    # 2. Carregar config para obter a porta
    # -------------------------------------------------------
    port = 5555
    if os.path.exists(CONFIG_JSON):
        try:
            with open(CONFIG_JSON, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                port = cfg.get("port", 5555)
        except Exception:
            pass

    # -------------------------------------------------------
    # 3. Detectar Windows e configurar ngrok
    #    Windows 7 = ngrok v2 (v3 nao suporta Win7)
    #    Windows 10+ = ngrok v3
    # -------------------------------------------------------
    import platform
    import subprocess as sp

    win_ver = platform.version()
    major_ver = int(win_ver.split(".")[0]) if win_ver else 10
    need_v2 = major_ver < 10

    if need_v2:
        safe_print("  Windows antigo detectado (version %s)" % win_ver)
        safe_print("  A usar ngrok v2 (compativel com Windows 7/8)")

        from pyngrok import installer
        default_ngrok_dir = installer.get_default_ngrok_dir()
        default_ngrok_bin = os.path.join(default_ngrok_dir, "ngrok.exe")

        if os.path.exists(default_ngrok_bin):
            safe_print("  ngrok encontrado: %s" % default_ngrok_bin)
            try:
                result = sp.run([default_ngrok_bin, "version"],
                               capture_output=True, text=True, timeout=10)
                ver_output = result.stdout.strip() + result.stderr.strip()
                safe_print("  Versao: %s" % ver_output)
                if "version 3" in ver_output or result.returncode != 0:
                    safe_print("  Incompativel (v3) - a apagar...")
                    os.remove(default_ngrok_bin)
                    safe_print("  [OK] Apagado. Vai descarregar v2.")
                else:
                    safe_print("  Ja e v2 - OK!")
            except Exception:
                safe_print("  Nao funciona - a apagar...")
                try:
                    os.remove(default_ngrok_bin)
                except Exception:
                    pass

        # Callback que vai para ficheiro, nunca para terminal
        def ngrok_log_cb(log):
            logging.info("[ngrok] %s", str(log.line).rstrip())

        pyngrok_config = conf.PyngrokConfig(
            ngrok_version="v2",
            ngrok_path=default_ngrok_bin,
            log_event_callback=ngrok_log_cb,
            monitor_thread=True
        )
        conf.set_default(pyngrok_config)
        safe_print("  ngrok path: %s" % default_ngrok_bin)
    else:
        safe_print("  Windows moderno (version %s) - ngrok v3" % win_ver)
    safe_print()

    # -------------------------------------------------------
    # 4. Arrancar Print Agent em thread separada
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

            # Redirigir logs do werkzeug para ficheiro
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
    # 5. Arrancar tunel ngrok
    # -------------------------------------------------------
    safe_print("  [3/4] A criar tunel ngrok...")
    safe_print()

    try:
        tunnel = ngrok.connect(port, "http")
        public_url = tunnel.public_url

        if public_url.startswith("http://"):
            public_url = public_url.replace("http://", "https://", 1)

        safe_print()
        safe_print("=" * 60)
        safe_print()
        safe_print("  TUDO A FUNCIONAR!")
        safe_print()
        safe_print("  URL DO NGROK (copie este URL):")
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
        safe_print("    - ngrok gratuito NAO tem timeout")
        safe_print("    - MAS o URL muda se reiniciar este script")
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
        url_file = os.path.join(BASE_DIR, "ngrok_url.txt")
        with open(url_file, "w") as f:
            f.write(public_url)
        safe_print("  [INFO] URL guardado em: %s" % url_file)

        safe_print()
        safe_print("  [4/4] Sistema completo. Prima Ctrl+C para parar.")
        safe_print()

        # Manter vivo
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            safe_print()
            safe_print("  A encerrar...")
            ngrok.kill()
            safe_print("  Sistema parado.")

    except Exception as e:
        safe_print()
        safe_print("  [ERRO] Falha ao criar tunel ngrok: %s" % str(e))
        safe_print()
        safe_print("  Possiveis causas:")
        safe_print("    - Sem ligacao a internet")
        safe_print("    - ngrok bloqueado pela firewall")
        safe_print()
        safe_print("  Print Agent continua em localhost:%d" % port)
        safe_print("  Log detalhado em: %s" % LOG_FILE)
        safe_print()

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            safe_print("  Sistema parado.")


if __name__ == "__main__":
    main()
