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

# Pasta base = pasta onde este script esta
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_JSON = os.path.join(BASE_DIR, "config.json")

def main():
    # Setup logging - usar encoding seguro para Win7
    # O terminal do Win7 (cp850) pode nao suportar certos chars do ngrok
    try:
        import io
        safe_stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True
        )
    except Exception:
        safe_stdout = sys.stdout

    handler = logging.StreamHandler(safe_stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    # Aplicar handler seguro a TODOS os loggers (incluindo pyngrok)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    # Remover handlers existentes
    for h in root_logger.handlers[:]:
        root_logger.removeHandler(h)
    root_logger.addHandler(handler)

    print()
    print("=" * 60)
    print("  PRINT AGENT NAUTILUS - LAUNCHER")
    print("=" * 60)
    print()

    # -------------------------------------------------------
    # 1. Verificar/instalar pyngrok
    # -------------------------------------------------------
    print("  [1/3] A verificar pyngrok...")
    try:
        from pyngrok import ngrok, conf
    except ImportError:
        print("  pyngrok nao encontrado. A instalar...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyngrok"])
        from pyngrok import ngrok, conf

    print("  [OK] pyngrok disponivel.")
    print()

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
    # 3. Detectar Windows e configurar ngrok adequadamente
    #    Windows 7 = ngrok v2 (v3 nao suporta Win7)
    #    Windows 10+ = ngrok v3
    # -------------------------------------------------------
    import platform
    import subprocess as sp

    win_ver = platform.version()  # e.g. "6.1.7601" for Win7, "10.0.xxxxx" for Win10+
    major_ver = int(win_ver.split(".")[0]) if win_ver else 10
    need_v2 = major_ver < 10

    if need_v2:
        print(f"  Windows antigo detectado (version {win_ver})")
        print("  A usar ngrok v2 (compativel com Windows 7/8)")

        # Caminho onde o pyngrok guarda o ngrok por defeito
        from pyngrok import installer
        default_ngrok_dir = installer.get_default_ngrok_dir()
        default_ngrok_bin = os.path.join(default_ngrok_dir, "ngrok.exe")

        # Se ja existe um ngrok no caminho padrao, verificar se e v3
        # Se for v3, apagar para forcar re-download da v2
        if os.path.exists(default_ngrok_bin):
            print(f"  ngrok existente encontrado em: {default_ngrok_bin}")
            try:
                result = sp.run([default_ngrok_bin, "version"],
                               capture_output=True, text=True, timeout=10)
                ver_output = result.stdout.strip() + result.stderr.strip()
                print(f"  Versao existente: {ver_output}")
                # ngrok v3 output: "ngrok version 3.x.x"
                # ngrok v2 output: "ngrok version 2.x.x"
                if "version 3" in ver_output or result.returncode != 0:
                    print("  E versao 3 (incompativel) - a apagar para forcar download v2...")
                    os.remove(default_ngrok_bin)
                    print("  [OK] ngrok v3 apagado.")
                else:
                    print("  Ja e v2 - OK!")
            except Exception as e:
                print(f"  ngrok existente nao funciona ({e}) - a apagar...")
                try:
                    os.remove(default_ngrok_bin)
                except Exception:
                    pass

        # Configurar pyngrok para v2 com caminho explicito
        # log_event_callback evita crash de encoding no Win7
        def safe_log_callback(log):
            try:
                print(f"  [ngrok] {log.msg}")
            except Exception:
                pass

        pyngrok_config = conf.PyngrokConfig(
            ngrok_version="v2",
            ngrok_path=default_ngrok_bin,
            log_event_callback=safe_log_callback
        )
        conf.set_default(pyngrok_config)
        print(f"  ngrok sera instalado em: {default_ngrok_bin}")
    else:
        print(f"  Windows moderno detectado (version {win_ver})")
        print("  A usar ngrok v3")
    print()

    # -------------------------------------------------------
    # 4. Arrancar Print Agent em thread separada
    # -------------------------------------------------------
    print("  [2/3] A arrancar Print Agent na porta", port, "...")
    print()

    def run_print_agent():
        """Importa e arranca o print_agent.py"""
        try:
            # Adicionar pasta base ao path
            if BASE_DIR not in sys.path:
                sys.path.insert(0, BASE_DIR)

            # Importar e configurar
            import print_agent
            print_agent.config = print_agent.load_config()
            print_agent.print_banner()

            # Arrancar Flask (sem reloader para nao interferir)
            print_agent.app.run(
                host="0.0.0.0",
                port=port,
                debug=False,
                use_reloader=False
            )
        except Exception as e:
            logging.error(f"Erro ao arrancar Print Agent: {e}")

    agent_thread = threading.Thread(target=run_print_agent, daemon=True)
    agent_thread.start()

    # Esperar que o Flask arranque
    time.sleep(3)

    # Verificar se arrancou
    import urllib.request
    try:
        resp = urllib.request.urlopen(f"http://localhost:{port}/health", timeout=5)
        if resp.getcode() == 200:
            print("  [OK] Print Agent a correr!")
        else:
            print("  [AVISO] Print Agent pode nao estar a responder...")
    except Exception:
        print("  [AVISO] Print Agent ainda a arrancar... (pode demorar)")
    print()

    # -------------------------------------------------------
    # 5. Arrancar tunel ngrok
    # -------------------------------------------------------
    print("  [3/3] A criar tunel ngrok...")
    print()

    try:
        # Abrir tunel
        tunnel = ngrok.connect(port, "http")
        public_url = tunnel.public_url

        # Garantir HTTPS
        if public_url.startswith("http://"):
            public_url = public_url.replace("http://", "https://", 1)

        print()
        print("=" * 60)
        print()
        print("  TUDO A FUNCIONAR!")
        print()
        print("  URL DO NGROK (copie este URL):")
        print()
        print(f"  >>> {public_url} <<<")
        print()
        print("=" * 60)
        print()
        print("  PROXIMO PASSO:")
        print("    1. Copie o URL acima")
        print("    2. Abra o Google Apps Script")
        print("    3. No Config.gs, altere:")
        print(f'       PRINT_AGENT_URL: "{public_url}"')
        print("    4. Guarde e EDITE a implementacao existente")
        print("       (NAO crie uma nova implementacao!)")
        print()
        print("  NOTAS:")
        print("    - ngrok gratuito NAO tem timeout (corre infinitamente)")
        print("    - MAS o URL muda se reiniciar este script")
        print("    - NAO feche esta janela!")
        print()
        print("=" * 60)

        # Tentar copiar para clipboard
        try:
            import subprocess
            process = subprocess.Popen(
                ["clip"],
                stdin=subprocess.PIPE
            )
            process.communicate(public_url.encode("utf-8"))
            print()
            print("  [INFO] URL copiado para a area de transferencia!")
        except Exception:
            pass

        # Guardar URL num ficheiro para referencia
        url_file = os.path.join(BASE_DIR, "ngrok_url.txt")
        with open(url_file, "w") as f:
            f.write(public_url)
        print(f"  [INFO] URL guardado em: {url_file}")

        print()
        print("  Prima Ctrl+C para parar o sistema.")
        print()

        # Manter vivo
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print()
            print("  A encerrar...")
            ngrok.kill()
            print("  Sistema parado.")

    except Exception as e:
        print()
        print(f"  [ERRO] Falha ao criar tunel ngrok: {e}")
        print()
        print("  Possiveis causas:")
        print("    - Sem ligacao a internet")
        print("    - ngrok bloqueado pela firewall")
        print("    - Versao do ngrok incompativel")
        print()
        print("  O Print Agent continua a correr em localhost:", port)
        print("  Pode tentar correr o ngrok manualmente.")
        print()

        # Manter o print agent a correr
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("  Sistema parado.")


if __name__ == "__main__":
    main()
