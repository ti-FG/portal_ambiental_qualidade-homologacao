# -*- coding: utf-8 -*-
"""
refresh_service.py — Mini serviço HTTP para refresh sob demanda do dashboard.

Roda no servidor 10.1.1.151 escutando a porta 8081.
Quando recebe GET / (ou /refresh), executa o gerar_dashboard_pbix.py
diretamente (sem passar pelo .bat) e responde com o resultado em texto.

GET /status → retorna JSON com a última execução (horário e resultado).

Iniciado automaticamente no boot via tarefa agendada.
"""

import http.server
import socketserver
import subprocess
import sys
import threading
import os
import json
from datetime import datetime

# ───── Configuração ─────────────────────────────────────────────
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
PYTHON_EXE     = sys.executable                                      # mesmo Python que roda o serviço
GENERATOR_PY   = os.path.join(SCRIPT_DIR, "gerar_dashboard_pbix.py")  # script de geração
PORT           = 8081
TIMEOUT        = 300   # 5 minutos máx por refresh

# ───── Estado compartilhado ─────────────────────────────────────
refresh_lock = threading.Lock()
last_run = {"time": None, "status": None, "duration_s": None}


def log(msg):
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


class RefreshHandler(http.server.BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path == "/status":
            return self._json(200, last_run)

        if self.path not in ("/", "/refresh"):
            return self._text(404, "Endpoint nao encontrado.")

        # Impede duas atualizacoes simultaneas
        if not refresh_lock.acquire(blocking=False):
            return self._text(429, "Outra atualizacao ja em andamento. Aguarde alguns segundos.")

        client = self.client_address[0]
        log(f"Refresh solicitado por {client}")
        t0 = datetime.now()

        try:
            # Chama o Python diretamente, sem passar pelo .bat
            # (mais robusto: evita problemas de PATH, encoding e parens em .bat)
            result = subprocess.run(
                [PYTHON_EXE, GENERATOR_PY],
                cwd=SCRIPT_DIR,
                capture_output=True,
                text=True,
                timeout=TIMEOUT,
                encoding="utf-8",
                errors="replace",
            )
            elapsed = (datetime.now() - t0).total_seconds()

            if result.returncode == 0:
                last_run.update({"time": datetime.now().isoformat(),
                                 "status": "ok",
                                 "duration_s": round(elapsed, 1)})
                log(f"  → OK ({elapsed:.1f}s)")
                self._text(200, f"Dashboard atualizado com sucesso em {elapsed:.1f}s.")
            else:
                last_run.update({"time": datetime.now().isoformat(),
                                 "status": "error",
                                 "duration_s": round(elapsed, 1)})
                err = (result.stderr or result.stdout or "")[-800:]
                log(f"  → ERRO rc={result.returncode}")
                self._text(500, f"Falha ao atualizar (rc={result.returncode}):\n\n{err}")

        except subprocess.TimeoutExpired:
            log("  → TIMEOUT")
            self._text(504, f"Atualizacao excedeu {TIMEOUT}s e foi cancelada.")
        except Exception as e:
            log(f"  → EXCECAO: {e}")
            self._text(500, f"Erro inesperado: {e}")
        finally:
            refresh_lock.release()

    # ── helpers ─────────────────────────────────────────────────
    def _text(self, code, msg):
        body = msg.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # silencia logs padrão (usamos log() próprio)


class ThreadingServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


if __name__ == "__main__":
    log(f"Refresh service iniciando em http://0.0.0.0:{PORT}/")
    log(f"Workdir:   {SCRIPT_DIR}")
    log(f"Python:    {PYTHON_EXE}")
    log(f"Generator: {GENERATOR_PY}")
    if not os.path.exists(GENERATOR_PY):
        log(f"[AVISO] gerar_dashboard_pbix.py nao encontrado em {GENERATOR_PY}")
    with ThreadingServer(("0.0.0.0", PORT), RefreshHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            log("Encerrando...")
