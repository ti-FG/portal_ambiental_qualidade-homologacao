# -*- coding: utf-8 -*-
"""
serve_dashboard_hml.py — Servidor HTTP estático simples para o dashboard
em HOMOLOGAÇÃO (servidor 10.1.1.155, porta 8090).

Não é IIS (não instalado nesta máquina) — é um servidor Python leve,
suficiente para servir o dashboard_fg.html estático gerado por
gerar_dashboard_pbix.py. Instalado como serviço Windows via NSSM
(serviço "FGDashAmbiental-HML-Web").

"/" e "" redirecionam para dashboard_fg.html (equivalente ao defaultDocument
do web.config usado em produção via IIS).
"""
import http.server
import socketserver
import os

PORT = 8090
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DOC = "dashboard_fg.html"


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        if self.path in ("/", ""):
            self.path = "/" + DEFAULT_DOC
        return super().do_GET()

    def log_message(self, fmt, *args):
        # Log simples no stdout (capturado pelo NSSM em logs/web-service.log)
        print(f"{self.client_address[0]} - {fmt % args}", flush=True)


class ThreadingServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


if __name__ == "__main__":
    print(f"Servindo {DIRECTORY} em http://0.0.0.0:{PORT}/  (default: {DEFAULT_DOC})", flush=True)
    with ThreadingServer(("0.0.0.0", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("Encerrando...", flush=True)
