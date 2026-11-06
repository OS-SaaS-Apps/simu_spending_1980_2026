"""Serveur de test : sert le site et les fichiers Excel du dossier public/.

Usage : python server.py [port]   (port par défaut : 8080)
"""
import http.server
import sys
from functools import partial
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "public"


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".js": "text/javascript",
    }

    def end_headers(self):
        # Pas de cache : une correction dans un fichier Excel est visible au rechargement.
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
    server = http.server.ThreadingHTTPServer(("", port), partial(Handler, directory=str(ROOT)))
    print(f"Budget 1980-2026 : http://localhost:{port}")
    server.serve_forever()
