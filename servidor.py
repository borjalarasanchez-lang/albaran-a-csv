"""
Servidor local ultraligero y robusto para la aplicacion Albaran a CSV Pro.
Procesa archivos PDF digitales y escaneados a traves de extractor.py con soporte OCR.
Genera archivos Excel (.xlsx) altamente estilizados con xlsxwriter.
"""

import os
import sys
import json
import urllib.parse
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
import extractor
import generador_excel

PORT = 8080
if len(sys.argv) > 1:
    try:
        PORT = int(sys.argv[1])
    except ValueError:
        pass

DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class AlbaranHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # Desactivar cache para que los cambios en js y css se reflejen inmediatamente
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_POST(self):
        # 1. API: Extraer datos de PDFs subidos
        if self.path == "/api/extraer":
            try:
                content_type = self.headers.get("Content-Type", "")
                if "multipart/form-data" in content_type:
                    boundary = None
                    for part in content_type.split(";"):
                        part = part.strip()
                        if part.startswith("boundary="):
                            boundary = part.split("=", 1)[1].strip('"').strip("'").encode("utf-8")
                            break

                    if not boundary:
                        raise ValueError("No se encontró el boundary multipart.")

                    content_length = int(self.headers.get("Content-Length", 0))
                    body = self.rfile.read(content_length)

                    delimiter = b"--" + boundary
                    parts = body.split(delimiter)

                    results = []
                    os.makedirs(os.path.join(DIRECTORY, "temp_uploads"), exist_ok=True)

                    for part in parts:
                        if not part or part == b"--\r\n" or part == b"--":
                            continue

                        headers_data, _, content = part.partition(b"\r\n\r\n")
                        if content.endswith(b"\r\n"):
                            content = content[:-2]

                        try:
                            headers_text = headers_data.decode("utf-8")
                        except UnicodeDecodeError:
                            headers_text = headers_data.decode("latin-1", errors="replace")

                        if 'filename="' in headers_text:
                            filename_match = None
                            for line in headers_text.splitlines():
                                if "Content-Disposition:" in line:
                                    import re
                                    m = re.search(r'filename="([^"]+)"', line)
                                    if m:
                                        filename_match = m.group(1)
                                        break

                            if not filename_match:
                                filename_match = "albaran_subido.pdf"

                            temp_path = os.path.join(DIRECTORY, "temp_uploads", f"up_{os.getpid()}_{os.path.basename(filename_match)}")
                            with open(temp_path, "wb") as f:
                                f.write(content)

                            try:
                                data = extractor.extract_albaran_data(temp_path)
                                data["archivo"] = filename_match
                                results.append(data)
                            finally:
                                if os.path.exists(temp_path):
                                    try:
                                        os.remove(temp_path)
                                    except Exception:
                                        pass

                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(json.dumps(results, ensure_ascii=False).encode("utf-8"))
                    return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 2. API: Exportar Excel Estilizado con xlsxwriter
        elif self.path == "/api/exportar_excel":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8"))

                items = payload.get("items", [])
                modelo = payload.get("modelo", "TODOS")
                filename = payload.get("filename", "albaranes_exportados.xlsx")

                if not filename.lower().endswith(".xlsx"):
                    filename += ".xlsx"

                excel_bytes = generador_excel.generar_excel_bytes(items, modelo=modelo)

                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(excel_bytes)
                return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 3. API: Exportar CSV codificado en Windows-1252 (ANSI) para que Excel lo abra con tildes perfectas
        elif self.path == "/api/exportar_csv":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8"))

                items = payload.get("items", [])
                modelo = payload.get("modelo", "TODOS")
                tipo = payload.get("tipo", "detalle")
                filename = payload.get("filename", "albaranes.csv")

                if not filename.lower().endswith(".csv"):
                    filename += ".csv"

                csv_text = generador_excel.generar_csv_texto(items, modelo=modelo, tipo=tipo)
                # UTF-8 con BOM puro (\xef\xbb\xbf) SIN sep=; (sep=; rompe la detección UTF-8 en Excel y genera iconos [][])
                csv_bytes = b"\xef\xbb\xbf" + csv_text.encode("utf-8")

                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(csv_bytes)
                return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        super().do_POST()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

def run_server(port=PORT):
    server_address = ("", port)
    httpd = ThreadingHTTPServer(server_address, AlbaranHandler)
    print("======================================================")
    print(f" Servidor iniciado en: http://localhost:{port}")
    print(" Motor OCR de Windows: ACTIVO (Documentos escaneados soportados)")
    print(" Generador Excel Estilizado (xlsxwriter): ACTIVO")
    print(" Presiona Ctrl + C para detener el servidor.")
    print("======================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")

if __name__ == "__main__":
    run_server(PORT)
