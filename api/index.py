import os
import sys
import json
import re
import tempfile
from http.server import BaseHTTPRequestHandler

# Asegurar que el directorio raíz está en sys.path para importar extractor y generador_excel
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import extractor
import generador_excel

class handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        url_path = self.path.split("?")[0]

        # 1. API: Extraer datos de PDFs
        if url_path.endswith("/api/extraer") or url_path == "/api/extraer":
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
                    temp_dir = os.path.join(tempfile.gettempdir(), "albaran_pro_uploads")
                    os.makedirs(temp_dir, exist_ok=True)

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
                                    m = re.search(r'filename="([^"]+)"', line)
                                    if m:
                                        filename_match = m.group(1)
                                        break

                            if not filename_match:
                                filename_match = "albaran_subido.pdf"

                            temp_path = os.path.join(temp_dir, f"up_{os.getpid()}_{os.path.basename(filename_match)}")
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
                    self.end_headers()
                    self.wfile.write(json.dumps(results, ensure_ascii=False).encode("utf-8"))
                    return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 2. API: Exportar Excel Estilizado
        elif url_path.endswith("/api/exportar_excel") or url_path == "/api/exportar_excel":
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
                self.end_headers()
                self.wfile.write(excel_bytes)
                return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 3. API: Exportar CSV UTF-8
        elif url_path.endswith("/api/exportar_csv") or url_path == "/api/exportar_csv":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8"))

                items = payload.get("items", [])
                modelo = payload.get("modelo", "TODOS")
                tipo = payload.get("tipo", "resumen")
                filename = payload.get("filename", "albaranes.csv")

                if not filename.lower().endswith(".csv"):
                    filename += ".csv"

                csv_text = generador_excel.generar_csv_texto(items, modelo=modelo, tipo=tipo)
                csv_bytes = b"\xef\xbb\xbf" + csv_text.encode("utf-8")

                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.end_headers()
                self.wfile.write(csv_bytes)
                return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()

