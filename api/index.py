import os
import sys
import json
import re
import tempfile
import urllib.parse
from http.server import BaseHTTPRequestHandler

# Asegurar que el directorio raíz está en sys.path para importar extractor, generador_excel y storage_manager
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import extractor
import generador_excel
import storage_manager


class handler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        url_parts = urllib.parse.urlparse(self.path)
        url_path = url_parts.path

        # API: Obtener albaranes compartidos por todos los usuarios
        if url_path.endswith("/api/albaranes") or url_path == "/api/albaranes":
            try:
                items = storage_manager.get_shared_items()
                backend = storage_manager.get_backend_info()
                res = {
                    "ok": True,
                    "items": items,
                    "count": len(items),
                    "backend": backend
                }
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # API: Servir archivo PDF original para previsualización
        if url_path.endswith("/api/pdf") or url_path == "/api/pdf":
            query = urllib.parse.parse_qs(url_parts.query)
            filename = query.get("file", [""])[0]
            if not filename:
                self.send_error(400, "Falta parámetro 'file'")
                return

            safe_name = os.path.basename(filename)
            candidate_dirs = [
                os.path.join(DIRECTORY, "archivos_pdf"),
                os.path.join(DIRECTORY, "muestras"),
                os.path.join(DIRECTORY, "temp_uploads")
            ]
            found_path = None
            clean_target = safe_name.lower().split(" (")[0].strip()

            for c_dir in candidate_dirs:
                if not os.path.exists(c_dir):
                    continue
                p = os.path.join(c_dir, safe_name)
                if os.path.exists(p) and os.path.isfile(p):
                    found_path = p
                    break
                for fname in os.listdir(c_dir):
                    if fname.lower() == clean_target or fname.lower() == safe_name.lower() or fname.lower().startswith(clean_target):
                        found_path = os.path.join(c_dir, fname)
                        break
                if found_path:
                    break

            if found_path:
                try:
                    with open(found_path, "rb") as f:
                        pdf_data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", "application/pdf")
                    self.send_header("Content-Disposition", f'inline; filename="{safe_name}"')
                    self.send_header("Content-Length", str(len(pdf_data)))
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(pdf_data)
                    return
                except Exception as e:
                    self.send_error(500, f"Error leyendo PDF: {e}")
                    return
            else:
                self.send_response(404)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": f"PDF '{safe_name}' no encontrado en el servidor"}, ensure_ascii=False).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        url_parts = urllib.parse.urlparse(self.path)
        url_path = url_parts.path

        # 0. API: Login / Autenticación
        if url_path.endswith("/api/login") or url_path == "/api/login":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8")) if body else {}

                username = str(payload.get("username", "")).strip().lower()
                password = str(payload.get("password", "")).strip()

                if username == "apepgroup" and password == "adminapep":
                    res = {
                        "ok": True,
                        "user": "apepgroup",
                        "token": "apep_session_valid_2026",
                        "message": "Autenticación exitosa"
                    }
                    self.send_response(200)
                else:
                    res = {
                        "ok": False,
                        "error": "Usuario o contraseña incorrectos"
                    }
                    self.send_response(401)

                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

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

                            # Guardar copia persistente en archivos_pdf para previsualizaciones futuras
                            try:
                                target_dir = os.path.join(DIRECTORY, "archivos_pdf")
                                os.makedirs(target_dir, exist_ok=True)
                                dest_path = os.path.join(target_dir, os.path.basename(filename_match))
                                with open(dest_path, "wb") as f_dest:
                                    f_dest.write(content)
                            except Exception as e_save:
                                print(f"Aviso guardando en archivos_pdf: {e_save}")

                            try:
                                items_data = extractor.extract_albaran_data(temp_path, filename=filename_match)
                                if isinstance(items_data, list):
                                    for it in items_data:
                                        if not it.get("nombre_archivo_origen"):
                                            it["nombre_archivo_origen"] = filename_match
                                        if it.get("total_albaranes_en_pdf", 1) <= 1:
                                            it["archivo"] = filename_match
                                        results.append(it)
                                else:
                                    items_data["archivo"] = filename_match
                                    results.append(items_data)
                            finally:
                                if os.path.exists(temp_path):
                                    try:
                                        os.remove(temp_path)
                                    except Exception:
                                        pass

                    # Guardar automáticamente en la memoria compartida
                    if results:
                        storage_manager.add_shared_items(results)
                        self.send_response(200)
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps(results, ensure_ascii=False).encode("utf-8"))
                        return
                    else:
                        # Si no se extrajeron datos por ser un PDF escaneado en servidor sin motor OCR
                        res = {
                            "ok": False,
                            "requires_client_ocr": True,
                            "message": "Documento escaneado sin texto digital. Procesando con motor OCR de alta precisión en navegador."
                        }
                        self.send_response(200)
                        self.send_header("Content-Type", "application/json; charset=utf-8")
                        self.end_headers()
                        self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                        return

            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 2. API: Almacenamiento Compartido (Añadir / Sincronizar)
        elif url_path.endswith("/api/albaranes") or url_path == "/api/albaranes":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8")) if body else {}

                action = payload.get("action")
                if action == "clear_all":
                    storage_manager.clear_shared_items()
                    res = {"ok": True, "count": 0, "items": []}
                elif action == "delete":
                    item_id = payload.get("id")
                    deleted = storage_manager.delete_shared_item(item_id) if item_id else False
                    res = {"ok": deleted, "items": storage_manager.get_shared_items()}
                elif action == "replace_all":
                    new_items = payload.get("items", [])
                    storage_manager.save_shared_items(new_items)
                    res = {"ok": True, "count": len(new_items), "items": new_items}
                else:
                    new_items = payload.get("items") or ([payload["item"]] if "item" in payload else [])
                    updated_list = storage_manager.add_shared_items(new_items)
                    res = {"ok": True, "count": len(updated_list), "items": updated_list}

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        # 3. API: Exportar Excel Estilizado
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

        # 4. API: Exportar CSV UTF-8
        elif url_path.endswith("/api/exportar_csv") or url_path == "/api/exportar_csv":
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

        # 5. API: Guardar PDF para visualizador multiusuario
        elif url_path.endswith("/api/guardar_pdf") or url_path == "/api/guardar_pdf":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)

                raw_filename = self.headers.get("X-Filename", "albaran.pdf")
                filename = urllib.parse.unquote(raw_filename)
                safe_name = os.path.basename(filename)

                target_dir = os.path.join(DIRECTORY, "archivos_pdf")
                os.makedirs(target_dir, exist_ok=True)
                dest_path = os.path.join(target_dir, safe_name)

                with open(dest_path, "wb") as f_dest:
                    f_dest.write(body)

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": True, "filename": safe_name}).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()

    def do_PUT(self):
        url_parts = urllib.parse.urlparse(self.path)
        url_path = url_parts.path

        # API: Actualizar un albarán en la memoria compartida
        if url_path.endswith("/api/albaranes") or url_path == "/api/albaranes":
            try:
                content_length = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_length)
                payload = json.loads(body.decode("utf-8")) if body else {}

                item = payload.get("item")
                if item:
                    updated = storage_manager.update_shared_item(item)
                    res = {"ok": updated, "items": storage_manager.get_shared_items()}
                else:
                    res = {"ok": False, "error": "No se envió el objeto 'item'"}

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()

    def do_DELETE(self):
        url_parts = urllib.parse.urlparse(self.path)
        url_path = url_parts.path
        query = urllib.parse.parse_qs(url_parts.query)

        # API: Eliminar un albarán o vaciar memoria compartida
        if url_path.endswith("/api/albaranes") or url_path == "/api/albaranes":
            try:
                action = query.get("action", [""])[0]
                item_id = query.get("id", [""])[0]

                if action == "clear_all":
                    storage_manager.clear_shared_items()
                    res = {"ok": True, "count": 0, "items": []}
                elif item_id:
                    deleted = storage_manager.delete_shared_item(item_id)
                    res = {"ok": deleted, "items": storage_manager.get_shared_items()}
                else:
                    res = {"ok": False, "error": "Falta parámetro 'id' o 'action=clear_all'"}

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))
                return
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode("utf-8"))
                return

        self.send_response(404)
        self.end_headers()
