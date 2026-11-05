"""
Gestor de almacenamiento compartido multiusuario ("Misma Memoria").
Permite que todos los usuarios compartan la misma base de albaranes en tiempo real:
1. Upstash Redis / Vercel KV (si UPSTASH_REDIS_REST_URL o KV_REST_API_URL están configurados)
2. Firebase Realtime Database (si FIREBASE_DB_URL está configurado)
3. Archivo JSON local concurrente y atómico (albaranes_compartidos.json) para red local y servidor
4. Memoria interna en proceso con expiración automática de albaranes > 2 meses
"""

import os
import json
import time
import random
import re
import base64
import urllib.request
import urllib.error
from typing import List, Dict, Any, Optional

DOS_MESES_MS = 60 * 24 * 3600 * 1000  # 60 días en milisegundos
SHARED_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "albaranes_compartidos.json")

# Memoria en proceso para entornos efímeros
_MEM_CACHE: List[Dict[str, Any]] = []
_MEM_LAST_UPDATED: float = 0


def _clean_expired(items: List[Dict[str, Any]]) -> tuple[List[Dict[str, Any]], int]:
    """Elimina automáticamente albaranes de más de 2 meses."""
    now_ms = time.time() * 1000
    valid = []
    expired = 0
    for it in items:
        ts = it.get("_timestamp") or (it.get("fecha_creacion_ts", 0) * 1000) or now_ms
        if now_ms - ts > DOS_MESES_MS:
            expired += 1
        else:
            valid.append(it)

    # Limpiar PDFs huérfanos o mayores a 60 días en archivos_pdf/
    try:
        pdf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "archivos_pdf")
        if os.path.isdir(pdf_dir):
            PROTECTED = {"albaran_SALTOKI.pdf", "Cmr_Getafe.pdf", "ejemplo_apep_group.pdf", "ejemplo_aquafer.pdf"}
            for fname in os.listdir(pdf_dir):
                if fname in PROTECTED:
                    continue
                fpath = os.path.join(pdf_dir, fname)
                if os.path.isfile(fpath) and (time.time() - os.path.getmtime(fpath) > (60 * 24 * 3600)):
                    try:
                        os.remove(fpath)
                    except Exception:
                        pass
    except Exception:
        pass

    return valid, expired


def _get_cloud_config() -> Optional[Dict[str, str]]:
    """Detecta proveedores de base de datos en la nube."""
    upstash_url = os.environ.get("UPSTASH_REDIS_REST_URL") or os.environ.get("KV_REST_API_URL")
    upstash_token = os.environ.get("UPSTASH_REDIS_REST_TOKEN") or os.environ.get("KV_REST_API_TOKEN")
    if upstash_url and upstash_token:
        return {
            "type": "upstash",
            "url": upstash_url.rstrip("/"),
            "token": upstash_token
        }

    firebase_url = os.environ.get("FIREBASE_DB_URL") or os.environ.get("SHARED_DB_URL")
    if firebase_url:
        return {
            "type": "firebase",
            "url": firebase_url.rstrip("/")
        }

    return None


def get_backend_info() -> Dict[str, Any]:
    cfg = _get_cloud_config()
    if cfg:
        return {
            "type": cfg["type"],
            "name": "Nube Multi-Usuario (" + cfg["type"].title() + ")",
            "connected": True
        }
    if os.path.exists(SHARED_FILE_PATH) or not os.environ.get("VERCEL"):
        return {
            "type": "local_file",
            "name": "Archivo Compartido en Servidor",
            "connected": True
        }
    return {
        "type": "memory",
        "name": "Memoria en la Nube",
        "connected": True
    }


def get_shared_items() -> List[Dict[str, Any]]:
    """Recupera la lista de todos los albaranes compartidos."""
    global _MEM_CACHE, _MEM_LAST_UPDATED
    cfg = _get_cloud_config()

    # 1. Upstash Redis
    if cfg and cfg["type"] == "upstash":
        try:
            req = urllib.request.Request(
                f"{cfg['url']}/get/albaranes_pro_shared",
                headers={"Authorization": f"Bearer {cfg['token']}"}
            )
            with urllib.request.urlopen(req, timeout=5) as res:
                body = json.loads(res.read().decode("utf-8"))
                val = body.get("result")
                if val:
                    data = json.loads(val) if isinstance(val, str) else val
                    if isinstance(data, list):
                        valid, exp = _clean_expired(data)
                        _MEM_CACHE = valid
                        _MEM_LAST_UPDATED = time.time()
                        return valid
                return []
        except Exception as e:
            print("Error leyendo Upstash:", e)

    # 2. Firebase Realtime Database
    elif cfg and cfg["type"] == "firebase":
        try:
            req = urllib.request.Request(f"{cfg['url']}/albaranes_pro.json")
            with urllib.request.urlopen(req, timeout=5) as res:
                body = json.loads(res.read().decode("utf-8"))
                if isinstance(body, list):
                    valid, exp = _clean_expired(body)
                    _MEM_CACHE = valid
                    _MEM_LAST_UPDATED = time.time()
                    return valid
                elif isinstance(body, dict):
                    valid, exp = _clean_expired(list(body.values()))
                    _MEM_CACHE = valid
                    _MEM_LAST_UPDATED = time.time()
                    return valid
                return []
        except Exception as e:
            print("Error leyendo Firebase:", e)

    # 3. Archivo local compartido en disco
    if os.path.exists(SHARED_FILE_PATH):
        try:
            with open(SHARED_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    for idx, it in enumerate(data):
                        if not it.get("_id"):
                            it["_id"] = f"item_{int(it.get('_timestamp', time.time() * 1000))}_{idx}_{random.randint(1000, 9999)}"
                    valid, exp = _clean_expired(data)
                    _MEM_CACHE = valid
                    _MEM_LAST_UPDATED = time.time()
                    return valid
        except Exception as e:
            print("Error leyendo archivo local:", e)

    return _MEM_CACHE


def save_shared_items(items: List[Dict[str, Any]]) -> bool:
    """Guarda la lista completa de albaranes compartidos de forma atómica."""
    global _MEM_CACHE, _MEM_LAST_UPDATED
    valid, _ = _clean_expired(items)
    _MEM_CACHE = valid
    _MEM_LAST_UPDATED = time.time()

    cfg = _get_cloud_config()

    # 1. Upstash Redis
    if cfg and cfg["type"] == "upstash":
        try:
            payload = json.dumps(valid, ensure_ascii=False)
            req = urllib.request.Request(
                f"{cfg['url']}/set/albaranes_pro_shared",
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Authorization": f"Bearer {cfg['token']}",
                    "Content-Type": "application/json"
                },
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=5) as res:
                return res.status == 200
        except Exception as e:
            print("Error guardando en Upstash:", e)

    # 2. Firebase Realtime Database
    elif cfg and cfg["type"] == "firebase":
        try:
            payload = json.dumps(valid, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                f"{cfg['url']}/albaranes_pro.json",
                data=payload,
                headers={"Content-Type": "application/json"},
                method="PUT"
            )
            with urllib.request.urlopen(req, timeout=5) as res:
                return res.status in (200, 204)
        except Exception as e:
            print("Error guardando en Firebase:", e)

    # 3. Archivo local compartido en disco (atómico con replace)
    try:
        tmp_path = SHARED_FILE_PATH + f".tmp_{os.getpid()}"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(valid, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, SHARED_FILE_PATH)
        return True
    except Exception as e:
        print("Error guardando archivo local:", e)
        return False


def add_shared_items(new_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Añade nuevos albaranes a la memoria compartida."""
    current = get_shared_items()
    existing_ids = {it.get("_id") for it in current if it.get("_id")}

    added = []
    for it in new_items:
        if not it.get("_id"):
            it["_id"] = f"item_{int(time.time() * 1000)}_{random.randint(1000, 9999)}"
        it_id = it["_id"]
        if it_id not in existing_ids:
            if not it.get("_timestamp"):
                it["_timestamp"] = int(time.time() * 1000)
            current.insert(0, it)
            existing_ids.add(it_id)
            added.append(it)

    save_shared_items(current)
    return current


def update_shared_item(updated_item: Dict[str, Any]) -> bool:
    """Actualiza los datos de un albarán existente en la memoria compartida."""
    item_id = updated_item.get("_id")
    if not item_id:
        return False

    current = get_shared_items()
    found = False
    for i, it in enumerate(current):
        if it.get("_id") == item_id:
            current[i] = updated_item
            found = True
            break

    if found:
        save_shared_items(current)
    return found


def delete_shared_item(item_id: str) -> bool:
    """Elimina un albarán por su ID de la memoria compartida."""
    if not item_id:
        return False
    current = get_shared_items()
    init_len = len(current)
    current = [it for it in current if it.get("_id") != item_id]

    if len(current) < init_len:
        save_shared_items(current)
        return True
    return False


def clear_shared_items() -> bool:
    """Vacía toda la memoria compartida para todos los usuarios."""
    return save_shared_items([])


def save_pdf(filename: str, data: bytes) -> bool:
    """Guarda un archivo PDF de forma persistente para acceso multiusuario (disco local y nube)."""
    if not filename or not data:
        return False

    safe_name = os.path.basename(filename).strip()
    saved_disk = False

    # 1. Guardar en disco local (archivos_pdf/) con fallback seguro a /tmp si es de solo lectura
    candidate_dirs = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "archivos_pdf"),
        "/tmp/archivos_pdf",
        "/tmp"
    ]
    for cdir in candidate_dirs:
        try:
            os.makedirs(cdir, exist_ok=True)
            fpath = os.path.join(cdir, safe_name)
            with open(fpath, "wb") as f:
                f.write(data)
            saved_disk = True
            break
        except Exception:
            continue

    # 2. Guardar en la nube si está configurado Upstash o Firebase
    cfg = _get_cloud_config()
    if cfg and cfg["type"] == "upstash":
        try:
            b64_str = base64.b64encode(data).decode("ascii")
            req = urllib.request.Request(
                f"{cfg['url']}/set/pdf_blob:{urllib.parse.quote(safe_name)}",
                data=json.dumps(b64_str).encode("utf-8"),
                headers={
                    "Authorization": f"Bearer {cfg['token']}",
                    "Content-Type": "application/json"
                },
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=10) as res:
                pass
        except Exception as e:
            print(f"Error guardando PDF en Upstash: {e}")

    elif cfg and cfg["type"] == "firebase":
        try:
            b64_str = base64.b64encode(data).decode("ascii")
            safe_key = re.sub(r'[^a-zA-Z0-9_-]', '_', safe_name)
            req = urllib.request.Request(
                f"{cfg['url']}/pdfs/{safe_key}.json",
                data=json.dumps({"name": safe_name, "b64": b64_str, "ts": time.time()}).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="PUT"
            )
            with urllib.request.urlopen(req, timeout=10) as res:
                pass
        except Exception as e:
            print(f"Error guardando PDF en Firebase: {e}")

    return saved_disk


def get_pdf(filename: str) -> Optional[bytes]:
    """Recupera los bytes de un PDF desde disco local o desde almacenamiento en la nube."""
    if not filename:
        return None
    safe_name = os.path.basename(filename).strip()
    clean_target = safe_name.lower().split(" (")[0].strip()

    candidate_dirs = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "archivos_pdf"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "muestras"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_uploads"),
        "/tmp/archivos_pdf",
        "/tmp"
    ]

    for cdir in candidate_dirs:
        if not os.path.isdir(cdir):
            continue
        p = os.path.join(cdir, safe_name)
        if os.path.isfile(p):
            try:
                with open(p, "rb") as f:
                    return f.read()
            except Exception:
                pass
        # Búsqueda por coincidencia insensible a mayúsculas
        try:
            for fname in os.listdir(cdir):
                fl = fname.lower()
                if fl == safe_name.lower() or fl == clean_target or fl.startswith(clean_target):
                    found_p = os.path.join(cdir, fname)
                    if os.path.isfile(found_p):
                        with open(found_p, "rb") as f:
                            return f.read()
        except Exception:
            pass

    # Si no está en disco, buscar en Upstash / Firebase
    cfg = _get_cloud_config()
    if cfg and cfg["type"] == "upstash":
        try:
            req = urllib.request.Request(
                f"{cfg['url']}/get/pdf_blob:{urllib.parse.quote(safe_name)}",
                headers={"Authorization": f"Bearer {cfg['token']}"}
            )
            with urllib.request.urlopen(req, timeout=8) as res:
                body = json.loads(res.read().decode("utf-8"))
                val = body.get("result")
                if val:
                    data = base64.b64decode(val)
                    try:
                        cdir = candidate_dirs[0]
                        os.makedirs(cdir, exist_ok=True)
                        with open(os.path.join(cdir, safe_name), "wb") as f:
                            f.write(data)
                    except Exception:
                        pass
                    return data
        except Exception as e:
            print(f"Error recuperando PDF de Upstash: {e}")

    elif cfg and cfg["type"] == "firebase":
        try:
            safe_key = re.sub(r'[^a-zA-Z0-9_-]', '_', safe_name)
            req = urllib.request.Request(f"{cfg['url']}/pdfs/{safe_key}.json")
            with urllib.request.urlopen(req, timeout=8) as res:
                body = json.loads(res.read().decode("utf-8"))
                if body and "b64" in body:
                    data = base64.b64decode(body["b64"])
                    try:
                        cdir = candidate_dirs[0]
                        os.makedirs(cdir, exist_ok=True)
                        with open(os.path.join(cdir, safe_name), "wb") as f:
                            f.write(data)
                    except Exception:
                        pass
                    return data
        except Exception as e:
            print(f"Error recuperando PDF de Firebase: {e}")

    return None

