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
        it_id = it.get("_id")
        if not it_id or it_id not in existing_ids:
            if not it.get("_timestamp"):
                it["_timestamp"] = int(time.time() * 1000)
            current.insert(0, it)
            if it_id:
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
    current = get_shared_items()
    init_len = len(current)
    current = [it for it in current if it.get("_id") != item_id and it.get("_id") != item_id.split("_")[0]]

    if len(current) < init_len:
        save_shared_items(current)
        return True
    return False


def clear_shared_items() -> bool:
    """Vacía toda la memoria compartida para todos los usuarios."""
    return save_shared_items([])
