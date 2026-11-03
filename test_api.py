"""Test rapido de la API de extraccion."""
import urllib.request
import json
import os

pdf_path = r'muestras\albaran_SALTOKI.pdf'
with open(pdf_path, 'rb') as f:
    pdf_bytes = f.read()

boundary = b'TestBoundary999'
crlf = b'\r\n'
body = (
    b'--' + boundary + crlf +
    b'Content-Disposition: form-data; name="files"; filename="albaran_SALTOKI.pdf"' + crlf +
    b'Content-Type: application/pdf' + crlf +
    crlf +
    pdf_bytes +
    crlf +
    b'--' + boundary + b'--' + crlf
)

req = urllib.request.Request(
    'http://localhost:8080/api/extraer',
    data=body,
    headers={
        'Content-Type': 'multipart/form-data; boundary=TestBoundary999',
        'Content-Length': str(len(body))
    }
)

try:
    resp = urllib.request.urlopen(req, timeout=15)
    data = json.loads(resp.read())
    item = data[0]
    print("=== SALTOKI ===")
    print("Modelo:", item.get('modelo'))
    print("Documento:", item.get('numero_documento'))
    print("Pedidos:", item.get('numeros_pedidos'))
    print("Total lineas:", item.get('total_lineas'))
    print("Total unidades:", item.get('total_unidades'))
    print("Mercancia:", item.get('mercancia_resumen'))
    prods = item.get('lineas_productos', [])
    print(f"Productos ({len(prods)}):")
    for p in prods:
        print(f"  L{p.get('linea')}: [{p.get('codigo_articulo')}] {p.get('descripcion','')[:60]} | {p.get('cantidad')} {p.get('unidad')}")
except Exception as e:
    print("ERROR:", e)

print()

# Test Aquafer
pdf_path2 = r'muestras\Cmr_Getafe.pdf'
with open(pdf_path2, 'rb') as f:
    pdf_bytes2 = f.read()

boundary2 = b'TestBoundary998'
body2 = (
    b'--' + boundary2 + crlf +
    b'Content-Disposition: form-data; name="files"; filename="Cmr_Getafe.pdf"' + crlf +
    b'Content-Type: application/pdf' + crlf +
    crlf +
    pdf_bytes2 +
    crlf +
    b'--' + boundary2 + b'--' + crlf
)
req2 = urllib.request.Request(
    'http://localhost:8080/api/extraer',
    data=body2,
    headers={
        'Content-Type': 'multipart/form-data; boundary=TestBoundary998',
        'Content-Length': str(len(body2))
    }
)
try:
    resp2 = urllib.request.urlopen(req2, timeout=15)
    data2 = json.loads(resp2.read())
    item2 = data2[0]
    print("=== AQUAFER ===")
    print("Modelo:", item2.get('modelo'))
    print("Documento:", item2.get('numero_documento'))
    print("Pedidos:", item2.get('numeros_pedidos'))
    print("Total lineas:", item2.get('total_lineas'))
    print("Total unidades:", item2.get('total_unidades'))
    print("Mercancia:", item2.get('mercancia_resumen'))
    prods2 = item2.get('lineas_productos', [])
    print(f"Productos ({len(prods2)}):")
    for p in prods2:
        print(f"  L{p.get('linea')} [{p.get('pedido')}]: [{p.get('codigo_articulo')}] {p.get('descripcion','')[:55]} | {p.get('cantidad')} {p.get('unidad')}")
except Exception as e:
    print("ERROR:", e)
