"""
Módulo para generación de archivos Excel (.xlsx) profesionales y estilizados.
Usa xlsxwriter con:
- Detalle completo de productos por pedido, código, descripción, cantidad y unidad
- Cabeceras en negrita con colores temáticos (Azul Saltoki, Verde/Teal Aquafer)
- Celdas subrayadas/resaltadas del PDF en amarillo pastel fluorescente (#FEF08A)
- Ajuste automático de anchos de columna sin texto cortado
- Bordes finos elegantes y filas alternas con tono suave
- Autofiltro en cabeceras y panel superior congelado (freeze panes)
- Hojas dedicadas para productos detallados y resúmenes de albarán
"""

import io
from typing import List, Dict, Any, Optional
import xlsxwriter


def clean_str(val: Any) -> str:
  if val is None:
    return ""
  s = str(val).replace("\r", " ").replace("\n", ", ")
  while "  " in s:
    s = s.replace("  ", " ")
  return s.strip()


# ==========================================
# COLUMNAS DETALLE DE PRODUCTOS (LÍNEAS)
# ==========================================
COLUMNAS_PRODUCTOS_SALTOKI = [
    ("archivo", "Archivo PDF", 22),
    ("numero_documento", "Nº Albarán SF", 16),
    ("prod_pedido", "Nº Pedido Cliente", 18),
    ("prod_linea", "Línea", 8),
    ("prod_codigo", "Código / Ref. Artículo", 18),
    ("prod_codigo_sec", "Código Saltoki", 16),
    ("prod_descripcion", "Descripción del Producto", 42),
    ("prod_cantidad", "Cantidad", 12),
    ("prod_unidad", "Unidad", 12),
    ("fecha_albaran_envio", "Fecha Albarán", 14),
    ("fecha_pedido_cliente", "Fecha Pedido", 14),
    ("lugar_descarga", "Lugar Entrega / Descarga", 38),
    ("poblacion_descarga", "Población / CP", 18),
    ("cliente_destinatario", "Cliente / Razón Social", 28),
    ("direccion_logistico", "Dirección Fiscal", 34),
    ("cif_cliente", "CIF Cliente", 14),
    ("proveedor_remitente", "Proveedor Remitente", 28)
]

COLUMNAS_PRODUCTOS_AQUAFER = [
    ("archivo", "Archivo PDF", 22),
    ("numero_documento", "Nº Guía Remesa (GREM)", 22),
    ("numero_cmr", "Nº Carta Porte (CMR)", 16),
    ("prod_pedido", "Nº Pedido Cliente", 18),
    ("prod_linea", "Línea", 8),
    ("prod_codigo", "Código Artículo / Ref.", 18),
    ("prod_descripcion", "Descripción del Producto", 45),
    ("prod_cantidad", "Cantidad", 12),
    ("prod_unidad", "Unidad", 12),
    ("lugar_descarga", "Lugar de Descarga (Entrega)", 36),
    ("poblacion_descarga", "Población / CP", 18),
    ("lugar_carga", "Lugar de Carga (Portugal)", 34),
    ("cliente_destinatario", "Cliente Destinatario", 28),
    ("direccion_logistico", "Dirección Fiscal", 32),
    ("cif_cliente", "NIF Destinatario", 14),
    ("proveedor_remitente", "Proveedor Remitente", 30),
    ("fecha_albaran_envio", "Fecha Emisión", 15),
    ("transportista", "Transportista", 26),
    ("instrucciones_notas", "Notas de Entrega (Español)", 24)
]

COLUMNAS_PRODUCTOS_TODOS = [
    ("archivo", "Archivo PDF", 22),
    ("modelo", "Empresa / Modelo", 16),
    ("numero_documento", "Nº Documento / Albarán", 20),
    ("numero_cmr", "Nº Carta Porte (CMR)", 16),
    ("prod_pedido", "Nº Pedido Cliente", 18),
    ("prod_linea", "Línea", 8),
    ("prod_codigo", "Código / Ref. Artículo", 18),
    ("prod_descripcion", "Descripción del Producto", 42),
    ("prod_cantidad", "Cantidad", 12),
    ("prod_unidad", "Unidad", 12),
    ("lugar_descarga", "Lugar Entrega / Descarga", 36),
    ("poblacion_descarga", "Población / CP", 18),
    ("cliente_destinatario", "Cliente / Destinatario", 28),
    ("direccion_logistico", "Dirección Fiscal", 32),
    ("cif_cliente", "CIF / NIF", 14),
    ("proveedor_remitente", "Proveedor Remitente", 28),
    ("fecha_albaran_envio", "Fecha Albarán / Emisión", 16),
    ("transportista", "Transportista", 24),
    ("instrucciones_notas", "Notas de Entrega", 24)
]

# ==========================================
# COLUMNAS RESUMEN CABECERAS
# ==========================================
COLUMNAS_RESUMEN_SALTOKI = [
    ("archivo", "Archivo PDF", 24),
    ("numero_documento", "Nº Albarán SF", 18),
    ("numeros_pedidos", "Nº Pedido Cliente", 18),
    ("fecha_albaran_envio", "Fecha Albarán", 14),
    ("fecha_pedido_cliente", "Fecha Pedido", 14),
    ("total_unidades", "Total Unidades", 14),
    ("total_lineas", "Total Líneas", 12),
    ("lugar_descarga", "Lugar Entrega / Descarga", 38),
    ("poblacion_descarga", "Población / CP", 18),
    ("cliente_destinatario", "Cliente / Razón Social", 28),
    ("direccion_logistico", "Dirección Fiscal", 34),
    ("cif_cliente", "CIF Cliente", 14),
    ("proveedor_remitente", "Proveedor Remitente", 28),
    ("mercancia_resumen", "Mercancía / Bultos", 35)
]

COLUMNAS_RESUMEN_AQUAFER = [
    ("archivo", "Archivo PDF", 24),
    ("numero_documento", "Nº Guía Remesa (GREM)", 22),
    ("numero_cmr", "Nº Carta de Porte (CMR)", 18),
    ("numeros_pedidos", "Nº Pedidos Cliente", 22),
    ("total_unidades", "Total Unidades", 14),
    ("total_lineas", "Total Líneas", 12),
    ("lugar_descarga", "Lugar de Descarga (Entrega)", 36),
    ("poblacion_descarga", "Población / CP Descarga", 18),
    ("lugar_carga", "Lugar de Carga (Origen Portugal)", 35),
    ("cliente_destinatario", "Cliente Destinatario", 28),
    ("direccion_logistico", "Dirección Fiscal Destinatario", 32),
    ("cif_cliente", "NIF Destinatario", 14),
    ("proveedor_remitente", "Proveedor Remitente", 32),
    ("fecha_albaran_envio", "Fecha Emisión Albarán", 16),
    ("fecha_expedicion_cmr", "Fecha Expedición CMR", 18),
    ("transportista", "Transportista", 28),
    ("instrucciones_notas", "Notas de Entrega (Español)", 26),
    ("mercancia_resumen", "Mercancía Resumen", 32)
]


def _expand_items_to_lines(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
  """Desglosa cada albarán en sus líneas de productos individuales."""
  expanded = []
  for item in items:
    lines = item.get("lineas_productos", [])
    if lines:
      for l in lines:
        row = dict(item)
        row["prod_pedido"] = l.get("pedido") or item.get("numeros_pedidos") or ""
        row["prod_linea"] = l.get("linea", 1)
        row["prod_codigo"] = l.get("codigo_articulo", "")
        row["prod_codigo_sec"] = l.get("codigo_secundario", "")
        row["prod_descripcion"] = l.get("descripcion", "")
        row["prod_cantidad"] = l.get("cantidad_str") or l.get("cantidad") or 1
        row["prod_unidad"] = l.get("unidad", "Unidades")
        expanded.append(row)
    else:
      # Si no tuviera líneas desglosadas, crear al menos una línea con la mercancía general
      row = dict(item)
      row["prod_pedido"] = item.get("numeros_pedidos", "")
      row["prod_linea"] = 1
      row["prod_codigo"] = "-"
      row["prod_codigo_sec"] = "-"
      row["prod_descripcion"] = item.get("mercancia_resumen", "Artículos de albarán")
      row["prod_cantidad"] = item.get("total_unidades") or 1
      row["prod_unidad"] = "Unidades"
      expanded.append(row)
  return expanded


def _crear_hoja(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    rows_data: List[Dict[str, Any]],
    columns: List[tuple],
    header_color: str,
    title_text: str,
):
  worksheet = workbook.add_worksheet(sheet_name)
  worksheet.set_paper(9)  # A4
  worksheet.set_margins(0.5, 0.5, 0.5, 0.5)

  # Formatos
  fmt_title = workbook.add_format({
      "bold": True,
      "font_size": 13,
      "font_name": "Segoe UI",
      "font_color": "#FFFFFF",
      "bg_color": header_color,
      "align": "left",
      "valign": "vcenter",
      "border": 1,
      "border_color": "#1E293B",
  })

  fmt_header = workbook.add_format({
      "bold": True,
      "font_size": 10,
      "font_name": "Segoe UI",
      "font_color": "#FFFFFF",
      "bg_color": header_color,
      "align": "center",
      "valign": "vcenter",
      "border": 1,
      "border_color": "#334155",
      "text_wrap": True,
  })

  fmt_data_normal = workbook.add_format({
      "font_size": 9,
      "font_name": "Segoe UI",
      "font_color": "#0F172A",
      "bg_color": "#FFFFFF",
      "valign": "vcenter",
      "border": 1,
      "border_color": "#E2E8F0",
  })

  fmt_data_alt = workbook.add_format({
      "font_size": 9,
      "font_name": "Segoe UI",
      "font_color": "#0F172A",
      "bg_color": "#F8FAFC",
      "valign": "vcenter",
      "border": 1,
      "border_color": "#E2E8F0",
  })

  fmt_highlighted = workbook.add_format({
      "bold": True,
      "font_size": 9,
      "font_name": "Segoe UI",
      "font_color": "#78350F",
      "bg_color": "#FEF08A",  # Amarillo suave fluorescente
      "valign": "vcenter",
      "border": 1,
      "border_color": "#FACC15",
  })

  # Fila 0: Título de la hoja
  worksheet.set_row(0, 30)
  worksheet.merge_range(0, 0, 0, len(columns) - 1, f"  {title_text}", fmt_title)

  # Fila 1: Cabeceras
  worksheet.set_row(1, 26)
  for col_idx, (key, label, default_w) in enumerate(columns):
    worksheet.write(1, col_idx, label, fmt_header)

  # Filas de datos
  max_lens = {col_idx: len(label) for col_idx, (_, label, _) in enumerate(columns)}
  row_idx = 2

  for item in rows_data:
    worksheet.set_row(row_idx, 21)
    is_alt = row_idx % 2 == 1
    campos_resaltados = item.get("campos_resaltados", []) or []

    for col_idx, (key, label, default_w) in enumerate(columns):
      val = clean_str(item.get(key, ""))
      if not val:
        val = "-"

      # Si el campo estaba subrayado/resaltado en el PDF original
      if (key in campos_resaltados or key in ("prod_pedido", "numero_documento")) and val != "-":
        cell_fmt = fmt_highlighted
      elif is_alt:
        cell_fmt = fmt_data_alt
      else:
        cell_fmt = fmt_data_normal

      worksheet.write(row_idx, col_idx, val, cell_fmt)

      val_len = min(len(str(val)), 45)
      if val_len > max_lens[col_idx]:
        max_lens[col_idx] = val_len

    row_idx += 1

  # Anchos de columna
  for col_idx, (key, label, default_w) in enumerate(columns):
    calc_w = max(default_w, max_lens[col_idx] + 3)
    worksheet.set_column(col_idx, col_idx, calc_w)

  worksheet.freeze_panes(2, 2)
  if rows_data:
    worksheet.autofilter(1, 0, row_idx - 1, len(columns) - 1)


def generar_excel_bytes(items: List[Dict[str, Any]], modelo: str = "TODOS") -> bytes:
  """Genera el libro Excel profesional con hojas detalladas de productos y resúmenes."""
  output = io.BytesIO()
  workbook = xlsxwriter.Workbook(output, {"in_memory": True})

  saltoki_items = [it for it in items if it.get("modelo", "").upper() == "SALTOKI"]
  aquafer_items = [it for it in items if it.get("modelo", "").upper() == "AQUAFER"]

  if not saltoki_items and not aquafer_items:
    for it in items:
      txt = str(it)
      if "AQUAFER" in txt.upper() or "GREM" in txt.upper() or "CMR" in txt.upper():
        aquafer_items.append(it)
      else:
        saltoki_items.append(it)

  if modelo == "SALTOKI" or (saltoki_items and not aquafer_items):
    target = saltoki_items if saltoki_items else items
    prod_lines = _expand_items_to_lines(target)
    # HOJA 1: Detalle Completo con todos los productos, pedidos y referencias
    _crear_hoja(
        workbook,
        "Detalle Completo",
        prod_lines,
        COLUMNAS_PRODUCTOS_SALTOKI,
        "#0284C7",
        "📦 SALTOKI: DETALLE COMPLETO DE PRODUCTOS, PEDIDOS Y CANTIDADES (Campos amarillos = Subrayados)",
    )
    # HOJA 2: Resumen de totales
    _crear_hoja(
        workbook,
        "Resumen Albaranes",
        target,
        COLUMNAS_RESUMEN_SALTOKI,
        "#1E293B",
        "📋 SALTOKI: RESUMEN GENERAL (1 Fila por Albarán)",
    )

  elif modelo == "AQUAFER" or (aquafer_items and not saltoki_items):
    target = aquafer_items if aquafer_items else items
    prod_lines = _expand_items_to_lines(target)
    # HOJA 1: Detalle Completo con todos los artículos y pedidos traducidos
    _crear_hoja(
        workbook,
        "Detalle Completo",
        prod_lines,
        COLUMNAS_PRODUCTOS_AQUAFER,
        "#0F766E",
        "📦 AQUAFER / CMR: DETALLE COMPLETO DE ARTÍCULOS, PEDIDOS Y CANTIDADES (Traducido al Español)",
    )
    # HOJA 2: Resumen de totales
    _crear_hoja(
        workbook,
        "Resumen Albaranes",
        target,
        COLUMNAS_RESUMEN_AQUAFER,
        "#1E293B",
        "📋 AQUAFER / CMR: RESUMEN GENERAL (1 Fila por Albarán)",
    )

  else:
    # AMBOS MODELOS: Detalle completo primero para cada empresa
    if saltoki_items:
      lines_s = _expand_items_to_lines(saltoki_items)
      _crear_hoja(
          workbook,
          "Detalle Saltoki",
          lines_s,
          COLUMNAS_PRODUCTOS_SALTOKI,
          "#0284C7",
          "🏢 SALTOKI: DETALLE COMPLETO DE ARTÍCULOS, PEDIDOS Y CANTIDADES",
      )
    if aquafer_items:
      lines_a = _expand_items_to_lines(aquafer_items)
      _crear_hoja(
          workbook,
          "Detalle Aquafer",
          lines_a,
          COLUMNAS_PRODUCTOS_AQUAFER,
          "#0F766E",
          "🚛 AQUAFER / CMR: DETALLE COMPLETO DE ARTÍCULOS Y CANTIDADES (Traducido al Español)",
      )
    if saltoki_items:
      _crear_hoja(
          workbook,
          "Resumen Saltoki",
          saltoki_items,
          COLUMNAS_RESUMEN_SALTOKI,
          "#1E293B",
          "📋 SALTOKI: RESUMEN DE ALBARANES",
      )
    if aquafer_items:
      _crear_hoja(
          workbook,
          "Resumen Aquafer",
          aquafer_items,
          COLUMNAS_RESUMEN_AQUAFER,
          "#1E293B",
          "📋 AQUAFER: RESUMEN DE GUÍAS DE REMESA Y CMR",
      )

  workbook.close()
  output.seek(0)
  return output.getvalue()


def generar_csv_texto(items: List[Dict[str, Any]], modelo: str = "TODOS", tipo: str = "detalle") -> str:
  """
  Genera texto CSV limpio con punto y coma (;).
  - tipo='detalle' (por defecto): Desglosa cada producto en una fila individual con código, descripción, cantidad y unidad.
  - tipo='resumen': Genera 1 fila limpia por albarán (sin repetición de datos, igual que en pantalla).
  """
  if tipo == "detalle":
    rows_data = _expand_items_to_lines(items)
    if modelo == "SALTOKI":
      cols = COLUMNAS_PRODUCTOS_SALTOKI
    elif modelo == "AQUAFER":
      cols = COLUMNAS_PRODUCTOS_AQUAFER
    else:
      cols = COLUMNAS_PRODUCTOS_TODOS
  else:
    rows_data = items
    if modelo == "SALTOKI":
      cols = COLUMNAS_RESUMEN_SALTOKI
    elif modelo == "AQUAFER":
      cols = COLUMNAS_RESUMEN_AQUAFER
    else:
      cols = COLUMNAS_RESUMEN_SALTOKI

  lines = []
  headers = [f'"{label}"' for (_, label, _) in cols]
  lines.append(";".join(headers))

  for row in rows_data:
    row_vals = []
    for (k, _, _) in cols:
      val = clean_str(row.get(k, ""))
      val_escaped = val.replace('"', '""')
      row_vals.append(f'"{val_escaped}"')
    lines.append(";".join(row_vals))

  return "\r\n".join(lines) + "\r\n"
