# 📑 Albarán a CSV Pro - Extractor Inteligente de Albaranes Logísticos

Aplicación para **convertir albaranes y guías de remesa/transporte en PDF a formato CSV y Excel**, extrayendo con precisión los campos logísticos clave:

1. **Número de Documento / Albarán** (ej. `SF 20260625`, `GREM 90A25/922 | CMR 32851`)
2. **Números de Pedido** (ej. `SF - 227018`, `473036, 473318`)
3. **Lugar de Descarga / Entrega** (ej. `CL RIO BEMBEZAR 4, 28906 GETAFE`, `EMPRESARIUM LOGISTICA...`)
4. **Dirección Logístico / Cliente** (ej. `SALTOKI LOGISTICO EMPRESARIUM...`)
5. **Fecha Albarán Envío** (ej. `02/10/26`, `2025-07-11`)
6. **Fecha Pedido Cliente** (ej. `25/05/2026`)
7. **Fechas Detectadas (Todas)** (ej. `Documento: 2025-07-11 | Vencimiento: 2025-09-09 | Expedición/CMR: 11 JULHO 2025`)
8. **Subrayados / Resaltados** (Identifica y extrae el texto destacado con marcador fluorescente o subrayado: lugares de descarga, pedidos resaltados, notas especiales de entrega, etc.)

---

## 🚀 Formas de Uso

Tienes **3 opciones súper cómodas** para usar la aplicación:

### Opción 1: Aplicación Visual con Servidor Local (Recomendada)
1. Haz doble clic en el archivo **`Iniciar_App.bat`**.
2. Se abrirá automáticamente la aplicación en tu navegador web en `http://localhost:8080`.
3. Arrastra tus archivos PDF al recuadro o haz clic en **"Seleccionar Archivos PDF"**.
4. ¡Listo! Verás los campos extraídos al instante en la tabla interactiva.

### Opción 2: Abrir Directamente en el Navegador (Sin servidor)
1. Haz doble clic en **`Abrir_Directo_Navegador.bat`** (o en `index.html`).
2. Funciona 100% en tu navegador de forma local, offline y sin necesidad de ejecutar comandos.

### Opción 3: Modo Línea de Comandos / Procesamiento por Lotes
Si prefieres procesar una carpeta completa con PDFs por consola:
```bash
# Procesar una carpeta y generar albaranes.csv (con separador ';' para Excel)
python extractor.py "ruta/a/tus/pdfs" --csv "albaranes.csv" --sep ";"

# O simplemente arrastra tu carpeta de PDFs sobre el archivo "Procesar_Carpeta_PDF.bat"
```

---

## ✨ Características Principales

- **Arrastrar y Soltar (Drag & Drop)**: Procesa 1, 10 o 100 albaranes de golpe.
- **Edición en tiempo real**: Si algún albarán escaneado tiene algún dato borroso, puedes hacer clic sobre cualquier celda de la tabla para editarlo o corregirlo antes de exportar.
- **Previsualizador de PDF integrado**: Al pulsar el botón 👁️ en cualquier fila, se abre un visor interactivo donde puedes contrastar los datos extraídos cara a cara con el PDF original.
- **Botón de Muestras**: Incluye un botón **"Cargar Ejemplos de Prueba"** para probar la aplicación con los ejemplos de APEP Group y Aquafer en un solo clic.
- **Múltiples formatos de Exportación**:
  - 📥 **Descargar CSV (Excel ;)**: Codificado en UTF-8 con BOM para que se abra en Excel en español sin descolocar tildes ni caracteres especiales.
  - 📥 **Descargar CSV (Coma ,)**: Estándar RFC 4180 internacional.
  - 📊 **Descargar Excel (.xlsx)**: Genera un libro nativo de Microsoft Excel formateado con anchos de columna ajustados.
  - 📋 **Copiar Todo**: Copia toda la tabla al portapapeles en formato tabulado para que puedas hacer `Ctrl + V` directo en Excel o Google Sheets.

---

## 📁 Estructura del Proyecto

- `index.html`: Interfaz web de usuario.
- `style.css`: Estilos visuales modernos y responsivos.
- `app.js`: Motor de extracción y manipulación en el navegador.
- `extractor.py`: Motor de extracción en Python para línea de comandos y servidor.
- `servidor.py`: Servidor local ligero.
- `lib/`: Librerías locales (`pdf.js`, `xlsx.full.min.js`) para funcionamiento offline garantizado.
- `muestras/`: Archivos PDF de prueba creados a partir de tus ejemplos.
- `Iniciar_App.bat`: Lanzador con doble clic para el servidor y la web.
- `Abrir_Directo_Navegador.bat`: Lanzador directo al navegador.
- `Procesar_Carpeta_PDF.bat`: Procesador directo por lotes para carpetas.
