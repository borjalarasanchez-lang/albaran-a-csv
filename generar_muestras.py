import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

os.makedirs("muestras", exist_ok=True)

def create_apep_pdf(filename):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Page 1
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 50, "APEP GROUP 2020 SL")
    c.setFont("Helvetica", 9)
    c.drawString(50, height - 65, "AVENIDA VILLA DE MADRID, 73")
    c.drawString(50, height - 77, "29700 VELEZ MALAGA")
    c.drawString(50, height - 89, "MALAGA")
    c.drawString(50, height - 101, "CIF: B93730240")
    c.drawString(50, height - 113, "info@apepgroup.com")

    # Destination / Logistico Box
    c.rect(340, height - 130, 200, 80)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(350, height - 65, "SALTOKI LOGISTICO EMPRESARIUM")
    c.setFont("Helvetica", 9)
    c.drawString(350, height - 80, "C/CAPITANA S/N P.IND. EMPRESARIUM")
    c.drawString(350, height - 95, "50720 CARTUJA BAJA")
    c.drawString(350, height - 110, "ZARAGOZA")

    # Header table
    c.rect(50, height - 180, 500, 35)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(70, height - 160, "FECHA")
    c.drawString(150, height - 160, "NUMERO")
    c.drawString(250, height - 160, "CLIENTE")
    c.drawString(350, height - 160, "N.I.F")
    c.drawString(450, height - 160, "PAG.")

    c.setFont("Helvetica", 9)
    c.drawString(70, height - 173, "02/10/26")
    c.drawString(150, height - 173, "SF 20260625")
    c.drawString(250, height - 173, "43000173")
    c.drawString(350, height - 173, "B56722226")
    c.drawString(450, height - 173, "1")

    # Content
    c.setFont("Helvetica-Bold", 9)
    c.drawString(50, height - 210, "Pedido nº SF - 227018 de fecha 25/05/2026")
    c.setFont("Helvetica", 8)
    c.drawString(50, height - 230, "TAP1500VS1PP INTERACUMULADOR 1500L INOX ESPECIAL AEROTERMIA   3,00")
    c.drawString(50, height - 245, "Termoacumulador con un serpentin fijo en inox 444")
    c.drawString(50, height - 260, "CODIGO 2310366062 , DESCRIPCION 2310366062 INTERACUMULADOR 444 DPI/I/BC 1500 AQ")

    c.showPage()

    # Page 2
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 50, "APEP GROUP 2020 SL")
    c.rect(50, height - 100, 500, 35)
    c.setFont("Helvetica", 9)
    c.drawString(70, height - 90, "02/10/26")
    c.drawString(150, height - 90, "SF 20260625")
    c.drawString(250, height - 90, "43000173")
    c.drawString(450, height - 90, "2")

    # Subrayado / Resaltado amarillo fluorescente en DIRECCION DE ENTREGA
    c.saveState()
    c.setFillColor(Color(1, 1, 0, alpha=0.45))
    c.rect(48, height - 143, 140, 15, fill=1, stroke=0)
    c.restoreState()

    c.setFont("Helvetica-Bold", 9)
    c.drawString(50, height - 140, "DIRECCION DE ENTREGA:")
    c.setFont("Helvetica", 9)
    c.drawString(50, height - 155, "EMPRESARIUM LOGISTICA")
    c.drawString(50, height - 170, "POLIGONO EMPRESARIUM")
    c.drawString(50, height - 185, "C/CAPITANA")
    c.drawString(50, height - 200, "50720 ZARAGOZA")

    c.showPage()
    c.save()

def create_aquafer_pdf(filename):
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    # Page 1
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 60, "AQUAFER")
    c.setFont("Helvetica", 8)
    c.drawString(50, height - 75, "INDUSTRIA")
    c.drawString(50, height - 90, "AV. DR. FRANCISCO SA CARNEIRO, 980")
    c.drawString(50, height - 100, "5350-005 ALFANDEGA DA FE, PORTUGAL")

    # Document Header Right con resaltador rojo/rosa en GUIA DE REMESSA y DATA DOC
    c.saveState()
    c.setFillColor(Color(1, 0.35, 0.35, alpha=0.35))
    c.roundRect(380, height - 90, 160, 50, 4, fill=1, stroke=0)
    c.restoreState()

    c.rect(380, height - 90, 160, 50)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(390, height - 55, "GUIA DE REMESSA")
    c.drawString(390, height - 70, "GREM 90A25/922")
    c.drawString(390, height - 85, "DATA DOC.: 2025-07-11")

    # Recipient
    c.drawString(280, height - 130, "Exmo./a(s) Sr./a(s)")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(280, height - 145, "SALTOKI LOGISTICA ZARAGOZA, SA")
    c.setFont("Helvetica", 9)
    c.drawString(280, height - 160, "C/ ALCALDE CABALLERO, Nº16")
    c.drawString(280, height - 175, "50014 ZARAGOZA")

    # Content - Items con resaltados rosas en los números de pedido (enc nº 473036 y enc nº 473318)
    c.setFont("Helvetica", 8)
    c.drawString(50, height - 230, "2310351150  ACUMULADOR 444 DPI/I/BC 25 AQ mod: DII 0025 VSQ P  5,000 Un")
    
    # Resaltado enc nº 473036
    c.saveState()
    c.setFillColor(Color(1, 0.4, 0.4, alpha=0.4))
    c.rect(98, height - 247, 85, 12, fill=1, stroke=0)
    c.restoreState()

    c.setFont("Helvetica-Bold", 8)
    c.drawString(100, height - 245, "enc nº 473036")

    c.setFont("Helvetica", 8)
    c.drawString(50, height - 265, "2310351151  ACUMULADOR 444 DPI/I/BC 60 AQ mod: DII 0060 VSQ P  10,000 Un")
    
    # Resaltado enc nº 473318
    c.saveState()
    c.setFillColor(Color(1, 0.4, 0.4, alpha=0.4))
    c.rect(98, height - 282, 85, 12, fill=1, stroke=0)
    c.restoreState()

    c.setFont("Helvetica-Bold", 8)
    c.drawString(100, height - 280, "enc nº 473318")

    c.drawString(50, height - 300, "DATA VENCIMENTO: 2025-09-09")
    c.drawString(50, height - 315, "ENTREGAR EM GETAFE")

    # Local de descarga box con resaltado rosa
    c.saveState()
    c.setFillColor(Color(1, 0.4, 0.4, alpha=0.35))
    c.roundRect(48, height - 422, 224, 74, 4, fill=1, stroke=0)
    c.restoreState()

    c.rect(50, height - 420, 220, 70)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(60, height - 365, "LOCAL DE DESCARGA:")
    c.setFont("Helvetica", 9)
    c.drawString(60, height - 380, "CL RIO BEMBEZAR 4")
    c.drawString(60, height - 395, "28906 GETAFE")
    c.drawString(60, height - 410, "Morada do destinatario")

    c.showPage()

    # Page 2: CMR
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 50, "DECLARAÇÃO DE EXPEDIÇÃO INTERNACIONAL / CMR Nº 32851")
    c.setFont("Helvetica", 9)
    c.drawString(50, height - 80, "1 Expedidor: Aquafer, Lda - Rua Monte Loscer 436, 4775-263 Viatodos")
    c.drawString(50, height - 100, "2 Destinatario: Saltoki SA C/ Alcalde Caballero 16, 50014 Zaragoza")
    c.drawString(50, height - 120, "3 Lugar de entrega da mercadoria: CL RIO BEMBEZAR 4, 28906 Getafe")
    c.drawString(50, height - 140, "5 Documentos anexos: GREM 90A25/922")
    c.drawString(50, height - 160, "Feito em: Viatodos 11 JULHO 2025")

    c.showPage()
    c.save()

create_apep_pdf("muestras/ejemplo_apep_group.pdf")
create_aquafer_pdf("muestras/ejemplo_aquafer.pdf")
print("PDFs actualizados con resaltados y campos completos.")
