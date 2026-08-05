import io
import os
import subprocess
import sys

# Forzar instalación automática de dependencias si faltan
try:
  import reportlab
  import openpyxl
  import pypdf
except ImportError:
  subprocess.check_call([
      sys.executable,
      "-m",
      "pip",
      "install",
      "streamlit==1.32.0",
      "pandas==2.2.1",
      "openpyxl==3.1.2",
      "reportlab==4.1.0",
      "pypdf==4.1.0",
  ])

import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# Configuración general de la página web
st.set_page_config(
    page_title="Herramientas SRI Ecuador", page_icon="🧾", layout="wide"
)

# Diccionario oficial de formas de pago del SRI
FORMAS_PAGO_SRI = {
    "01": "SIN UTILIZACION DEL SISTEMA FINANCIERO",
    "15": "COMPENSACIÓN DE DEUDAS",
    "16": "TARJETA DE DÉBITO",
    "17": "DINERO ELECTRÓNICO",
    "18": "TARJETA DE PREPAGO",
    "19": "TARJETA DE CRÉDITO",
    "20": "OTROS CON UTILIZACION DEL SISTEMA FINANCIERO",
    "21": "ENDOSO DE TÍTULOS",
}


def obtener_valor(nodo, tag, default=""):
  if nodo is None:
    return default
  res = nodo.findtext(tag)
  return res if res else default


def formatear_num_doc(numero_crudo):
  if not numero_crudo:
    return ""
  limpio = str(numero_crudo).replace("-", "").strip()
  if len(limpio) == 15:
    return f"{limpio[:3]}-{limpio[3:6]}-{limpio[6:]}"
  return numero_crudo


def formatear_moneda(valor):
  try:
    return "{:.2f}".format(float(valor))
  except:
    return "0.00"


# --- FUNCIONES DE PROCESAMIENTO EXCEL ---
def procesar_comprobante_excel(contenido_xml):
  try:
    root = ET.fromstring(contenido_xml)
    comprobante_node = root.find("comprobante")
    if comprobante_node is not None:
      xml_data = ET.fromstring(comprobante_node.text)
      fecha_aut = obtener_valor(root, "fechaAutorizacion", "N/A")
    else:
      xml_data = root
      fecha_aut = "N/A"

    info_trib = xml_data.find("infoTributaria")
    if info_trib is None:
      return None, None

    cod_doc = info_trib.find("codDoc").text
    ruc_emisor = obtener_valor(info_trib, "ruc")
    nombre_emisor = obtener_valor(info_trib, "razonSocial")
    secuencial = f"{obtener_valor(info_trib, 'estab')}-{obtener_valor(info_trib, 'ptoEmi')}-{obtener_valor(info_trib, 'secuencial')}"
    clave_acc = obtener_valor(info_trib, "claveAcceso")

    if cod_doc in ["01", "04"]:
      es_nc = cod_doc == "04"
      info_sec = (
          xml_data.find("infoNotaCredito")
          if es_nc
          else xml_data.find("infoFactura")
      )
      datos = {
          "tipo": "NOTA DE CREDITO" if es_nc else "FACTURA",
          "secuencial": secuencial,
          "clave_autorizacion": clave_acc,
          "ruc_emisor": ruc_emisor,
          "nombre_emisor": nombre_emisor,
          "fecha_emision": obtener_valor(info_sec, "fechaEmision"),
          "fecha_autorizacion": fecha_aut,
          "comprador_ruc": obtener_valor(info_sec, "identificacionComprador"),
          "subtotal_cero": 0.0,
          "subtotal_iva_12": 0.0,
          "subtotal_iva_13": 0.0,
          "subtotal_iva_14": 0.0,
          "subtotal_iva_15": 0.0,
          "subtotal_iva_5": 0.0,
          "iva": 0.0,
          "importe_total": float(
              obtener_valor(info_sec, "importeTotal") or 0.0
          ),
      }
      if es_nc:
        num_mod = obtener_valor(info_sec, "numDocModificado")
        datos["factura_que_modifica"] = formatear_num_doc(num_mod)

      total_imp = info_sec.find("totalConImpuestos")
      if total_imp is not None:
        for imp in total_imp.findall("totalImpuesto"):
          cp = obtener_valor(imp, "codigoPorcentaje")
          base = float(obtener_valor(imp, "baseImponible") or 0)
          valor_imp = float(obtener_valor(imp, "valor") or 0)
          if cp == "0":
            datos["subtotal_cero"] += base
          elif cp in ["2", "10"]:
            datos["subtotal_iva_12"] += base
          elif cp == "8":
            datos["subtotal_iva_13"] += base
          elif cp == "4":
            datos["subtotal_iva_15"] += base
          elif cp == "3":
            datos["subtotal_iva_14"] += base
          elif cp == "5":
            datos["subtotal_iva_5"] += base
          datos["iva"] += valor_imp
      return cod_doc, datos

    elif cod_doc == "07":
      info_ret = xml_data.find("infoCompRetencion")
      datos = {
          "tipo": "RETENCION",
          "secuencial": secuencial,
          "clave_autorizacion": clave_acc,
          "ruc_emisor": ruc_emisor,
          "nombre_emisor": nombre_emisor,
          "fecha_emision": obtener_valor(info_ret, "fechaEmision"),
          "identificacion_retenido": obtener_valor(
              info_ret, "identificacionSujetoRetenido"
          ),
          "numDocSustento": "",
          "RIR": 0.0,
          "RIVA": 0.0,
      }
      docs_sustento = xml_data.find("docsSustento")
      lista_ret = []
      if docs_sustento is not None:
        doc = docs_sustento.find("docSustento")
        if doc is not None:
          num_crudo = obtener_valor(doc, "numDocSustento") or f"{obtener_valor(doc, 'estab')}{obtener_valor(doc, 'ptoEmi')}{obtener_valor(doc, 'secuencial')}"
          datos["numDocSustento"] = formatear_num_doc(num_crudo)
          rets = doc.find("retenciones")
          if rets is not None:
            lista_ret = rets.findall("retencion")
      else:
        imp_nodo = xml_data.find("impuestos")
        if imp_nodo is not None:
          lista_ret = imp_nodo.findall("impuesto")
          if lista_ret:
            datos["numDocSustento"] = formatear_num_doc(
                obtener_valor(lista_ret[0], "numDocSustento")
            )

      for idx, ret in enumerate(lista_ret[:3]):
        i = idx + 1
        cod = obtener_valor(ret, "codigo")
        val = float(obtener_valor(ret, "valorRetenido") or 0)
        if cod == "1":
          datos["RIR"] += val
        elif cod == "2":
          datos["RIVA"] += val
        datos[f"Ret{i}_CodRetencion"] = obtener_valor(ret, "codigoRetencion")
        datos[f"Ret{i}_Base"] = obtener_valor(ret, "baseImponible")
        datos[f"Ret{i}_Porcentaje"] = obtener_valor(ret, "porcentajeRetener")
        datos[f"Ret{i}_Valor"] = val
      return "07", datos
  except Exception:
    return None, None
  return None, None


# --- FUNCIONES DE PROCESAMIENTO PDF ---
def leer_xml_completo_pdf(contenido_xml):
  try:
    root = ET.fromstring(contenido_xml)
    clave = root.findtext(".//numeroAutorizacion", "")
    fecha_aut = root.findtext(".//fechaAutorizacion", "")

    comprobante_node = root.find(".//comprobante")
    if comprobante_node is not None:
      xml_inner = comprobante_node.text.strip()
      doc = ET.fromstring(xml_inner)
    else:
      doc = root

    tag = doc.tag.split("}")[-1]
    info_t = doc.find("infoTributaria")

    dir_matriz = info_t.findtext("dirMatriz", "")
    dir_sucursal = info_t.findtext("dirSucursal", "") or doc.findtext(
        ".//dirSucursal", ""
    )
    contrib_especial = doc.findtext(".//contribuyenteEspecial", "")
    agente_retencion = info_t.findtext("agenteRetencion", "") or doc.findtext(
        ".//agenteRetencion", ""
    )
    rimpe = info_t.findtext("contribuyenteRimpe", "") or info_t.findtext(
        "regimenRimpe", ""
    )

    datos = {
        "tipo": tag,
        "clave": clave,
        "fecha_aut": fecha_aut,
        "empresa": info_t.findtext("razonSocial", ""),
        "ruc": info_t.findtext("ruc", ""),
        "numero": f"{info_t.findtext('estab')}-{info_t.findtext('ptoEmi')}-{info_t.findtext('secuencial')}",
        "dirMatriz": dir_matriz,
        "dirSucursal": dir_sucursal,
        "contribuyenteEspecial": contrib_especial,
        "agenteRetencion": agente_retencion,
        "rimpe": rimpe.strip() if rimpe else "",
        "contabilidad": doc.findtext(".//obligadoContabilidad", "NO"),
        "info_adicional": [],
        "pagos": [],
    }

    for pago in doc.findall(".//pago"):
      cod_pago = pago.findtext("formaPago", "20")
      forma_txt = FORMAS_PAGO_SRI.get(
          cod_pago, "OTROS CON UTILIZACION DEL SISTEMA FINANCIERO"
      )
      total_pago = formatear_moneda(pago.findtext("total", "0.00"))
      datos["pagos"].append([forma_txt, total_pago])

    if tag == "factura":
      info = doc.find("infoFactura")
      datos.update({
          "fecha": info.findtext("fechaEmision"),
          "receptor": info.findtext("razonSocialComprador"),
          "id_receptor": info.findtext("identificacionComprador"),
      })
      items = [["Cod.", "Descripción", "Cant.", "P. Unit", "Desc.", "Total"]]
      for d in doc.findall(".//detalle"):
        items.append([
            obtener_valor(d, "codigoPrincipal", "-"),
            d.findtext("descripcion")[:50],
            formatear_moneda(obtener_valor(d, "cantidad")),
            formatear_moneda(obtener_valor(d, "precioUnitario")),
            formatear_moneda(obtener_valor(d, "descuento")),
            formatear_moneda(obtener_valor(d, "precioTotalSinImpuesto")),
        ])
      datos["items"] = items

      irbpnr_val = (
          info.findtext(".//totalImpuesto[codigo='5']/valor")
          or info.findtext("valorIRBPNR")
          or "0.00"
      )
      tots = [
          [
              "SUBTOTAL 15%",
              formatear_moneda(
                  obtener_valor(
                      info, ".//totalImpuesto[codigoPorcentaje='4']/baseImponible"
                  )
              ),
          ],
          [
              "SUBTOTAL 12%",
              formatear_moneda(
                  obtener_valor(
                      info, ".//totalImpuesto[codigoPorcentaje='2']/baseImponible"
                  )
              ),
          ],
          [
              "SUBTOTAL 0%",
              formatear_moneda(
                  obtener_valor(
                      info, ".//totalImpuesto[codigoPorcentaje='0']/baseImponible"
                  )
              ),
          ],
          [
              "IVA 15%",
              formatear_moneda(
                  obtener_valor(
                      info, ".//totalImpuesto[codigoPorcentaje='4']/valor"
                  )
              ),
          ],
          [
              "IVA 12%",
              formatear_moneda(
                  obtener_valor(
                      info, ".//totalImpuesto[codigoPorcentaje='2']/valor"
                  )
              ),
          ],
      ]
      if float(irbpnr_val) > 0:
        tots.append(["IRBPNR", formatear_moneda(irbpnr_val)])
      tots.append(["TOTAL", formatear_moneda(info.findtext("importeTotal"))])
      datos["resumen"] = tots
      datos["cw"] = [50, 210, 50, 60, 60, 70]

    elif tag == "comprobanteRetencion":
      info = doc.find("infoCompRetencion")
      datos.update({
          "fecha": info.findtext("fechaEmision"),
          "receptor": info.findtext("razonSocialSujetoRetenido"),
          "id_receptor": info.findtext("identificacionSujetoRetenido"),
      })
      items = [["Ejercicio", "Factura Sustento", "Cód.", "%", "Base Imp.", "Valor Ret."]]
      total_r = 0.0
      for ret in doc.findall(".//retencion"):
        num_doc = ret.findtext("numDocSustento")
        if not num_doc:
          padre_doc = doc.find(".//docSustento")
          num_doc = (
              padre_doc.findtext("numDocSustento")
              if padre_doc is not None
              else "-"
          )
        if num_doc and num_doc != "-" and len(num_doc) == 15:
          num_doc = f"{num_doc[:3]}-{num_doc[3:6]}-{num_doc[6:]}"
        v = float(obtener_valor(ret, "valorRetenido", "0"))
        total_r += v
        ejercicio = (
            ret.findtext("ejercicioFiscal")
            or info.findtext("periodoFiscal")
            or "-"
        )
        items.append([
            ejercicio,
            num_doc,
            obtener_valor(ret, "codigoRetencion", "-"),
            obtener_valor(ret, "porcentajeRetener", "0") + "%",
            formatear_moneda(obtener_valor(ret, "baseImponible", "0")),
            formatear_moneda(v),
        ])
      datos["items"] = items
      datos["resumen"] = [["TOTAL RETENIDO", formatear_moneda(total_r)]]
      datos["cw"] = [70, 130, 60, 50, 90, 90]

    elif tag == "notaCredito":
      info = doc.find("infoNotaCredito")
      datos.update({
          "fecha": info.findtext("fechaEmision"),
          "receptor": info.findtext("razonSocialComprador"),
          "id_receptor": info.findtext("identificacionComprador"),
          "modifica": (
              f"Doc. Modifica: {info.findtext('numDocModificado')} | Fecha Emisión"
              f" Doc. Sustento: {info.findtext('fechaEmisionDocSustento')}"
          ),
      })
      items = [["Descripción", "Cant.", "P. Unit", "Desc.", "Total"]]
      for d in doc.findall(".//detalle"):
        items.append([
            d.findtext("descripcion")[:60],
            d.findtext("cantidad"),
            formatear_moneda(d.findtext("precioUnitario")),
            "0.00",
            formatear_moneda(d.findtext("precioTotalSinImpuesto")),
        ])
      datos["items"] = items

      sub15 = float(
          obtener_valor(
              info, ".//totalImpuesto[codigoPorcentaje='4']/baseImponible", "0"
          )
      )
      sub12 = float(
          obtener_valor(
              info, ".//totalImpuesto[codigoPorcentaje='2']/baseImponible", "0"
          )
      )
      sub0 = float(
          obtener_valor(
              info, ".//totalImpuesto[codigoPorcentaje='0']/baseImponible", "0"
          )
      )
      iva15 = float(
          obtener_valor(
              info, ".//totalImpuesto[codigoPorcentaje='4']/valor", "0"
          )
      )
      iva12 = float(
          obtener_valor(
              info, ".//totalImpuesto[codigoPorcentaje='2']/valor", "0"
          )
      )
      irbpnr_nc = float(
          info.findtext(".//totalImpuesto[codigo='5']/valor") or "0.00"
      )
      total_nc = sub15 + sub12 + sub0 + iva15 + iva12 + irbpnr_nc

      tots_nc = [
          ["SUBTOTAL 15%", formatear_moneda(sub15)],
          ["SUBTOTAL 12%", formatear_moneda(sub12)],
          ["SUBTOTAL 0%", formatear_moneda(sub0)],
          ["IVA 15%", formatear_moneda(iva15)],
          ["IVA 12%", formatear_moneda(iva12)],
      ]
      if irbpnr_nc > 0:
        tots_nc.append(["IRBPNR", formatear_moneda(irbpnr_nc)])
      tots_nc.append(["TOTAL NC", formatear_moneda(total_nc)])
      datos["resumen"] = tots_nc
      datos["cw"] = [280, 60, 60, 60, 70]

    nodo_ad = doc.find("infoAdicional")
    if nodo_ad is not None:
      for campo in nodo_ad.findall("campoAdicional"):
        datos["info_adicional"].append([f"{campo.get('nombre')}:", campo.text])

    return datos
  except Exception:
    return None


def generar_pdf_bytes(datos):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=15
  )
  styles = getSampleStyleSheet()

  estilo_titulo_empresa = ParagraphStyle(
      "TE", parent=styles["Heading3"], fontSize=9.5, leading=11.5
  )
  estilo_ruc = ParagraphStyle(
      "ERUC", parent=styles["Normal"], fontSize=9.5, leading=11.5
  )
  estilo_clave = ParagraphStyle(
      "C", parent=styles["Normal"], fontSize=8.5, leading=10.5
  )
  estilo_titulo = ParagraphStyle(
      "T", parent=styles["Heading2"], fontSize=10, leading=12
  )
  estilo_datos_cab = ParagraphStyle(
      "DC", parent=styles["Normal"], fontSize=7, leading=8.5
  )
  estilo_general = ParagraphStyle(
      "G", parent=styles["Normal"], fontSize=7, leading=9
  )

  elementos = []
  titulos = {
      "factura": "FACTURA",
      "comprobanteRetencion": "COMPROBANTE DE RETENCIÓN",
      "notaCredito": "NOTA DE CRÉDITO",
  }

  col_izq = [
      Spacer(1, 2),
      Paragraph(f"<b>{datos['empresa']}</b>", estilo_titulo_empresa),
      Paragraph(f"<b>R.U.C.:</b> {datos['ruc']}", estilo_ruc),
      Spacer(1, 2),
      Paragraph(f"<b>Dirección Matriz:</b> {datos['dirMatriz']}", estilo_datos_cab),
  ]
  if datos["dirSucursal"]:
    col_izq.append(
        Paragraph(
            f"<b>Dirección Sucursal:</b> {datos['dirSucursal']}", estilo_datos_cab
        )
    )
  if datos["contribuyenteEspecial"]:
    col_izq.append(
        Paragraph(
            f"<b>Contribuyente Especial Nro:</b>"
            f" {datos['contribuyenteEspecial']}",
            estilo_datos_cab,
        )
    )
  if datos["agenteRetencion"]:
    col_izq.append(
        Paragraph(
            f"<b>Agente de Retención Resolución No.:</b>"
            f" {datos['agenteRetencion']}",
            estilo_datos_cab,
        )
    )
  col_izq.append(
      Paragraph(
          f"<b>OBLIGADO A LLEVAR CONTABILIDAD:</b> {datos['contabilidad']}",
          estilo_datos_cab,
      )
  )
  if datos["rimpe"]:
    col_izq.append(
        Paragraph(f"<b>{datos['rimpe'].upper()}</b>", estilo_datos_cab)
    )
  col_izq.append(Spacer(1, 2))

  col_der = [
      Spacer(1, 2),
      Paragraph(f"<b>{titulos.get(datos['tipo'])}</b>", estilo_titulo),
      Paragraph(f"<b>No.:</b> {datos['numero']}", estilo_datos_cab),
      Paragraph(
          f"<b>FECHA Y HORA DE AUTORIZACIÓN:</b> {datos['fecha_aut']}",
          estilo_datos_cab,
      ),
      Paragraph(f"<b>AMBIENTE:</b> PRODUCCIÓN", estilo_datos_cab),
      Paragraph(f"<b>EMISIÓN:</b> NORMAL", estilo_datos_cab),
      Paragraph(
          f"<b>NÚMERO DE AUTORIZACIÓN / CLAVE DE ACCESO:</b>", estilo_datos_cab
      ),
      Paragraph(f"{datos['clave']}", estilo_clave),
      Spacer(1, 2),
  ]

  tabla_cabecera = Table([[col_izq, col_der]], colWidths=[265, 275])
  tabla_cabecera.setStyle(
      TableStyle([
          ("BOX", (0, 0), (-1, -1), 1, colors.black),
          ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("LEFTPADDING", (0, 0), (-1, -1), 5),
          ("RIGHTPADDING", (0, 0), (-1, -1), 5),
          ("TOPPADDING", (0, 0), (-1, -1), 3),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
      ])
  )
  elementos.append(tabla_cabecera)
  elementos.append(Spacer(1, 6))

  cli = [
      [
          Paragraph(
              f"<b>Razón Social / Nombres y Apellidos:</b> {datos['receptor']}",
              estilo_general,
          ),
          "",
      ],
      [
          Paragraph(f"<b>Identificación:</b> {datos['id_receptor']}", estilo_general),
          Paragraph(f"<b>Fecha Emisión:</b> {datos['fecha']}", estilo_general),
      ],
  ]
  if "modifica" in datos:
    cli.append([Paragraph(datos["modifica"], estilo_general), ""])

  tabla_cliente = Table(cli, colWidths=[270, 270])
  tabla_cliente.setStyle(
      TableStyle([
          ("SPAN", (0, 0), (1, 0)),
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
          ("TOPPADDING", (0, 0), (-1, -1), 2),
      ])
  )
  elementos.append(tabla_cliente)
  elementos.append(Spacer(1, 5))

  t_main = Table(datos["items"], colWidths=datos["cw"])
  t_main.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
          ("FONTSIZE", (0, 0), (-1, -1), 6.5),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
          ("TOPPADDING", (0, 0), (-1, -1), 2),
          ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
      ])
  )
  elementos.append(t_main)
  elementos.append(Spacer(1, 4))

  col_izquierda_bloque = ""
  if datos["pagos"]:
    estructura_pagos = [["Forma de Pago", "Valor"]]
    for p in datos["pagos"]:
      estructura_pagos.append([Paragraph(p[0], estilo_general), p[1]])
    t_pago = Table(estructura_pagos, colWidths=[180, 50])
    t_pago.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 6.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ])
    )
    col_izquierda_bloque = t_pago

  t_tot = Table(datos["resumen"], colWidths=[110, 70])
  t_tot.setStyle(
      TableStyle([
          ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
          ("ALIGN", (1, 0), (1, -1), "RIGHT"),
          ("FONTSIZE", (0, 0), (-1, -1), 6.5),
          ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
          ("TOPPADDING", (0, 0), (-1, -1), 1.5),
          ("FONTNAME", (-1, -1), (-1, -1), "Helvetica-Bold"),
      ])
  )

  tabla_contenedor_totales = Table(
      [[col_izquierda_bloque, t_tot]], colWidths=[360, 180]
  )
  tabla_contenedor_totales.setStyle(
      TableStyle([
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("ALIGN", (1, 0), (1, 0), "RIGHT"),
      ])
  )
  elementos.append(tabla_contenedor_totales)
  elementos.append(Spacer(1, 6))

  if datos["info_adicional"]:
    estilo_info_ad = ParagraphStyle(
        "IA", parent=styles["Normal"], fontSize=6.5, leading=8.5
    )
    filas_info_ad = []
    for nombre, contenido in datos["info_adicional"]:
      filas_info_ad.append([
          Paragraph(f"<b>{nombre}</b>", estilo_info_ad),
          Paragraph(contenido if contenido else "", estilo_info_ad),
      ])
    info_ad_t = Table(filas_info_ad, colWidths=[110, 410])
    info_ad_t.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ])
    )
    recuadro_adicional = Table(
        [
            [Paragraph("<b>Información Adicional</b>", estilo_general)],
            [Spacer(1, 1)],
            [info_ad_t],
        ],
        colWidths=[540],
    )
    recuadro_adicional.setStyle(
        TableStyle([
            ("BOX", (0, 0), (-1, -1), 1, colors.grey),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ])
    )
    elementos.append(recuadro_adicional)

  doc.build(elementos)
  buffer.seek(0)
  return buffer.getvalue()


# --- INTERFAZ WEB UNIFICADA CON PESTAÑAS ---
st.title("🧾 Sistema de Gestión de Comprobantes Electrónicos SRI")
st.write(
    "Selecciona la herramienta que deseas utilizar en las pestañas de abajo."
)

tab1, tab2 = st.tabs(["📊 Conversor XML a Excel", "📄 Generador de RIDE / PDF"])

# --- CONTENIDO DE LA PESTAÑA 1: EXCEL ---
with tab1:
  st.header("Convertidor Masivo de XML a Excel")
  st.write(
      "Sube múltiples archivos XML para unificarlos en un reporte consolidado."
  )

  archivos_excel = st.file_uploader(
      "Selecciona tus archivos XML para Excel",
      type=["xml"],
      accept_multiple_files=True,
      key="excel_uploader",
  )

  if archivos_excel:
    if st.button("🚀 Procesar y Generar Excel"):
      f_list, nc_list, r_list = [], [], []
      with st.spinner("Procesando comprobantes..."):
        for archivo in archivos_excel:
          tipo, res = procesar_comprobante_excel(archivo.getvalue())
          if tipo == "01":
            f_list.append(res)
          elif tipo == "04":
            nc_list.append(res)
          elif tipo == "07":
            r_list.append(res)

      if f_list or nc_list or r_list:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          if f_list:
            pd.DataFrame(f_list).to_excel(
                writer, sheet_name="Facturas", index=False
            )
          if nc_list:
            pd.DataFrame(nc_list).to_excel(
                writer, sheet_name="Notas de Credito", index=False
            )
          if r_list:
            pd.DataFrame(r_list).to_excel(
                writer, sheet_name="Retenciones", index=False
            )
        st.success("¡Reporte Excel generado correctamente!")
        st.download_button(
            label="📥 Descargar Reporte_SRI_Unificado.xlsx",
            data=output.getvalue(),
            file_name="Reporte_SRI_Unificado.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning(
            "No se encontró información válida en los archivos seleccionados."
        )

# --- CONTENIDO DE LA PESTAÑA 2: PDF ---
with tab2:
  st.header("Generador de RIDE / PDF Individual")
  st.write("Sube tus archivos XML para visualizarlos y descargarlos en PDF.")

  archivos_pdf = st.file_uploader(
      "Selecciona tus archivos XML para PDF",
      type=["xml"],
      accept_multiple_files=True,
      key="pdf_uploader",
  )

  if archivos_pdf:
    st.write(f"### Total archivos listos: {len(archivos_pdf)}")
    for archivo in archivos_pdf:
      datos = leer_xml_completo_pdf(archivo.getvalue())
      if datos:
        col1, col2 = st.columns([3, 1])
        with col1:
          st.info(
              f"**{datos['tipo'].upper()}** | RUC: {datos['ruc']} | No.:"
              f" {datos['numero']} | Cliente: {datos['receptor']}"
          )
        with col2:
          pdf_bytes = generar_pdf_bytes(datos)
          nombre_archivo = f"{datos['tipo']}_{datos['numero']}.pdf"
          st.download_button(
              label="📥 Descargar PDF",
              data=pdf_bytes,
              file_name=nombre_archivo,
              mime="application/pdf",
              key=f"btn_{archivo.name}",
          )
      else:
        st.warning(f"No se pudo procesar el archivo: {archivo.name}")
