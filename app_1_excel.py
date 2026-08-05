import io
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st

# Configuración de la página web
st.set_page_config(
    page_title="Convertidor XML a Excel - SRI Ecuador",
    page_icon="📊",
    layout="wide",
)


def obtener_valor(nodo, tag):
  if nodo is None:
    return ""
  element = nodo.find(tag)
  return element.text if element is not None else ""


def formatear_num_doc(numero_crudo):
  """Transforma 001010000000255 en 001-010-000000255"""
  if not numero_crudo:
    return ""
  limpio = str(numero_crudo).replace("-", "").strip()
  if len(limpio) == 15:
    return f"{limpio[:3]}-{limpio[3:6]}-{limpio[6:]}"
  return numero_crudo


def procesar_comprobante_contenido(contenido_xml):
  try:
    root = ET.fromstring(contenido_xml)

    comprobante_node = root.find("comprobante")
    if comprobante_node is not None:
      xml_data = ET.fromstring(comprobante_node.text)
      fecha_aut = obtener_valor(root, "fechaAutorizacion")
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

    # --- FACTURAS (01) Y NOTAS DE CRÉDITO (04) ---
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

    # --- RETENCIONES (07) ---
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

  except Exception as e:
    print(f"Error procesando XML: {e}")
    return None, None
  return None, None


# --- INTERFAZ WEB STREAMLIT ---
st.title("📊 Procesador de XML del SRI a Excel")
st.write(
    "Sube tus comprobantes electrónicos en formato XML (Facturas, Notas de"
    " Crédito y Retenciones) para unificarlos en un solo archivo de Excel."
)

archivos_subidos = st.file_uploader(
    "Selecciona o arrastra tus archivos XML",
    type=["xml"],
    accept_multiple_files=True,
)

if archivos_subidos:
  if st.button("🚀 Procesar Comprobantes"):
    f_list, nc_list, r_list = [], [], []

    with st.spinner("Procesando archivos..."):
      for archivo in archivos_subidos:
        contenido = archivo.getvalue()
        tipo, res = procesar_comprobante_contenido(contenido)
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
          pd.DataFrame(f_list).to_excel(writer, sheet_name="Facturas", index=False)
        if nc_list:
          pd.DataFrame(nc_list).to_excel(
              writer, sheet_name="Notas de Credito", index=False
          )
        if r_list:
          pd.DataFrame(r_list).to_excel(
              writer, sheet_name="Retenciones", index=False
          )
      excel_data = output.getvalue()

      st.success("¡Reporte generado con éxito!")
      st.download_button(
          label="📥 Descargar Reporte Excel Unificado",
          data=excel_data,
          file_name="Reporte_SRI_Unificado.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning(
          "No se encontró información válida en los archivos subidos o formato"
          " no compatible."
      )