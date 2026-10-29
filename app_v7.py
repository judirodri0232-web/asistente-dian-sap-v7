import re
import unicodedata
from difflib import SequenceMatcher
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Asistente DIAN - SAP Business One V7", page_icon="📊", layout="wide")
st.title("Asistente de Validación de Facturas DIAN e Impuestos en SAP Business One")
st.caption("V7 · Búsqueda en TT + LEG para facturas y NC para notas crédito · Referencias flexibles · Auditoría")

def norm_text(value):
    if pd.isna(value): return ""
    s = str(value).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]", "", s)

def norm_name(value):
    if pd.isna(value): return ""
    s = str(value).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def norm_number(value):
    if pd.isna(value) or value == "": return None
    if isinstance(value, (int, float)): return float(value)
    s = str(value).strip().replace("$", "").replace(" ", "")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s: return None
    try:
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
        elif "," in s:
            parts = s.split(",")
            s = s.replace(",", ".") if len(parts[-1]) in (1, 2) else s.replace(",", "")
        elif s.count(".") > 1: s = s.replace(".", "")
        return float(s)
    except Exception: return None

def norm_date(value):
    if pd.isna(value) or value == "": return pd.NaT
    return pd.to_datetime(value, errors="coerce", dayfirst=True)

def find_col(df, candidates):
    normalized = {norm_text(c): c for c in df.columns}
    for candidate in candidates:
        if norm_text(candidate) in normalized: return normalized[norm_text(candidate)]
    for c in df.columns:
        nc = norm_text(c)
        for candidate in candidates:
            kc = norm_text(candidate)
            if kc in nc or nc in kc: return c
    return None

def provider_score(a, b):
    a, b = norm_name(a), norm_name(b)
    if not a or not b: return 0
    if a == b: return 1.0
    if a in b or b in a: return 0.92
    ta, tb = set(a.split()), set(b.split())
    common = len(ta & tb) / max(1, len(ta | tb))
    return max(common, SequenceMatcher(None, a, b).ratio() * 0.85)

def ref_tokens(value):
    s = norm_text(value)
    nums = re.findall(r"\d+", s)
    return s, nums

def safe_columns(df):
    df = df.copy()
    for c in df.select_dtypes(include=["object"]).columns:
        df[c] = df[c].fillna("").astype(str)
    return df

def load_sap(sap_file):
    xls = pd.ExcelFile(sap_file)
    parts = []
    for sheet in ["TT", "LEG", "NC"]:
        if sheet in xls.sheet_names:
            part = pd.read_excel(sap_file, sheet_name=sheet)
            part["Tipo SAP"] = sheet
            parts.append(part)
    if not parts:
        raise ValueError("El archivo SAP debe tener al menos una hoja llamada TT, LEG o NC.")
    return pd.concat(parts, ignore_index=True)

def reconcile(dian, sap):
    d_type = find_col(dian, ["Tipo de documento"])
    d_prefix = find_col(dian, ["Prefijo"])
    d_folio = find_col(dian, ["Folio", "Número", "Numero"])
    d_date = find_col(dian, ["Fecha de emisión", "Fecha Emision", "Fecha"])
    d_nit = find_col(dian, ["NIT emisor", "NIT Emisor", "NIT"])
    d_name = find_col(dian, ["Nombre emisor", "Nombre Emisor", "Proveedor", "Razón social"])
    d_total = find_col(dian, ["Total", "Total factura", "Valor total"])
    d_iva = find_col(dian, ["IVA"])
    s_doc = find_col(sap, ["#", "Número del documento", "Numero del documento", "Documento", "Número"])
    s_date = find_col(sap, ["Fecha", "Fecha del documento"])
    s_name = find_col(sap, ["Proveedor", "Nombre proveedor", "Nombre del proveedor"])
    s_total = find_col(sap, ["Total del documento", "Total documento", "Total"])
    s_ref = find_col(sap, ["Referencia", "Referencia TT", "Ref"])
    if not all([d_folio, d_name, d_total, s_doc, s_date, s_name, s_total]):
        raise ValueError("No identifiqué todas las columnas obligatorias. Verifica los encabezados de los Excel.")

    # Detectar posibles TT repetidos/anulaciones, sin borrarlos del análisis.
    dup_groups = {}
    for i, row in sap.iterrows():
        ref = norm_text(row.get(s_ref, "")) if s_ref else ""
        prov = norm_name(row.get(s_name, ""))
        total = norm_number(row.get(s_total, ""))
        if ref and prov and total is not None:
            dup_groups.setdefault((str(row.get("Tipo SAP", "")), ref, prov, round(total, 2)), []).append(i)
    dup_ids = {i for ids in dup_groups.values() if len(ids) > 1 for i in ids}
    dup_rows = []
    for key, ids in dup_groups.items():
        if len(ids) > 1:
            for i in ids:
                r = sap.loc[i]
                dup_rows.append({"Grupo revisión": f"REV-{len(dup_rows)+1:03d}", "Tipo SAP": r.get("Tipo SAP",""),
                    "Documento SAP": r.get(s_doc,""), "Referencia SAP": r.get(s_ref,""),
                    "Proveedor SAP": r.get(s_name,""), "Fecha SAP": r.get(s_date,""),
                    "Total SAP": norm_number(r.get(s_total,"")), "Documentos del grupo": ", ".join(str(sap.loc[j,s_doc]) for j in ids),
                    "Estado": "POSIBLE DUPLICADO / ANULACIÓN - VALIDAR EN SAP"})
    duplicate_df = pd.DataFrame(dup_rows)
    used, rows = set(), []
    for di, dr in dian.iterrows():
        dtype = str(dr.get(d_type, "")) if d_type else ""
        is_credit = "nota de credito" in norm_name(dtype).lower()
        expected = {"NC"} if is_credit else {"TT", "LEG"} if "factura electronica" in norm_name(dtype).lower() else {"TT", "LEG"}
        folio = norm_text(dr.get(d_folio, ""))
        prefix = norm_text(dr.get(d_prefix, "")) if d_prefix else ""
        dref = prefix + folio
        ddate = norm_date(dr.get(d_date, "")) if d_date else pd.NaT
        dname = dr.get(d_name, "")
        dnit = norm_text(dr.get(d_nit, "")) if d_nit else ""
        total = norm_number(dr.get(d_total, ""))
        iva = norm_number(dr.get(d_iva, "")) if d_iva else None
        candidates = []
        for si, sr in sap.iterrows():
            if si in used or str(sr.get("Tipo SAP","")) not in expected: continue
            sref_raw = sr.get(s_ref, "") if s_ref else ""
            sref = norm_text(sref_raw)
            sdoc = sr.get(s_doc, "")
            sdate = norm_date(sr.get(s_date, ""))
            stotal = norm_number(sr.get(s_total, ""))
            ps = provider_score(dname, sr.get(s_name, ""))
            exact = bool(dref and sref and dref == sref)
            folio_match = bool(folio and folio in sref)
            # Referencia flexible: números coincidentes incluso si prefijo/texto no coincide.
            _, dnums = ref_tokens(folio)
            _, snums = ref_tokens(sref_raw)
            number_match = bool(set(dnums) & set(snums)) if dnums else False
            date_diff = abs((ddate - sdate).days) if pd.notna(ddate) and pd.notna(sdate) else None
            value_diff = abs(total - stotal) if total is not None and stotal is not None else None
            score = 0
            reasons = []
            if exact: score += 48; reasons.append("referencia exacta")
            elif folio_match: score += 42; reasons.append("folio contenido en referencia SAP")
            elif number_match: score += 30; reasons.append("número coincidente dentro de referencia")
            if ps >= .9: score += 27; reasons.append("proveedor muy similar")
            elif ps >= .65: score += 19; reasons.append("proveedor parcialmente similar")
            elif ps >= .45: score += 9; reasons.append("similitud parcial del proveedor")
            if date_diff is not None:
                if date_diff == 0: score += 12; reasons.append("misma fecha")
                elif date_diff <= 3: score += 9; reasons.append(f"fecha a {date_diff} días")
                elif date_diff <= 10: score += 5; reasons.append(f"fecha a {date_diff} días")
                elif date_diff <= 30: score += 1
            if value_diff is not None:
                # El valor ayuda a ordenar, pero no descarta por sí solo; retenciones pueden explicar diferencias.
                rel = value_diff / max(abs(total), 1)
                if value_diff <= 1: score += 13; reasons.append("total igual")
                elif rel <= .02: score += 10; reasons.append("diferencia hasta 2%")
                elif rel <= .10: score += 6; reasons.append("diferencia hasta 10%")
                elif rel <= .25: score += 2; reasons.append("diferencia de valor relevante")
            if not exact and not folio_match and not number_match and ps < .45:
                continue
            # Para NOVALENE permitir búsqueda por tercero + fecha + importe, aunque referencia no coincida.
            if "NOVALENE" in norm_name(dname) and ps >= .65 and date_diff is not None and date_diff <= 10:
                score += 8
                reasons.append("búsqueda ampliada por proveedor NOVALENE")
            candidates.append({"idx":si,"doc":sdoc,"ref":sref_raw,"provider":sr.get(s_name,""),"date":sr.get(s_date,""),
                "total":stotal,"diff":value_diff,"score":score,"reasons":reasons,"tipo":sr.get("Tipo SAP",""),
                "ps":ps,"date_diff":date_diff,"dup":si in dup_ids})
        candidates.sort(key=lambda c: (c["score"], -(c["diff"] if c["diff"] is not None else 1e99)), reverse=True)
        base = {"DIAN índice":di,"Tipo documento DIAN":dtype,"Factura DIAN":dref,"Folio DIAN":dr.get(d_folio,""),
                "Prefijo DIAN":dr.get(d_prefix,"") if d_prefix else "","NIT DIAN":dnit,"Proveedor DIAN":dname,
                "Fecha DIAN":dr.get(d_date,"") if d_date else "","Total DIAN":total,"IVA DIAN":iva,
                "Hojas SAP permitidas":", ".join(sorted(expected))}
        if not candidates:
            base.update({"Tipo SAP":"","Documento SAP":"","Referencia SAP":"","Proveedor SAP":"","Fecha SAP":"",
                "Total SAP":None,"Diferencia valor":None,"Puntaje":0,"Confianza":"BAJA","Estado":"SIN CORRESPONDENCIA",
                "Criterio":"No se encontró candidato en las hojas permitidas para el tipo de documento."})
        else:
            best = candidates[0]
            second = candidates[1]["score"] if len(candidates)>1 else -1
            # Evitar asignación si candidatos son casi equivalentes; marcar revisión.
            min_score = 42 if best["ps"] >= .65 else 52
            ambiguous = len(candidates)>1 and best["score"]-second < 5
            if best["dup"]:
                base.update({"Tipo SAP":best["tipo"],"Documento SAP":"","Referencia SAP":"","Proveedor SAP":"","Fecha SAP":"",
                    "Total SAP":None,"Diferencia valor":None,"Puntaje":best["score"],"Confianza":"REVISIÓN",
                    "Estado":"POSIBLE DUPLICADO / ANULACIÓN","Criterio":"El mejor candidato pertenece a un grupo repetido; validar en SAP antes de asignar."})
            elif best["score"] < min_score or ambiguous:
                base.update({"Tipo SAP":"","Documento SAP":"","Referencia SAP":"","Proveedor SAP":"","Fecha SAP":"",
                    "Total SAP":None,"Diferencia valor":None,"Puntaje":best["score"],"Confianza":"REVISIÓN",
                    "Estado":"CANDIDATO PARA REVISAR","Criterio":"Candidato no suficientemente claro o hay candidatos con puntajes similares. Mejor candidato: "+str(best["tipo"])+" "+str(best["doc"])+"; "+ " + ".join(best["reasons"])})
            else:
                used.add(best["idx"])
                diff = best["diff"]
                if diff is not None and diff <= 1: state, conf = "CONCILIADA", "ALTA"
                elif best["score"] >= 70 and best["ps"] >= .65: state, conf = "ASOCIADA CON DIFERENCIA", "ALTA"
                else: state, conf = "ASOCIADA CON DIFERENCIA", "MEDIA"
                base.update({"Tipo SAP":best["tipo"],"Documento SAP":best["doc"],"Referencia SAP":best["ref"],
                    "Proveedor SAP":best["provider"],"Fecha SAP":best["date"],"Total SAP":best["total"],
                    "Diferencia valor":diff,"Puntaje":best["score"],"Confianza":conf,"Estado":state,
                    "Criterio":" + ".join(best["reasons"])})
        rows.append(base)

    result = pd.DataFrame(rows)
    un = []
    for si, sr in sap.iterrows():
        if si not in used:
            un.append({"Tipo SAP":sr.get("Tipo SAP",""),"Documento SAP":sr.get(s_doc,""),"Referencia SAP":sr.get(s_ref,"") if s_ref else "",
                "Fecha SAP":sr.get(s_date,""),"Proveedor SAP":sr.get(s_name,""),"Total SAP":norm_number(sr.get(s_total,"")),
                "Estado":"SAP SIN ASOCIAR","Observación":"No se utilizó en una asociación automática; puede requerir revisión."})
    return safe_columns(result), safe_columns(pd.DataFrame(un)), safe_columns(duplicate_df)

def make_excel(result, unmatched, duplicates):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="DIAN_asociacion")
        unmatched.to_excel(writer, index=False, sheet_name="SAP_no_asociados")
        duplicates.to_excel(writer, index=False, sheet_name="Duplicados_anulaciones")
        result["Estado"].value_counts(dropna=False).rename_axis("Estado").reset_index(name="Cantidad").to_excel(writer,index=False,sheet_name="Resumen")
    return output.getvalue()

with st.sidebar:
    st.header("📁 Archivos")
    dian_file = st.file_uploader("1. Archivo DIAN 2025", type=["xlsx","xls"], key="dian")
    sap_file = st.file_uploader("2. Archivo SAP 2025 (hojas TT, LEG, NC)", type=["xlsx","xls"], key="sap")
    st.info("Facturas electrónicas → TT y LEG. Notas de crédito electrónicas → NC. La asociación la determina el asistente.")
if not dian_file or not sap_file:
    st.info("Carga el archivo DIAN y el archivo SAP para ejecutar la V7.")
    st.stop()
try:
    dian_xls = pd.ExcelFile(dian_file)
    dsheet = "FC DIAN" if "FC DIAN" in dian_xls.sheet_names else dian_xls.sheet_names[0]
    dian = pd.read_excel(dian_file, sheet_name=dsheet)
    sap = load_sap(sap_file)
    result, unmatched, duplicates = reconcile(dian, sap)
except Exception as e:
    st.error(f"No fue posible procesar los archivos: {e}")
    st.stop()

states = result["Estado"].value_counts()
n = len(result)
associated = int(result["Documento SAP"].fillna("").astype(str).ne("").sum())
conc = int((result["Estado"]=="CONCILIADA").sum())
no_match = int((result["Estado"]=="SIN CORRESPONDENCIA").sum())
review = int(result["Estado"].isin(["CANDIDATO PARA REVISAR","POSIBLE DUPLICADO / ANULACIÓN"]).sum())
st.header("📊 Resultado de asociación automática")
a,b,c,d,e,f = st.columns(6)
a.metric("Documentos DIAN", n)
b.metric("Asociados", associated)
c.metric("Conciliados", conc)
d.metric("Con diferencia", int((result["Estado"]=="ASOCIADA CON DIFERENCIA").sum()))
e.metric("Sin correspondencia", no_match)
f.metric("Revisar", review)
st.caption(f"DIAN: hoja {dsheet} · SAP: TT, LEG y NC. Las diferencias de valor no descartan automáticamente una asociación.")
tab1,tab2,tab3,tab4 = st.tabs(["📈 Resumen","📄 Asociación DIAN","🟠 Posibles anulaciones","🧾 SAP no asociados"])
with tab1:
    st.dataframe(states.rename_axis("Estado").reset_index(name="Cantidad"), hide_index=True, use_container_width=True)
    fig=px.pie(states.rename_axis("Estado").reset_index(name="Cantidad"), names="Estado", values="Cantidad", hole=.35)
    st.plotly_chart(fig, use_container_width=True)
with tab2:
    selected=st.multiselect("Filtrar estados", list(states.index), default=list(states.index))
    st.dataframe(result[result["Estado"].isin(selected)], hide_index=True, use_container_width=True)
with tab3:
    st.warning("Los documentos repetidos se muestran como posibles anulaciones. No se eliminan automáticamente; verifica en SAP.")
    st.dataframe(duplicates, hide_index=True, use_container_width=True)
with tab4:
    st.dataframe(unmatched, hide_index=True, use_container_width=True)
st.subheader("⬇️ Exportar resultados")
st.download_button("Descargar conciliación V7 (Excel)", data=make_excel(result,unmatched,duplicates),
    file_name="conciliacion_DIAN_SAP_V7.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.download_button("Descargar asociaciones V7 (CSV)", data=result.to_csv(index=False,encoding="utf-8-sig").encode("utf-8-sig"),
    file_name="DIAN_asociacion_V7.csv", mime="text/csv")
