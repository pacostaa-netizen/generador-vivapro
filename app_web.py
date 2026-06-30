# -*- coding: utf-8 -*-
"""Generador de Documentos VIVA PRO — versión web (Streamlit Community Cloud)."""
import os, sys, json, tempfile, subprocess, datetime, zipfile, io
import streamlit as st

APPDIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLAS = APPDIR
PLANOS = APPDIR
DEPTOS = json.load(open(os.path.join(PLANTILLAS, "deptos.json"), encoding="utf-8"))

def run_engine(script, cfg):
    fd, p = tempfile.mkstemp(suffix=".json"); os.close(fd)
    json.dump(cfg, open(p, "w", encoding="utf-8"))
    try:
        r = subprocess.run([sys.executable, os.path.join(APPDIR, script), p],
                           capture_output=True, text=True)
        return r.returncode == 0, (r.stdout + "\n" + r.stderr).strip()
    finally:
        os.remove(p)

st.set_page_config(page_title="Generador VIVA PRO", page_icon="🏢", layout="centered")
st.title("🏢 Generador de Documentos — VIVA PRO")
st.caption("LA HAUS CONSTRUCTORA S.A.C. · Llena los datos y descarga los PDFs del cliente.")

codigos = sorted(DEPTOS.keys())
with st.form("frm"):
    st.subheader("Datos del cliente")
    c1, c2 = st.columns(2)
    nombre = c1.text_input("Nombre completo *")
    apellido = c2.text_input("Apellido (para Sr./Sra.) *")
    sexo = c1.selectbox("Sexo", ["F", "M"])
    dni = c2.text_input("DNI *")
    estado_civil = c1.selectbox("Estado civil", ["Soltero","Soltera","Casado","Casada","Conviviente"])
    domicilio = c2.text_input("Domicilio")
    conyuge = c1.text_input("Cónyuge (si aplica)")
    conyuge_dni = c2.text_input("DNI cónyuge")
    telefono = c1.text_input("Teléfono")
    correo = c2.text_input("Correo")

    st.subheader("Departamento")
    cod = st.selectbox("Código de departamento *", codigos,
                       format_func=lambda k: f"{k} — {DEPTOS[k]['tipologia']} · {DEPTOS[k]['area_m2']} m² · {DEPTOS[k]['piso']}")
    dep = DEPTOS[cod]
    st.info(f"Tipología {dep['tipologia']} · {dep['area_m2']} m² · {dep['piso']} · BBP estimado S/ {dep.get('bbp') or 0:,}")
    c3, c4 = st.columns(2)
    precio = c3.number_input("Precio negociado (S/)", value=int(dep.get("precio_final_soles") or 0), step=1000)
    unidad_n = c4.text_input("N° unidad registral")
    fecha = c3.date_input("Fecha", value=datetime.date.today())
    n_sep = c4.text_input("N° separación", value="001")

    st.subheader("¿Qué generar?")
    c5, c6 = st.columns(2)
    opA = c5.checkbox("Propuesta Opción A (directo)", value=True)
    opB = c6.checkbox("Propuesta Opción B (50/50)", value=False)
    sim = c5.checkbox("Simulación de crédito", value=True)
    prof = c6.checkbox("Proforma (cotización)", value=True)

    enviar = st.form_submit_button("⚙️ GENERAR DOCUMENTOS", use_container_width=True)

if enviar:
    if not nombre or not dni or not cod:
        st.warning("Completa al menos Nombre, DNI y Departamento.")
        st.stop()
    ape = (apellido or nombre.split()[0]).replace(" ", "")
    with tempfile.TemporaryDirectory() as out:
        base = dict(carpeta_salida=out, nombre=nombre, apellido=ape, sexo=sexo, dni=dni,
                    estado_civil=estado_civil, conyuge=conyuge or None, conyuge_dni=conyuge_dni,
                    domicilio=domicilio, telefono=telefono, correo=correo, codigo_depto=cod,
                    unidad_n=unidad_n, precio_soles=int(precio) or None, separacion_soles=3500,
                    separacion_usd=1000, fecha=fecha.isoformat(), n_sep=n_sep, planos_dir=PLANOS)
        errores = []
        with st.spinner("Generando documentos..."):
            ops = (["A"] if opA else []) + (["B"] if opB else [])
            if ops or sim:
                cfg = dict(base, opciones=ops, incluir_simulacion=sim,
                           plazo_anios=20, tipo_cuota="Simple", tea_mivivienda=0.09,
                           tea_tradicional=0.09, inicial_pct=0.20)
                ok, log = run_engine("generar_paquete.py", cfg)
                if not ok: errores.append("Paquete: " + log)
            if prof:
                cfg = dict(base, forma_pago="Crédito Directo", inicial_pct=0.40)
                ok, log = run_engine("generar_proforma.py", cfg)
                if not ok: errores.append("Proforma: " + log)
            pdfs = [f for f in os.listdir(out) if f.lower().endswith(".pdf")]
        if errores:
            st.error("Ocurrió un problema:\n\n" + "\n\n".join(errores))
        if pdfs:
            st.success(f"✅ {len(pdfs)} documento(s) generado(s).")
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for f in sorted(pdfs): z.write(os.path.join(out, f), f)
            st.download_button("⬇️ Descargar todo (ZIP)", buf.getvalue(),
                               file_name=f"Documentos_{ape}_{cod}.zip", mime="application/zip",
                               use_container_width=True)
            for f in sorted(pdfs):
                with open(os.path.join(out, f), "rb") as fh:
                    st.download_button("⬇️ " + f, fh.read(), file_name=f, mime="application/pdf")
