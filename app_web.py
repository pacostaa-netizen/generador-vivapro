# -*- coding: utf-8 -*-
"""Generador de Documentos VIVA PRO — web (Streamlit). Cronograma de pago editable."""
import os, sys, json, tempfile, subprocess, datetime, zipfile, io
import streamlit as st
import pandas as pd
import cronograma as C
import generar_propuesta as GP

APPDIR=os.path.dirname(os.path.abspath(__file__))
TPL=os.path.join(APPDIR,"plantillas")
if not os.path.exists(os.path.join(TPL,"deptos.json")): TPL=APPDIR
PLANOS=TPL
DEPTOS=json.load(open(os.path.join(TPL,"deptos.json"),encoding="utf-8"))
PLANTILLAS=["30% directo + 70% hipotecario","20% directo + 80% hipotecario",
            "50% directo + 50% hipotecario","40% directo + 4 armadas","Personalizada (en blanco)"]

def run_engine(script,cfg):
    fd,p=tempfile.mkstemp(suffix=".json"); os.close(fd); json.dump(cfg,open(p,"w",encoding="utf-8"))
    try:
        r=subprocess.run([sys.executable,os.path.join(APPDIR,script),p],capture_output=True,text=True)
        return r.returncode==0,(r.stdout+"\n"+r.stderr).strip()
    finally: os.remove(p)

st.set_page_config(page_title="Generador VIVA PRO",page_icon="🏢",layout="centered")
st.title("🏢 Generador de Documentos — VIVA PRO")
st.caption("LA HAUS CONSTRUCTORA S.A.C. · Llena los datos y descarga los PDFs del cliente.")

st.subheader("Datos del cliente")
c1,c2=st.columns(2)
nombre=c1.text_input("Nombre completo *")
apellido=c2.text_input("Apellido (para Sr./Sra.) *")
sexo=c1.selectbox("Sexo",["F","M"])
dni=c2.text_input("DNI *")
estado_civil=c1.selectbox("Estado civil",["Soltero","Soltera","Casado","Casada","Conviviente"])
domicilio=c2.text_input("Domicilio")
conyuge=c1.text_input("Cónyuge (si aplica)")
conyuge_dni=c2.text_input("DNI cónyuge")
telefono=c1.text_input("Teléfono"); correo=c2.text_input("Correo")

st.subheader("Departamento")
cod=st.selectbox("Código de departamento *",sorted(DEPTOS.keys()),
    format_func=lambda k:f"{k} — {DEPTOS[k]['tipologia']} · {DEPTOS[k]['area_m2']} m² · {DEPTOS[k]['piso']}")
dep=DEPTOS[cod]
st.info(f"Tipología {dep['tipologia']} · {dep['area_m2']} m² · {dep['piso']} · BBP estimado S/ {dep.get('bbp') or 0:,}")
c3,c4=st.columns(2)
precio=c3.number_input("Precio negociado (S/)",value=int(dep.get("precio_final_soles") or 0),step=1000)
unidad_n=c4.text_input("N° unidad registral")
fecha=c3.date_input("Fecha",value=datetime.date.today()); n_sep=c4.text_input("N° separación",value="001")

st.subheader("Estructura de pago")
colp1,colp2=st.columns([2,1])
plant=colp1.selectbox("Plantilla de pago",PLANTILLAS)
incluir_hip=colp2.checkbox("Incluye hipotecario",value=("hipotecario" in plant))
if st.session_state.get("_plant")!=plant or st.session_state.get("_precio")!=precio:
    rows,_=C.plantilla(plant,precio); 
    st.session_state["_df"]=pd.DataFrame([{"Concepto":r["concepto"],"Fecha":r.get("fecha",""),
        "% precio":round(C.monto_de(r,precio)/precio*100,2)} for r in rows])
    st.session_state["_plant"]=plant; st.session_state["_precio"]=precio
st.caption("Edita Concepto, Fecha y % del precio. Agrega o quita armadas con + / 🗑. La separación queda fija en S/ 3,500.")
df=st.data_editor(st.session_state["_df"],num_rows="dynamic",use_container_width=True,key="ed",
    column_config={"% precio":st.column_config.NumberColumn(format="%.2f %%",min_value=0.0,max_value=100.0)})
directo_pct=float(df["% precio"].fillna(0).sum())
if incluir_hip:
    st.success(f"Aporte directo {directo_pct:.2f}% · Crédito hipotecario {100-directo_pct:.2f}% (toma el saldo) · cuadra al 100%")
else:
    estado="✅ cuadra 100%" if abs(directo_pct-100)<0.5 else f"⚠️ suma {directo_pct:.2f}% (debe ser 100%)"
    (st.success if abs(directo_pct-100)<0.5 else st.warning)(f"Sin hipotecario · {estado}")

st.subheader("¿Qué generar?")
g1,g2=st.columns(2)
prop=g1.checkbox("Propuesta (cronograma de arriba)",value=True)
sim=g2.checkbox("Simulación de crédito",value=True)
fic=g1.checkbox("Ficha Técnica",value=True)
con=g2.checkbox("Contrato de Separación",value=True)
prof=g1.checkbox("Proforma (cotización)",value=False)

if st.button("⚙️ GENERAR DOCUMENTOS",use_container_width=True,type="primary"):
    if not nombre or not dni or not cod:
        st.warning("Completa al menos Nombre, DNI y Departamento."); st.stop()
    if not incluir_hip and prop and abs(directo_pct-100)>0.5:
        st.warning("Sin hipotecario, el cronograma debe sumar 100%. Ajusta los %."); st.stop()
    ape=(apellido or nombre.split()[0]).replace(" ","")
    with tempfile.TemporaryDirectory() as out:
        base=dict(carpeta_salida=out,nombre=nombre,apellido=ape,sexo=sexo,dni=dni,estado_civil=estado_civil,
            conyuge=conyuge or None,conyuge_dni=conyuge_dni,domicilio=domicilio,telefono=telefono,correo=correo,
            codigo_depto=cod,unidad_n=unidad_n,precio_soles=int(precio) or None,separacion_soles=3500,separacion_usd=1000,
            fecha=fecha.isoformat(),n_sep=n_sep,planos_dir=PLANOS)
        errs=[]
        with st.spinner("Generando documentos..."):
            # cronograma a filas para el motor
            rows=[]
            for _,r in df.iterrows():
                con=r["Concepto"]
                if pd.isna(con) or not str(con).strip(): continue
                pv=r["% precio"]; pctv=0.0 if pd.isna(pv) else float(pv)
                fec="" if pd.isna(r["Fecha"]) else str(r["Fecha"]).strip()
                monto=3500.0 if str(con).strip().lower().startswith("separaci") else precio*pctv/100
                rows.append({"concepto":str(con).strip(),"fecha":fec,"monto":monto})
            hip={"concepto":"Saldo con crédito hipotecario","fecha":"Contra entrega (diciembre de 2027).",
                 "sub2":"Tasa, plazo y cuota los define el banco."} if incluir_hip else None
            inicial_pct=round(directo_pct/100,4) if incluir_hip else 0.20
            if prop:
                try: GP.build_propuesta(dict(base,cronograma=rows,hipotecario=hip))
                except Exception as ex: errs.append("Propuesta: "+str(ex))
            if sim or fic or con:
                cfg=dict(base,opciones=[],incluir_simulacion=sim,plazo_anios=20,tipo_cuota="Simple",
                         tea_mivivienda=0.09,tea_tradicional=0.09,inicial_pct=inicial_pct)
                # ficha/contrato siempre; si no se quiere alguno se borra luego
                ok,log=run_engine("generar_paquete.py",cfg)
                if not ok: errs.append("Paquete: "+log)
                if not fic:
                    for f in list(os.listdir(out)):
                        if f.startswith("Ficha_"): os.remove(os.path.join(out,f))
                if not con:
                    for f in list(os.listdir(out)):
                        if f.startswith("Contrato_"): os.remove(os.path.join(out,f))
            if prof:
                ok,log=run_engine("generar_proforma.py",dict(base,forma_pago="Crédito Directo",inicial_pct=0.40))
                if not ok: errs.append("Proforma: "+log)
            pdfs=[f for f in os.listdir(out) if f.lower().endswith(".pdf")]
        if errs: st.error("Problemas:\n\n"+"\n\n".join(errs))
        if pdfs:
            st.success(f"✅ {len(pdfs)} documento(s) generado(s).")
            buf=io.BytesIO()
            with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
                for f in sorted(pdfs): z.write(os.path.join(out,f),f)
            st.download_button("⬇️ Descargar todo (ZIP)",buf.getvalue(),file_name=f"Documentos_{ape}_{cod}.zip",
                mime="application/zip",use_container_width=True)
            for f in sorted(pdfs):
                st.download_button("⬇️ "+f,open(os.path.join(out,f),"rb").read(),file_name=f,mime="application/pdf")
