# -*- coding: utf-8 -*-
"""Generador de Documentos VIVA PRO — web (Streamlit). Cronograma editable + cochera."""
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
PLANTILLAS=["10% directo + 90% hipotecario","20% directo + 80% hipotecario",
            "30% directo + 70% hipotecario","40% directo + 60% hipotecario",
            "50% directo + 50% hipotecario","40% directo + 4 armadas","Personalizada (en blanco)"]
COCH_LISTA=59500

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
nombre=c1.text_input("Nombre completo *"); apellido=c2.text_input("Apellido (para Sr./Sra.) *")
sexo=c1.selectbox("Sexo",["F","M"]); dni=c2.text_input("DNI *")
estado_civil=c1.selectbox("Estado civil",["Soltero","Soltera","Casado","Casada","Conviviente"])
domicilio=c2.text_input("Domicilio")
conyuge=c1.text_input("Cónyuge (si aplica)"); conyuge_dni=c2.text_input("DNI cónyuge")
telefono=c1.text_input("Teléfono"); correo=c2.text_input("Correo")

st.subheader("Departamento")
cod=st.selectbox("Código de departamento *",sorted(DEPTOS.keys()),
    format_func=lambda k:f"{k} — {DEPTOS[k]['tipologia']} · {DEPTOS[k]['area_m2']} m² · {DEPTOS[k]['piso']}")
dep=DEPTOS[cod]
st.info(f"Tipología {dep['tipologia']} · {dep['area_m2']} m² · {dep['piso']} · BBP estimado S/ {dep.get('bbp') or 0:,}")
c3,c4=st.columns(2)
precio=c3.number_input("Precio negociado depto (S/)",value=int(dep.get("precio_lista_soles") or dep.get("precio_final_soles") or 0),step=1000)
unidad_n=c4.text_input("N° unidad registral")
fecha=c3.date_input("Fecha",value=datetime.date.today()); n_sep=c4.text_input("N° separación",value="001")

st.subheader("Cochera (opcional)")
inc_coch=st.checkbox("Incluye cochera")
cochera=None; coch_mode=None; coch_precio=0
if inc_coch:
    cc1,cc2=st.columns(2)
    est=cc1.selectbox("Estacionamiento",["E01","E02","E03","E04","E05","E06"])
    coch_precio=cc2.number_input("Precio cochera (S/)",value=COCH_LISTA,step=500)
    forma=st.selectbox("Forma de pago de la cochera",
        ["Se incorpora al crédito hipotecario","Se incorpora al crédito directo",
         "Al contado — a la firma del contrato de bien futuro","Al contado — a la firma de la escritura pública"])
    if forma.startswith("Se incorpora"):
        coch_mode="sumada"; coch_fecha=None
    else:
        coch_mode="contado"; coch_fecha=("A la firma del contrato de bien futuro." if "bien futuro" in forma else "A la firma de la escritura pública.")
    cochera={"est":est,"precio":int(coch_precio),"mode":coch_mode,"nota":forma,"fecha":coch_fecha}
    st.caption(f"Cochera {est} · 16 m² · reja corrediza (no elevadiza) · partida registral independiente · **Total con cochera: S/ {int(precio)+int(coch_precio):,}**")

# base de precio para la tabla: si la cochera va SUMADA, la tabla es sobre el total
precio_basis=int(precio)+int(coch_precio) if (cochera and coch_mode=="sumada") else int(precio)

st.subheader("Estructura de pago")
colp1,colp2=st.columns([2,1])
plant=colp1.selectbox("Plantilla de pago",PLANTILLAS)
incluir_hip=colp2.checkbox("Incluye hipotecario",value=("hipotecario" in plant))
if st.session_state.get("_plant")!=plant or st.session_state.get("_precio")!=precio_basis:
    rows0,_=C.plantilla(plant,precio_basis)
    st.session_state["_df"]=pd.DataFrame([{"Concepto":r["concepto"],"Fecha":r.get("fecha",""),
        "% precio":round(C.monto_de(r,precio_basis)/precio_basis*100,2)} for r in rows0])
    st.session_state["_plant"]=plant; st.session_state["_precio"]=precio_basis
st.caption("Edita Concepto, Fecha y % del precio. Agrega o quita armadas. La separación queda fija en S/ 3,500.")
df=st.data_editor(st.session_state["_df"],num_rows="dynamic",use_container_width=True,key="ed",
    column_config={"% precio":st.column_config.NumberColumn(format="%.2f %%",min_value=0.0,max_value=100.0)})
directo_pct=float(df["% precio"].fillna(0).sum())
if incluir_hip:
    st.success(f"Aporte directo {directo_pct:.2f}% · Crédito hipotecario {100-directo_pct:.2f}% (toma el saldo) · cuadra al 100%")
else:
    (st.success if abs(directo_pct-100)<0.5 else st.warning)(f"Sin hipotecario · suma {directo_pct:.2f}% (debe ser 100%)")

st.subheader("¿Qué generar?")
g1,g2=st.columns(2)
prop=g1.checkbox("Propuesta (cronograma de arriba)",value=True); sim=g2.checkbox("Simulación de crédito",value=True)
fic=g1.checkbox("Ficha Técnica",value=True); con=g2.checkbox("Contrato de Separación",value=True)
prof=g1.checkbox("Proforma (cotización)",value=False)

if st.button("⚙️ GENERAR DOCUMENTOS",use_container_width=True,type="primary"):
    if not nombre or not dni or not cod:
        st.warning("Completa al menos Nombre, DNI y Departamento."); st.stop()
    if not incluir_hip and prop and abs(directo_pct-100)>0.5:
        st.warning("Sin hipotecario, el cronograma debe sumar 100%. Ajusta los %."); st.stop()
    ape=(apellido or nombre.split()[0]).strip()
    with tempfile.TemporaryDirectory() as out:
        base=dict(carpeta_salida=out,nombre=nombre,apellido=ape,sexo=sexo,dni=dni,estado_civil=estado_civil,
            conyuge=conyuge or None,conyuge_dni=conyuge_dni,domicilio=domicilio,telefono=telefono,correo=correo,
            codigo_depto=cod,unidad_n=unidad_n,precio_soles=int(precio),separacion_soles=3500,separacion_usd=1000,
            fecha=fecha.isoformat(),n_sep=n_sep,planos_dir=PLANOS)
        if cochera: base["cochera"]=cochera
        errs=[]
        with st.spinner("Generando documentos..."):
            rows=[]
            for _,r in df.iterrows():
                c_=r["Concepto"]
                if pd.isna(c_) or not str(c_).strip(): continue
                pv=r["% precio"]; pctv=0.0 if pd.isna(pv) else float(pv)
                fec="" if pd.isna(r["Fecha"]) else str(r["Fecha"]).strip()
                es_sep=str(c_).strip().lower().startswith("separaci")
                monto=3500.0 if es_sep else precio_basis*pctv/100
                if monto<=0 and not es_sep: continue
                rows.append({"concepto":str(c_).strip(),"fecha":fec,"monto":monto})
            # cochera APARTE: agregar línea(s)
            if cochera and coch_mode=="contado":
                sub2="Cochera 16 m² · reja corrediza (no elevadiza) · partida registral independiente."
                rows.append({"concepto":f"Estacionamiento N° {cochera['est']} (al contado)","fecha":cochera.get("fecha") or "A la firma del contrato.","monto":float(coch_precio),"sub2":sub2})
            hip={"concepto":"Saldo con crédito hipotecario","fecha":"Contra entrega (diciembre de 2027).",
                 "sub2":"Tasa, plazo y cuota los define el banco."} if incluir_hip else None
            inicial_pct=round(directo_pct/100,4) if incluir_hip else 0.20
            if prop:
                try: GP.build_propuesta(dict(base,cronograma=rows,hipotecario=hip))
                except Exception as ex: errs.append("Propuesta: "+str(ex))
            if sim or fic or con:
                cfg=dict(base,opciones=[],incluir_simulacion=sim,plazo_anios=20,tipo_cuota="Simple",
                         tea_mivivienda=0.09,tea_tradicional=0.09,inicial_pct=inicial_pct)
                ok,log=run_engine("generar_paquete.py",cfg)
                if not ok: errs.append("Paquete: "+log)
                if not fic:
                    for f in list(os.listdir(out)):
                        if f.startswith("Ficha_"): os.remove(os.path.join(out,f))
                if not con:
                    for f in list(os.listdir(out)):
                        if f.startswith("Contrato_"): os.remove(os.path.join(out,f))
            if prof:
                forma_lbl=plant if not str(plant).startswith("Personalizada") else "Plan personalizado"
                grand=int(precio)+(int(coch_precio) if cochera else 0)
                crono_p=[[r["concepto"], f"S/ {round(r['monto']):,}"] for r in rows]
                if hip:
                    saldo=grand-sum(r["monto"] for r in rows)
                    crono_p.append(["Saldo con crédito hipotecario (contra entrega)", f"S/ {round(saldo):,}"])
                ok,log=run_engine("generar_proforma.py",dict(base,forma_pago=forma_lbl,inicial_pct=round(directo_pct/100,4),cronograma=crono_p))
                if not ok: errs.append("Proforma: "+log)
            # numerar nombres para que se ordenen: 1.Proforma 2.Ficha 3.Propuesta 4.Simulacion 5.Contrato
            _ordn=[("proforma","1"),("ficha","2"),("propuesta","3"),("simulacion","4"),("simulación","4"),("contrato","5")]
            for f in list(os.listdir(out)):
                if not f.lower().endswith(".pdf"): continue
                base_f=f.split(". ",1)[1] if f[:2].strip().rstrip(".").isdigit() and ". " in f[:4] else f
                pref=next((n for k,n in _ordn if base_f.lower().startswith(k)),None)
                if pref and f!=f"{pref}. {base_f}":
                    try: os.rename(os.path.join(out,f),os.path.join(out,f"{pref}. {base_f}"))
                    except Exception: pass
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
