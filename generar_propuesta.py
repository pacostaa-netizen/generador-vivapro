# -*- coding: utf-8 -*-
"""Genera una PROPUESTA con cronograma personalizado (lista de filas % del precio).
Uso interno: build_propuesta(cfg). cfg incluye cronograma=[{concepto,fecha,pct,sub2?}], hipotecario={...}|None."""
import os, json, zipfile, shutil, subprocess, datetime
import cronograma as C

BASE=os.path.dirname(os.path.abspath(__file__)); TPL=os.path.join(BASE,"plantillas")
if not os.path.exists(os.path.join(TPL,"deptos.json")): TPL=BASE
MES=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
def flarga(d): return f"{d.day} de {MES[d.month-1]} de {d.year}"

def _findsoffice():
    import shutil as sh
    for c in ("libreoffice","soffice","soffice.exe"):
        if sh.which(c): return sh.which(c)
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p): return p
    raise RuntimeError("No se encontró LibreOffice.")

def build_propuesta(cfg):
    cat=json.load(open(os.path.join(TPL,"deptos.json"),encoding="utf-8"))
    dep=cat[str(cfg["codigo_depto"])]; P=float(cfg["precio_soles"])
    rows=cfg["cronograma"]; hip=cfg.get("hipotecario")
    directo=sum(C.monto_de(r,float(cfg["precio_soles"]))/float(cfg["precio_soles"])*100 for r in rows)
    hip_pct=100-directo
    d=datetime.date.fromisoformat(cfg["fecha"]); num=dep["codigo"]; piso=dep["piso"]; tip=dep["tipologia"]; area=dep["area_m2"]
    nd={"un (01) dormitorio":1,"dos (02) dormitorios":2,"tres (03) dormitorios":3}.get(dep["dormitorios_txt"],3)
    sexo=cfg.get("sexo","F").upper(); ape=cfg.get("apellido",""); ec=cfg.get("estado_civil","")
    trato="Señora:" if sexo=="F" else "Señor:"; estim=(f"Estimada Sra. {ape}" if sexo=="F" else f"Estimado Sr. {ape}")
    # textos dinámicos
    if hip:
        paren=f"{directo:.0f}% aporte directo + {hip_pct:.0f}% crédito hipotecario contra entrega"
        introtxt=f"{directo:.0f}% del precio cancelado directamente a LA HAUS en aportes fraccionados durante la construcción, y el {hip_pct:.0f}% restante mediante crédito hipotecario contra entrega del inmueble."
        secbody=f"El {directo:.0f}% del precio se cancela directamente a LA HAUS en aportes fraccionados durante la construcción, y el {hip_pct:.0f}% restante se desembolsa al momento de la entrega mediante crédito hipotecario que {'la clienta' if sexo=='F' else 'el cliente'} tramitará con el banco de su elección:"
        head="PLAN PERSONALIZADO (CON CRÉDITO HIPOTECARIO)"
    else:
        narm=sum(1 for r in rows if "Armada" in r["concepto"])
        paren=f"financiamiento directo con LA HAUS en {narm} armadas"
        introtxt=f"el precio se cancela íntegramente a LA HAUS en aportes fraccionados durante la construcción ({directo:.0f}% directo), sin intervención bancaria."
        secbody="El precio de venta se cancela directamente a LA HAUS, sin intervención bancaria, conforme al siguiente cronograma:"
        head="PLAN PERSONALIZADO (FINANCIAMIENTO DIRECTO)"
    table=C.build_table(rows,P,hip)
    out=cfg["carpeta_salida"]; os.makedirs(out,exist_ok=True)
    name=f"Propuesta_VIVA_PRO_Depa{num}_{ape.replace(' ','')}_Personalizada.docx"
    outdocx=os.path.join(out,name)
    shutil.copyfile(os.path.join(TPL,"TPL_Propuesta_OpcionB.docx"),outdocx)
    tmp=outdocx+".tmp"
    import re
    with zipfile.ZipFile(outdocx) as zin, zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data=zin.read(it.filename)
            if it.filename=="word/document.xml":
                xml=data.decode("utf-8")
                tbls=[(m.start(),m.end()) for m in re.finditer(r'<w:tbl>.*?</w:tbl>',xml,re.S)]
                s,e=tbls[1]; xml=xml[:s]+table+xml[e:]
                R={"YADIRA LISSET CABALLERO NOEL":cfg["nombre"].upper(),
                   "DNI: 44578531 — Estado civil: Soltera":f"DNI: {cfg['dni']} — Estado civil: {ec}",
                   "Señora:":trato,"Estimada Sra. Caballero":estim,
                   "Lima, 11 de junio de 2026":f"Lima, {flarga(d)}",
                   "PROP-VP-203-2026-01B":f"PROP-VP-{num}-{d.year}-01P",
                   "Departamento N.° 203, Piso 2.":f"Departamento N.° {num}, {piso}.",
                   "3 dormitorios (3D-A) — 76 m² de área techada.":f"{nd} dormitorio{'s' if nd!=1 else ''} ({tip}) — {area} m² de área techada.",
                   "Departamento 203":f"Departamento {num}","320,000.00":f"{P:,.2f}",
                   "(50% directo + 50% crédito hipotecario contra entrega)":f"({paren})",
                   ", bajo un esquema mixto: 50% del precio cancelado directamente a LA HAUS en aportes fraccionados durante la construcción, y el 50% restante mediante crédito hipotecario contra entrega del inmueble.":f", bajo un esquema personalizado: {introtxt}",
                   "El 50% del precio se cancela directamente a LA HAUS en aportes fraccionados durante la construcción, y el 50% restante se desembolsa al momento de la entrega mediante crédito hipotecario que la clienta tramitará con el banco de su elección:":secbody,
                   "OPCIÓN B (MIXTA CON HIPOTECARIO)":head,"— Opción B —":"— Plan de pago —","Opción B":"propuesta de pago"}
                if sexo=="M": R["la clienta"]="el cliente"
                for a,b in sorted(R.items(),key=lambda kv:-len(kv[0])): xml=xml.replace(a,b)
                data=xml.encode("utf-8")
            zout.writestr(it,data)
    os.replace(tmp,outdocx)
    subprocess.run([_findsoffice(),"--headless","--convert-to","pdf","--outdir",out,outdocx],
        check=True,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,env={**os.environ,"HOME":os.path.expanduser("~") or "/tmp"})
    return outdocx
