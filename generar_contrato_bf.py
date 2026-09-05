# -*- coding: utf-8 -*-
"""Genera el CONTRATO DE COMPRAVENTA DE BIEN FUTURO (estándar, comprador directo).
Basado en el modelo vAA (203). Rellena comprador/unidad/fechas/precio, genera la
cláusula 4.2 desde el cronograma e inserta la cláusula de moneda recomendada.
El caso con PODER/representación se maneja de forma particular (no aquí).
Uso: build_contrato(cfg) — cfg como el de la propuesta + cronograma/hipotecario.
"""
import os, json, zipfile, shutil, subprocess, datetime, re
from num2words import num2words
import cronograma as C

BASE=os.path.dirname(os.path.abspath(__file__)); TPL=os.path.join(BASE,"plantillas")
if not os.path.exists(os.path.join(TPL,"deptos.json")): TPL=BASE
MES=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
ORD=["","primer","segundo","tercer","cuarto","quinto","sexto","séptimo","octavo","noveno"]
ORDF=["","primero","segundo","tercero","cuarto","quinto","sexto","séptimo","octavo","noveno"]

def flarga(d): return f"{d.day:02d} de {MES[d.month-1]} de {d.year}"
def money_words(x, moneda):
    w=num2words(int(round(x)), lang="es").upper()
    return w + (" CON 00/100 SOLES" if moneda=="S" else " CON 00/100 DÓLARES AMERICANOS")
def money_num(x, moneda):
    return ("S/ " if moneda=="S" else "US$ ")+f"{round(x):,.2f}"
def _findsoffice():
    import shutil as sh
    for c in ("libreoffice","soffice","soffice.exe"):
        if sh.which(c): return sh.which(c)
    raise RuntimeError("No se encontró LibreOffice.")
def _piso_txt(piso_str):
    m=re.search(r"(\d+)", piso_str or ""); n=int(m.group(1)) if m else 0
    pal=ORDF[n] if 0<n<len(ORDF) else f"{n}°"
    return f"{pal} ({n}°)"
def _esc(s): return (s or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def _para(text, *, bold_lead=None, indent=False, size="22"):
    ind='<w:ind w:left="360" w:hanging="360"/>' if indent else ''
    pPr=f'<w:pPr><w:spacing w:before="60" w:after="60"/>{ind}<w:jc w:val="both"/></w:pPr>'
    runs=''
    if bold_lead:
        runs+=f'<w:r><w:rPr><w:b/><w:sz w:val="{size}"/></w:rPr><w:t xml:space="preserve">{_esc(bold_lead)}</w:t></w:r>'
    runs+=f'<w:r><w:rPr><w:sz w:val="{size}"/></w:rPr><w:t xml:space="preserve">{_esc(text)}</w:t></w:r>'
    return f'<w:p>{pPr}{runs}</w:p>'

def _build_42(rows, hip, P, moneda_mode, tc, fecha_firma):
    """Genera los párrafos de la cláusula 4.2 desde el cronograma.
    moneda_mode: 'PEN' (dual: hasta firma soles, después dólares) | 'PEN_FIJO' (todo soles) | 'USD' (todo dólares)."""
    dual = (moneda_mode=="PEN")
    def cur_for(concepto):
        if moneda_mode=="USD": return "D"
        if moneda_mode=="PEN_FIJO": return "S"
        # dual: soles hasta la firma; dólares después
        return "S" if C._es_fila_soles(concepto) else "D"
    def conv(soles, cur):
        return soles/tc if cur=="D" else soles
    out=[ _para("La forma de pago se efectuará de la siguiente manera:", bold_lead="4.2 ") ]
    letters="abcdefghij"; i=0; acum=0.0
    for r in rows:
        c=r["concepto"]; mt=C.monto_de(r,P); acum+=mt; cur=cur_for(c)
        val=conv(mt,cur); pctv=mt/P*100
        L=letters[i]; i+=1
        low=c.lower()
        if "separ" in low:
            txt=(f"SEPARACIÓN (ya abonada): la suma de {money_num(val,cur)} ({money_words(val,cur)}), "
                 f"abonada por EL COMPRADOR con anterioridad a la firma como reserva de EL INMUEBLE, la cual se "
                 f"imputa al precio total y forma parte del primer aporte correspondiente a la firma del presente contrato.")
        elif "firma" in low:
            txt=(f"APORTE A LA FIRMA: a la suscripción del presente contrato, EL COMPRADOR cancelará la suma de "
                 f"{money_num(val,cur)} ({money_words(val,cur)}). Esta suma se imputa al precio como ARRAS "
                 f"CONFIRMATORIAS conforme a los artículos 1477 y 1478 del Código Civil, y se cancela mediante depósito "
                 f"o transferencia bancaria a la cuenta de LA HAUS CONSTRUCTORA S.A.C.")
        elif "hipotec" in low or "saldo con" in low or (hip and r is rows[-1] and False):
            txt=(f"SALDO FINAL ({pctv:.0f}%): la suma de {money_num(val,cur)} ({money_words(val,cur)}), que EL "
                 f"COMPRADOR cancelará a la entrega de EL INMUEBLE mediante crédito hipotecario, crédito personal o "
                 f"cualquier otro medio de pago bancarizado conforme al TUO de la Ley N° 28194, simultáneamente con la "
                 f"suscripción de la Escritura Pública de Compraventa.")
        elif "saldo final" in low or ("saldo" in low and i==len(rows)):
            txt=(f"SALDO FINAL ({pctv:.0f}%): la suma de {money_num(val,cur)} ({money_words(val,cur)}), que EL "
                 f"COMPRADOR cancelará a la entrega de EL INMUEBLE, simultáneamente con la suscripción de la Escritura "
                 f"Pública de Compraventa.")
        else:
            fec=r.get("fecha","").strip().rstrip(".")
            txt=(f"{c.upper()} ({pctv:.0f}%): la suma de {money_num(val,cur)} ({money_words(val,cur)})"
                 + (f", que EL COMPRADOR abonará en {fec}" if fec else "")
                 + " en la cuenta señalada.")
        out.append(_para(txt, bold_lead=f"({L}) ", indent=True))
    # cláusula de saldo hipotecario aparte (si viene como hip separado)
    if hip:
        hm=round(P)-round(acum); cur="D" if moneda_mode!="PEN_FIJO" else "S"
        val=conv(hm,cur); L=letters[i]; i+=1
        txt=(f"SALDO CON CRÉDITO HIPOTECARIO ({hm/P*100:.0f}%): la suma de {money_num(val,cur)} "
             f"({money_words(val,cur)}), que EL COMPRADOR cancelará a la entrega de EL INMUEBLE mediante crédito "
             f"hipotecario, crédito personal o cualquier otro medio de pago bancarizado conforme al TUO de la Ley "
             f"N° 28194, simultáneamente con la suscripción de la Escritura Pública de Compraventa.")
        out.append(_para(txt, bold_lead=f"({L}) ", indent=True))
    return "".join(out)

def _moneda_clause(mode, tc):
    if mode=="PEN":
        return _para(f"Los pagos hasta la firma en notaría se cancelan en soles a un tipo de cambio fijo de "
                     f"S/ {tc:.2f} por dólar; los pagos posteriores (armadas y saldo final) se expresan en dólares y "
                     f"se cancelan en soles al tipo de cambio venta publicado por la SBS el día de pago, el cual en "
                     f"ningún caso será menor a S/ {tc:.2f} por dólar.", bold_lead="4.5 MONEDA DE PAGO. ")
    if mode=="USD":
        return _para("Todos los pagos del presente contrato se pactan y cancelan en dólares americanos (US$).",
                     bold_lead="4.5 MONEDA DE PAGO. ")
    return _para(f"Todos los pagos del presente contrato se pactan y cancelan en soles.", bold_lead="4.5 MONEDA DE PAGO. ")

def build_contrato(cfg):
    cat=json.load(open(os.path.join(TPL,"deptos.json"),encoding="utf-8"))
    dep=cat[str(cfg["codigo_depto"])]; coch=cfg.get("cochera")
    P=float(cfg["precio_soles"])+(float(coch["precio"]) if coch else 0)
    rows=cfg["cronograma"]; hip=cfg.get("hipotecario")
    tc=float(cfg.get("tc") or 3.5); moneda=cfg.get("moneda","PEN")
    saldo_usd=bool(cfg.get("saldo_usd",False))
    mode = "USD" if moneda=="USD" else ("PEN" if saldo_usd else "PEN_FIJO")
    # datos comprador
    sexo=cfg.get("sexo","F").upper()
    nom=cfg["nombre"].strip().upper()
    ec=(cfg.get("estado_civil") or "").lower()
    trato="la señora" if sexo=="F" else "el señor"
    dni=cfg["dni"]; dom=cfg.get("domicilio") or "[●]"; corr=cfg.get("correo") or "[●]"; tel=cfg.get("telefono") or "[●]"
    # unidad
    num=dep["codigo"]; ui=cfg.get("unidad_n") or str(dep.get("ui","")); piso=_piso_txt(dep["piso"])
    area=dep["area_m2"]; area_w=num2words(int(area),lang="es").upper()+" METROS CUADRADOS"
    tip=dep["tipologia"]; ndd={"un (01) dormitorio":("un (1)","uno"),"dos (02) dormitorios":("dos (2)","dos"),"tres (03) dormitorios":("tres (3)","tres")}.get(dep["dormitorios_txt"],("tres (3)","tres"))
    dorm=f"{ndd[0]} dormitorio" + ("s" if ndd[1]!="uno" else "")
    # fechas
    d_firma=datetime.date.fromisoformat(cfg["fecha"])
    entrega=cfg.get("entrega_str","30 de noviembre de 2027")
    # precio 4.1 en la moneda base del contrato
    if mode=="USD":
        pmon=money_num(P/tc,"D"); pwords=money_words(P/tc,"D")
    else:
        pmon=money_num(P,"S"); pwords=money_words(P,"S")

    out=cfg["carpeta_salida"]; os.makedirs(out,exist_ok=True)
    ape=(cfg.get("apellido") or "").replace(" ","")
    name=f"Contrato_BienFuturo_Depa{num}_{(cfg['nombre'].split()[0] if cfg['nombre'].split() else 'Cliente')}_{ape}.docx"
    outdocx=os.path.join(out,name)
    shutil.copyfile(os.path.join(TPL,"TPL_Contrato_BienFuturo.docx"),outdocx)
    tmp=outdocx+".tmp"
    bloque42=_build_42(rows,hip,P,mode,tc,d_firma)

    R={
      # comprador (intro)
      "la señora YADIRA LISSET CABALLERO NOEL, de nacionalidad peruana, identificada con DNI N° 44578531, de estado civil soltera, con domicilio en Prolongación Arequipa N° 250, distrito de Barranca, provincia de Barranca, departamento de Lima, correo electrónico residencialvivapro203@gmail.com, teléfono +51 984 230 507":
        f"{trato} {nom}, de nacionalidad peruana, identificad{'a' if sexo=='F' else 'o'} con DNI N° {dni}, de estado civil {ec}, con domicilio en {dom}, correo electrónico {corr}, teléfono {tel}",
      # fechas entrega
      "30 de enero de 2028": entrega,
      # unidad 3.1
      "Departamento N° 203, que en el Reglamento Interno y en la partida registral independiente se identificará como Unidad Inmobiliaria N° 7, ubicado en el segundo (2°) piso de EL EDIFICIO, con un área techada aproximada de 76 m² (SETENTA Y SEIS METROS CUADRADOS), tipología 3D-A, de tres (3) dormitorios":
        f"Departamento N° {num}, que en el Reglamento Interno y en la partida registral independiente se identificará como Unidad Inmobiliaria N° {ui}, ubicado en el {piso} piso de EL EDIFICIO, con un área techada aproximada de {area} m² ({area_w}), tipología {tip}, de {dorm}",
      # precio 4.1
      "S/ 320,000.00 (TRESCIENTOS VEINTE MIL CON 00/100 SOLES)": f"{pmon} ({pwords})",
      # firma
      "YADIRA LISSET CABALLERO NOEL": nom, "DNI 44578531": f"DNI {dni}",
      # cierre
      "a los 01 días del mes de septiembre del año 2026": f"a los {d_firma.day:02d} días del mes de {MES[d_firma.month-1]} del año {d_firma.year}",
    }
    if coch:
        R["El presente contrato no comprende estacionamiento ni depósito."]=(
            f"El presente contrato comprende además el Estacionamiento N° {coch['est']} (16 m², reja corrediza no elevadiza, "
            f"con partida registral independiente).")

    with zipfile.ZipFile(outdocx) as zin, zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
        seen=set()
        for it in zin.infolist():
            if it.filename in seen: continue
            seen.add(it.filename)
            data=zin.read(it.filename)
            if it.filename=="word/document.xml":
                xml=data.decode("utf-8")
                # reemplazar el bloque 4.2 (desde su párrafo hasta antes de 4.3)
                i=xml.find("La forma de pago se efectuará de la siguiente manera")
                s=max(xml.rfind("<w:p>",0,i),xml.rfind("<w:p ",0,i))
                j=xml.find("Si EL COMPRADOR no logra el financiamiento del SALDO FINAL")
                e=max(xml.rfind("<w:p>",0,j),xml.rfind("<w:p ",0,j))
                if s!=-1 and e!=-1 and e>s:
                    xml=xml[:s]+bloque42+xml[e:]
                # insertar cláusula 4.5 MONEDA DE PAGO antes de la Cláusula Quinta
                q=xml.find("FINANCIAMIENTO DEL SALDO FINAL Y BONOS")
                if q!=-1:
                    qs=max(xml.rfind("<w:p>",0,q),xml.rfind("<w:p ",0,q))
                    # subir hasta el inicio del párrafo del título (que incluye 'CLÁUSULA QUINTA')
                    xml=xml[:qs]+_moneda_clause(mode,tc)+xml[qs:]
                for a,b in sorted(R.items(),key=lambda kv:-len(kv[0])):
                    xml=xml.replace(a,b)
                data=xml.encode("utf-8")
            zout.writestr(it,data)
    os.replace(tmp,outdocx)
    subprocess.run([_findsoffice(),"--headless","--convert-to","pdf","--outdir",out,outdocx],
        check=True,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,env={**os.environ,"HOME":os.path.expanduser("~") or "/tmp"})
    return outdocx
