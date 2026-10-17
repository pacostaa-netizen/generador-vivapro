#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generador de PROFORMA VIVA PRO — formato ligero 2 hojas (datos + plano).
Deptos con terraza: plano del depto + azotea a la MISMA ESCALA. ReportLab.
Uso: python3 generar_proforma.py proforma.json"""
import sys, os, json
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Spacer,
                                Paragraph, Image as RLImage, PageBreak, Frame, PageTemplate)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from PIL import Image as PILImage

BASE=os.path.dirname(os.path.abspath(__file__)); TPL=os.path.join(BASE,"plantillas")
if not os.path.exists(os.path.join(TPL,"deptos.json")): TPL=BASE
DEFAULT_PLANOS=os.path.normpath(os.path.join(BASE,"..","..","05_MARKETING_Y_MARCA","04_Renders","Plantas"))
GRAY=colors.HexColor("#EAEAEA"); BORDER=colors.HexColor("#C9C9C9")
DARK=colors.HexColor("#1B1B1B"); REQBG=colors.HexColor("#FBF7F2"); REQBD=colors.HexColor("#E8C9B0")
PW,PH=A4; MX=42; FOOT_H=30; TOPM=64; BOTM=FOOT_H

def img_dims(p):
    im=PILImage.open(p); return im.size

def sec_table(title, rows, w1, w2):
    data=[[title,""]]+rows
    t=Table(data,colWidths=[w1,w2])
    t.setStyle(TableStyle([("SPAN",(0,0),(1,0)),("BACKGROUND",(0,0),(1,0),GRAY),
        ("FONTNAME",(0,0),(1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.black),("GRID",(0,0),(-1,-1),0.6,BORDER),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("FONTNAME",(0,1),(0,-1),"Helvetica")]))
    return t

def framed(img):
    """RLImage dentro de un marco negro."""
    t=Table([[img]])
    t.setStyle(TableStyle([("BOX",(0,0),(-1,-1),1.3,colors.black),("BACKGROUND",(0,0),(-1,-1),colors.white),
        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    return t

def header_footer(canvas, doc):
    canvas.saveState()
    # pie: barra negra con texto blanco (TODAS las hojas)
    canvas.setFillColor(DARK); canvas.rect(0,0,PW,FOOT_H,fill=1,stroke=0)
    canvas.setFillColor(colors.white); canvas.setFont("Helvetica-Bold",7.6)
    canvas.drawString(MX,10,"JR PERSEVERANCIA 7907, URB PRO - LOS OLIVOS")
    canvas.drawRightString(PW-MX,10,"https://www.instagram.com/lahausperu/")
    # cabecera: logos + PROFORMA (TODAS las hojas)
    lh=os.path.join(TPL,"logo_lahaus.png"); vp=os.path.join(TPL,"logo_vivapro.png")
    if os.path.exists(lh):
        w,h=img_dims(lh); lw=120; lh2=lw*h/w
        canvas.drawImage(lh,MX,PH-46-lh2/2,width=lw,height=lh2,mask='auto')
    if os.path.exists(vp):
        w,h=img_dims(vp); vw=150; vh=vw*h/w
        canvas.drawImage(vp,PW-MX-vw,PH-44-vh/2,width=vw,height=vh,mask='auto')
    canvas.setFillColor(colors.black); canvas.setFont("Helvetica-Bold",21)
    canvas.drawCentredString(PW/2,PH-48,"PROFORMA")
    canvas.restoreState()

def build(cfg):
    cat=json.load(open(os.path.join(TPL,"deptos.json"),encoding="utf-8"))
    cod=str(cfg["codigo_depto"]); dep=cat[cod]
    P=int(cfg.get("precio_soles") or dep["precio_lista_soles"])
    nd={"un (01) dormitorio":1,"dos (02) dormitorios":2,"tres (03) dormitorios":3}.get(dep.get("dormitorios_txt",""),None)
    if nd is None:
        try: nd=int(''.join(ch for ch in str(dep.get("tipologia","")) if ch.isdigit()) or 0)
        except: nd=0
    nb=1 if "un (01)" in dep.get("banos_txt","") else 2
    deptxt=f"{nd} Dormitorio{'s' if nd!=1 else ''} / {nb} Baño{'s' if nb!=1 else ''}"
    if dep.get("terraza_m2"): deptxt+=" + Azotea"
    planos=cfg.get("planos_dir",DEFAULT_PLANOS)
    out=cfg["carpeta_salida"]; os.makedirs(out,exist_ok=True)
    ini_pct=cfg.get("inicial_pct",0.40)
    _coch=cfg.get("cochera")
    _base_ini=P+int(_coch["precio"]) if (_coch and _coch.get("mode")=="sumada") else P
    inicial=round(_base_ini*ini_pct)
    crono=cfg.get("cronograma")
    if not crono:
        a=round(P*0.15)
        crono=[[f"Inicial {int(ini_pct*100)}%",f"S/ {inicial:,}"],
               ["15% – Julio 2026",f"S/ {a:,}"],["15% – Diciembre 2026",f"S/ {a:,}"],
               ["15% – Julio 2027",f"S/ {a:,}"],["15% – Contra entrega 2027",f"S/ {a:,}"]]
    forma=cfg.get("forma_pago","Crédito Directo")
    W=PW-2*MX
    elems=[Spacer(1,46)]
    _n=str(cfg.get("nombre","")).strip(); _a=str(cfg.get("apellido","")).strip()
    _nom=_n if (_a and _a.upper() in _n.upper()) else (_n+" "+_a).strip()
    elems.append(sec_table("DATOS DEL CLIENTE",[["Nombre",_nom]],W*0.42,W*0.58)); elems.append(Spacer(1,10))
    _terr=dep.get("terraza_m2")
    cot=[["Departamento",deptxt],["Área techada" if _terr else "Área",f"{dep['area_m2']} m²"]]
    if _terr: cot.append(["Área de terraza (no techada)",f"{_terr} m² (azotea, incluida en el precio)"])
    cot+=[["Nivel / Código",f"{dep['piso']} - {cod}"],["Precio del departamento",f"S/ {P:,}"],["Forma de pago",forma]]
    if _coch:
        cot.append(["Estacionamiento", f"N° {_coch['est']} · S/ {int(_coch['precio']):,} · 16 m² · partida independiente"])
        cot.append(["Precio total (depto + cochera)", f"S/ {P+int(_coch['precio']):,}"])
    cot.append([f"Inicial {int(ini_pct*100)}%",f"S/ {inicial:,}"])
    elems.append(sec_table("COTIZACIÓN",cot,W*0.42,W*0.58)); elems.append(Spacer(1,10))
    elems.append(sec_table(f"CRONOGRAMA DE PAGOS – {forma.upper()}",crono,W*0.62,W*0.38)); elems.append(Spacer(1,11))
    # Requisitos del cliente
    nb_=ParagraphStyle("nb",fontName="Helvetica-Bold",fontSize=8,leading=12)
    n_=ParagraphStyle("n",fontName="Helvetica",fontSize=8,leading=12)
    req=[Paragraph("Requisitos del cliente (para reservar y avanzar):",nb_),
         Paragraph("• Copia de DNI del titular (y del cónyuge, si aplica).",n_),
         Paragraph("• Correo y número de contacto.",n_),
         Paragraph("• Separación de S/ 3,500 para reservar la unidad (forma parte de la cuota inicial).",n_),
         Paragraph("• Si el pago es con crédito hipotecario: constancia de precalificación del banco.",n_)]
    rt=Table([[req]],colWidths=[W]); rt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),REQBG),("BOX",(0,0),(-1,-1),0.6,REQBD),
        ("LEFTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    elems.append(rt); elems.append(Spacer(1,9))
    # Nota
    nota=[Paragraph("Nota:",nb_),
          Paragraph("• Las imágenes, renders y el plano son <b>referenciales</b>.",n_),
          Paragraph("• Los precios están sujetos a cambios.",n_),
          Paragraph("• Validez de la proforma: 7 días.",n_),
          Paragraph("• La separación forma parte de la cuota inicial.",n_),
          Paragraph("• El área del departamento puede variar hasta un 5%.",n_)]
    nt=Table([[nota]],colWidths=[W]); nt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),GRAY),("BOX",(0,0),(-1,-1),0.6,BORDER),
        ("LEFTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    elems.append(nt)

    # ---------- HOJA(S) DE IMAGEN ----------
    elems.append(PageBreak())
    cen=ParagraphStyle("c",fontName="Helvetica-Bold",fontSize=12,alignment=TA_CENTER,spaceAfter=14)
    cap=ParagraphStyle("cap",fontName="Helvetica-Bold",fontSize=9,alignment=TA_CENTER)
    small=ParagraphStyle("sm",fontName="Helvetica",fontSize=8.5,alignment=TA_CENTER,textColor=colors.HexColor("#555555"))
    avail_h=PH-TOPM-BOTM-30  # alto util aprox bajo la cabecera
    plano_f=os.path.join(planos,dep["plano"]) if dep.get("plano") else None
    azo_f=os.path.join(planos,dep["azotea"]) if dep.get("azotea") else None
    tiene_azo=bool(azo_f and os.path.exists(azo_f) and plano_f and os.path.exists(plano_f))

    if tiene_azo:
        w1,h1=img_dims(plano_f); w2,h2=img_dims(azo_f)
        gap=20.0; maxW=W-gap; maxH=avail_h-70
        s=min(maxW/(w1+w2), maxH/max(h1,h2))
        im1=RLImage(plano_f,width=w1*s,height=h1*s); im2=RLImage(azo_f,width=w2*s,height=h2*s)
        duo=Table([[framed(im1),framed(im2)],[Paragraph("Departamento",cap),Paragraph("Azotea",cap)]],
                  colWidths=[w1*s+14,w2*s+14])
        duo.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,0),"TOP"),
            ("TOPPADDING",(0,1),(-1,1),7),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
        duo.hAlign='CENTER'
        block_h=h1*s+40
        elems.append(Spacer(1,max(10,(avail_h-block_h)/2-20)))
        elems.append(Paragraph("Vista en planta — departamento y azotea",cen))
        elems.append(duo); elems.append(Spacer(1,10))
        elems.append(Paragraph("Ambas vistas a la misma escala. Plano e imágenes de carácter <b>referencial</b>; la distribución final puede variar.",small))
    elif plano_f and os.path.exists(plano_f):
        w,h=img_dims(plano_f); maxw=W*0.82; maxh=avail_h-30
        sc=min(maxw/w,maxh/h); iw,ih=w*sc,h*sc
        elems.append(Spacer(1,max(8,(avail_h-ih-40)/2-14)))
        elems.append(Paragraph("Vista en planta del departamento",cen))
        fr=framed(RLImage(plano_f,width=iw,height=ih)); fr.hAlign='CENTER'; elems.append(fr)
        elems.append(Spacer(1,10))
        elems.append(Paragraph("Plano e imágenes de carácter <b>referencial</b>; la distribución final puede variar.",small))

    name=f"Proforma_{(_n.split()[0] if _n.split() else 'Cliente')}_{(_a or 'Cliente').replace(' ','')}_{cod}_VIVAPRO.pdf"
    path=os.path.join(out,name)
    doc=SimpleDocTemplate(path,pagesize=A4,leftMargin=MX,rightMargin=MX,topMargin=TOPM,bottomMargin=BOTM)
    frame=Frame(MX,BOTM,PW-2*MX,PH-TOPM-BOTM,id='n')
    doc.addPageTemplates([PageTemplate(id='all',frames=[frame],onPage=header_footer)])
    doc.build(elems)
    print("OK",name); return path

if __name__=="__main__":
    build(json.load(open(sys.argv[1],encoding="utf-8")))
