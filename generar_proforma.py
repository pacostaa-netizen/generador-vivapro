#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generador de PROFORMA VIVA PRO (formato cotización 1 hoja + plano). ReportLab.
Uso: python3 generar_proforma.py proforma.json"""
import sys, os, json
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Spacer,
                                Paragraph, Image as RLImage, PageBreak, Frame, PageTemplate)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from PIL import Image as PILImage

BASE=os.path.dirname(os.path.abspath(__file__)); TPL=BASE
DEFAULT_PLANOS=BASE
GRAY=colors.HexColor("#EAEAEA"); BORDER=colors.HexColor("#C9C9C9"); DARK=colors.HexColor("#2B2B2B")
PW,PH=A4; MX=42

def img_dims(p):
    im=PILImage.open(p); return im.size

def sec_table(title, rows, w1, w2):
    data=[[title,""]]+rows
    t=Table(data,colWidths=[w1,w2])
    st=[("SPAN",(0,0),(1,0)),("BACKGROUND",(0,0),(1,0),GRAY),
        ("FONTNAME",(0,0),(1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.black),
        ("GRID",(0,0),(-1,-1),0.6,BORDER),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("FONTNAME",(0,1),(0,-1),"Helvetica")]
    t.setStyle(TableStyle(st)); return t

def header_footer(canvas, doc, first):
    canvas.saveState()
    # footer dark bar
    canvas.setFillColor(DARK); canvas.rect(0,0,PW,26,fill=1,stroke=0)
    canvas.setFillColor(colors.white); canvas.setFont("Helvetica-Bold",7.5)
    canvas.drawString(MX,9,"JR PERSEVERANCIA 7907, URB PRO - LOS OLIVOS")
    canvas.drawRightString(PW-MX,9,"@lahausperu")
    if first:
        lh=os.path.join(TPL,"logo_lahaus.png"); vp=os.path.join(TPL,"logo_vivapro.png")
        w,h=img_dims(lh); lw=120; lh2=lw*h/w
        canvas.drawImage(lh,MX,PH-50-lh2/2,width=lw,height=lh2,mask='auto')
        w,h=img_dims(vp); vw=150; vh=vw*h/w
        canvas.drawImage(vp,PW-MX-vw,PH-48-vh/2,width=vw,height=vh,mask='auto')
        canvas.setFillColor(colors.black); canvas.setFont("Helvetica-Bold",21)
        canvas.drawCentredString(PW/2,PH-52,"PROFORMA")
    canvas.restoreState()

def build(cfg):
    cat=json.load(open(os.path.join(TPL,"deptos.json"),encoding="utf-8"))
    cod=str(cfg["codigo_depto"]); dep=cat[cod]
    P=int(cfg.get("precio_soles") or dep["precio_lista_soles"])
    nd={"un (01) dormitorio":1,"dos (02) dormitorios":2,"tres (03) dormitorios":3}.get(dep["dormitorios_txt"],3)
    nb=1 if "un (01)" in dep["banos_txt"] else 2
    deptxt=f"{nd} Dormitorio{'s' if nd!=1 else ''} / {nb} Baño{'s' if nb!=1 else ''}"
    planos=cfg.get("planos_dir",DEFAULT_PLANOS)
    out=cfg["carpeta_salida"]; os.makedirs(out,exist_ok=True)
    ini_pct=cfg.get("inicial_pct",0.40); inicial=round(P*ini_pct)
    # cronograma: usar el dado o el estándar Crédito Directo 40%+4x15%
    crono=cfg.get("cronograma")
    if not crono:
        a=round(P*0.15)
        crono=[[f"Inicial {int(ini_pct*100)}%",f"S/ {inicial:,}"],
               ["15% – Julio 2026",f"S/ {a:,}"],["15% – Diciembre 2026",f"S/ {a:,}"],
               ["15% – Julio 2027",f"S/ {a:,}"],["15% – Contra entrega 2027",f"S/ {a:,}"]]
    forma=cfg.get("forma_pago","Crédito Directo")
    W=PW-2*MX
    elems=[Spacer(1,52)]
    elems.append(sec_table("DATOS DEL CLIENTE",[["Nombre",cfg["nombre"]]],W*0.42,W*0.58))
    elems.append(Spacer(1,10))
    cot=[["Departamento",deptxt],["Área",f"{dep['area_m2']} m²"],
         ["Nivel / Código",f"{dep['piso']} - {cod}"],["Precio del departamento",f"S/ {P:,}"],
         ["Forma de pago",forma],[f"Inicial {int(ini_pct*100)}%",f"S/ {inicial:,}"]]
    elems.append(sec_table("COTIZACIÓN",cot,W*0.42,W*0.58))
    elems.append(Spacer(1,10))
    elems.append(sec_table(f"CRONOGRAMA DE PAGOS – {forma.upper()}",crono,W*0.62,W*0.38))
    elems.append(Spacer(1,12))
    nota=ParagraphStyle("n",fontName="Helvetica",fontSize=8,leading=12)
    notab=ParagraphStyle("nb",fontName="Helvetica-Bold",fontSize=8,leading=12)
    ncell=[Paragraph("Nota:",notab),
           Paragraph("• Los precios están sujetos a cambios.",nota),
           Paragraph("• Validez de la proforma: 7 días.",nota),
           Paragraph("• La separación forma parte de la cuota inicial.",nota),
           Paragraph("• El área del departamento puede variar hasta un 5%.",nota)]
    nt=Table([[ncell]],colWidths=[W])
    nt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),GRAY),("BOX",(0,0),(-1,-1),0.6,BORDER),
        ("LEFTPADDING",(0,0),(-1,-1),8),("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]))
    elems.append(nt)
    # pagina 2: plano
    elems.append(PageBreak())
    cen=ParagraphStyle("c",fontName="Helvetica-Bold",fontSize=11,alignment=TA_CENTER,spaceAfter=14)
    elems.append(Spacer(1,8)); elems.append(Paragraph("Vista en planta del departamento",cen))
    if dep.get("plano"):
        pf=os.path.join(planos,dep["plano"])
        if os.path.exists(pf):
            w,h=img_dims(pf); maxw=W*0.62; maxh=PH-200
            sc=min(maxw/w,maxh/h); elems.append(RLImage(pf,width=w*sc,height=h*sc))
    name=f"Proforma_{cfg.get('apellido','Cliente').replace(' ','')}_{cod}_VIVAPRO.pdf"
    path=os.path.join(out,name)
    doc=SimpleDocTemplate(path,pagesize=A4,leftMargin=MX,rightMargin=MX,topMargin=70,bottomMargin=36)
    frame=Frame(MX,36,PW-2*MX,PH-70-36,id='n')
    doc.addPageTemplates([PageTemplate(id='all',frames=[frame],
        onPage=lambda c,d:header_footer(c,d,d.page==1))])
    # centrar imagen en pagina 2
    for e in elems:
        if isinstance(e,RLImage): e.hAlign='CENTER'
    doc.build(elems)
    print("OK",name); return path

if __name__=="__main__":
    build(json.load(open(sys.argv[1],encoding="utf-8")))
