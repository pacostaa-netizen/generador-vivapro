# -*- coding: utf-8 -*-
"""Tabla de cronograma dinámica para propuestas VIVA PRO.
rows: lista de {concepto, fecha, pct?, monto?, sub2?}. Si 'monto' está, manda; si no, monto=precio*pct/100.
hipotecario: {concepto, fecha, sub2?} | None  -> su monto absorbe el resto para que el total cuadre exacto."""
W={"n":"700","con":"3300","fec":"2372","mon":"1500","pct":"1200"}
def _cell(w,text,*,fill=None,bold=False,white=False,jc=None,sz="20",sub2=None):
    shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ''
    j=f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ''
    rb=('<w:b/>' if bold else '')+('<w:color w:val="FFFFFF"/>' if white else '')+f'<w:sz w:val="{sz}"/>'
    p=f'<w:p>{j}<w:r><w:rPr>{rb}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    if sub2: p+=f'<w:p><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="595959"/></w:rPr><w:t xml:space="preserve">{sub2}</w:t></w:r></w:p>'
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}</w:tcPr>{p}</w:tc>'
def _row(c): return '<w:tr>'+''.join(c)+'</w:tr>'
def m2(x): return f"{round(x):,.2f}"
def pc(x): return f"{x:.2f} %"
def monto_de(r,P): return float(r["monto"]) if r.get("monto") not in (None,"") else P*float(r["pct"])/100

def build_table(rows,precio,hipotecario=None):
    P=float(precio); out=[]
    hdr=[_cell(W["n"],"#",fill="EA5A29",bold=True,white=True,jc="center",sz="18"),
         _cell(W["con"],"Concepto",fill="EA5A29",bold=True,white=True,sz="18"),
         _cell(W["fec"],"Fecha de pago",fill="EA5A29",bold=True,white=True,sz="18"),
         _cell(W["mon"],"Monto (S/)",fill="EA5A29",bold=True,white=True,jc="right",sz="18"),
         _cell(W["pct"],"% precio",fill="EA5A29",bold=True,white=True,jc="right",sz="18")]
    out.append('<w:tr><w:trPr><w:tblHeader/></w:trPr>'+''.join(hdr)+'</w:tr>')
    directo=0.0; i=0
    for r in rows:
        i+=1; mt=monto_de(r,P); directo+=mt
        out.append(_row([_cell(W["n"],str(i),jc="center"),
            _cell(W["con"],r["concepto"],bold=True,sub2=r.get("sub2")),
            _cell(W["fec"],r.get("fecha","")),
            _cell(W["mon"],m2(mt),bold=True,jc="right"),
            _cell(W["pct"],pc(mt/P*100),jc="right")]))
    if hipotecario:
        out.append(_row([_cell(W["n"],"=",fill="FFF4EE",bold=True,jc="center"),
            _cell(W["con"],f"Subtotal aporte directo ({directo/P*100:.0f}%)",fill="FFF4EE",bold=True),
            _cell(W["fec"],"Durante la construcción.",fill="FFF4EE"),
            _cell(W["mon"],m2(directo),fill="FFF4EE",bold=True,jc="right"),
            _cell(W["pct"],pc(directo/P*100),fill="FFF4EE",bold=True,jc="right")]))
        i+=1; hm=P-directo
        out.append(_row([_cell(W["n"],str(i),jc="center"),
            _cell(W["con"],hipotecario["concepto"],bold=True,sub2=hipotecario.get("sub2")),
            _cell(W["fec"],hipotecario.get("fecha","")),
            _cell(W["mon"],m2(hm),bold=True,jc="right"),
            _cell(W["pct"],pc(hm/P*100),jc="right")]))
    out.append(_row([_cell(W["n"],"$",fill="000000",bold=True,white=True,jc="center"),
        _cell(W["con"],"PRECIO TOTAL DEL INMUEBLE",fill="000000",bold=True,white=True),
        _cell(W["fec"],"Total.",fill="000000",white=True),
        _cell(W["mon"],m2(P),fill="000000",bold=True,white=True,jc="right"),
        _cell(W["pct"],"100.00 %",fill="000000",bold=True,white=True,jc="right")]))
    g='<w:tblPr><w:tblW w:w="9072" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="3300"/><w:gridCol w:w="2372"/><w:gridCol w:w="1500"/><w:gridCol w:w="1200"/></w:tblGrid>'
    return '<w:tbl>'+g+''.join(out)+'</w:tbl>'

def plantilla(nombre, precio, sep=3500.0):
    """Devuelve (rows, hipotecario) por defecto según plantilla. Textos genéricos editables."""
    P=float(precio)
    seprow={"concepto":"Separación","fecha":"A la suscripción de la presente.","monto":sep,"sub2":"Forma parte de la cuota inicial."}
    if nombre=="40% directo + 4 armadas":
        rows=[seprow,
          {"concepto":"Saldo a la firma","fecha":"A la firma del contrato (notaría).","monto":P*0.20-sep},
          {"concepto":"Saldo al término del casco","fecha":"Abril de 2027.","monto":P*0.20},
          {"concepto":"1.ª Armada","fecha":"Junio de 2027.","monto":P*0.15},
          {"concepto":"2.ª Armada","fecha":"Agosto de 2027.","monto":P*0.15},
          {"concepto":"3.ª Armada","fecha":"Octubre de 2027.","monto":P*0.15},
          {"concepto":"4.ª Armada (contra entrega)","fecha":"Diciembre de 2027.","monto":P*0.15}]
        return rows,None
    pcts={"30% directo + 70% hipotecario":(0.30,0.70),"20% directo + 80% hipotecario":(0.20,0.80),
          "50% directo + 50% hipotecario":(0.50,0.50)}
    if nombre in pcts:
        di,hi=pcts[nombre]; firma=P*0.15 if di>=0.30 else P*0.10
        rows=[seprow,{"concepto":"Saldo a la firma","fecha":"A la firma (notaría / inicio de obra).","monto":firma-sep}]
        rest=P*di-firma; narm=3
        fechas=["Diciembre de 2026.","Junio de 2027.","Diciembre de 2027."]
        for k in range(narm):
            rows.append({"concepto":f"{k+1}.ª Armada","fecha":fechas[k],"monto":rest/narm})
        return rows,{"concepto":"Saldo con crédito hipotecario","fecha":"Contra entrega (diciembre de 2027).","sub2":"Tasa, plazo y cuota los define el banco."}
    # personalizada (en blanco)
    return [seprow,{"concepto":"Saldo a la firma","fecha":"A la firma.","monto":P*0.20-sep}],None
