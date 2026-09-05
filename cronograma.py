# -*- coding: utf-8 -*-
"""Tabla de cronograma dinámica para propuestas VIVA PRO.
rows: lista de {concepto, fecha, pct?, monto?, sub2?}. Si 'monto' está, manda; si no, monto=precio*pct/100.
hipotecario: {concepto, fecha, sub2?} | None  -> su monto absorbe el resto para que el total cuadre exacto.

Modelo de moneda (tc + saldo_usd):
  - Si saldo_usd=True y tc>0, el TRAMO FINAL (el saldo hipotecario si existe, o la última armada) se
    expresa en DÓLARES al tipo de cambio del día, con referencia en soles. El resto queda en soles fijo.
  - Se agrega una nota cambiaria debajo de la tabla."""
W={"n":"700","con":"3300","fec":"2372","mon":"1500","pct":"1200"}
def _cell(w,text,*,fill=None,bold=False,white=False,jc=None,sz="20",sub2=None,sub2_color="595959"):
    shd=f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ''
    j=f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ''
    rb=('<w:b/>' if bold else '')+('<w:color w:val="FFFFFF"/>' if white else '')+f'<w:sz w:val="{sz}"/>'
    p=f'<w:p>{j}<w:r><w:rPr>{rb}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    if sub2:
        j2=f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ''
        p+=f'<w:p>{j2}<w:r><w:rPr><w:sz w:val="17"/><w:color w:val="{sub2_color}"/></w:rPr><w:t xml:space="preserve">{sub2}</w:t></w:r></w:p>'
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}</w:tcPr>{p}</w:tc>'
def _row(c): return '<w:tr><w:trPr><w:cantSplit/></w:trPr>'+''.join(c)+'</w:tr>'
def m2(x): return f"{round(x):,.2f}"
def pc(x): return f"{x:.2f} %"
def monto_de(r,P): return float(r["monto"]) if r.get("monto") not in (None,"") else P*float(r["pct"])/100

def _nota_cambiaria_xml(tc):
    t=f"{tc:.2f}"
    txt=(f"El aporte directo (inicial y armadas previas) está expresado y se paga en soles a un tipo de "
         f"cambio fijo de S/ {t} por dólar. El saldo final (diciembre 2027) se expresa en dólares (US$) y se "
         f"cancela en soles al tipo de cambio venta SBS del día de pago, el cual en ningún caso será menor a "
         f"S/ {t} por dólar. Los importes en soles mostrados son referenciales a S/ {t}.")
    return ('<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="FBF6E7"/>'
            '<w:pBdr><w:top w:val="single" w:sz="4" w:color="D8B24A"/><w:left w:val="single" w:sz="4" w:color="D8B24A"/>'
            '<w:bottom w:val="single" w:sz="4" w:color="D8B24A"/><w:right w:val="single" w:sz="4" w:color="D8B24A"/></w:pBdr>'
            '<w:spacing w:before="120" w:after="40"/></w:pPr>'
            '<w:r><w:rPr><w:b/><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve">Nota sobre moneda y tipo de cambio: </w:t></w:r>'
            f'<w:r><w:rPr><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve">{txt}</w:t></w:r></w:p>')

def _mon_dual(w, soles, tc, *, fill=None, white=False, ref=False):
    """Celda de monto: US$ arriba, S/ debajo (en ese orden)."""
    usd=soles/tc
    sub_color="EEEEEE" if white else "888888"
    pre="ref S/ " if ref else "S/ "
    return _cell(w, f"US$ {usd:,.2f}", fill=fill, bold=True, white=white, jc="right",
                 sub2=f"{pre}{soles:,.2f}", sub2_color=sub_color)

def _es_fila_soles(concepto):
    """Pagos 'hasta la firma en notaría' se muestran solo en soles."""
    c=(concepto or "").lower()
    return ("separ" in c) or ("firma" in c)

def build_table(rows,precio,hipotecario=None,tc=None,moneda='PEN'):
    P=float(precio); Pr=round(P); out=[]
    # no mostrar filas en 0 (p. ej. armadas vacías); la separación siempre se conserva
    rows=[r for r in rows if (round(monto_de(r,P))>0 or "separ" in (r.get("concepto","") or "").lower())]
    # PEN: cada monto se muestra en US$ y su equivalente en S/ (referencial al tc). USD: solo US$.
    dual = (moneda!='USD')
    TC = float(tc) if (tc and float(tc)>0) else 3.5
    # montos redondeados a soles; una fila absorbe el residual para que TODO cuadre exacto
    m_rows=[round(monto_de(r,P)) for r in rows]
    # --- fix "S/ 4": si hay hipotecario pero su monto es residual (directo ~100%), tratar como sin hipotecario ---
    if hipotecario:
        hm_check=Pr-sum(m_rows)
        if hm_check < max(100, 0.005*Pr):
            hipotecario=None
    if not hipotecario and m_rows:
        idx=next((i for i,r in enumerate(rows) if r.get("residual")),None)
        if idx is None:
            idx=next((i for i,r in enumerate(rows) if ("firma" in r["concepto"].lower() and "separ" not in r["concepto"].lower())),None)
        if idx is None: idx=len(m_rows)-1
        m_rows[idx]+=Pr-sum(m_rows)
    mon_label = "Monto (US$)" if moneda=='USD' else "Monto (US$ / S/)"
    hdr=[_cell(W["n"],"#",fill="EA5A29",bold=True,white=True,jc="center",sz="18"),
         _cell(W["con"],"Concepto",fill="EA5A29",bold=True,white=True,sz="18"),
         _cell(W["fec"],"Fecha de pago",fill="EA5A29",bold=True,white=True,sz="18"),
         _cell(W["mon"],mon_label,fill="EA5A29",bold=True,white=True,jc="right",sz="18"),
         _cell(W["pct"],"% precio",fill="EA5A29",bold=True,white=True,jc="right",sz="18")]
    out.append('<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>'+''.join(hdr)+'</w:tr>')
    directo=0.0; i=0
    for j,(r,mt) in enumerate(zip(rows,m_rows)):
        i+=1; directo+=mt
        mon_cell = _mon_dual(W["mon"], mt, TC) if dual else _cell(W["mon"],m2(mt),bold=True,jc="right")
        out.append(_row([_cell(W["n"],str(i),jc="center"),
            _cell(W["con"],r["concepto"],bold=True,sub2=r.get("sub2")),
            _cell(W["fec"],r.get("fecha","")),
            mon_cell,
            _cell(W["pct"],pc(mt/P*100),jc="right")]))
    if hipotecario:
        out.append(_row([_cell(W["n"],"=",fill="FFF4EE",bold=True,jc="center"),
            _cell(W["con"],f"Subtotal aporte directo ({directo/P*100:.0f}%)",fill="FFF4EE",bold=True),
            _cell(W["fec"],"Durante la construcción.",fill="FFF4EE"),
            (_mon_dual(W["mon"],directo,TC,fill="FFF4EE") if dual else _cell(W["mon"],m2(directo),fill="FFF4EE",bold=True,jc="right")),
            _cell(W["pct"],pc(directo/P*100),fill="FFF4EE",bold=True,jc="right")]))
        i+=1; hm=Pr-directo
        hip_mon = _mon_dual(W["mon"], hm, TC) if dual else _cell(W["mon"],m2(hm),bold=True,jc="right")
        out.append(_row([_cell(W["n"],str(i),jc="center"),
            _cell(W["con"],hipotecario["concepto"],bold=True,sub2=hipotecario.get("sub2")),
            _cell(W["fec"],hipotecario.get("fecha","")),
            hip_mon,
            _cell(W["pct"],pc(hm/P*100),jc="right")]))
    tot_mon = (_mon_dual(W["mon"],Pr,TC,fill="000000",white=True) if dual
               else _cell(W["mon"],m2(Pr),fill="000000",bold=True,white=True,jc="right"))
    out.append(_row([_cell(W["n"],"$",fill="000000",bold=True,white=True,jc="center"),
        _cell(W["con"],"PRECIO TOTAL DEL INMUEBLE",fill="000000",bold=True,white=True),
        _cell(W["fec"],"Total.",fill="000000",white=True),
        tot_mon,
        _cell(W["pct"],"100.00 %",fill="000000",bold=True,white=True,jc="right")]))
    g='<w:tblPr><w:tblW w:w="9072" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="CCCCCC"/><w:left w:val="single" w:sz="4" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/><w:right w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="3300"/><w:gridCol w:w="2372"/><w:gridCol w:w="1500"/><w:gridCol w:w="1200"/></w:tblGrid>'
    tbl='<w:tbl>'+g+''.join(out)+'</w:tbl>'
    return tbl

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
    if nombre=="20% inicial + 2 armadas de 40%":
        rows=[seprow,
          {"concepto":"Saldo a la firma","fecha":"A la firma (notaría / inicio de obra).","monto":P*0.20-sep},
          {"concepto":"1.ª Armada","fecha":"Junio de 2027.","monto":P*0.40},
          {"concepto":"2.ª Armada (saldo final)","fecha":"Diciembre de 2027.","monto":P*0.40}]
        return rows,None
    pcts={"10% directo + 90% hipotecario":(0.10,0.90),"20% directo + 80% hipotecario":(0.20,0.80),
          "30% directo + 70% hipotecario":(0.30,0.70),"40% directo + 60% hipotecario":(0.40,0.60),"50% directo + 50% hipotecario":(0.50,0.50)}
    if nombre in pcts:
        di,hi=pcts[nombre]; firma=P*0.15 if di>=0.30 else P*0.10
        rows=[seprow,{"concepto":"Saldo a la firma","fecha":"A la firma (notaría / inicio de obra).","monto":firma-sep}]
        rest=P*di-firma
        if rest>1:  # solo agregar armadas si queda saldo directo por repartir
            narm=3; fechas=["Diciembre de 2026.","Junio de 2027.","Diciembre de 2027."]
            for k in range(narm):
                rows.append({"concepto":f"{k+1}.ª Armada","fecha":fechas[k],"monto":rest/narm})
        return rows,{"concepto":"Saldo con crédito hipotecario","fecha":"Contra entrega (diciembre de 2027).","sub2":"Tasa, plazo y cuota los define el banco."}
    # personalizada (en blanco)
    return [seprow,{"concepto":"Saldo a la firma","fecha":"A la firma.","monto":P*0.20-sep}],None
