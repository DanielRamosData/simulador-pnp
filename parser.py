import docx
import re

def extraer_lineas_puras(doc):
    lineas = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            txt = p.text.strip()
            if txt:
                lineas.append(txt)
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    txt = cell.text.strip()
                    if txt:
                        for sub_linea in txt.split('\n'):
                            if sub_linea.strip():
                                lineas.append(sub_linea.strip())
    return lineas

def limpiar_texto_respuesta(cadena):
    """
    Recibe un texto tipo 'ALLANAMIENTO ILEGAL DE DOMICILIO. (ART: 160)** [LIBRO...]'
    y devuelve solo el nombre del delito: 'ALLANAMIENTO ILEGAL DE DOMICILIO'
    """
    if not cadena:
        return ""
    # Cortar antes de (ART:, [, **, o el primer paréntesis/corchete de metadatos
    partes = re.split(r'(\(ART:|\[|\*\*|\()', cadena, flags=re.IGNORECASE)
    delito = partes[0].strip().rstrip('.')
    return re.sub(r'^[»>•\-\*\s]+', '', delito).strip()

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    lineas = extraer_lineas_puras(doc)
    
    questions = []
    q_actual = None
    
    regex_num = re.compile(r'^(\d+)[\.\)\-]\s*(.*)')
    regex_resp = re.compile(r'^\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\b\s*[\.:\-\_]*\s*(.*)', re.IGNORECASE)
    
    # Captura lineas como: UBICACIÓN: CÓDIGO: LESIONES CULPOSAS. (ART: 124)...
    regex_ubic_cod = re.compile(r'^UBICACI[ÓO]N\s*:?\s*C[ÓO]DIGO\s*:?\s*(.*)', re.IGNORECASE)
    regex_ubic = re.compile(r'^UBICACI[ÓO]N\s*:?\s*(.*)', re.IGNORECASE)
    regex_cod = re.compile(r'^C[ÓO]DIGO\s*:?\s*(.*)', re.IGNORECASE)

    for linea in lineas:
        match_num = regex_num.match(linea)
        
        # 1. Nueva pregunta
        if match_num:
            if q_actual:
                questions.append(q_actual)
            
            num_p = int(match_num.group(1))
            resto_texto = match_num.group(2).strip()
            
            q_actual = {
                "num": num_p,
                "pregunta": resto_texto,
                "opciones": [],
                "correcta": "",
                "modulo": "",
                "codigo_id": f"PNP-{num_p}",
                "_leyendo_resp": False
            }
            continue

        if not q_actual:
            continue

        # 2. Si es una línea que empieza con RESPUESTA:
        match_resp = regex_resp.match(linea)
        if match_resp:
            val_resp = match_resp.group(2).strip()
            q_actual["correcta"] = limpiar_texto_respuesta(val_resp)
            q_actual["_leyendo_resp"] = True
            continue

        # 3. Si es una línea combinada UBICACIÓN: CÓDIGO:
        match_comb = regex_ubic_cod.match(linea)
        if match_comb:
            contenido = match_comb.group(1).strip()
            # Si NO se capturó una respuesta antes, la extraemos de esta línea
            if not q_actual["correcta"]:
                q_actual["correcta"] = limpiar_texto_respuesta(contenido)
            
            q_actual["modulo"] = contenido
            q_actual["_leyendo_resp"] = False
            continue

        # 4. Ubicación o Código individual (por si vienen separados)
        match_ubic = regex_ubic.match(linea)
        if match_ubic:
            q_actual["modulo"] = match_ubic.group(1).strip()
            q_actual["_leyendo_resp"] = False
            continue

        match_cod = regex_cod.match(linea)
        if match_cod:
            val_c = match_cod.group(1).strip()
            if val_c:
                q_actual["codigo_id"] = val_c
            q_actual["_leyendo_resp"] = False
            continue

        # 5. Acumular Enunciado u Opciones
        if not q_actual["correcta"]:
            if not q_actual["pregunta"]:
                q_actual["pregunta"] = linea
            else:
                opc_limpia = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
                if opc_limpia:
                    q_actual["opciones"].append(opc_limpia)
        elif q_actual["_leyendo_resp"] and not match_ubic and not match_cod:
            # Concatenar si la respuesta era de varias líneas
            q_actual["correcta"] += " " + limpiar_texto_respuesta(linea)

    if q_actual:
        questions.append(q_actual)

    for q in questions:
        q.pop("_leyendo_resp", None)

    return questions

def estructurar_15_modulos(preguntas, preguntas_por_modulo=100):
    modulos = {}
    total = len(preguntas)
    if total == 0:
        return modulos
    
    for i in range(0, total, preguntas_por_modulo):
        num_modulo = (i // preguntas_por_modulo) + 1
        modulos[f"Módulo {num_modulo}"] = preguntas[i:i + preguntas_por_modulo]
    return modulos
