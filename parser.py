import docx
import re

def extraer_lineas_puras(doc):
    lineas = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            txt = p.text.replace('\xa0', ' ').strip()
            if txt:
                lineas.append(txt)
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    txt = cell.text.replace('\xa0', ' ').strip()
                    if txt:
                        for sub_linea in txt.split('\n'):
                            sub_txt = sub_linea.strip()
                            if sub_txt:
                                lineas.append(sub_txt)
    return lineas

def limpiar_texto_respuesta(cadena):
    if not cadena:
        return ""
    
    # 1. Remover prefijos de UBICACIÓN / CÓDIGO al inicio
    cadena = re.sub(r'^(UBICACI[ÓO]N\s*:?\s*|C[ÓO]DIGO\s*:?\s*)+', '', cadena, flags=re.IGNORECASE).strip()
    
    # 2. Cortar antes de referencias legales como (ART:, [, **, etc.
    partes = re.split(r'(\(ART:|\[|\*\*|\()', cadena, flags=re.IGNORECASE)
    delito = partes[0].strip()
    
    # 3. Remover viñetas, puntos finales y espacios
    delito = re.sub(r'^[»>•\-\*\s]+', '', delito)
    return delito.rstrip('.').strip()

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    lineas = extraer_lineas_puras(doc)
    
    questions = []
    q_actual = None
    
    regex_num = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)')
    regex_resp = re.compile(r'\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\b\s*[\.:\-\_]*\s*(.*)', re.IGNORECASE)
    regex_ubic_cod = re.compile(r'\bUBICACI[ÓO]N\b\s*:?\s*(C[ÓO]DIGO\s*:?)?\s*(.*)', re.IGNORECASE)
    regex_cod = re.compile(r'\bC[ÓO]DIGO\b\s*:?\s*(.*)', re.IGNORECASE)

    for linea in lineas:
        match_num = regex_num.search(linea)
        
        # 1. Inicio de nueva pregunta
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

        # 2. Detección de RESPUESTA:
        match_resp = regex_resp.search(linea)
        if match_resp:
            val_resp = match_resp.group(2).strip()
            q_actual["correcta"] = limpiar_texto_respuesta(val_resp)
            q_actual["_leyendo_resp"] = True
            continue

        # 3. Detección de UBICACIÓN: / CÓDIGO:
        match_comb = regex_ubic_cod.search(linea)
        if match_comb:
            contenido = match_comb.group(2).strip() if match_comb.group(2) else linea
            if not q_actual["correcta"]:
                q_actual["correcta"] = limpiar_texto_respuesta(contenido)
            
            q_actual["modulo"] = contenido
            q_actual["_leyendo_resp"] = False
            continue

        # 4. Detección de CÓDIGO: individual
        match_cod = regex_cod.search(linea)
        if match_cod:
            val_c = match_cod.group(1).strip()
            if val_c:
                q_actual["codigo_id"] = val_c
            q_actual["_leyendo_resp"] = False
            continue

        # 5. Lectura de Enunciado, Opciones o Respuesta multilínea
        if not q_actual["correcta"]:
            if not q_actual["pregunta"]:
                q_actual["pregunta"] = linea
            else:
                opc_limpia = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
                if opc_limpia:
                    q_actual["opciones"].append(opc_limpia)
        elif q_actual["_leyendo_resp"] and not match_comb and not match_cod:
            # Concatenar respuesta si venía en líneas subsiguientes
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
