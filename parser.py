import docx
import re

def extraer_lineas_puras(doc):
    """
    Extrae todas las líneas de texto del documento (de párrafos y tablas)
    manteniendo el orden secuencial estricto.
    """
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
                        # Si hay varias líneas dentro de la celda, las agregamos individualmente
                        for sub_linea in txt.split('\n'):
                            if sub_linea.strip():
                                lineas.append(sub_linea.strip())
    return lineas

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    lineas = extraer_lineas_puras(doc)
    
    questions = []
    q_actual = None
    
    # Expresión para detectar el inicio de una pregunta (ej: "1000.", "1001.-")
    regex_num = re.compile(r'^(\d+)[\.\)\-]\s*(.*)')
    # Expresión para detectar respuesta
    regex_resp = re.compile(r'^\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\b\s*[\.:\-\_]*\s*(.*)', re.IGNORECASE)
    # Expresiones para metadatos
    regex_ubic = re.compile(r'^UBICACI[ÓO]N\s*:?\s*(.*)', re.IGNORECASE)
    regex_cod = re.compile(r'^C[ÓO]DIGO\s*:?\s*(.*)', re.IGNORECASE)

    for linea in lineas:
        match_num = regex_num.match(linea)
        
        # 1. Si encontramos el inicio de una NUEVA pregunta
        if match_num:
            # Si ya teníamos una pregunta en proceso, la guardamos
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

        # 2. Si detectamos la etiqueta RESPUESTA:
        match_resp = regex_resp.match(linea)
        if match_resp:
            val_resp = match_resp.group(2).strip()
            q_actual["correcta"] = re.sub(r'^[»>•\-\*\s]+', '', val_resp).strip()
            q_actual["_leyendo_resp"] = True
            continue

        # 3. Si detectamos UBICACIÓN:
        match_ubic = regex_ubic.match(linea)
        if match_ubic:
            q_actual["modulo"] = match_ubic.group(1).strip()
            q_actual["_leyendo_resp"] = False
            continue

        # 4. Si detectamos CÓDIGO:
        match_cod = regex_cod.match(linea)
        if match_cod:
            val_c = match_cod.group(1).strip()
            if val_c:
                q_actual["codigo_id"] = val_c
            q_actual["_leyendo_resp"] = False
            continue

        # 5. Si no es etiqueta y aún no hay respuesta -> Es enunciado u opción
        if not q_actual["correcta"]:
            # Si el enunciado estaba vacío lo asignamos, de lo contrario es una opción
            if not q_actual["pregunta"]:
                q_actual["pregunta"] = linea
            else:
                opc_limpia = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
                if opc_limpia:
                    q_actual["opciones"].append(opc_limpia)
        # Si ya estábamos leyendo la respuesta y esta ocupa más de una línea (multilínea)
        elif q_actual["_leyendo_resp"] and not match_ubic and not match_cod:
            # Si la respuesta es de varias líneas, concatenamos
            q_actual["correcta"] += " " + linea.strip()

    # Guardar la última pregunta del archivo
    if q_actual:
        questions.append(q_actual)

    # Limpieza final de llaves auxiliares
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
