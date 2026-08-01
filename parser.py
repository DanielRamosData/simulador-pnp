import docx
import re

def extraer_y_limpiar_lineas(doc):
    """
    Extrae las líneas del documento Word y descarta de inmediato
    cualquier línea que contenga UBICACIÓN o CÓDIGO.
    """
    lineas_limpias = []
    
    # 1. Función auxiliar para filtrar líneas no deseadas
    def es_linea_valida(texto):
        t = texto.strip().upper()
        # Si contiene UBICACIÓN o CÓDIGO (o variaciones), se descarta
        if "UBICACIÓN:" in t or "UBICACION:" in t or "CÓDIGO:" in t or "CODIGO:" in t:
            return False
        return True

    # 2. Recorrer párrafos y tablas secuencialmente
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            txt = p.text.strip()
            if txt and es_linea_valida(txt):
                lineas_limpias.append(txt)
                
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    txt = cell.text.strip()
                    if txt:
                        for sub_linea in txt.split('\n'):
                            sub_l = sub_linea.strip()
                            if sub_l and es_linea_valida(sub_l):
                                lineas_limpias.append(sub_l)
                                
    return lineas_limpias

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    # Extraemos solo las líneas útiles (sin ubicación ni código)
    lineas = extraer_y_limpiar_lineas(doc)
    
    questions = []
    q_actual = None
    
    # Detectar número de pregunta (ej: "1000.", "1001.-")
    regex_num = re.compile(r'^(\d+)[\.\)\-]\s*(.*)')
    # Detectar encabezado de respuesta
    regex_resp = re.compile(r'^\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\b\s*[\.:\-\_]*\s*(.*)', re.IGNORECASE)

    for linea in lineas:
        match_num = regex_num.match(linea)
        
        # A) Inicio de nueva pregunta
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

        # B) Si encontramos la línea de RESPUESTA:
        match_resp = regex_resp.match(linea)
        if match_resp:
            val_resp = match_resp.group(2).strip()
            # Limpiar viñetas si las hay
            q_actual["correcta"] = re.sub(r'^[»>•\-\*\s]+', '', val_resp).strip()
            q_actual["_leyendo_resp"] = True
            continue

        # C) Si ya estamos capturando datos de la pregunta
        if not q_actual["correcta"]:
            if not q_actual["pregunta"]:
                q_actual["pregunta"] = linea
            else:
                opc_limpia = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
                if opc_limpia:
                    q_actual["opciones"].append(opc_limpia)
        elif q_actual["_leyendo_resp"]:
            # Si la respuesta ocupa más de una línea, la concatenamos
            q_actual["correcta"] += " " + linea.strip()

    # Guardar la última pregunta
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
