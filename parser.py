import docx
import re

def extraer_texto_secuencial(doc):
    """
    Lee el documento Word respetando el ORDEN REAL en el que aparecen 
    los párrafos y las tablas en la página.
    """
    texto = []
    for element in doc.element.body:
        # Si el elemento es un Párrafo
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            if p.text.strip():
                texto.append(p.text.strip())
        # Si el elemento es una Tabla
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                # Unir el texto de las celdas de la fila
                celdas = [c.text.strip() for c in row.cells if c.text.strip()]
                if celdas:
                    # Si la fila tiene varias celdas, las unimos con espacio
                    texto.append(" ".join(celdas))
                    
    return "\n".join(texto)

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    
    # 1. Extraer todo el texto en orden secuencial estricto
    full_text = extraer_texto_secuencial(doc)
    
    # 2. Separar bloques por número de pregunta (ej: "1000.", "1001.", "1000.-")
    raw_blocks = re.split(r'\n(?=\d+[\.\)\-])', full_text)
    
    questions = []
    
    for block in raw_blocks:
        block = block.strip()
        if not block:
            continue
            
        q_match = re.match(r'^(\d+)[\.\)\-]\s*(.*)', block, re.DOTALL)
        if not q_match:
            continue
            
        q_num = int(q_match.group(1))
        content = q_match.group(2).strip()
        
        # -------------------------------------------------------------
        # 3. BUSCAR LA RESPUESTA (Soporta RESPUESTA, RPTA, CLAVE, SOLUCION, RESP)
        # -------------------------------------------------------------
        correcta_val = ""
        
        # Buscar variantes comunes
        pos_resp = re.search(r'\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\b', content, re.IGNORECASE)
        
        if pos_resp:
            # Tomar desde la palabra encontrada en adelante
            sub_str = content[pos_resp.start():]
            
            # Limpiar la etiqueta encontrada y los símbolos que la acompañan (: . - _)
            sub_str = re.sub(r'^(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA)\s*[\.:\-\_]*\s*', '', sub_str, flags=re.IGNORECASE).strip()
            
            # Cortar si aparecen etiquetas de metadatos posteriores (UBICACIÓN, CÓDIGO, TEMA)
            corte_meta = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO|TEMA|MODULO):', sub_str, flags=re.IGNORECASE)
            sub_str = corte_meta[0].strip()
            
            # Tomar la primera línea limpia
            lineas_resp = [l.strip() for l in sub_str.split('\n') if l.strip()]
            if lineas_resp:
                correcta_val = re.sub(r'^[»>•\-\*\s]+', '', lineas_resp[0]).strip()

        # -------------------------------------------------------------
        # 4. EXTRAER METADATOS (UBICACIÓN Y CÓDIGO)
        # -------------------------------------------------------------
        modulo_val = ""
        ubic_match = re.search(r'UBICACI[ÓO]N\s*:?\s*([^C\n]+)', content, re.IGNORECASE)
        if ubic_match:
            modulo_val = ubic_match.group(1).strip()

        codigo_val = f"PNP-{q_num}"
        cod_match = re.search(r'C[ÓO]DIGO\s*:?\s*(\d+)', content, re.IGNORECASE)
        if cod_match:
            codigo_val = cod_match.group(1).strip()

        # -------------------------------------------------------------
        # 5. EXTRAER ENUNCIADO Y OPCIONES
        # -------------------------------------------------------------
        texto_antes = content[:pos_resp.start()].strip() if pos_resp else content
        lineas = [l.strip() for l in texto_antes.split('\n') if l.strip()]
        
        pregunta_texto = lineas[0] if lineas else ""
        opciones = [re.sub(r'^[»>•\-\*\s]+', '', l).strip() for l in lineas[1:] if l.strip()]

        questions.append({
            "num": q_num,
            "pregunta": pregunta_texto,
            "opciones": opciones,
            "correcta": correcta_val,
            "modulo": modulo_val,
            "codigo_id": codigo_val
        })
        
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
