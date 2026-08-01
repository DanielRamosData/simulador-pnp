import docx
import re

def extraer_texto_completo(doc):
    texto = []
    # 1. Párrafos normales
    for p in doc.paragraphs:
        if p.text.strip():
            texto.append(p.text.strip())
            
    # 2. Contenido dentro de tablas (si existen)
    for table in doc.tables:
        for row in table.rows:
            fila_txt = []
            for cell in row.cells:
                if cell.text.strip():
                    fila_txt.append(cell.text.strip())
            if fila_txt:
                texto.append(" | ".join(fila_txt))
                    
    return "\n".join(texto)

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    full_text = extraer_texto_completo(doc)
    
    # Separar por número de pregunta (ej: "1000.", "1001.", "1000.-", etc.)
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
        # 1. EXTRAER LA RESPUESTA CON PATRÓN AMPLIADO (MULTIVARIANTE)
        # -------------------------------------------------------------
        correcta_val = ""
        
        # Busca: RESPUESTA, RPTA, RPTA., RESP, CORRECTA (mayúsculas o minúsculas)
        pos_resp = re.search(r'\b(RESPUESTA|RPTA|RESP|CORRECTA)\b', content, re.IGNORECASE)
        
        if pos_resp:
            # Tomar texto desde donde se encuentra la etiqueta
            sub_str = content[pos_resp.start():]
            
            # Quitar la etiqueta encontrada y símbolos como ':', '.', '-'
            sub_str = re.sub(r'^(RESPUESTA|RPTA|RESP|CORRECTA)\s*[\.:\-\_]*\s*', '', sub_str, flags=re.IGNORECASE).strip()
            
            # Cortar si aparecen etiquetas de metadatos posteriores (UBICACIÓN, CÓDIGO, TEMA)
            corte_meta = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO|TEMA|MODULO):', sub_str, flags=re.IGNORECASE)
            sub_str = corte_meta[0].strip()
            
            # Tomar la primera línea válida o elemento
            lineas_resp = [l.strip() for l in sub_str.split('\n') if l.strip()]
            if lineas_resp:
                val_candidato = lineas_resp[0]
                # Si en esa línea hay separadores de tabla '|', tomar la parte relevante
                if '|' in val_candidato:
                    val_candidato = val_candidato.split('|')[0].strip()
                
                # Limpiar viñetas o caracteres extra
                correcta_val = re.sub(r'^[»>•\-\*\s]+', '', val_candidato).strip()

        # -------------------------------------------------------------
        # 2. EXTRAER METADATOS (UBICACIÓN / CÓDIGO)
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
        # 3. EXTRAER ENUNCIADO Y OPCIONES
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
