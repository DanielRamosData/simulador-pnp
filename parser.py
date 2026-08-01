import docx
import re

def extraer_texto_completo(doc):
    texto = []
    for p in doc.paragraphs:
        if p.text.strip():
            texto.append(p.text.strip())
            
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    texto.append(cell.text.strip())
                    
    return "\n".join(texto)

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    full_text = extraer_texto_completo(doc)
    
    # Separar bloques por número (1000., 1001., etc.)
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
        
        # 1. BUSCAR LA RESPUESTA
        correcta_val = ""
        # Buscar RESPUESTA o RPTA
        pos_resp = re.search(r'\b(RESPUESTA|RPTA)\b', content, re.IGNORECASE)
        
        if pos_resp:
            # Tomar desde "RESPUESTA" en adelante
            sub_str = content[pos_resp.start():]
            
            # Quitar la palabra RESPUESTA/RPTA y los dos puntos si existen
            sub_str = re.sub(r'^(RESPUESTA|RPTA)\s*:?\s*', '', sub_str, flags=re.IGNORECASE).strip()
            
            # CORTAR INMEDIATAMENTE si aparece UBICACIÓN o CÓDIGO
            corte_meta = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO):', sub_str, flags=re.IGNORECASE)
            sub_str = corte_meta[0].strip()
            
            # Tomar solo la primera línea limpia
            primera_linea = sub_str.split('\n')[0].strip()
            correcta_val = re.sub(r'^[»>•\-\*\s]+', '', primera_linea).strip()

        # 2. BUSCAR METADATOS DE FORMA AISLADA
        modulo_val = ""
        ubic_match = re.search(r'UBICACI[ÓO]N\s*:?\s*([^C\n]+)', content, re.IGNORECASE)
        if ubic_match:
            modulo_val = ubic_match.group(1).strip()

        codigo_val = f"PNP-{q_num}"
        cod_match = re.search(r'C[ÓO]DIGO\s*:?\s*(\d+)', content, re.IGNORECASE)
        if cod_match:
            codigo_val = cod_match.group(1).strip()

        # 3. EXTRAER ENUNCIADO Y OPCIONES
        texto_antes = content[:pos_resp.start()].strip() if pos_resp else content
        lineas = [l.strip() for l in texto_antes.split('\n') if l.strip()]
        
        pregunta_texto = lineas[0] if lineas else ""
        opciones = [re.sub(r'^[»>•\-\*\s]+', '', l).strip() for l in lineas[1:] if l.strip()]

        # Retornar el diccionario con llaves exactas
        questions.append({
            "num": q_num,
            "pregunta": pregunta_texto,
            "opciones": opciones,
            "correcta": correcta_val,  # <-- 'ROBO AGRAVADO.' quedará aquí
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
