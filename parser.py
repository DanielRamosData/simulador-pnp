import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    
    # 1. Unir todo el texto del documento
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # 2. Separar por bloques de preguntas usando el número (ej: 1000., 1001.)
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
        # 3. EXTRAER LA RESPUESTA DE MANERA ROBUSTA
        # -------------------------------------------------------------
        correcta_val = ""
        # Buscar variantes: RESPUESTA:, RESPUESTA :, RPTA:, RPTA :
        # Captura todo el texto hasta encontrar UBICACIÓN:, CÓDIGO:, otra etiqueta o fin del bloque
        ans_match = re.search(
            r'(?:RESPUESTA|RPTA)\s*:?\s*(.*?)(?=\s*(?:UBICACI[ÓO]N|C[ÓO]DIGO|\n\d+[\.\)\-]|$))', 
            content, 
            re.IGNORECASE | re.DOTALL
        )
        
        if ans_match:
            raw_ans = ans_match.group(1).strip()
            # Si la captura tiene saltos internos, tomar solo el primer renglón
            raw_ans = raw_ans.split('\n')[0].strip()
            # Limpiar viñetas o caracteres residuales
            correcta_val = re.sub(r'^[»>•\-\*\s]+', '', raw_ans).strip()
            # Eliminar punto final si viene pegado a etiquetas
            correcta_val = re.sub(r'\s*(UBICACI[ÓO]N|C[ÓO]DIGO).*$', '', correcta_val, flags=re.IGNORECASE).strip()

        # -------------------------------------------------------------
        # 4. EXTRAER ENUNCIADO Y OPCIONES
        # -------------------------------------------------------------
        # Dividir el texto antes de la palabra RESPUESTA/RPTA
        partes = re.split(r'\b(?:RESPUESTA|RPTA)\b', content, flags=re.IGNORECASE)
        texto_pregunta_y_opciones = partes[0].strip() if partes else content
        
        lineas = [l.strip() for l in texto_pregunta_y_opciones.split('\n') if l.strip()]
        
        pregunta_texto = lineas[0] if lineas else ""
        opciones = []
        
        for linea in lineas[1:]:
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
            if clean_opt:
                opciones.append(clean_opt)
                
        # -------------------------------------------------------------
        # 5. EXTRAER METADATOS (UBICACIÓN Y CÓDIGO)
        # -------------------------------------------------------------
        ubic_match = re.search(r'UBICACI[ÓO]N\s*:?\s*(.*?)(?=\s*(?:C[ÓO]DIGO|\n|$))', content, re.IGNORECASE | re.DOTALL)
        cod_match = re.search(r'C[ÓO]DIGO\s*:?\s*(\d+)', content, re.IGNORECASE)
        
        modulo_val = ubic_match.group(1).split('\n')[0].strip() if ubic_match else ""
        codigo_val = cod_match.group(1).strip() if cod_match else f"PNP-{q_num}"
        
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
