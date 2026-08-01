import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    
    # 1. Unir todo el texto del archivo en un solo bloque continuo
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # 2. Separar por el número de pregunta (ej: "1000.", "1001.")
    raw_blocks = re.split(r'\n(?=\d+[\.\)\-])', full_text)
    
    questions = []
    
    for block in raw_blocks:
        block = block.strip()
        if not block:
            continue
            
        # Detectar el número de pregunta
        q_match = re.match(r'^(\d+)[\.\)\-]\s*(.*)', block, re.DOTALL)
        if not q_match:
            continue
            
        q_num = int(q_match.group(1))
        content = q_match.group(2).strip()
        
        # -------------------------------------------------------------
        # 3. LEER DESDE "RESPUESTA:" EN ADELANTE (Tu idea exacta)
        # -------------------------------------------------------------
        correcta_val = ""
        
        # Buscamos la palabra RESPUESTA o RPTA
        pos_respuesta = re.search(r'\b(RESPUESTA|RPTA)\b', content, re.IGNORECASE)
        
        if pos_respuesta:
            # Tomamos todo el texto DESDE donde dice RESPUESTA en adelante
            texto_desde_respuesta = content[pos_respuesta.start():]
            
            # Removemos la etiqueta "RESPUESTA:" o "RPTA:" del inicio
            texto_limpio = re.sub(r'^(RESPUESTA|RPTA)\s*:?\s*', '', texto_desde_respuesta, flags=re.IGNORECASE).strip()
            
            # Si en esa misma sección vienen UBICACIÓN o CÓDIGO, cortamos antes de llegar a ellos
            texto_limpio = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO):', texto_limpio, flags=re.IGNORECASE)[0].strip()
            
            # Tomamos la primera línea limpia y le quitamos viñetas o puntos de sobra
            primera_linea = texto_limpio.split('\n')[0].strip()
            correcta_val = re.sub(r'^[»>•\-\*\s]+', '', primera_linea).strip()
            
        # -------------------------------------------------------------
        # 4. EXTRAER ENUNCIADO Y OPCIONES (Todo lo que esté ANTES de RESPUESTA)
        # -------------------------------------------------------------
        if pos_respuesta:
            texto_antes_respuesta = content[:pos_respuesta.start()].strip()
        else:
            texto_antes_respuesta = content
            
        lineas = [l.strip() for l in texto_antes_respuesta.split('\n') if l.strip()]
        pregunta_texto = lineas[0] if lineas else ""
        
        opciones = []
        for linea in lineas[1:]:
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
            if clean_opt:
                opciones.append(clean_opt)
                
        # -------------------------------------------------------------
        # 5. EXTRAER METADATOS (UBICACIÓN / CÓDIGO)
        # -------------------------------------------------------------
        ubic_match = re.search(r'UBICACI[ÓO]N\s*:?\s*(.*)', content, re.IGNORECASE)
        cod_match = re.search(r'C[ÓO]DIGO\s*:?\s*(.*)', content, re.IGNORECASE)
        
        modulo_val = ubic_match.group(1).split('\n')[0].strip() if ubic_match else ""
        codigo_val = cod_match.group(1).split('\n')[0].strip() if cod_match else f"PNP-{q_num}"
        
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
