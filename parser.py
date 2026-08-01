import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    
    # 1. Unir todos los párrafos en un solo texto con saltos de línea
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # 2. Separar el texto completo por cada número de pregunta (ej: "1000.", "1001.")
    raw_blocks = re.split(r'\n(?=\d+[\.\)\-])', full_text)
    
    questions = []
    
    for block in raw_blocks:
        block = block.strip()
        if not block:
            continue
            
        # Validar si el bloque empieza con número de pregunta (ej: "1000. COMETE...")
        q_match = re.match(r'^(\d+)[\.\)\-]\s*(.*)', block, re.DOTALL)
        if not q_match:
            continue
            
        q_num = int(q_match.group(1))
        content = q_match.group(2).strip()
        
        # 3. EXTRAER RESPUESTA (Cortando estrictamente antes de UBICACIÓN o CÓDIGO)
        ans_match = re.search(r'RESPUESTA:\s*(.*?)(?=\s*(?:UBICACI[ÓO]N:|C[ÓO]DIGO:|$))', content, re.IGNORECASE | re.DOTALL)
        correcta_val = ""
        if ans_match:
            # Tomamos la captura limpia y eliminamos saltos de línea o caracteres raros
            linea_respuesta = ans_match.group(1).split('\n')[0].strip()
            correcta_val = re.sub(r'^[»>•\-\*\s]+', '', linea_respuesta).strip()
        
        # 4. EXTRAER ENUNCIADO Y OPCIONES (Todo lo que está antes de "RESPUESTA:")
        partes_antes_respuesta = re.split(r'RESPUESTA:', content, flags=re.IGNORECASE)[0].strip()
        lineas = [l.strip() for l in partes_antes_respuesta.split('\n') if l.strip()]
        
        pregunta_texto = lineas[0] if lineas else ""
        opciones = []
        
        for linea in lineas[1:]:
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', linea).strip()
            if clean_opt:
                opciones.append(clean_opt)
                
        # 5. Extraer UBICACIÓN y CÓDIGO
        ubic_match = re.search(r'UBICACI[ÓO]N:\s*(.*?)(?=\s*(?:C[ÓO]DIGO:|$))', content, re.IGNORECASE | re.DOTALL)
        cod_match = re.search(r'C[ÓO]DIGO:\s*(.*)', content, re.IGNORECASE)
        
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
