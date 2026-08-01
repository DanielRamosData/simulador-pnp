import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    
    # 1. Unir todo el texto en una sola cadena
    full_text = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    
    # 2. Separar por bloques usando los números de pregunta (ej: "1000.", "1062.")
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
        lines = [l.strip() for l in content.split('\n') if l.strip()]
        
        pregunta_texto = lines[0] if lines else ""
        opciones = []
        correcta_val = ""
        modulo_val = ""
        codigo_val = f"PNP-{q_num}"
        
        # 3. Procesar línea por línea dentro del bloque para máxima precisión
        for line in lines[1:]:
            l_upper = line.upper()
            
            # Detectar variantes de RESPUESTA / RPTA
            if "RESPUESTA" in l_upper or "RPTA" in l_upper:
                # Extraer todo lo que está después del primer dos puntos ':'
                if ":" in line:
                    val = line.split(":", 1)[1].strip()
                else:
                    val = re.sub(r'^(RESPUESTA|RPTA)\s*', '', line, flags=re.IGNORECASE).strip()
                
                # Cortar si viene pegado UBICACIÓN o CÓDIGO en la misma línea
                val_clean = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO):', val, flags=re.IGNORECASE)[0].strip()
                correcta_val = re.sub(r'^[»>•\-\*\s]+', '', val_clean).strip()
                continue
                
            # Detectar UBICACIÓN
            if "UBICACI" in l_upper:
                if ":" in line:
                    modulo_val = line.split(":", 1)[1].strip()
                continue
                
            # Detectar CÓDIGO
            if "CÓDIGO" in l_upper or "CODIGO" in l_upper:
                if ":" in line:
                    codigo_val = line.split(":", 1)[1].strip()
                continue
                
            # Si no es ninguna etiqueta de metadato, es una opción de respuesta
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', line).strip()
            if clean_opt:
                opciones.append(clean_opt)
        
        # Respaldo: Si no se detectó "RESPUESTA:", tomar la última opción si existe
        if not correcta_val and opciones:
            # En algunos formatos la última opción/línea marcada es la correcta
            pass

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
