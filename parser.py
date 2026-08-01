import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Regex para el número de pregunta (ej: "1. TODA PERSONA...", "1000. COMETE DELITO...")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar si inicia una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            # Guardar la pregunta anterior
            if current_q and current_q.get("pregunta"):
                # Si la respuesta quedó vacía o corrupta, usar la primera opción
                if not current_q["correcta"] or "UBICACI" in current_q["correcta"].upper():
                    if current_q["opciones"]:
                        current_q["correcta"] = current_q["opciones"][0]
                questions.append(current_q)
            
            q_num = int(q_match.group(1))
            q_text = q_match.group(2).strip()
            
            current_q = {
                "num": q_num,
                "pregunta": q_text,
                "opciones": [],
                "correcta": "",
                "modulo": "",
                "codigo_id": f"PNP-{q_num}"
            }
            continue
            
        if current_q is not None:
            text_upper = text.upper()
            
            # 2. Capturar RESPUESTA y amputar metadatos posteriores si están en la misma línea
            if text_upper.startswith("RESPUESTA:"):
                # Extraer lo que está después de "RESPUESTA:"
                raw_ans = text.split(":", 1)[1].strip()
                
                # Eliminar todo desde UBICACIÓN: o CÓDIGO: hasta el final si vienen en la misma línea
                clean_ans = re.sub(r'\s*(UBICACI[ÓO]N|C[ÓO]DIGO):.*$', '', raw_ans, flags=re.IGNORECASE).strip()
                
                current_q["correcta"] = clean_ans
                continue
                
            # 3. Capturar UBICACIÓN solo para el campo de módulo
            if text_upper.startswith("UBICACIÓN:") or text_upper.startswith("UBICACION:"):
                current_q["modulo"] = text.split(":", 1)[1].strip()
                continue
                
            # 4. Capturar CÓDIGO solo para la ID
            if text_upper.startswith("CÓDIGO:") or text_upper.startswith("CODIGO:"):
                current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
            
            # 5. Si no es un encabezado de metadatos, procesar como Opción
            if not any(text_upper.startswith(k) for k in ["RESPUESTA", "UBICACI", "CÓDIGO", "CODIGO"]):
                clean_opt = re.sub(r'^[»>•\-\*\s]+', '', text).strip()
                if clean_opt:
                    current_q["opciones"].append(clean_opt)

    # Guardar la última pregunta del documento
    if current_q and current_q.get("pregunta"):
        if not current_q["correcta"] or "UBICACI" in current_q["correcta"].upper():
            if current_q["opciones"]:
                current_q["correcta"] = current_q["opciones"][0]
        questions.append(current_q)
        
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
