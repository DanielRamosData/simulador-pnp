import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Expresión regular para detectar el número de pregunta (ej: "1. TODA PERSONA...", "1000. COMETE DELITO...")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar si inicia una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            # Guardar pregunta anterior
            if current_q and current_q.get("pregunta"):
                # Limpieza de seguridad si la respuesta quedó vacía o apuntó a metadatos
                if not current_q["correcta"] or current_q["correcta"].upper().startswith(("UBICAC", "CÓDIG", "CODIG")):
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
            
            # 2. Capturar RESPUESTA:
            if text_upper.startswith("RESPUESTA:"):
                ans_content = text.split(":", 1)[1].strip()
                # Si hay contenido en la misma línea, guardarlo limpiando metadatos pegados
                ans_clean = re.split(r'\b(UBICACIÓ|UBICACIO|CÓDIGO|CODIGO):', ans_content, flags=re.IGNORECASE)[0].strip()
                if ans_clean:
                    current_q["correcta"] = ans_clean
                continue
                
            # 3. Capturar UBICACIÓN:
            if text_upper.startswith("UBICACIÓN:") or text_upper.startswith("UBICACION:"):
                current_q["modulo"] = text.split(":", 1)[1].strip()
                continue
                
            # 4. Capturar CÓDIGO:
            if text_upper.startswith("CÓDIGO:") or text_upper.startswith("CODIGO:"):
                current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
            
            # 5. Guardar Opciones (Excluyendo explícitamente cualquier metadato)
            if not text_upper.startswith(("RESPUESTA:", "UBICACIÓN:", "UBICACION:", "CÓDIGO:", "CODIGO:")):
                # Limpiar viñetas especiales (», >, •, -, *, etc.)
                clean_opt = re.sub(r'^[»>•\-\*\s]+', '', text).strip()
                if clean_opt:
                    current_q["opciones"].append(clean_opt)

    # Guardar la última pregunta del archivo Word
    if current_q and current_q.get("pregunta"):
        if not current_q["correcta"] or current_q["correcta"].upper().startswith(("UBICAC", "CÓDIG", "CODIG")):
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
