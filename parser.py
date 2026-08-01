import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Expresión regular para el número de pregunta (ej: "1000. COMETE DELITO...")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar inicio de una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            # Guardar la pregunta anterior si existía
            if current_q and current_q.get("pregunta"):
                # Si por alguna razón la respuesta quedó vacía, usar la primera opción
                if not current_q["correcta"] and current_q["opciones"]:
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
            
            # 2. Capturar línea de RESPUESTA:
            if text_upper.startswith("RESPUESTA:"):
                # Extraer el texto tras "RESPUESTA:"
                ans_text = text.split(":", 1)[1].strip()
                # Eliminar metadatos si vienen pegados en la misma línea (UBICACIÓN, CÓDIGO)
                ans_clean = re.split(r'\b(UBICACIÓ|UBICACIO|CÓDIGO|CODIGO):', ans_text, flags=re.IGNORECASE)[0].strip()
                current_q["correcta"] = ans_clean
                continue
                
            # 3. Capturar o descartar UBICACIÓN:
            if text_upper.startswith("UBICACIÓN:") or text_upper.startswith("UBICACION:"):
                current_q["modulo"] = text.split(":", 1)[1].strip()
                continue
                
            # 4. Capturar o descartar CÓDIGO:
            if text_upper.startswith("CÓDIGO:") or text_upper.startswith("CODIGO:"):
                current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
            
            # 5. Capturar alternativas (excluyendo cualquier metadato)
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', text).strip()
            if clean_opt:
                current_q["opciones"].append(clean_opt)

    # Guardar la última pregunta del archivo
    if current_q and current_q.get("pregunta"):
        if not current_q["correcta"] and current_q["opciones"]:
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
