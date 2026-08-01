import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Expresión regular para capturar la numeración inicial (ej: "1.", "1.-", "1)")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar inicio de una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            # Guardar la pregunta anterior procesada
            if current_q and current_q.get("pregunta"):
                # Si no capturó respuesta explícita, usa la primera alternativa como fallback
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
            
            # 2. Detectar línea de RESPUESTA:
            if text_upper.startswith("RESPUESTA:"):
                current_q["correcta"] = text.split(":", 1)[1].strip()
                continue
                
            # 3. Detectar datos de UBICACIÓN:
            if text_upper.startswith("UBICACIÓN:") or text_upper.startswith("UBICACION:"):
                current_q["modulo"] = text.split(":", 1)[1].strip()
                continue
                
            # 4. Detectar datos de CÓDIGO:
            if text_upper.startswith("CÓDIGO:") or text_upper.startswith("CODIGO:"):
                current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
            
            # 5. Detectar alternativas eliminando viñetas (», >, -, •, etc.)
            opt_clean = re.sub(r'^[»>•\-\*\s]+', '', text).strip()
            
            # Si se eliminó una viñeta o si ya estábamos recolectando opciones
            if opt_clean != text or len(current_q["opciones"]) > 0:
                current_q["opciones"].append(opt_clean)
            else:
                # Si aún no hay opciones, es continuación del texto de la pregunta
                current_q["pregunta"] += " " + text

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
