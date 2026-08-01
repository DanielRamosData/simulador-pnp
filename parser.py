import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    q_pattern = re.compile(r'^\s*(\d+)[\.\)]\s*(.*)', re.IGNORECASE)
    opt_pattern = re.compile(r'^\s*([A-E])[\.\)]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # Detectar inicio de pregunta (ej: "1.", "2)")
        q_match = q_pattern.match(text)
        if q_match:
            if current_q and current_q.get("pregunta") and len(current_q.get("opciones", [])) >= 2:
                # Si no se detectó respuesta explícita, usar la primera alternativa como referencia
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
            # Detectar si es una alternativa (A, B, C, D, E)
            opt_match = opt_pattern.match(text)
            if opt_match:
                opt_text = opt_match.group(2).strip()
                current_q["opciones"].append(opt_text)
                
                # Si la alternativa tiene marca (*) o dice CORRECTA
                if '*' in text or 'CORRECTA' in text.upper():
                    current_q["correcta"] = opt_text
                continue
            
            # Detectar clave en línea independiente
            if text.upper().startswith(('RPTA:', 'RESPUESTA:', 'CLAVE:')):
                partes = text.split(':', 1)
                if len(partes) > 1:
                    current_q["correcta"] = partes[1].strip()
                continue

            # Si es continuación de la pregunta
            if not current_q["opciones"]:
                current_q["pregunta"] += " " + text

    if current_q and current_q.get("pregunta") and len(current_q.get("opciones", [])) >= 2:
        if not current_q["correcta"] and current_q["opciones"]:
            current_q["correcta"] = current_q["opciones"][0]
        questions.append(current_q)
        
    return questions

def estructurar_15_modulos(preguntas, preguntas_por_modulo=100):
    modulos = {}
    for i in range(0, len(preguntas), preguntas_por_modulo):
        num_modulo = (i // preguntas_por_modulo) + 1
        modulos[f"Módulo {num_modulo}"] = preguntas[i:i + preguntas_por_modulo]
    return modulos
