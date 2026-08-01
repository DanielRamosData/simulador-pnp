import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Detecta números de cualquier longitud (ej: "1.", "1000.", "1001.")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)\-]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar inicio de nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            if current_q and current_q.get("pregunta"):
                # Si no se capturó la respuesta, usar la primera opción
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
            
            # 2. Ignorar líneas de UBICACIÓN y CÓDIGO para que no entren como alternativas ni respuestas
            if text_upper.startswith(("UBICACIÓN:", "UBICACION:", "CÓDIGO:", "CODIGO:")):
                if text_upper.startswith(("UBICACIÓN:", "UBICACION:")):
                    current_q["modulo"] = text.split(":", 1)[1].strip()
                elif text_upper.startswith(("CÓDIGO:", "CODIGO:")):
                    current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
                
            # 3. Capturar RESPUESTA:
            if text_upper.startswith("RESPUESTA:"):
                raw_ans = text.split(":", 1)[1].strip()
                # Limpiar viñetas si la respuesta las traía
                clean_ans = re.sub(r'^[»>•\-\*\s]+', '', raw_ans).strip()
                current_q["correcta"] = clean_ans
                continue
            
            # 4. Capturar Alternativas
            clean_opt = re.sub(r'^[»>•\-\*\s]+', '', text).strip()
            if clean_opt:
                current_q["opciones"].append(clean_opt)

    # Guardar la última pregunta procesada
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
