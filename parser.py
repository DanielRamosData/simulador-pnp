import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Patrón para el número de pregunta (ej: "1. TODA PERSONA...")
    q_pattern = re.compile(r'^\s*(\d+)[\.\)]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar si inicia una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            # Si ya teníamos una pregunta en proceso, la guardamos
            if current_q and current_q.get("pregunta") and len(current_q.get("opciones", [])) >= 2:
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
            # 2. Detectar línea de RESPUESTA:
            if text.upper().startswith("RESPUESTA:"):
                ans_text = text.split(":", 1)[1].strip()
                current_q["correcta"] = ans_text
                continue
                
            # 3. Detectar datos de UBICACIÓN:
            if text.upper().startswith("UBICACIÓN:") or text.upper().startswith("UBICACION:"):
                current_q["modulo"] = text.split(":", 1)[1].strip()
                continue
                
            # 4. Detectar datos de CÓDIGO:
            if text.upper().startswith("CÓDIGO:") or text.upper().startswith("CODIGO:"):
                current_q["codigo_id"] = text.split(":", 1)[1].strip()
                continue
            
            # 5. Detectar si es una alternativa (empieza con » o similar)
            if text.startswith("»") or text.startswith(">") or text.startswith("-"):
                opt_text = text.lstrip("»>- ").strip()
                current_q["opciones"].append(opt_text)
                continue
                
            # 6. Si es texto de la pregunta (si aún no se han leído opciones)
            if not current_q["opciones"] and not current_q["correcta"]:
                current_q["pregunta"] += " " + text

    # Guardar la última pregunta del archivo
    if current_q and current_q.get("pregunta") and len(current_q.get("opciones", [])) >= 2:
        questions.append(current_q)
        
    return questions

def estructurar_15_modulos(preguntas, preguntas_por_modulo=100):
    modulos = {}
    for i in range(0, len(preguntas), preguntas_por_modulo):
        num_modulo = (i // preguntas_por_modulo) + 1
        modulos[f"Módulo {num_modulo}"] = preguntas[i:i + preguntas_por_modulo]
    return modulos
