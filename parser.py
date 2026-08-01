import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Patrones para detectar preguntas y alternativas
    q_pattern = re.compile(r'^\s*(\d+)[\.\)]\s*(.*)', re.IGNORECASE)
    opt_pattern = re.compile(r'^\s*([A-E])[\.\)]\s*(.*)', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar inicio de una nueva pregunta (ej: "1.", "2)")
        q_match = q_pattern.match(text)
        if q_match:
            if current_q and current_q.get('question') and len(current_q.get('options', [])) >= 2:
                questions.append(current_q)
            
            q_num = q_match.group(1)
            q_text = q_match.group(2)
            current_q = {
                'id': int(q_num),
                'question': q_text,
                'options': [],       # Lista de alternativas en texto
                'correct_text': ''   # Texto exacto de la respuesta correcta
            }
            continue
            
        if current_q is not None:
            # 2. Detectar si es una alternativa (A, B, C, D, E)
            opt_match = opt_pattern.match(text)
            if opt_match:
                opt_text = opt_match.group(2).strip()
                current_q['options'].append(opt_text)
                
                # Si la alternativa tiene una marca (*, negrita o la palabra CORRECTA)
                if '*' in text or 'CORRECTA' in text.upper():
                    current_q['correct_text'] = opt_text
                continue
            
            # 3. Detectar si la respuesta correcta está escrita en una línea aparte (ej: "RPTA: ...")
            if text.upper().startswith(('RPTA:', 'RESPUESTA:', 'CLAVE:')):
                partes = text.split(':', 1)
                if len(partes) > 1:
                    current_q['correct_text'] = partes[1].strip()
                continue

            # 4. Si es continuación del texto de la pregunta previa
            if not current_q['options']:
                current_q['question'] += " " + text

    # Si la pregunta no tenía marca explícita, tomamos la primera opción como referencia predeterminada
    if current_q and current_q.get('question') and len(current_q.get('options', [])) >= 2:
        if not current_q['correct_text'] and current_q['options']:
            current_q['correct_text'] = current_q['options'][0]
        questions.append(current_q)
        
    return questions

def estructurar_15_modulos(preguntas, tamaño_modulo=100):
    modulos = {}
    for i in range(0, len(preguntas), tamaño_modulo):
        num_modulo = (i // tamaño_modulo) + 1
        modulos[f"Módulo {num_modulo}"] = preguntas[i:i + tamaño_modulo]
    return modulos
