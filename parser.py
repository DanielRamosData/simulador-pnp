import docx
import re

def parse_docx(file_path):
    doc = docx.Document(file_path)
    questions = []
    
    current_q = None
    
    # Expresiones regulares para detectar preguntas y alternativas
    q_pattern = re.compile(r'^\s*(\d+)[\.\)]\s*(.*)', re.IGNORECASE)
    opt_pattern = re.compile(r'^\s*([A-E])[\.\)]\s*(.*)', re.IGNORECASE)
    ans_pattern = re.compile(r'(?:RPTA|RESPUESTA|CLAVE)[:\s]*([A-E])', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Detectar si es el inicio de una nueva pregunta
        q_match = q_pattern.match(text)
        if q_match:
            if current_q and current_q['question'] and len(current_q['options']) >= 2:
                questions.append(current_q)
            
            q_num = q_match.group(1)
            q_text = q_match.group(2)
            current_q = {
                'id': int(q_num),
                'question': q_text,
                'options': {},
                'correct': 'A'  # Por defecto A si no hay clave especificada
            }
            continue
            
        if current_q is not None:
            # 2. Detectar si es una alternativa (A, B, C, D, E)
            opt_match = opt_pattern.match(text)
            if opt_match:
                letter = opt_match.group(1).upper()
                opt_text = opt_match.group(2).strip()
                current_q['options'][letter] = opt_text
                
                # Si la alternativa tiene algún formato especial (negrita/subrayado) o marca (*)
                # o si la línea contiene la clave
                if '*' in text or 'CORRECTA' in text.upper():
                    current_q['correct'] = letter
                continue
            
            # 3. Detectar si la línea especifica la respuesta correcta (ej: RPTA: B)
            ans_match = ans_pattern.search(text)
            if ans_match:
                current_q['correct'] = ans_match.group(1).upper()
                continue
                
            # 4. Si es texto continuo de la pregunta previa
            if not current_q['options']:
                current_q['question'] += " " + text

    # Agregar la última pregunta procesada
    if current_q and current_q['question'] and len(current_q['options']) >= 2:
        questions.append(current_q)
        
    return questions
