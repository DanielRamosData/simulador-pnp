import docx
import re

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    questions = []
    current_q = None
    
    # Expresiones regulares
    q_pattern = re.compile(r'^\s*(\d+)[\.\)]\s*(.*)', re.IGNORECASE)
    opt_pattern = re.compile(r'^\s*([A-E])[\.\)]\s*(.*)', re.IGNORECASE)
    ans_pattern = re.compile(r'(?:RPTA|RESPUESTA|CLAVE)[:\s]*([A-E])', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # Detectar inicio de pregunta
        q_match = q_pattern.match(text)
        if q_match:
            if current_q and current_q.get('question') and len(current_q.get('options', {})) >= 2:
                questions.append(current_q)
            
            q_num = q_match.group(1)
            q_text = q_match.group(2)
            current_q = {
                'id': int(q_num),
                'question': q_text,
                'options': {},
                'correct': 'A'
            }
            continue
            
        if current_q is not None:
            # Detectar opción A-E
            opt_match = opt_pattern.match(text)
            if opt_match:
                letter = opt_match.group(1).upper()
                opt_text = opt_match.group(2).strip()
                current_q['options'][letter] = opt_text
                
                if '*' in text or 'CORRECTA' in text.upper():
                    current_q['correct'] = letter
                continue
            
            # Detectar clave explícita
            ans_match = ans_pattern.search(text)
            if ans_match:
                current_q['correct'] = ans_match.group(1).upper()
                continue
                
            # Continuación de texto del enunciado
            if not current_q['options']:
                current_q['question'] += " " + text

    if current_q and current_q.get('question') and len(current_q.get('options', {})) >= 2:
        questions.append(current_q)
        
    return questions

def estructurar_15_modulos(preguntas, tamaño_modulo=100):
    modulos = {}
    for i in range(0, len(preguntas), tamaño_modulo):
        num_modulo = (i // tamaño_modulo) + 1
        modulos[f"Módulo {num_modulo}"] = preguntas[i:i + tamaño_modulo]
    return modulos
