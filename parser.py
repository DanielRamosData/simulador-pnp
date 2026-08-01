import docx
import re

def extraer_texto_secuencial(doc):
    texto = []
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            if p.text.strip():
                texto.append(p.text.strip())
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                celdas = [c.text.strip() for c in row.cells if c.text.strip()]
                if celdas:
                    texto.append(" | ".join(celdas))
                    
    return "\n".join(texto)

def cargar_banco_preguntas(file_path):
    doc = docx.Document(file_path)
    full_text = extraer_texto_secuencial(doc)
    
    # Separar bloques por número de pregunta (ej: "1000.", "1001.", "1000.-")
    raw_blocks = re.split(r'\n(?=\d+[\.\)\-])', full_text)
    
    questions = []
    
    for block in raw_blocks:
        block = block.strip()
        if not block:
            continue
            
        q_match = re.match(r'^(\d+)[\.\)\-]\s*(.*)', block, re.DOTALL)
        if not q_match:
            continue
            
        q_num = int(q_match.group(1))
        content = q_match.group(2).strip()
        
        # -------------------------------------------------------------
        # 1. BÚSQUEDA MULTI-PATRÓN DE LA RESPUESTA
        # -------------------------------------------------------------
        correcta_val = ""
        
        # Patrón amplio para detectar cualquier encabezado de respuesta
        patron_ans = r'\b(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA|RPTA\.|RPTA\s*:)\b'
        pos_resp = re.search(patron_ans, content, re.IGNORECASE)
        
        if pos_resp:
            sub_str = content[pos_resp.start():]
            # Limpiar etiquetas y puntuaciones iniciales
            sub_str = re.sub(r'^(RESPUESTA|RPTA|CLAVE|SOLUCI[ÓO]N|RESP|CORRECTA|RPTA\.|RPTA\s*:)\s*[\.:\-\_]*\s*', '', sub_str, flags=re.IGNORECASE).strip()
            
            # Cortar si topamos con metadatos de ubicación o código
            corte_meta = re.split(r'\b(UBICACI[ÓO]N|C[ÓO]DIGO|TEMA|MODULO):', sub_str, flags=re.IGNORECASE)
            sub_str = corte_meta[0].strip()
            
            lineas_resp = [l.strip() for l in sub_str.split('\n') if l.strip()]
            if lineas_resp:
                val = lineas_resp[0]
                if '|' in val:
                    val = val.split('|')[0].strip()
                correcta_val = re.sub(r'^[»>•\-\*\s]+', '', val).strip()
        else:
            # Estrategia de respaldo: Buscar si alguna línea interna empieza con "R:" o similar
            lineas_bloque = [l.strip() for l in content.split('\n') if l.strip()]
            for l in lineas_bloque:
                m_r = re.match(r'^(?:R|Rpta|Resp)\s*[\.:\-]\s*(.*)', l, re.IGNORECASE)
                if m_r:
                    correcta_val = m_r.group(1).strip()
                    break

        # -------------------------------------------------------------
        # 2. EXTRAER METADATOS (UBICACIÓN Y CÓDIGO)
        # -------------------------------------------------------------
        modulo_val = ""
        ubic_match = re.search(r'UBICACI[ÓO]N\s*:?\s*([^C\n]+)', content, re.IGNORECASE)
        if ubic_match:
            modulo_val = ubic_match.group(1).strip()

        codigo_val = f"PNP-{q_num}"
        cod_match = re.search(r'C[ÓO]DIGO\s*:?\s*(\d+)', content, re.IGNORECASE)
        if cod_match:
            codigo_val = cod_match.group(1).strip()

        # -------------------------------------------------------------
        # 3. EXTRAER ENUNCIADO Y OPCIONES
        # -------------------------------------------------------------
        texto_antes = content[:pos_resp.start()].strip() if pos_resp else content
        lineas = [l.strip() for l in texto_antes.split('\n') if l.strip()]
        
        pregunta_texto = lineas[0] if lineas else ""
        opciones = [re.sub(r'^[»>•\-\*\s]+', '', l).strip() for l in lineas[1:] if l.strip()]

        questions.append({
            "num": q_num,
            "pregunta": pregunta_texto,
            "opciones": opciones,
            "correcta": correcta_val,
            "modulo": modulo_val,
            "codigo_id": codigo_val
        })
        
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
