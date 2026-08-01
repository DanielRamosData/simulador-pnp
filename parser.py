import re
import random
from docx import Document

def cargar_banco_preguntas(ruta_o_stream_docx):
    """
    Parsea el Word delimitando limpiamente cada pregunta para evitar
    que el texto de la Pregunta 1 se contamine con el resto.
    """
    doc = Document(ruta_o_stream_docx)
    
    lineas = []
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            lineas.append(txt)
            
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    txt = p.text.strip()
                    if txt:
                        lineas.append(txt)

    preguntas = []
    enunciado_lineas = []
    opciones = []
    modulo_actual = "General"
    codigo_actual = ""
    correcta_actual = "A"
    
    for linea in lineas:
        linea_upper = linea.upper()
        
        # Si la línea contiene metadatos de cierre de pregunta
        if any(kw in linea_upper for kw in ["RESPUESTA", "RESPUETA", "RPTA:"]):
            match_resp = re.search(r'(?:RESPUESTA|RESPUETA|RPTA):\s*([^UBICACI[ÓO]N|C[ÓO]DIGO|\n]+)', linea, re.IGNORECASE)
            if match_resp:
                correcta_actual = match_resp.group(1).strip()
            
            match_ubic = re.search(r'UBICACI[ÓO]N:\s*([^C[ÓO]DIGO|\n]+)', linea, re.IGNORECASE)
            if match_ubic:
                modulo_actual = match_ubic.group(1).strip()
                
            match_cod = re.search(r'C[ÓO]DIGO:\s*(\d+)', linea, re.IGNORECASE)
            if match_cod:
                codigo_actual = match_cod.group(1).strip()
            
            # Construir enunciado
            enunciado_completo = " ".join(enunciado_lineas).strip()
            enunciado_limpio = re.sub(r'^\s*\d+[\.\)\-\s]*', '', enunciado_completo).strip()
            
            if enunciado_limpio or opciones:
                preguntas.append({
                    "codigo_id": codigo_actual if codigo_actual else str(len(preguntas) + 1),
                    "pregunta": enunciado_limpio if enunciado_limpio else f"Pregunta {len(preguntas) + 1}",
                    "opciones": opciones.copy(),
                    "correcta": correcta_actual,
                    "modulo": modulo_actual
                })
                
            # REBOOT OBLIGATORIO: Se limpian los acumuladores para la SIGUIENTE pregunta
            enunciado_lineas = []
            opciones = []
            correcta_actual = "A"
            codigo_actual = ""
            continue

        # Si es una opción
        if linea.startswith("»") or linea.startswith(">") or linea.startswith("•"):
            opc_limpia = re.sub(r'^[»>•]\s*', '', linea).strip()
            if opc_limpia:
                opciones.append(opc_limpia)
            continue
            
        # Si es parte del enunciado
        if not any(tag in linea_upper for tag in ["UBICACIÓN:", "UBICACION:", "CÓDIGO:", "CODIGO:"]):
            # Si entramos a un nuevo enunciado y ya teníamos opciones cargadas sin haber procesado RESPUESTA, reseteamos
            enunciado_lineas.append(linea)

    return preguntas


def estructurar_15_modulos(banco_preguntas, preguntas_por_modulo=100):
    """
    Divide el banco global procesado en 15 módulos ordenados de 100 preguntas cada uno.
    """
    banco_copia = banco_preguntas.copy()
    
    modulos_dict = {}
    for i in range(1, 16):
        inicio = (i - 1) * preguntas_por_modulo
        fin = inicio + preguntas_por_modulo
        
        sub_pool = banco_copia[inicio:fin] if inicio < len(banco_copia) else []
        
        preguntas_modulo = []
        for idx, q in enumerate(sub_pool, 1):
            q_copy = q.copy()
            q_copy["num"] = idx
            preguntas_modulo.append(q_copy)
            
        modulos_dict[f"Módulo {i}"] = preguntas_modulo
        
    return modulos_dict