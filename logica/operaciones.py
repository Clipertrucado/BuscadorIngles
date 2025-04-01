from openpyxl import load_workbook, Workbook
import random

def cargar_datos(archivo_entrada):
    wb = load_workbook(archivo_entrada)
    sheet = wb.active
    encabezados = [cell.value for cell in sheet[1]]
    return wb, sheet, encabezados

def verificar_columnas(encabezados):
    if 'TEMA' not in encabezados or not any(col.strip() in encabezados for col in ['CABO 24-25', 'SARGENTO 24-25', 'OFICIAL 24-25', 'CAMBIO ESCALA 24-25']):
        raise ValueError("El archivo no contiene todas las columnas necesarias: TEMA y alguna columna de curso")

def filtrar_filas(sheet, temas, idx_tema, idx_cursos):
    filas_filtradas = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[idx_tema] in temas and all(row[i] is None or row[i] == "" for i in idx_cursos):
            filas_filtradas.append(row)
    return filas_filtradas

def seleccionar_preguntas(filas_filtradas, num_preguntas):
    if num_preguntas and num_preguntas < len(filas_filtradas):
        return random.sample(filas_filtradas, num_preguntas)
    return filas_filtradas

def obtener_preguntas(archivo_entrada, temas, cursos, num_preguntas):
    temas = [int(t.strip()) for t in temas if t.strip().isdigit()]
    wb, sheet, encabezados = cargar_datos(archivo_entrada)
    verificar_columnas(encabezados)
    idx_tema = encabezados.index('TEMA')
    idx_cursos = [encabezados.index(col.strip()) for col in cursos if col.strip() in encabezados]

    filas_disponibles_totales = filtrar_filas(sheet, temas, idx_tema, idx_cursos)
    filas_seleccionadas = seleccionar_preguntas(filas_disponibles_totales, num_preguntas)

    return encabezados, filas_seleccionadas, filas_disponibles_totales

def reemplazar_pregunta(filas_disponibles_totales, filas_seleccionadas, indice):
    seleccionadas_set = set(tuple(row) for row in filas_seleccionadas)
    disponibles = [row for row in filas_disponibles_totales if tuple(row) not in seleccionadas_set]
    if not disponibles:
        return None
    nueva = random.choice(disponibles)
    filas_seleccionadas[indice] = nueva
    return nueva

def guardar_preguntas(encabezados, filas, ruta_destino):
    wb_nuevo = Workbook()
    sheet_nuevo = wb_nuevo.active
    sheet_nuevo.append(encabezados)
    for fila in filas:
        sheet_nuevo.append(fila)
    wb_nuevo.save(ruta_destino)
    return True