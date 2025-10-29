import streamlit as st
import pandas as pd
import random
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import letter # Usaremos 'letter' como base

# --- CONSTANTES GLOBALES ---

# Diccionario de tamaños KDP (ancho, alto en cm)
TAMAÑOS_KDP = {
    '5" x 8" (12.7 x 20.32 cm)': (12.7, 20.32),
    '5.06" x 7.81" (12.85 x 19.84 cm)': (12.85, 19.84),
    '5.25" x 8" (13.34 x 20.32 cm)': (13.34, 20.32),
    '5.5" x 8.5" (13.97 x 21.59 cm)': (13.97, 21.59),
    '6" x 9" (15.24 x 22.86 cm)': (15.24, 22.86),
    '6.14" x 9.21" (15.6 x 23.39 cm)': (15.6, 23.39),
    '6.69" x 9.61" (16.99 x 24.41 cm)': (16.99, 24.41),
    '7" x 10" (17.78 x 25.4 cm)': (17.78, 25.4),
    '8" x 10" (20.32 x 25.4 cm)': (20.32, 25.4),
    '8.25" x 6" (20.96 x 15.24 cm)': (20.96, 15.24),
    '8.25" x 8.25" (20.96 x 20.96 cm)': (20.96, 20.96),
    '8.27" x 11.69" (21.01 x 29.69 cm)': (21.01, 29.69),
    '8.5" x 8.5" (21.59 x 21.59 cm)': (21.59, 21.59),
    '8.5" x 11" (21.59 x 27.94 cm)': (21.59, 27.94)
}

# --- NUEVAS CONSTANTES DE FUENTE ---
FONT_NAME_REGULAR = "RobotoMono-Regular"
FONT_NAME_BOLD = "RobotoMono-Bold"
TTF_FILE_REGULAR = "RobotoMono-Regular.ttf"
TTF_FILE_BOLD = "RobotoMono-Bold.ttf"

# Constantes de Diseño
PADDING = 0.75 * inch  # Margen de seguridad interior
TITLE_FONT_SIZE = 18   # Tamaño fijo para el título de la sopa
WORDS_FONT_SIZE = 11   # Tamaño fijo para la lista de palabras
SPACE_AFTER_TITLE = 0.2 * inch
INTERLINEADO_PALABRAS = 1.3
PROPORCION_LETRA_CUADRICULA = 0.7  # 70% del tamaño de la celda
NUM_COLUMNAS_PALABRAS = 4
SOLUCIONES_POR_PAGINA = 4
LETRAS_ALFABETO = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

# NUEVAS constantes para el recuadro
BOX_PADDING = 0.15 * inch        # Relleno interno del recuadro de palabras
BOX_CORNER_RADIUS = 0.1 * inch   # Radio de las esquinas redondeadas

# --- FUNCIONES DE LÓGICA (Sin cambios) ---

def cm_to_pt(cm):
    """Convierte centímetros a puntos de ReportLab."""
    return cm * 28.3465

def crear_sopa_letras(palabras, dimension=20):
    sopa = [[' ' for _ in range(dimension)] for _ in range(dimension)]
    ubicaciones = {}
    palabras_no_colocadas = []
    
    direcciones = [(1,0), (0,1), (1,1), (-1,1), (-1,0), (0,-1), (-1,-1), (1,-1)]
    
    def chequear_fit(palabra, fila, col, dy, dx):
        for i in range(len(palabra)):
            fila_actual, col_actual = fila + i * dy, col + i * dx
            if not (0 <= fila_actual < dimension and 0 <= col_actual < dimension):
                return False
            celda = sopa[fila_actual][col_actual]
            if celda != ' ' and celda != palabra[i]:
                return False
        return True

    def colocar_palabra(palabra):
        palabra_upper = palabra.upper()
        posibles_ubicaciones = []
        
        for dy, dx in direcciones:
            for fila_inicio in range(dimension):
                for col_inicio in range(dimension):
                    if chequear_fit(palabra_upper, fila_inicio, col_inicio, dy, dx):
                        posibles_ubicaciones.append((fila_inicio, col_inicio, dy, dx))
        
        if posibles_ubicaciones:
            random.shuffle(posibles_ubicaciones)
            fila_inicio, col_inicio, dy, dx = posibles_ubicaciones[0]
            
            for i in range(len(palabra_upper)):
                fila_actual = fila_inicio + i * dy
                col_actual = col_inicio + i * dx
                sopa[fila_actual][col_actual] = palabra_upper[i]
            
            return (fila_inicio, col_inicio, dx, dy)
        
        return None

    for palabra in palabras:
        ubicacion = colocar_palabra(palabra)
        if ubicacion:
            ubicaciones[palabra.upper()] = ubicacion
        else:
            palabras_no_colocadas.append(palabra)
            
    for i in range(dimension):
        for j in range(dimension):
            if sopa[i][j] == ' ':
                sopa[i][j] = random.choice(LETRAS_ALFABETO)
                
    return sopa, ubicaciones, palabras_no_colocadas

def procesar_excel(excel_file, words_per_puzzle):
    df = pd.read_excel(excel_file)
    word_lists = []
    themes = []
    current_words = []
    current_theme = ""

    for index, row in df.iterrows():
        cell_value = str(row.iloc[0]).strip()
        
        if ',' in cell_value:
            word, theme = cell_value.split(',', 1)
            word, theme = word.strip(), theme.strip()
            
            if word and word.lower() != 'nan':
                current_words.append(word)
            if theme and not current_theme:
                current_theme = theme
        else:
            if cell_value and cell_value.lower() != 'nan':
                current_words.append(cell_value)
        
        if len(current_words) == words_per_puzzle or index == len(df) - 1:
            if current_words:
                word_lists.append(current_words.copy())
                themes.append(current_theme if current_theme else f"Sopa de letras {len(themes) + 1}")
                current_words.clear()
                current_theme = ""
                
    return word_lists, themes

# --- ESTA ES LA FUNCIÓN 'dibujar_pagina_sopa' (CORREGIDA Y ÚNICA) ---
# Resuelve el problema del centrado vertical y el padding superior/inferior.

def dibujar_pagina_sopa(c, sopa, palabras, theme, config):
    """Dibuja una (1) página de sopa de letras en el canvas."""
    
    page_width, page_height = config['page_size']
    dimension = config['dimension']
    
    ancho_usable = page_width - 2 * PADDING
    y_cursor = page_height - PADDING

    # 1. Dibujar Título (en negrita)
    c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
    c.drawString(PADDING, y_cursor - TITLE_FONT_SIZE, theme[:60])
    y_cursor -= (TITLE_FONT_SIZE + SPACE_AFTER_TITLE)

    # 2. Dibujar Recuadro y Lista de palabras
    c.setFont(FONT_NAME_REGULAR, WORDS_FONT_SIZE)
    line_height = WORDS_FONT_SIZE * INTERLINEADO_PALABRAS
    
    # --- LÓGICA DE CÁLCULO ---
    n_palabras = len(palabras)
    n_cols = NUM_COLUMNAS_PALABRAS
    
    # Esta es la fórmula clave (equivale a math.ceil)
    max_words_per_col = (n_palabras + n_cols - 1) // n_cols
    
    # --- NUEVA LÓGICA DE CENTRADO ---
    # 1. Averiguamos cuántas palabras hay en CADA columna
    col_word_counts = [0] * n_cols
    for i in range(n_palabras):
        # Usamos la misma lógica de división para asignar
        col_index = i // max_words_per_col
        if col_index < n_cols:
            col_word_counts[col_index] += 1
            
    # 2. Estimamos la altura de la fuente para el padding superior
    font_ascent_guess = WORDS_FONT_SIZE * 0.8
    # --- FIN NUEVA LÓGICA ---

    # Calcular altura total del recuadro (basada en la columna más alta)
    word_list_height_inner = max_words_per_col * line_height
    word_list_height_total = word_list_height_inner + (2 * BOX_PADDING)

    # Coordenadas del recuadro
    y_box_bottom = y_cursor - word_list_height_total
    
    # --- DIBUJAR EL RECUADRO ---
    c.setLineWidth(1)
    c.setStrokeColor(black)
    c.roundRect(
        PADDING,                  # x
        y_box_bottom,             # y
        ancho_usable,             # width
        word_list_height_total,   # height
        BOX_CORNER_RADIUS,        # radius
        stroke=1, 
        fill=0
    )

    # Ancho de columna
    col_width = (ancho_usable - (2 * BOX_PADDING)) / n_cols
    
    # --- BUCLE DE DIBUJO (ESTABLE) ---
    for i, word in enumerate(palabras):
        col_index = i // max_words_per_col
        row_index = i % max_words_per_col
        
        # --- CÁLCULO DE 'Y' CON CENTRADO ---
        # 1. Obtenemos el total de palabras de esta columna
        words_in_this_col = col_word_counts[col_index]
        
        # 2. Calculamos el offset vertical para esta columna
        this_col_height_inner = words_in_this_col * line_height
        vertical_offset = (word_list_height_inner - this_col_height_inner) / 2
        
        # 3. Posición X
        cx = PADDING + BOX_PADDING + (col_index * col_width)
        
        # 4. Posición Y (CORREGIDA)
        y_top_of_centered_block = (y_cursor - BOX_PADDING - vertical_offset)
        cy_base = y_top_of_centered_block - font_ascent_guess
        cy = cy_base - (row_index * line_height)
        # --- FIN CÁLCULO 'Y' ---

        c.drawString(cx, cy, word.upper()[:18])

    # Bajar el cursor hasta debajo del recuadro
    y_cursor -= (word_list_height_total + PADDING)

    # 3. Dibujar la Cuadrícula (en el espacio restante)
    grid_available_height = y_cursor - PADDING
    grid_available_width = ancho_usable
    cell_size = min(grid_available_width / dimension, grid_available_height / dimension)
    
    grid_width = cell_size * dimension
    grid_height = cell_size * dimension
    
    x_grid = PADDING + (grid_available_width - grid_width) / 2
    y_grid = PADDING + (grid_available_height - grid_height) / 2
    
    c.setLineWidth(1)
    c.setStrokeColor(black)
    c.rect(x_grid, y_grid, grid_width, grid_height, stroke=1, fill=0)
    
    grid_font_size = int(cell_size * PROPORCION_LETRA_CUADRICULA)
    c.setFont(FONT_NAME_REGULAR, grid_font_size)
    
    for i in range(dimension):
        for j in range(dimension):
            letra = sopa[i][j].upper()
            x_letra = x_grid + j * cell_size + cell_size / 2
            y_letra = y_grid + grid_height - (i + 0.5) * cell_size
            y_letra_ajustada = y_letra - (grid_font_size * 0.3) 
            c.drawCentredString(x_letra, y_letra_ajustada, letra)

# --- ESTA ES LA FUNCIÓN 'dibujar_paginas_soluciones' (ÚNICA) ---

def dibujar_paginas_soluciones(c, sopas_list, palabras_list, ubicaciones_list, themes, config):
    """Dibuja todas las páginas de soluciones."""
    
    page_width, page_height = config['page_size']
    dimension = config['dimension']

    if not sopas_list:
        return

    c.showPage()
    pagina_solucion_actual = 1
    
    ancho_usable = page_width - 2 * PADDING
    alto_usable = page_height - 2 * PADDING
    
    # Título inicial (en negrita)
    c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
    c.drawString(PADDING, page_height - PADDING - TITLE_FONT_SIZE, "Soluciones")

    sol_area_width = ancho_usable / 2
    sol_area_height = (alto_usable - (TITLE_FONT_SIZE + PADDING)) / 2
    cell_size_sol = min(sol_area_width / dimension, sol_area_height / dimension) * 0.85
    mini_grid_w = cell_size_sol * dimension
    mini_grid_h = cell_size_sol * dimension

    for sol_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        
        pos_en_pagina = sol_index % SOLUCIONES_POR_PAGINA
        
        if sol_index > 0 and pos_en_pagina == 0:
            c.showPage()
            pagina_solucion_actual += 1
            c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
            c.drawString(PADDING, page_height - PADDING - TITLE_FONT_SIZE, f"Soluciones (pág. {pagina_solucion_actual})")

        col_sol = pos_en_pagina % 2
        row_sol = pos_en_pagina // 2
        
        x_offset = PADDING + col_sol * sol_area_width
        left_sol = x_offset + (sol_area_width - mini_grid_w) / 2
        
        y_area_base = PADDING + (1 - row_sol) * sol_area_height
        bot_sol = y_area_base + (sol_area_height - mini_grid_h) / 2
        
        if row_sol == 0:
             bot_sol = (page_height - PADDING - TITLE_FONT_SIZE - PADDING - sol_area_height) + (sol_area_height - mini_grid_h) / 2

        celdas_solucion = set()
        for p in palabras:
            ubicacion = ubicaciones.get(p.upper())
            if ubicacion:
                fila_ini, col_ini, dx, dy = ubicacion
                for k in range(len(p)):
                    celdas_solucion.add((fila_ini + k*dy, col_ini + k*dx))

        # Cuadrícula y letras (en regular)
        c.setFont(FONT_NAME_REGULAR, int(cell_size_sol * 0.7))
        c.setLineWidth(0.5)
        for i in range(dimension):
            for j in range(dimension):
                x = left_sol + j * cell_size_sol
                y = bot_sol + mini_grid_h - (i + 1) * cell_size_sol
                c.rect(x, y, cell_size_sol, cell_size_sol, stroke=1, fill=0)
                
                if (i, j) in celdas_solucion:
                    c.setFillColor(black)
                    c.drawCentredString(x + cell_size_sol / 2, y + (cell_size_sol * 0.2), sopa[i][j].upper())

        # Título de la mini-sopa (en regular)
        c.setFont(FONT_NAME_REGULAR, int(cell_size_sol * 0.6))
        c.setFillColor(black)
        c.drawCentredString(left_sol + mini_grid_w / 2, bot_sol - 14, theme)

# --- FUNCIÓN 'exportar_pdf' (ÚNICA) ---

def exportar_pdf(sopas_list, palabras_list, ubicaciones_list, themes, dimension, page_size):
    """
    Crea el archivo PDF completo llamando a las funciones de dibujo.
    """
    buffer = BytesIO()
    
    # Registramos las fuentes aquí.
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME_REGULAR, TTF_FILE_REGULAR))
        pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, TTF_FILE_BOLD))
    except Exception as e:
        st.error(f"¡ERROR DE FUENTE! No se pudieron cargar los archivos .ttf.")
        st.error(f"Asegúrate de que 'RobotoMono-Regular.ttf' y 'RobotoMono-Bold.ttf' estén en la misma carpeta que el script.")
        st.error(f"Detalle del error: {e}")
        return None # Devuelve None para que Streamlit sepa que falló

    c = canvas.Canvas(buffer, pagesize=page_size)
    
    config = {
        "page_size": page_size,
        "dimension": dimension
    }

    # Dibujar páginas de sopas
    for puzzle_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        if puzzle_index > 0:
            c.showPage()
        dibujar_pagina_sopa(c, sopa, palabras, theme, config)

    # Dibujar páginas de soluciones
    dibujar_paginas_soluciones(c, sopas_list, palabras_list, ubicaciones_list, themes, config)

    c.save()
    buffer.seek(0)
    return buffer


# --- INTERFAZ DE STREAMLIT (UI) ---

st.title("Generador de Sopas de Letras KDP")
st.write("¡Bienvenido! Sube tu archivo Excel de palabras para empezar.")

excel_file = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])
dimension = st.number_input("Tamaño de cuadrícula (ej. 20x20)", min_value=10, max_value=30, value=20)
words_per_puzzle = st.number_input("Número de palabras por sopa", min_value=5, max_value=40, value=20)
nombre_tamaño = st.selectbox("Tamaño de página KDP", list(TAMAÑOS_KDP.keys()), index=12)

ancho_cm, alto_cm = TAMAÑOS_KDP[nombre_tamaño]
page_width = cm_to_pt(ancho_cm)
page_height = cm_to_pt(alto_cm)

if st.button("Generar PDF de Sopas de Letras"):
    if excel_file is not None:
        
        word_lists, themes = procesar_excel(excel_file, words_per_puzzle)
        
        if not word_lists:
            st.error("No se encontraron palabras en el archivo Excel. Revisa el formato.")
        else:
            st.info(f"Se van a generar {len(word_lists)} sopas de letras.")
        
            sopas_list = []
            palabras_list_filled = []
            ubicaciones_list = []
            
            with st.spinner("Generando sopas..."):
                for words, theme in zip(word_lists, themes):
                    sopa, ubicaciones, no_colocadas = crear_sopa_letras(words, dimension=dimension)
                    
                    if no_colocadas:
                        st.warning(f"En la sopa '{theme}', no se pudieron colocar las siguientes palabras: {', '.join(no_colocadas)}")
                    
                    sopas_list.append(sopa)
                    palabras_list_filled.append(words)
                    ubicaciones_list.append(ubicaciones)

            with st.spinner("Creando archivo PDF..."):
                buffer = exportar_pdf(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension, (page_width, page_height))
            
            if buffer: # Solo mostrar el botón de descarga si el buffer se creó
                st.success("¡PDF generado! Descárgalo aquí:")
                st.download_button(
                    label="Descargar PDF",
                    data=buffer,
                    file_name='sopas_de_letras_kdp.pdf',
                    mime='application/pdf'
                )
    else:
        st.error("Por favor, sube un archivo Excel.")