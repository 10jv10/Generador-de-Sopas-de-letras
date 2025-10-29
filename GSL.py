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

<<<<<<< HEAD
# --- CONSTANTES GLOBALES DE DISEÑO ---
=======
# --- CONSTANTES GLOBALES ---
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd

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

<<<<<<< HEAD
# --- CONSTANTES DE FUENTES ---
=======
# --- NUEVAS CONSTANTES DE FUENTE ---
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
FONT_NAME_REGULAR = "RobotoMono-Regular"
FONT_NAME_BOLD = "RobotoMono-Bold"
TTF_FILE_REGULAR = "RobotoMono-Regular.ttf"
TTF_FILE_BOLD = "RobotoMono-Bold.ttf"

<<<<<<< HEAD
# --- CONSTANTES DE DISEÑO Y LAYOUT ---
PADDING = 0.75 * inch              # Margen de seguridad interior (sangrado)
TITLE_FONT_SIZE = 18               # Tamaño fijo para el título de la sopa
WORDS_FONT_SIZE = 11               # Tamaño fijo para la lista de palabras
SPACE_AFTER_TITLE = 0.2 * inch     # Espacio entre título y lista de palabras
INTERLINEADO_PALABRAS = 1.3        # Multiplicador de altura de línea
PROPORCION_LETRA_CUADRICULA = 0.7   # Letra ocupa el 70% de la celda
NUM_COLUMNAS_PALABRAS = 4          # Columnas para la lista de palabras
SOLUCIONES_POR_PAGINA = 4          # Cuadrícula de 2x2 en páginas de solución
LETRAS_ALFABETO = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

# Constantes para el recuadro de palabras
BOX_PADDING = 0.15 * inch          # Relleno interno del recuadro
BOX_CORNER_RADIUS = 0.1 * inch     # Radio de las esquinas redondeadas

# --- FUNCIONES DE LÓGICA ---

def cm_to_pt(cm):
    """
    Convierte una medida de centímetros (cm) a puntos (pt) de ReportLab.
    
    Args:
        cm (float): La medida en centímetros.
        
    Returns:
        float: La medida equivalente en puntos.
    """
=======
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
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    return cm * 28.3465

def crear_sopa_letras(palabras, dimension=20):
    """
    Genera una sopa de letras (cuadrícula) e intenta colocar todas las palabras.
    
    *** Optimización Clave ***:
    Esta función ordena las palabras de más larga a más corta antes de
    intentar colocarlas. Esto aumenta drásticamente la probabilidad de que
    todas las palabras quepan, ya que las palabras más difíciles (largas)
    encuentran espacio primero.
    
    Args:
        palabras (list[str]): Lista de palabras a incluir en la sopa.
        dimension (int): Tamaño de la cuadrícula (dimension x dimension).
        
    Returns:
        tuple: Contiene:
            - sopa (list[list[str]]): La cuadrícula 2D completa con letras.
            - ubicaciones (dict): Diccionario {palabra: (fila, col, dx, dy)}
                                  con la ubicación de las palabras colocadas.
            - palabras_no_colocadas (list[str]): Lista de palabras que
                                                  no se pudieron colocar.
    """
    sopa = [[' ' for _ in range(dimension)] for _ in range(dimension)]
    ubicaciones = {}
    palabras_no_colocadas = []
    
<<<<<<< HEAD
    # 8 direcciones: horizontal, vertical y diagonales
    direcciones = [(1,0), (0,1), (1,1), (-1,1), (-1,0), (0,-1), (-1,-1), (1,-1)]
    
    def chequear_fit(palabra, fila, col, dy, dx):
        """Helper interno: Verifica si una palabra cabe en una ubicación."""
        for i in range(len(palabra)):
            fila_actual, col_actual = fila + i * dy, col + i * dx
            # 1. Comprobar límites de la cuadrícula
            if not (0 <= fila_actual < dimension and 0 <= col_actual < dimension):
                return False
            # 2. Comprobar colisiones con otras letras
=======
    direcciones = [(1,0), (0,1), (1,1), (-1,1), (-1,0), (0,-1), (-1,-1), (1,-1)]
    
    def chequear_fit(palabra, fila, col, dy, dx):
        for i in range(len(palabra)):
            fila_actual, col_actual = fila + i * dy, col + i * dx
            if not (0 <= fila_actual < dimension and 0 <= col_actual < dimension):
                return False
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
            celda = sopa[fila_actual][col_actual]
            if celda != ' ' and celda != palabra[i]:
                return False
        return True

    def colocar_palabra(palabra):
<<<<<<< HEAD
        """Helper interno: Intenta colocar una palabra en la cuadrícula."""
        palabra_upper = palabra.upper()
        posibles_ubicaciones = []
        
        # Buscar todas las ubicaciones válidas
=======
        palabra_upper = palabra.upper()
        posibles_ubicaciones = []
        
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        for dy, dx in direcciones:
            for fila_inicio in range(dimension):
                for col_inicio in range(dimension):
                    if chequear_fit(palabra_upper, fila_inicio, col_inicio, dy, dx):
                        posibles_ubicaciones.append((fila_inicio, col_inicio, dy, dx))
        
        if posibles_ubicaciones:
<<<<<<< HEAD
            # Si hay ubicaciones, elegir una al azar
            random.shuffle(posibles_ubicaciones)
            fila_inicio, col_inicio, dy, dx = posibles_ubicaciones[0]
            
            # Colocar la palabra en la cuadrícula
=======
            random.shuffle(posibles_ubicaciones)
            fila_inicio, col_inicio, dy, dx = posibles_ubicaciones[0]
            
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
            for i in range(len(palabra_upper)):
                fila_actual = fila_inicio + i * dy
                col_actual = col_inicio + i * dx
                sopa[fila_actual][col_actual] = palabra_upper[i]
            
<<<<<<< HEAD
            # Devolver la ubicación para las soluciones
            return (fila_inicio, col_inicio, dx, dy)
        
        return None # No se pudo colocar

    # --- ¡OPTIMIZACIÓN! ---
    # Ordenar palabras de más larga a más corta.
    palabras_ordenadas = sorted(palabras, key=len, reverse=True)

    for palabra in palabras_ordenadas:
=======
            return (fila_inicio, col_inicio, dx, dy)
        
        return None

    for palabra in palabras:
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        ubicacion = colocar_palabra(palabra)
        if ubicacion:
            ubicaciones[palabra.upper()] = ubicacion
        else:
            palabras_no_colocadas.append(palabra)
            
<<<<<<< HEAD
    # Rellenar los espacios vacíos con letras aleatorias
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    for i in range(dimension):
        for j in range(dimension):
            if sopa[i][j] == ' ':
                sopa[i][j] = random.choice(LETRAS_ALFABETO)
                
    return sopa, ubicaciones, palabras_no_colocadas

def procesar_excel(excel_file, words_per_puzzle):
<<<<<<< HEAD
    """
    Lee un archivo Excel y lo divide en múltiples listas de palabras.
    
    Formato esperado del Excel (en la primera columna):
    - "Palabra1, TemaA"
    - "Palabra2"
    - "Palabra3"
    ...
    - "PalabraN, TemaB"
    - "PalabraN+1"
    
    El tema se extrae de la primera palabra que lo define.
    Agrupa las palabras en listas del tamaño 'words_per_puzzle'.
    
    Args:
        excel_file (UploadedFile): El archivo Excel subido desde Streamlit.
        words_per_puzzle (int): El número de palabras para cada sopa.
        
    Returns:
        tuple: Contiene:
            - word_lists (list[list[str]]): Lista de listas de palabras.
            - themes (list[str]): Lista de temas correspondientes.
    """
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    df = pd.read_excel(excel_file)
    word_lists = []
    themes = []
    current_words = []
    current_theme = ""

    for index, row in df.iterrows():
        cell_value = str(row.iloc[0]).strip()
        
<<<<<<< HEAD
        # Si la celda contiene una coma, separamos palabra y tema
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        if ',' in cell_value:
            word, theme = cell_value.split(',', 1)
            word, theme = word.strip(), theme.strip()
            
            if word and word.lower() != 'nan':
                current_words.append(word)
<<<<<<< HEAD
            # El tema solo se asigna si no tenemos ya uno
            if theme and not current_theme:
                current_theme = theme
        else:
            # Si no hay coma, es solo una palabra
            if cell_value and cell_value.lower() != 'nan':
                current_words.append(cell_value)
        
        # Si alcanzamos el límite de palabras o es la última fila...
        if len(current_words) == words_per_puzzle or index == len(df) - 1:
            if current_words:
                word_lists.append(current_words.copy())
                # Asignar tema o un título genérico
                themes.append(current_theme if current_theme else f"Sopa de letras {len(themes) + 1}")
                # Resetear para la siguiente sopa
=======
            if theme and not current_theme:
                current_theme = theme
        else:
            if cell_value and cell_value.lower() != 'nan':
                current_words.append(cell_value)
        
        if len(current_words) == words_per_puzzle or index == len(df) - 1:
            if current_words:
                word_lists.append(current_words.copy())
                themes.append(current_theme if current_theme else f"Sopa de letras {len(themes) + 1}")
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
                current_words.clear()
                current_theme = ""
                
    return word_lists, themes

<<<<<<< HEAD
# --- FUNCIONES DE DIBUJO PDF (REPORTLAB) ---

def dibujar_pagina_sopa(c, sopa, palabras, theme, config):
    """
    Dibuja una (1) página completa de sopa de letras en el canvas de ReportLab.
    
    Organiza la página en tres secciones (de arriba a abajo):
    1. Título
    2. Recuadro con la lista de palabras
    3. Cuadrícula de la sopa
    
    Args:
        c (canvas.Canvas): El canvas de ReportLab sobre el que dibujar.
        sopa (list[list[str]]): La cuadrícula 2D de letras.
        palabras (list[str]): La lista de palabras a mostrar.
        theme (str): El título de la sopa de letras.
        config (dict): Diccionario con 'page_size' y 'dimension'.
    """
=======
# --- ESTA ES LA FUNCIÓN 'dibujar_pagina_sopa' (CORREGIDA Y ÚNICA) ---
# Resuelve el problema del centrado vertical y el padding superior/inferior.

def dibujar_pagina_sopa(c, sopa, palabras, theme, config):
    """Dibuja una (1) página de sopa de letras en el canvas."""
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    
    page_width, page_height = config['page_size']
    dimension = config['dimension']
    
<<<<<<< HEAD
    # Área usable (descontando márgenes/padding)
    ancho_usable = page_width - 2 * PADDING
    # 'y_cursor' rastrea nuestra posición vertical, empezando desde arriba
    y_cursor = page_height - PADDING

    # --- 1. Dibujar Título (en negrita) ---
    c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
    # Dibuja el título en el margen izquierdo, por debajo del padding superior
    c.drawString(PADDING, y_cursor - TITLE_FONT_SIZE, theme[:60])
    # Bajar el cursor
    y_cursor -= (TITLE_FONT_SIZE + SPACE_AFTER_TITLE)

    # --- 2. Dibujar Recuadro y Lista de palabras ---
    c.setFont(FONT_NAME_REGULAR, WORDS_FONT_SIZE)
    line_height = WORDS_FONT_SIZE * INTERLINEADO_PALABRAS
    
    # --- Lógica de Columnas y Centrado Vertical ---
    n_palabras = len(palabras)
    n_cols = NUM_COLUMNAS_PALABRAS
    
    # Calcular cuántas filas tendrá la columna más alta (equivale a math.ceil)
    max_words_per_col = (n_palabras + n_cols - 1) // n_cols
    
    # Contar cuántas palabras reales hay en CADA columna
    col_word_counts = [0] * n_cols
    for i in range(n_palabras):
=======
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
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        col_index = i // max_words_per_col
        if col_index < n_cols:
            col_word_counts[col_index] += 1
            
<<<<<<< HEAD
    # Estimación de la altura de la fuente (para alinear texto verticalmente)
    font_ascent_guess = WORDS_FONT_SIZE * 0.8

    # Calcular altura total del recuadro
    word_list_height_inner = max_words_per_col * line_height
    word_list_height_total = word_list_height_inner + (2 * BOX_PADDING)

    # Coordenadas Y del recuadro
    y_box_bottom = y_cursor - word_list_height_total
    
    # Dibujar el recuadro redondeado
=======
    # 2. Estimamos la altura de la fuente para el padding superior
    font_ascent_guess = WORDS_FONT_SIZE * 0.8
    # --- FIN NUEVA LÓGICA ---

    # Calcular altura total del recuadro (basada en la columna más alta)
    word_list_height_inner = max_words_per_col * line_height
    word_list_height_total = word_list_height_inner + (2 * BOX_PADDING)

    # Coordenadas del recuadro
    y_box_bottom = y_cursor - word_list_height_total
    
    # --- DIBUJAR EL RECUADRO ---
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    c.setLineWidth(1)
    c.setStrokeColor(black)
    c.roundRect(
        PADDING,                  # x
        y_box_bottom,             # y
        ancho_usable,             # width
        word_list_height_total,   # height
        BOX_CORNER_RADIUS,        # radius
<<<<<<< HEAD
        stroke=1, fill=0
    )

    # Ancho de cada columna de palabras
    col_width = (ancho_usable - (2 * BOX_PADDING)) / n_cols
    
    # --- Bucle para dibujar cada palabra ---
=======
        stroke=1, 
        fill=0
    )

    # Ancho de columna
    col_width = (ancho_usable - (2 * BOX_PADDING)) / n_cols
    
    # --- BUCLE DE DIBUJO (ESTABLE) ---
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    for i, word in enumerate(palabras):
        col_index = i // max_words_per_col
        row_index = i % max_words_per_col
        
<<<<<<< HEAD
        # --- Lógica de Centrado Vertical ---
        # 1. Total de palabras en esta columna específica
        words_in_this_col = col_word_counts[col_index]
        
        # 2. Offset para centrar esta columna (si es más corta que la más alta)
        this_col_height_inner = words_in_this_col * line_height
        vertical_offset = (word_list_height_inner - this_col_height_inner) / 2
        
        # 3. Posición X (columna)
        cx = PADDING + BOX_PADDING + (col_index * col_width)
        
        # 4. Posición Y (fila, ajustada por el centrado)
        y_top_of_centered_block = (y_cursor - BOX_PADDING - vertical_offset)
        cy_base = y_top_of_centered_block - font_ascent_guess
        cy = cy_base - (row_index * line_height)

        c.drawString(cx, cy, word.upper()[:18]) # Limitar longitud de palabra

    # Bajar el cursor hasta debajo del recuadro de palabras
    y_cursor -= (word_list_height_total + PADDING)

    # --- 3. Dibujar la Cuadrícula (en el espacio restante) ---
    grid_available_height = y_cursor - PADDING
    grid_available_width = ancho_usable
    
    # El tamaño de celda se ajusta al espacio disponible (el menor entre ancho y alto)
=======
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
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    cell_size = min(grid_available_width / dimension, grid_available_height / dimension)
    
    grid_width = cell_size * dimension
    grid_height = cell_size * dimension
    
<<<<<<< HEAD
    # Centrar la cuadrícula en el espacio disponible
    x_grid = PADDING + (grid_available_width - grid_width) / 2
    y_grid = PADDING + (grid_available_height - grid_height) / 2
    
    # Dibujar el borde exterior de la cuadrícula
=======
    x_grid = PADDING + (grid_available_width - grid_width) / 2
    y_grid = PADDING + (grid_available_height - grid_height) / 2
    
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    c.setLineWidth(1)
    c.setStrokeColor(black)
    c.rect(x_grid, y_grid, grid_width, grid_height, stroke=1, fill=0)
    
<<<<<<< HEAD
    # Tamaño de fuente proporcional al tamaño de la celda
    grid_font_size = int(cell_size * PROPORCION_LETRA_CUADRICULA)
    c.setFont(FONT_NAME_REGULAR, grid_font_size)
    
    # Dibujar cada letra en la cuadrícula
    for i in range(dimension):
        for j in range(dimension):
            letra = sopa[i][j].upper()
            
            # Calcular el centro de la celda
            x_letra = x_grid + j * cell_size + cell_size / 2
            y_letra = y_grid + grid_height - (i + 0.5) * cell_size
            
            # Ajuste vertical fino para centrar la letra (las fuentes no se centran en su 'y' base)
            y_letra_ajustada = y_letra - (grid_font_size * 0.3) 
            c.drawCentredString(x_letra, y_letra_ajustada, letra)

def dibujar_paginas_soluciones(c, sopas_list, palabras_list, ubicaciones_list, themes, config):
    """
    Dibuja todas las páginas de soluciones, colocando varias (4) por página.
    
    Args:
        c (canvas.Canvas): El canvas de ReportLab.
        sopas_list (list): Lista de todas las cuadrículas generadas.
        palabras_list (list): Lista de todas las listas de palabras.
        ubicaciones_list (list): Lista de todos los diccionarios de ubicaciones.
        themes (list): Lista de todos los temas.
        config (dict): Diccionario con 'page_size' y 'dimension'.
    """
=======
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
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    
    page_width, page_height = config['page_size']
    dimension = config['dimension']

    if not sopas_list:
<<<<<<< HEAD
        return # No hay nada que dibujar

    # Iniciar la primera página de soluciones
=======
        return

>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    c.showPage()
    pagina_solucion_actual = 1
    
    ancho_usable = page_width - 2 * PADDING
    alto_usable = page_height - 2 * PADDING
    
<<<<<<< HEAD
    # Título principal de la sección de soluciones
    c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
    c.drawString(PADDING, page_height - PADDING - TITLE_FONT_SIZE, "Soluciones")

    # Calcular dimensiones para las mini-sopas (layout 2x2)
    sol_area_width = ancho_usable / 2
    sol_area_height = (alto_usable - (TITLE_FONT_SIZE + PADDING)) / 2
    
    # Calcular tamaño de celda para las mini-cuadrículas
=======
    # Título inicial (en negrita)
    c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
    c.drawString(PADDING, page_height - PADDING - TITLE_FONT_SIZE, "Soluciones")

    sol_area_width = ancho_usable / 2
    sol_area_height = (alto_usable - (TITLE_FONT_SIZE + PADDING)) / 2
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    cell_size_sol = min(sol_area_width / dimension, sol_area_height / dimension) * 0.85
    mini_grid_w = cell_size_sol * dimension
    mini_grid_h = cell_size_sol * dimension

    for sol_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        
<<<<<<< HEAD
        # Posición dentro de la página actual (0, 1, 2, o 3)
        pos_en_pagina = sol_index % SOLUCIONES_POR_PAGINA
        
        # Si es la primera solución (índice 0) de una NUEVA página...
        if sol_index > 0 and pos_en_pagina == 0:
            c.showPage() # Crear nueva página
            pagina_solucion_actual += 1
            # Añadir título a la nueva página
            c.setFont(FONT_NAME_BOLD, TITLE_FONT_SIZE)
            c.drawString(PADDING, page_height - PADDING - TITLE_FONT_SIZE, f"Soluciones (pág. {pagina_solucion_actual})")

        # Calcular fila (0 o 1) y columna (0 o 1) en la cuadrícula 2x2
        col_sol = pos_en_pagina % 2
        row_sol = pos_en_pagina // 2
        
        # Calcular coordenadas X e Y para esta mini-sopa
        # Centrar la mini-cuadrícula dentro de su cuadrante
        x_offset = PADDING + col_sol * sol_area_width
        left_sol = x_offset + (sol_area_width - mini_grid_w) / 2
        
        # El cálculo de Y es más complejo porque ReportLab empieza desde abajo
        y_area_base = PADDING + (1 - row_sol) * sol_area_height
        bot_sol = y_area_base + (sol_area_height - mini_grid_h) / 2
        
        # Ajuste especial para la fila superior (row_sol == 0) para
        # que quede debajo del título "Soluciones"
        if row_sol == 0:
             bot_sol = (page_height - PADDING - TITLE_FONT_SIZE - PADDING - sol_area_height) + (sol_area_height - mini_grid_h) / 2

        # --- Encontrar todas las celdas que son parte de una solución ---
=======
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

>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        celdas_solucion = set()
        for p in palabras:
            ubicacion = ubicaciones.get(p.upper())
            if ubicacion:
                fila_ini, col_ini, dx, dy = ubicacion
                for k in range(len(p)):
                    celdas_solucion.add((fila_ini + k*dy, col_ini + k*dx))

<<<<<<< HEAD
        # --- Dibujar la mini-cuadrícula ---
=======
        # Cuadrícula y letras (en regular)
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        c.setFont(FONT_NAME_REGULAR, int(cell_size_sol * 0.7))
        c.setLineWidth(0.5)
        for i in range(dimension):
            for j in range(dimension):
                x = left_sol + j * cell_size_sol
<<<<<<< HEAD
                # 'y' se calcula desde abajo hacia arriba
                y = bot_sol + mini_grid_h - (i + 1) * cell_size_sol 
                
                # Dibujar la celda
                c.rect(x, y, cell_size_sol, cell_size_sol, stroke=1, fill=0)
                
                # Si la celda es una solución, rellenarla
=======
                y = bot_sol + mini_grid_h - (i + 1) * cell_size_sol
                c.rect(x, y, cell_size_sol, cell_size_sol, stroke=1, fill=0)
                
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
                if (i, j) in celdas_solucion:
                    c.setFillColor(black)
                    c.drawCentredString(x + cell_size_sol / 2, y + (cell_size_sol * 0.2), sopa[i][j].upper())

<<<<<<< HEAD
        # Dibujar el título de la mini-sopa (debajo de ella)
=======
        # Título de la mini-sopa (en regular)
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        c.setFont(FONT_NAME_REGULAR, int(cell_size_sol * 0.6))
        c.setFillColor(black)
        c.drawCentredString(left_sol + mini_grid_w / 2, bot_sol - 14, theme)

<<<<<<< HEAD
def exportar_pdf(sopas_list, palabras_list, ubicaciones_list, themes, dimension, page_size):
    """
    Crea el archivo PDF completo en memoria.
    
    Esta función inicializa el canvas, registra las fuentes y luego llama
    a las funciones de dibujo para las páginas de sopas y las de soluciones.
    
    Args:
        sopas_list (list): Lista de todas las cuadrículas.
        palabras_list (list): Lista de todas las listas de palabras.
        ubicaciones_list (list): Lista de todos los diccionarios de ubicaciones.
        themes (list): Lista de todos los temas.
        dimension (int): Tamaño de la cuadrícula.
        page_size (tuple): (ancho, alto) en puntos para el PDF.
        
    Returns:
        BytesIO: Un buffer en memoria que contiene el archivo PDF completo,
                 listo para ser descargado. O None si falla el registro de fuentes.
    """
    buffer = BytesIO()
    
    # --- Registro de Fuentes ---
    # Es crucial registrar las fuentes ANTES de usarlas.
=======
# --- FUNCIÓN 'exportar_pdf' (ÚNICA) ---

def exportar_pdf(sopas_list, palabras_list, ubicaciones_list, themes, dimension, page_size):
    """
    Crea el archivo PDF completo llamando a las funciones de dibujo.
    """
    buffer = BytesIO()
    
    # Registramos las fuentes aquí.
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    try:
        pdfmetrics.registerFont(TTFont(FONT_NAME_REGULAR, TTF_FILE_REGULAR))
        pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, TTF_FILE_BOLD))
    except Exception as e:
        st.error(f"¡ERROR DE FUENTE! No se pudieron cargar los archivos .ttf.")
        st.error(f"Asegúrate de que 'RobotoMono-Regular.ttf' y 'RobotoMono-Bold.ttf' estén en la misma carpeta que el script.")
        st.error(f"Detalle del error: {e}")
        return None # Devuelve None para que Streamlit sepa que falló

<<<<<<< HEAD
    # Crear el canvas (el "lienzo" del PDF)
    c = canvas.Canvas(buffer, pagesize=page_size)
    
    # Crear un diccionario de configuración para pasarlo fácilmente
=======
    c = canvas.Canvas(buffer, pagesize=page_size)
    
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    config = {
        "page_size": page_size,
        "dimension": dimension
    }

<<<<<<< HEAD
    # --- 1. Dibujar páginas de sopas ---
    for puzzle_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        if puzzle_index > 0:
            c.showPage() # Crear una nueva página para cada sopa
        dibujar_pagina_sopa(c, sopa, palabras, theme, config)

    # --- 2. Dibujar páginas de soluciones ---
    # Esta función se encarga internamente de crear sus propias páginas
    dibujar_paginas_soluciones(c, sopas_list, palabras_list, ubicaciones_list, themes, config)

    # Guardar y finalizar el PDF
=======
    # Dibujar páginas de sopas
    for puzzle_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        if puzzle_index > 0:
            c.showPage()
        dibujar_pagina_sopa(c, sopa, palabras, theme, config)

    # Dibujar páginas de soluciones
    dibujar_paginas_soluciones(c, sopas_list, palabras_list, ubicaciones_list, themes, config)

>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
    c.save()
    buffer.seek(0)
    return buffer


# --- INTERFAZ DE STREAMLIT (UI) ---
<<<<<<< HEAD
# Esta sección controla la página web interactiva.
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd

st.title("Generador de Sopas de Letras KDP")
st.write("¡Bienvenido! Sube tu archivo Excel de palabras para empezar.")

<<<<<<< HEAD
# --- Controles de la Barra Lateral ---
# st.sidebar.header("Configuración") # Descomenta si prefieres los controles en la barra lateral

# 1. Carga de archivo
excel_file = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])

# 2. Configuración de la Sopa
dimension = st.number_input("Tamaño de cuadrícula (ej. 20x20)", min_value=10, max_value=30, value=20)
words_per_puzzle = st.number_input("Número de palabras por sopa", min_value=5, max_value=40, value=20)

# 3. Configuración del PDF
nombre_tamaño = st.selectbox("Tamaño de página KDP", list(TAMAÑOS_KDP.keys()), index=12) # '8.5" x 8.5"' por defecto

# Convertir el tamaño seleccionado a puntos
=======
excel_file = st.file_uploader("Selecciona tu archivo Excel", type=["xlsx"])
dimension = st.number_input("Tamaño de cuadrícula (ej. 20x20)", min_value=10, max_value=30, value=20)
words_per_puzzle = st.number_input("Número de palabras por sopa", min_value=5, max_value=40, value=20)
nombre_tamaño = st.selectbox("Tamaño de página KDP", list(TAMAÑOS_KDP.keys()), index=12)

>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
ancho_cm, alto_cm = TAMAÑOS_KDP[nombre_tamaño]
page_width = cm_to_pt(ancho_cm)
page_height = cm_to_pt(alto_cm)

<<<<<<< HEAD
# --- Lógica de Generación (Botón Principal) ---
if st.button("Generar PDF de Sopas de Letras"):
    if excel_file is not None:
        
        # 1. Procesar el Excel
=======
if st.button("Generar PDF de Sopas de Letras"):
    if excel_file is not None:
        
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
        word_lists, themes = procesar_excel(excel_file, words_per_puzzle)
        
        if not word_lists:
            st.error("No se encontraron palabras en el archivo Excel. Revisa el formato.")
        else:
            st.info(f"Se van a generar {len(word_lists)} sopas de letras.")
        
<<<<<<< HEAD
            # Listas para almacenar los resultados de cada sopa
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
            sopas_list = []
            palabras_list_filled = []
            ubicaciones_list = []
            
<<<<<<< HEAD
            # 2. Generar cada Sopa (Lógica)
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
            with st.spinner("Generando sopas..."):
                for words, theme in zip(word_lists, themes):
                    sopa, ubicaciones, no_colocadas = crear_sopa_letras(words, dimension=dimension)
                    
<<<<<<< HEAD
                    # Informar al usuario si algunas palabras no cupieron
=======
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
                    if no_colocadas:
                        st.warning(f"En la sopa '{theme}', no se pudieron colocar las siguientes palabras: {', '.join(no_colocadas)}")
                    
                    sopas_list.append(sopa)
<<<<<<< HEAD
                    palabras_list_filled.append(words) # Guardamos la lista original de palabras
                    ubicaciones_list.append(ubicaciones)

            # 3. Dibujar el PDF
            with st.spinner("Creando archivo PDF..."):
                buffer = exportar_pdf(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension, (page_width, page_height))
            
            # 4. Ofrecer la descarga
            if buffer: # Si el buffer se creó correctamente
=======
                    palabras_list_filled.append(words)
                    ubicaciones_list.append(ubicaciones)

            with st.spinner("Creando archivo PDF..."):
                buffer = exportar_pdf(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension, (page_width, page_height))
            
            if buffer: # Solo mostrar el botón de descarga si el buffer se creó
>>>>>>> e25d8800b1ff2e6a18f59f7b9bb566455847cffd
                st.success("¡PDF generado! Descárgalo aquí:")
                st.download_button(
                    label="Descargar PDF",
                    data=buffer,
                    file_name='sopas_de_letras_kdp.pdf',
                    mime='application/pdf'
                )
    else:
        st.error("Por favor, sube un archivo Excel.")