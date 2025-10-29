import streamlit as st
import pandas as pd
import random
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.title("[translate:Generador de sopas de letras KDP]")
st.write("[translate:¡Bienvenido! Sube tu archivo Excel de palabras para empezar.]")

excel_file = st.file_uploader("[translate:Selecciona tu archivo Excel]", type=["xlsx"])
dimension = st.number_input("[translate:Tamaño de cuadrícula]", min_value=10, max_value=30, value=20)

def crear_sopa_letras(palabras, dimension=20):
    sopa = [[' ' for _ in range(dimension)] for _ in range(dimension)]
    ubicaciones = {}
    def colocar_palabra(sopa, palabra):
        direcciones = [(1,0), (0,1), (1,1), (-1,1), (-1,0), (0,-1), (-1,-1), (1,-1)]
        random.shuffle(direcciones)
        for dy, dx in direcciones:
            for _ in range(dimension * dimension * 2):
                fila_inicio = random.randint(0, dimension - 1)
                col_inicio = random.randint(0, dimension - 1)
                fit = True
                for i in range(len(palabra)):
                    fila_actual = fila_inicio + i * dy
                    col_actual = col_inicio + i * dx
                    if not (0 <= fila_actual < dimension and 0 <= col_actual < dimension):
                        fit = False
                        break
                    if sopa[fila_actual][col_actual] != ' ' and sopa[fila_actual][col_actual] != palabra[i]:
                        fit = False
                        break
                if fit:
                    for i in range(len(palabra)):
                        fila_actual = fila_inicio + i * dy
                        col_actual = col_inicio + i * dx
                        sopa[fila_actual][col_actual] = palabra[i]
                    return (fila_inicio, col_inicio, dx, dy)
        return None
    for palabra in palabras:
        ubicacion = colocar_palabra(sopa, palabra.upper())
        if ubicacion:
            ubicaciones[palabra.upper()] = ubicacion
    letras_aleatorias = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    for i in range(dimension):
        for j in range(dimension):
            if sopa[i][j] == ' ':
                sopa[i][j] = random.choice(letras_aleatorias)
    return sopa, ubicaciones

def exportar_pdf(sopas_list, palabras_list, ubicaciones_list, themes, dimension, filename='sopas_de_letras.pdf'):
    buffer = BytesIO()
    try:
        pdfmetrics.registerFont(TTFont('Aleo', 'Aleo-VariableFont_wght.ttf'))
        aleo_font = "Aleo"
    except:
        aleo_font = "Helvetica"
    c = canvas.Canvas(buffer, pagesize=A4)
    left_margin = 0.5 * inch
    right_margin = A4[0] - 0.5 * inch
    top_margin = A4[1] - 0.75 * inch
    line_height_words = 14
    num_col_words = 4
    words_per_column = 5
    col_width = (right_margin - left_margin) / num_col_words
    cell_size_grid = min((right_margin - left_margin) / dimension, 18)  # máximo tamaño visual estable
    grid_width = dimension * cell_size_grid
    grid_height = dimension * cell_size_grid
    left_grid = left_margin
    top_grid = top_margin - 30 - words_per_column * num_col_words * line_height_words - 40
    for puzzle_index, (sopa, palabras, ubicaciones, theme) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
        if puzzle_index > 0:
            c.showPage()
        c.setFont(aleo_font, 15)
        c.drawString(left_margin, top_margin - 20, theme)
        c.setFont(aleo_font, 9)
        yp_start_words = top_margin - 40
        for i, word in enumerate(palabras):
            col_index = i // words_per_column
            row_index = i % words_per_column
            c.drawString(left_margin + col_index * col_width, yp_start_words - row_index * line_height_words, word.upper())
        c.setLineWidth(1)
        c.setStrokeColor(black)
        c.rect(left_grid, top_grid - grid_height, grid_width, grid_height, stroke=1, fill=0)
        c.setFont(aleo_font, 12)
        for i in range(dimension):
            for j in range(dimension):
                letra = sopa[i][j].upper()
                c.drawCentredString(left_grid + j * cell_size_grid + cell_size_grid / 2,
                                  top_grid - (i + 1) * cell_size_grid + cell_size_grid / 2 - 4, letra)
    # Soluciones
    if sopas_list:
        c.showPage()
        c.setFont(aleo_font, 14)
        c.drawString(left_margin, top_margin - 20, "[translate:Soluciones]")
        solutions_per_page = 4
        sol_area_width = (right_margin - left_margin) / 2
        sol_area_height = (top_margin - left_margin) / 2
        for sol_index, (sopa_sol, palabras_sol, ubicaciones_sol, theme_sol) in enumerate(zip(sopas_list, palabras_list, ubicaciones_list, themes)):
            if sol_index > 0 and sol_index % solutions_per_page == 0:
                c.showPage()
                c.setFont(aleo_font, 14)
                c.drawString(left_margin, top_margin - 20, "[translate:Soluciones (cont.)]")
            cell_size_sol = min(sol_area_width / dimension, sol_area_height / dimension) * 0.9
            mini_grid_width = dimension * cell_size_sol
            mini_grid_height = dimension * cell_size_sol
            col_sol = (sol_index % solutions_per_page) % 2
            row_sol = (sol_index % solutions_per_page) // 2
            left_sol = left_margin + col_sol * sol_area_width + (sol_area_width - mini_grid_width) / 2
            top_sol_draw = top_margin - 20 - row_sol * sol_area_height - (sol_area_height - mini_grid_height) / 2
            c.setFont(aleo_font, 8)
            c.setLineWidth(0.5)
            for i in range(dimension):
                for j in range(dimension):
                    x = left_sol + j * cell_size_sol
                    y = top_sol_draw - (i + 1) * cell_size_sol
                    c.rect(x, y, cell_size_sol, cell_size_sol, stroke=1, fill=0)
                    is_word_letter = any(
                        (i == u[0] + k * u[3] and j == u[1] + k * u[2])
                        for p in palabras_sol
                        if (u := ubicaciones_sol.get(p.upper()))
                        for k in range(len(p))
                    )
                    if is_word_letter:
                        c.setFillColor(black)
                        c.drawCentredString(x + cell_size_sol / 2, y + cell_size_sol / 2 - 2, sopa_sol[i][j].upper())
            c.setFont(aleo_font, 10)
            c.setFillColor(black)
            c.drawString(left_sol, top_sol_draw - mini_grid_height - 15, theme_sol)
    c.save()
    buffer.seek(0)
    return buffer

if st.button("[translate:Procesar palabras]"):
    if excel_file is not None:
        df = pd.read_excel(excel_file)
        word_lists = []
        themes = []
        current_words = []
        current_theme = ""
        words_per_puzzle = 20
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
                    themes.append(current_theme if current_theme else f"[translate:Sopa de letras] {len(themes) + 1}")
                current_words.clear()
                current_theme = ""
        sopas_list = []
        palabras_list_filled = []
        ubicaciones_list = []
        for words, theme in zip(word_lists, themes):
            sopa, ubicaciones = crear_sopa_letras(words, dimension=dimension)
            sopas_list.append(sopa)
            palabras_list_filled.append(words)
            ubicaciones_list.append(ubicaciones)
        buffer = exportar_pdf(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension)
        st.success("[translate:¡PDF generado con puzles y soluciones! Descárgalo aquí:]")
        st.download_button(
            label="[translate:Descargar PDF]",
            data=buffer,
            file_name='sopas_de_letras.pdf',
            mime='application/pdf'
        )
    else:
        st.error("[translate:Por favor, sube un archivo Excel.]")
