# app_es.py
import streamlit as st
import core_logic  # Importa el "cerebro"
from ui_text import TEXT_ES as T  # Importa los textos en ESPAÑOL
from core_logic import TAMAÑOS_KDP, cm_to_pt

# --- Inicializar el Estado de Sesión ---
# Esto es crucial para "recordar" que el usuario está validado
if 'license_valid' not in st.session_state:
    st.session_state['license_valid'] = False

# --- Función para mostrar el formulario de licencia ---
def show_license_form():
    st.title(T['title'])
    st.write(T['welcome'])
    st.warning("Necesitas una clave de licencia válida para usar esta aplicación.")
    
    with st.form("license_form"):
        license_key = st.text_input("Ingresa tu clave de licencia de Lemon Squeezy:", type="password")
        submitted = st.form_submit_button("Activar Licencia")
        
        if submitted:
            if not license_key:
                st.error("Por favor, introduce una clave.")
            else:
                with st.spinner("Validando clave..."):
                    # Llamamos a la nueva función de core_logic
                    is_valid = core_logic.check_license_key(license_key)
                    
                    if is_valid:
                        # ¡ÉXITO! Marcamos el estado de sesión como True
                        st.session_state['license_valid'] = True
                        # Volvemos a ejecutar la app para mostrar la app principal
                        st.experimental_rerun()
                    else:
                        # El error ya se muestra dentro de check_license_key
                        pass

# --- Función para mostrar la aplicación principal (Tu código existente) ---
def show_main_app():
    st.title(T['title'])
    
    # 1. Carga de archivo
    excel_file = st.file_uploader(T['uploader_label'], type=["xlsx"])

    # 2. Configuración de la Sopa
    dimension = st.number_input(T['grid_size'], min_value=10, max_value=30, value=20)
    words_per_puzzle = st.number_input(T['words_per_puzzle'], min_value=5, max_value=40, value=20)

    # 3. Nivel de Dificultad
    dificultad = st.selectbox(
        T['difficulty'],
        T['difficulty_options'],
        index=2  # 'Hard' por defecto
    )

    # 4. Configuración del PDF
    nombre_tamaño = st.selectbox(T['kdp_size'], list(TAMAÑOS_KDP.keys()), index=12) # '8.5" x 8.5"' por defecto
    formato_salida = st.radio(T['output_format'], ("PDF", "PPTX"))

    # Convertir el tamaño seleccionado a puntos
    ancho_cm, alto_cm = TAMAÑOS_KDP[nombre_tamaño]
    page_width = cm_to_pt(ancho_cm)
    page_height = cm_to_pt(alto_cm)

    # --- Lógica de Generación (Botón Principal) ---
    if st.button(T['generate_button']):
        if excel_file is not None:
            
            try:
                # 1. Procesar el Excel
                word_lists, themes = core_logic.procesar_excel(excel_file, words_per_puzzle)
            except Exception as e:
                st.error(T['error_processing'])
                st.error(f"Detalle: {e}") # Mostramos el error de procesamiento de Excel
                return # Detenemos la ejecución aquí

            if not word_lists:
                st.error(T['error_processing'])
            else:
                st.info(f"Generando {len(word_lists)} sopas (dificultad: {dificultad}).")
                
                sopas_list = []
                palabras_list_filled = []
                ubicaciones_list = []
                
                # 2. Generar cada Sopa (Lógica)
                with st.spinner(T['spinner_generating']):
                    for words, theme in zip(word_lists, themes):
                        sopa, ubicaciones, no_colocadas = core_logic.crear_sopa_letras(
                            words, 
                            dimension=dimension, 
                            dificultad=dificultad
                        )
                        
                        if no_colocadas:
                            st.warning(T['warning_words_not_placed'].format(theme=theme, words=', '.join(no_colocadas)))
                        
                        sopas_list.append(sopa)
                        palabras_list_filled.append(words)
                        ubicaciones_list.append(ubicaciones)

                # 3. Generar el archivo de salida (PDF o PPTX)
                if formato_salida == "PDF":
                    with st.spinner(T['spinner_pdf']):
                        buffer = core_logic.exportar_pdf(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension, (page_width, page_height), T)

                    if buffer:
                        st.success(T['success_pdf'])
                        st.download_button(
                            label=T['download_pdf'],
                            data=buffer,
                            file_name='kdp_wordsearch_book.pdf',
                            mime='application/pdf'
                        )

                elif formato_salida == "PPTX":
                    with st.spinner(T['spinner_pptx']):
                        buffer = core_logic.crear_presentacion_ppt(sopas_list, palabras_list_filled, ubicaciones_list, themes, dimension, (page_width, page_height), T)

                    if buffer:
                        st.success(T['success_pptx'])
                        st.download_button(
                            label=T['download_pptx'],
                            data=buffer,
                            file_name='kdp_wordsearch_book.pptx',
                            mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
                        )
        else:
            st.error(T['error_no_file'])

# --- LÓGICA PRINCIPAL DE VISUALIZACIÓN ---
# Esto decide qué mostrar: el formulario de licencia o la app.
if st.session_state['license_valid']:
    show_main_app()
else:
    show_license_form()