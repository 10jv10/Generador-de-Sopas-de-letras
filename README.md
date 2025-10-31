Generador de Sopas de Letras KDP
Genera libros de sopas de letras de alta calidad (PDF y PPTX) a partir de tus propias listas de palabras, listos para publicar en Amazon KDP. Utiliza una interfaz gráfica de Streamlit, ajusta dimensiones, dificultad y personaliza los tamaños estándar de KDP fácilmente.

Características
Interfaz amigable con Streamlit para subir archivos y configurar el libro.

Soporte para exportación a PDF (alta calidad) y PPTX (editable en PowerPoint).

Dimensiones de página adaptadas a los formatos oficiales de KDP.

Control de dificultad: fácil, medio o difícil (definiendo las direcciones de palabras en la cuadrícula).

Genera automáticamente páginas de soluciones, organizadas en cuadrícula.

Soporte multitema: cada sopa de letras puede tener su propio tema.

Compatible con fuentes personalizadas para impresión profesional.

Instalación
Clona el repositorio:

bash
git clone https://github.com/tuusuario/generador-sopas-letras-kdp.git
cd generador-sopas-letras-kdp
Instala las dependencias:

bash
pip install -r requirements.txt
Las dependencias principales son:

streamlit

pandas

python-pptx

reportlab

Asegúrate de tener los archivos de fuente (RobotoMono-Regular.ttf y RobotoMono-Bold.ttf) en la raíz del proyecto.

Uso rápido
Ejecuta la aplicación:

bash
streamlit run app.py
Carga tu archivo Excel:

Columna A: Tema (debe repetirse para cada palabra)

Columna B: Palabra

Sin encabezados.

Configura el tamaño del libro y la dificultad.

Elige PDF o PPTX y genera.

Descarga tu archivo listo para KDP.

Ejemplo de archivo Excel
Tema	Palabra
ANIMALES	Gato
ANIMALES	Perro
FRUTAS	Manzana
FRUTAS	Naranja
Configuración avanzada
Cambia dimensiones de cuadrícula según lo necesites (recomendado 20x20).

Cada tema puede abarcar múltiples sopas si hay muchas palabras.

El sistema alerta si alguna palabra no cabe en la cuadrícula seleccionada.

Contribución
¿Ideas o mejoras?
¡Las contribuciones son bienvenidas! Realiza un fork, crea tu rama y haz un Pull Request siguiendo la guía estándar de GitHub.

Licencia
MIT License. Ver archivo LICENSE para más detalles.

FAQ
¿Este generador es válido para cualquier idioma?
Sí, siempre que utilices letras estándar. Los acentos pueden requerir fuentes alternativas.

¿El archivo PPTX es completamente editable?
Sí, puedes modificar palabras, reemplazar fuentes o ajustar cuadrículas en PowerPoint.

¿Funciona para todos los tamaños de KDP?
Sí, se incluyen los más usados, pero puedes agregar más tamaños en el diccionario TAMAÑOS_KDP.
