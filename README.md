# 📄 Convertidor de PDF a Word (.docx) Editable - Soporte Masivo

Una potente aplicación web construida con **HTML5, CSS3 y JavaScript moderno** que permite convertir documentos PDF a archivos Microsoft Word (`.docx`) completamente editables. Funciona **100% en el navegador (lado del cliente)**, sin servidores intermedios, garantizando máxima velocidad, cero límites y privacidad absoluta.

![Licencia](https://img.shields.io/badge/Licencia-MIT-blue.svg)
![Tecnología](https://img.shields.io/badge/Tecnología-HTML5%20%7C%20CSS3%20%7C%20JS-brightgreen.svg)
![Privacidad](https://img.shields.io/badge/Privacidad-100%25%20Local-success.svg)

---

## 🚀 Características Principales

- **⚡ Conversión en Masa (Batch Processing)**:
  - Sube **10, 20 o más archivos PDF simultáneamente** mediante arrastrar y soltar (Drag & Drop) o selector de archivos.
  - Barra de progreso global y control individual por documento.
  - Botón maestro **"Convertir Todos"**.
  - Botón maestro **"Descargar Todos (.ZIP)"** para obtener todos los archivos Word empaquetados en un solo clic.
- **📐 Motor de Maquetación Organizacional e Institucional**:
  - **Reflujo de Texto Continuo y Justificado**: Agrupa automáticamente las líneas pertenecientes al mismo párrafo institucional, preservando la justificación completa (`AlignmentType.BOTH`) sin saltos de línea huérfanos.
  - **Formato Mixto en el Mismo Renglón**: Detecta y conserva negritas y cursivas intercaladas dentro de un mismo párrafo (ej. nombres de dependencias o funcionarios en negrita dentro de texto regular) usando `docx.TextRun`.
  - **Listas y Viñetas Nativas con Sangría Francesa**: Identifica viñetas institucionales (bullets •, números, letras) y las formatea con sangría colgante profesional (`indent: { left: 720, hanging: 360 }`).
  - **Preservación de Logotipos y Membretes**: Extrae automáticamente imágenes vectoriales y de mapa de bits (escudos universitarios, logotipos gubernamentales) y los inserta como encabezados gráficos en Word.
  - **Detección Automática de Tablas**: Identifica bloques con columnas y los convierte en tablas nativas de Word (`docx.Table`), evitando que el texto se descuadre.
  - **Jerarquía Tipográfica**: Reconoce títulos principales (H1, H2) según tamaño y estilo original.
- **🔍 Auto-OCR Inteligente (Tesseract.js)**:
  - Detección automática de páginas escaneadas o fotocopias sin capa de texto seleccionable.
  - Reconocimiento óptico de caracteres en segundo plano en múltiples idiomas (Español / Inglés).
- **📝 Formato Word (.docx) 100% Editable**:
  - Generación de documentos OpenXML nativos compatibles con **Microsoft Word, Microsoft 365, Google Docs, LibreOffice y WPS Office**.
- **👁️ Visor y Editor Integrado**:
  - Vista previa de las páginas del PDF original renderizadas en Canvas de alta resolución con zoom y navegación.
  - Editor de texto interactivo para revisar o hacer ajustes al contenido antes de generar el Word final.
- **🔒 Privacidad y Seguridad Total**:
  - Todo el procesamiento se realiza en la memoria RAM de tu propio navegador. Ningún documento confidencial sale de tu equipo.

---

## 🛠️ Tecnologías y Librerías Utilizadas

- **[PDF.js](https://mozilla.github.io/pdf.js/)** (Mozilla): Motor para parsear capas de texto y renderizar páginas en canvas.
- **[Tesseract.js v5](https://tesseract.projectnaptha.com/)**: Motor de reconocimiento óptico de caracteres (OCR) ejecutado en WebAssembly/JS.
- **[docx.js v8.5](https://docx.js.org/)**: Generador de documentos Word nativos (`.docx`), tablas, párrafos, sangrías y estilos.
- **[JSZip](https://stuk.github.io/jszip/)**: Empaquetador masivo para descarga en archivo `.zip`.
- **[FileSaver.js](https://github.com/eligrey/FileSaver.js/)**: Manejo confiable de descargas binarias en navegadores.
- **[FontAwesome](https://fontawesome.com/)**: Iconografía moderna y limpia.

---

## 📂 Estructura del Repositorio

```
Convertidor-pdf-a-word-multidocs/
├── index.html       # Estructura principal y controles de usuario
├── style.css        # Diseño moderno, responsivo y temas de interfaz
├── app.js           # Lógica de extracción, conversión y descarga
├── ejemplo_1.pdf    # Archivo de prueba 1
├── ejemplo_2.pdf    # Archivo de prueba 2
├── ejemplo_3.pdf    # Archivo de prueba 3
└── README.md        # Documentación oficial
```

---

## 💻 Instrucciones de Uso

1. Clona este repositorio o descarga los archivos:
   ```bash
   git clone https://github.com/Vickman123/Convertidor-pdf-a-word-multidocs.git
   ```
2. Abre el archivo `index.html` en cualquier navegador web moderno (Google Chrome, Microsoft Edge, Mozilla Firefox, Safari, Brave, etc.).
3. Arrastra tus archivos PDF a la zona punteada o presiona **"Seleccionar Archivos PDF"**.
4. Haz clic en **"Convertir Todos"** para procesarlos en lote.
5. Descarga los documentos individuales o haz clic en **"Descargar Todos (.ZIP)"**.

---

## 🌐 Despliegue en GitHub Pages

Este proyecto es una Single Page Application (SPA) sin dependencias de backend, por lo que puede desplegarse gratuitamente en **GitHub Pages**:
1. Ve a **Settings > Pages** en el repositorio de GitHub.
2. En **Build and deployment > Source**, selecciona **Deploy from a branch**.
3. Elige la rama `main` y la carpeta `/ (root)`.
4. Guarda los cambios. Tu convertidor estará accesible en línea en:
   `https://vickman123.github.io/Convertidor-pdf-a-word-multidocs/`

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Siéntete libre de utilizarlo, modificarlo y distribuirlo.
