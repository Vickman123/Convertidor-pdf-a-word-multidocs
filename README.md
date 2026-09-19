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
- **📝 Formato Word (.docx) 100% Editable**:
  - Generación de documentos OpenXML nativos compatibles con **Microsoft Word, Microsoft 365, Google Docs, LibreOffice y WPS Office**.
  - Reconstrucción de párrafos, encabezados (H1, H2), pesos de fuente (negrita, normal) y saltos de página.
- **👁️ Visor y Editor Integrado**:
  - Vista previa de las páginas del PDF original renderizadas en Canvas de alta resolución con zoom y navegación.
  - Editor de texto interactivo para revisar o hacer ajustes al contenido antes de generar el Word final.
- **🔒 Privacidad y Seguridad Total**:
  - Todo el procesamiento se realiza en la memoria RAM de tu propio navegador. Ningún documento se envía a internet.

---

## 🛠️ Tecnologías y Librerías Utilizadas

- **[PDF.js](https://mozilla.github.io/pdf.js/)** (Mozilla): Motor para parsear capas de texto y renderizar páginas en canvas.
- **[docx.js](https://docx.js.org/)**: Generador de archivos `.docx` nativos en JavaScript.
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
