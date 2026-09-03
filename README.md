# Plataforma Digital - Memoria Descriptiva

Aplicación web interna diseñada para capturar y digitalizar el levantamiento de datos en sitio de obras y remodelaciones comerciales, generando automáticamente el reporte formal en formato Microsoft Word (.docx).

## Características principales

- **Formulario Inteligente en 10 Pasos**:
  1. Datos Generales (Tienda, ubicación y superficies).
  2. Croquis de Localización (Carga y vista previa de mapa).
  3. Descripción General (Texto y lista dinámica de áreas a intervenir).
  4. Remodelación Exterior (Cubierta, Estacionamiento, Anuncio, Fachadas, Andén, Servicio).
  5. Remodelación Interiores (11 áreas con soporte de fotografías).
  6. Giros de Negocio (Consultorio Médico u otros giros).
  7. Descripción de los Trabajos y Afectación Estructural.
  8. Estructura.
  9. Instalaciones (Refrigeración, Aire, Eléctrica, Hidro-Sanitaria, Gas, Filtrado, Incendio).
  10. Medidas de Seguridad y Generación del Documento.

- **Fotografías Dinámicas**:
  - Subida ilimitada de fotografías por sección (drag & drop o selector de archivos).
  - Pie de foto / descripción individual por cada imagen.
  - Maquetación automática en cuadrícula de 2 columnas en el documento Word.

- **Autoguardado**:
  - Guardado progresivo automático cada 1.5 segundos para evitar pérdida de datos.

- **Panel de Control (/admin)**:
  - Listado de todos los levantamientos (borradores y completados).
  - Edición posterior y descarga de archivos generados.

## Requisitos

- Python 3.10 o superior
- Paquetes listados en `requirements.txt` (`flask`, `python-docx`, `pillow`)

## Instalación y Uso Rápido (Windows)

1. Clona este repositorio o descarga los archivos:
   ```bash
   git clone https://github.com/TU_USUARIO/memoria-descriptiva.git
   cd memoria-descriptiva
   ```
2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Ejecuta la aplicación:
   - Haz doble clic en `iniciar.bat`, o corre en consola:
     ```bash
     python app.py
     ```
4. Abre tu navegador en: [http://localhost:5000](http://localhost:5000)
