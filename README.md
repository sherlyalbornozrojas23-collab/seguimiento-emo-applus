 # Sistema de seguimiento de vencimientos — Applus+

Aplicación desarrollada con Streamlit para importar, organizar y controlar documentos con fechas de vencimiento.

## Funcionalidades

- Importación de archivos Excel y CSV.
- Procesamiento de múltiples hojas de cálculo.
- Clasificación de documentos por categoría y proyecto.
- Seguimiento de fechas de vencimiento y días restantes.
- Identificación de documentos vencidos, próximos a vencer y vigentes.
- Filtros por categoría y proyecto.
- Gráficos y resúmenes ejecutivos.
- Exportación de informes Excel con base consolidada, acciones prioritarias y gráficos.

## Archivos del proyecto

- `app.py`: código principal de la aplicación.
- `requirements.txt`: dependencias de Python.
- `README.md`: documentación del proyecto.

## Instalación local

Se recomienda utilizar Python 3.10 o una versión compatible con las dependencias.

Instala las dependencias:

```bash
pip install -r requirements.txt
```

Inicia la aplicación:

```bash
streamlit run app.py
```

## Preparación de los archivos

Los archivos importados deben incluir datos que permitan identificar los documentos y sus vencimientos. Los nombres de las columnas pueden variar; la aplicación debe reconocer los encabezados equivalentes.

Se recomienda incluir, cuando corresponda:

- Proyecto o área.
- Categoría o tipo de documento.
- Persona, equipo o activo asociado.
- Identificador del documento.
- Fecha de emisión.
- Fecha de vencimiento.
- Responsable.
- Correo de contacto.

El estado de seguimiento y los días restantes deben calcularse a partir de la fecha de vencimiento. No debería ser obligatorio incluir columnas de estado calculado en el Excel original.

Antes de utilizar una base corporativa, comprueba que las fechas se interpreten correctamente y que los documentos queden clasificados en la categoría y proyecto correspondientes.

## Despliegue en Streamlit Community Cloud

1. Sube `app.py`, `requirements.txt` y `README.md` al repositorio de GitHub.
2. Abre la aplicación vinculada al repositorio en Streamlit Community Cloud.
3. Comprueba que la aplicación instale las dependencias y se inicie sin errores.
4. Importa un archivo de prueba y verifica los filtros, los estados y la descarga de informes.

## Almacenamiento y seguridad

Si la aplicación utiliza SQLite o guarda archivos en el disco local, los datos podrían no conservarse después de determinados reinicios o despliegues en Streamlit Community Cloud.

Antes de utilizar información corporativa, valida con TI, SST y el responsable de privacidad los permisos de acceso, la protección de los archivos y el almacenamiento persistente.

No cargues diagnósticos médicos ni historias clínicas en una aplicación que no haya sido autorizada para tratar esa información.

## Validación recomendada

Antes de usar el sistema en producción, verifica con datos de prueba:

- Documentos con vencimientos pasados.
- Documentos que vencen hoy.
- Documentos que vencerán próximamente.
- Documentos sin fecha de vencimiento.
- Archivos con varias hojas.
- Diferentes proyectos y categorías.
- Generación y apertura del informe Excel.
