Applus+ | Seguimiento de vencimientos
Aplicación Streamlit para importar bases Excel/CSV, calcular vencimientos desde las fechas y revisar documentos por proyecto y categoría.
Funcionalidades
- Carga de uno o varios archivos Excel (.xlsx, .xlsm) o CSV desde el Panel de control.
- Lectura de las hojas de cada libro y selección de columnas para mapear proyecto, categoría, persona/activo y fecha de vencimiento.
- Cálculo del estado a partir de la fecha de vencimiento: vencido, vence hoy, próximo a vencer o vigente/por revisar.
- Filtros por proyecto y categoría que actualizan los indicadores, tablas y gráficos del Panel de control.
- Pestaña separada de Reportes para generar informes descargables.
Archivos
- app.py: aplicación.
- requirements.txt: dependencias.
- README.md: instrucciones.
- data/: se crea para la base SQLite y archivos originales importados.
Ejecutar localmente
pip install -r requirements.txt
streamlit run app.py
Despliegue en Streamlit Community Cloud
Sube app.py, requirements.txt y README.md al repositorio conectado a Streamlit Community Cloud. Tras cada cambio, espera a que termine el redeploy y revisa los logs si la aplicación no inicia.
Importación y comprobaciones
1. En el Panel de control, selecciona uno o varios Excel/CSV.
2. Revisa cada hoja y confirma que la columna seleccionada como fecha de vencimiento sea la correcta.
3. Pulsa Importar hoja en cada hoja que quieras incorporar.
4. Comprueba los indicadores, los filtros por proyecto/categoría y la tabla de vencimientos.
5. Prueba con una copia de Excel que contenga fechas pasadas, de hoy, dentro de los próximos 30 días y sin fecha.
No es necesario que la base original contenga una columna de estado; la aplicación calcula el seguimiento usando la fecha de vencimiento seleccionada. Si una hoja no tiene esa columna identificable, asígnala manualmente en el mapeo de columnas.
Almacenamiento y seguridad
La aplicación utiliza SQLite y archivos locales bajo data/. El almacenamiento local de Streamlit Community Cloud puede no ser persistente tras reinicios o despliegues. Para un uso corporativo multiusuario, configura almacenamiento persistente y acceso autenticado con TI. No cargues datos médicos sensibles sin autorización y controles de acceso aprobados.
