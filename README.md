🧸 Sistema de Gestión Logística y Financiera de Juguetes (POO & BD Relacional)

Un sistema integral en Python para la administración logística, control de inventarios y análisis financiero de una empresa multinacional de juguetes. El proyecto implementa arquitectura orientada a objetos (POO), modelado relacional con claves primarias y foráneas sobre almacenamiento en Excel (Gestion.xlsx), validación estricta de integridad referencial y reportes analíticos interactivos con Matplotlib.

🚀 Tecnologías y Herramientas

Python 3.x: Núcleo de la aplicación estructurado en arquitectura orientada a objetos (POO) y control avanzado de excepciones (try-except).

Pandas & NumPy: Procesamiento vectorial de datos, manejo de estructuras relacionales y cálculos de márgenes financieros evitando división por cero.

Matplotlib: Generación de reportes gráficos interactivos (histogramas, torta, barras agrupadas).

Openpyxl: Persistencia atómica de datos sobre múltiples hojas de Excel.

🗂️ Modelo Relacional (Gestion.xlsx)

El sistema gestiona 3 entidades relacionales interconectadas mediante claves primarias (PK) y claves foráneas (FK), garantizando la integridad de datos antes de cada inserción:

  ┌─────────────────┐       ┌─────────────────┐       ┌─────────────────┐
  │     tb_pais     │       │   tb_fabrica    │       │   tb_juguete    │
  ├─────────────────┤       ├─────────────────┤       ├─────────────────┤
  │ PK  id_pais     │◄──────┤ FK  id_pais_orig│       │ PK  id_juguete  │
  │     nombre      │       │ PK  id_fabrica  │◄──────┤ FK  id_fabrica  │
  │     impuesto    │       │     nombre      │       │     nombre      │
  │     tipo        │       │     capacidad   │       │     stock       │
  └─────────────────┘       └─────────────────┘       │     precio_venta│
           ▲                                          │     costo_prod  │
           └──────────────────────────────────────────┤ FK  id_pais_dest│
                                                      └─────────────────┘


tb_pais: Almacena los países registrados como Origen (plantas) o Destino (mercados de exportación) junto con sus tasas de impuesto.

tb_fabrica: Almacena las plantas de producción, su capacidad máxima anual y el país de origen asociado.

tb_juguete: Contiene el inventario de productos, costos, precios, stock y sus vinculaciones con la fábrica de origen y el país de destino.

💻 Módulos del Sistema

El programa se organiza en un menú interactivo con submenús por entidad:

Módulo de Países (tb_pais):

Listado ordenado con formato de porcentaje impositivo.

Registro de nuevos países con autoincremento de ID.

📊 Gráfico: Comparativo de tasas impositivas por país (Origen vs. Destino).

Módulo de Juguetes (tb_juguete):

Inventario completo con cruce de información (JOIN) de fábrica y país.

Registro de nuevos juguetes con comprobación de claves foráneas obligatorias.

Cálculos financieros: Ingresos totales por producto e impuestos de destino.

📊 Gráficos: Valorización total del inventario en almacén ($) y distribución porcentual de stock por país de exportación.

Módulo de Fábricas (tb_fabrica):

Listado de plantas y cálculo de capacidad utilizada.

Cálculo de ingresos agrupados por fábrica e impuestos de origen.

Análisis financiero de Margen de Ganancia (%) por producto ((Ganancia / Ingresos) × 100).

Asignación interactiva de juguetes en la memoria interna del objeto Fabrica.

📊 Gráficos: Impuestos de origen por fábrica y comparativo de capacidad anual vs. stock real.

📸 Demostración y Capturas del Sistema

1. Interfaz de Consola y Tablas Formateadas

La consola formatea automáticamente los datos numéricos a moneda ($) y porcentajes (%) ajustando el ancho de columnas.

2. Menú Interactivo

Estructura modular con gestión de excepciones ante errores de escritura por consola.

3. Visualizaciones Analíticas e Interactivas (Matplotlib)

Reporte de Impuestos por País:

Valorización del Inventario en Almacén ($):

Distribución de Stock por País de Destino:

⚙️ Instalación y Ejecución

Clonar o descargar el repositorio:

git clone https://github.com/TU_USUARIO/Gestion_De_Fabricas_Juguetes.git
cd Gestion_De_Fabricas_Juguetes


Instalar las dependencias de Python:

pip install pandas numpy openpyxl matplotlib


Ejecutar la aplicación:
Asegúrate de que el archivo Gestion.xlsx se encuentre en la misma carpeta que Gestion.py.

python Gestion.py
