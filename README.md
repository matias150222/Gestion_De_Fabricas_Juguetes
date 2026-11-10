# 🧸 Sistema de Gestión Logística y Financiera de Juguetes

![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-data-150458?logo=pandas&logoColor=white)
![Matplotlib](https://img.shields.io/badge/Matplotlib-charts-11557c)
![Excel](https://img.shields.io/badge/Storage-Excel-217346?logo=microsoft-excel&logoColor=white)

Sistema integral en Python para la administración logística, el control de inventarios y el análisis financiero de una empresa multinacional de juguetes.

El proyecto implementa **programación orientada a objetos (POO)**, un **modelo relacional** con claves primarias y foráneas sobre almacenamiento en Excel (`Gestion.xlsx`), **validación estricta de integridad referencial** y **reportes analíticos interactivos** con Matplotlib.

---

## 📑 Tabla de contenidos

- [Tecnologías y herramientas](#-tecnologías-y-herramientas)
- [Modelo relacional](#️-modelo-relacional-gestionxlsx)
- [Módulos del sistema](#-módulos-del-sistema)
- [Demostración y capturas](#-demostración-y-capturas)
- [Instalación y ejecución](#️-instalación-y-ejecución)
- [Estructura del proyecto](#-estructura-del-proyecto)
- [Licencia y autor](#-licencia-y-autor)

---

## 🚀 Tecnologías y herramientas

| Herramienta | Uso en el proyecto |
|---|---|
| **Python 3.x** | Núcleo de la aplicación con arquitectura orientada a objetos y control de excepciones (`try-except`). |
| **Pandas & NumPy** | Procesamiento vectorial, manejo de estructuras relacionales y cálculo de márgenes financieros evitando división por cero. |
| **Matplotlib** | Reportes gráficos interactivos: histogramas, gráficos de torta y barras agrupadas. |
| **Openpyxl** | Persistencia de datos sobre múltiples hojas de Excel. |

---

## 🗂️ Modelo relacional (`Gestion.xlsx`)

El sistema gestiona **3 entidades** interconectadas mediante claves primarias (PK) y foráneas (FK), validando la integridad de los datos antes de cada inserción.

```text
┌─────────────────┐       ┌─────────────────┐       ┌──────────────────┐
│     tb_pais     │       │   tb_fabrica    │       │    tb_juguete    │
├─────────────────┤       ├─────────────────┤       ├──────────────────┤
│ PK  id_pais     │◄──────┤ FK  id_pais_orig│       │ PK  id_juguete   │
│     nombre      │       │ PK  id_fabrica  │◄──────┤ FK  id_fabrica   │
│     impuesto    │       │     nombre      │       │     nombre       │
│     tipo        │       │     capacidad   │       │     stock        │
└─────────────────┘       └─────────────────┘       │     precio_venta │
         ▲                                          │     costo_prod   │
         └──────────────────────────────────────────┤ FK  id_pais_dest │
                                                    └──────────────────┘
```

| Tabla | Descripción |
|---|---|
| `tb_pais` | Países registrados como **Origen** (plantas) o **Destino** (mercados de exportación), con su tasa de impuesto. |
| `tb_fabrica` | Plantas de producción, capacidad máxima anual y país de origen asociado. |
| `tb_juguete` | Inventario de productos: costos, precios, stock y vínculos con la fábrica de origen y el país de destino. |

---

## 💻 Módulos del sistema

El programa se organiza en un **menú interactivo** con submenús por entidad.

### 🌍 Módulo de Países (`tb_pais`)
- Listado ordenado con formato de porcentaje impositivo.
- Registro de nuevos países con **autoincremento de ID**.
- 📊 **Gráfico:** comparativo de tasas impositivas por país (Origen vs. Destino).

### 🧩 Módulo de Juguetes (`tb_juguete`)
- Inventario completo con cruce de información (**JOIN**) de fábrica y país.
- Registro de nuevos juguetes con comprobación de claves foráneas obligatorias.
- Cálculos financieros: ingresos totales por producto e impuestos de destino.
- 📊 **Gráficos:** valorización total del inventario en almacén ($) y distribución porcentual de stock por país de exportación.

### 🏭 Módulo de Fábricas (`tb_fabrica`)
- Listado de plantas y cálculo de capacidad utilizada.
- Cálculo de ingresos agrupados por fábrica e impuestos de origen.
- Análisis de **margen de ganancia (%)** por producto: `(Ganancia / Ingresos) × 100`.
- Asignación interactiva de juguetes en la memoria interna del objeto `Fabrica`.
- 📊 **Gráficos:** impuestos de origen por fábrica y comparativo de capacidad anual vs. stock real.

---

## 📸 Demostración y capturas

> 💡 Guarda tus imágenes en una carpeta `capturas/` y ajusta los nombres de archivo si es necesario.

### 1. Interfaz de consola y tablas formateadas
La consola formatea automáticamente los datos numéricos a moneda (`$`) y porcentajes (`%`), ajustando el ancho de las columnas.

![Consola](capturas/consola.png)

### 2. Menú interactivo
Estructura modular con gestión de excepciones ante errores de escritura por consola.

![Menú](capturas/menu.png)

### 3. Visualizaciones analíticas (Matplotlib)

**Reporte de impuestos por país**

![Impuestos por país](capturas/impuestos_pais.png)

**Valorización del inventario en almacén ($)**

![Valorización del inventario](capturas/valorizacion_inventario.png)

**Distribución de stock por país de destino**

![Distribución de stock](capturas/distribucion_stock.png)

---

## ⚙️ Instalación y ejecución

1. **Clonar o descargar el repositorio:**

   ```bash
   git clone https://github.com/TU_USUARIO/Gestion_De_Fabricas_Juguetes.git
   cd Gestion_De_Fabricas_Juguetes
   ```

2. **Instalar las dependencias:**

   ```bash
   pip install pandas numpy openpyxl matplotlib
   ```

3. **Ejecutar la aplicación:**

   Asegúrate de que `Gestion.xlsx` esté en la misma carpeta que `Gestion.py`.

   ```bash
   python Gestion.py
   ```

---

## 📁 Estructura del proyecto

```text
Gestion_De_Fabricas_Juguetes/
├── Gestion.py        # Aplicación principal (POO + menús)
├── Gestion.xlsx      # Base de datos relacional (tb_pais, tb_fabrica, tb_juguete)
├── capturas/         # Imágenes para el README
└── README.md
```

---

## 📄 Licencia y autor

Proyecto distribuido bajo la licencia **MIT** (ajusta según corresponda).

Desarrollado por **TU_NOMBRE** · [GitHub](https://github.com/TU_USUARIO)
