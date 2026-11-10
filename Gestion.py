import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

FILE_PATH = "Gestion.xlsx"

# Configuración global para formatear DataFrames de Pandas en la terminal
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.unicode.east_asian_width', True)


def cargar_datos(file_path):
    """Carga los DataFrames desde las hojas relacionales del archivo Excel."""
    try:
        sheets = pd.read_excel(file_path, sheet_name=["tb_pais", "tb_fabrica", "tb_juguete"])
        return sheets["tb_pais"], sheets["tb_fabrica"], sheets["tb_juguete"]
    except Exception as e:
        print(f"❌ Error al cargar el archivo Excel: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def guardar_hoja(df, sheet_name, file_path):
    """Actualiza una hoja en el archivo Excel manteniendo intactas las demás."""
    try:
        with pd.ExcelFile(file_path) as xls:
            all_sheets = {s: pd.read_excel(file_path, sheet_name=s) for s in xls.sheet_names}
        
        all_sheets[sheet_name] = df

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for s, d in all_sheets.items():
                d.to_excel(writer, sheet_name=s, index=False)
        print(f"\n✓ Cambios guardados correctamente en la hoja '{sheet_name}'.")
    except Exception as e:
        print(f"❌ Error al guardar los datos en Excel: {e}")


def imprimir_titulo(titulo):
    """Muestra un encabezado estilizado en la consola."""
    ancho = 80
    print("\n" + "=" * ancho)
    print(f"{titulo.center(ancho)}")
    print("=" * ancho)


def formatear_tabla(df, columnas_moneda=[], columnas_porcentaje=[], columnas_renombradas={}):
    """Formatea valores numéricos y renombra columnas para una visualización ordenada en consola."""
    df_temp = df.copy()
    
    for col in columnas_moneda:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].apply(lambda x: f"${x:,.2f}" if pd.notnull(x) else "$0.00")
            
    for col in columnas_porcentaje:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].apply(lambda x: f"{x * 100:.2f}%" if pd.notnull(x) and x <= 1 else f"{x:.2f}%")
            
    if columnas_renombradas:
        df_temp = df_temp.rename(columns=columnas_renombradas)
        
    return df_temp.to_string(index=False)


class Pais:
    def __init__(self, id_pais=None, nombre=None, impuesto=None, tipo=None): 
        self.id_pais = id_pais
        self.nombre = nombre
        self.impuesto = impuesto
        self.tipo = tipo

    def mostrar_pais(self, df_pais):
        imprimir_titulo("TABLA DE PAÍSES REGISTRADOS (tb_pais)")
        if df_pais.empty:
            print("⚠️ No hay países registrados.")
        else:
            renombrar = {
                'id_pais': 'ID',
                'nombre': 'Nombre País',
                'impuesto': 'Tasa Impuesto',
                'tipo': 'Tipo País'
            }
            print(formatear_tabla(df_pais, columnas_porcentaje=['impuesto'], columnas_renombradas=renombrar))

    def agregar_pais(self, df_pais, file_path):
        imprimir_titulo("REGISTRAR NUEVO PAÍS")
        try:
            nuevo_id = 1 if df_pais.empty else int(df_pais['id_pais'].max()) + 1
            print(f"📌 ID asignado automáticamente: {nuevo_id}\n")
            
            nombre = input("▶ Ingrese el nombre del país: ").strip()
            if not nombre:
                print("❌ El nombre no puede estar vacío.")
                return df_pais
            
            impuesto = float(input("▶ Ingrese el impuesto (ej. 0.05 para 5%): "))
            tipo = input("▶ Ingrese el tipo (origen / destino): ").strip().lower()
            if tipo not in ['origen', 'destino']:
                print("❌ Tipo no válido. Debe ser 'origen' o 'destino'.")
                return df_pais

            nuevo_registro = pd.DataFrame([{
                "id_pais": nuevo_id,
                "nombre": nombre,
                "impuesto": impuesto,
                "tipo": tipo
            }])

            df_pais = pd.concat([df_pais, nuevo_registro], ignore_index=True)
            guardar_hoja(df_pais, "tb_pais", file_path)
            print(f"✓ País '{nombre}' (ID: {nuevo_id}) agregado exitosamente.")
            return df_pais
        except ValueError:
            print("❌ Error: Entrada numérica no válida.")
            return df_pais

    def generar_reporte_impuesto_paises(self, df_pais):
        """GRÁFICO 1: Impuestos por País"""
        if df_pais.empty:
            print("⚠️ No hay datos de países para graficar.")
            return
        plt.figure(figsize=(10, 6))
        colores = ['#1f77b4' if str(t).lower() == 'origen' else '#ff7f0e' for t in df_pais['tipo']]
        bars = plt.bar(df_pais['nombre'], df_pais['impuesto'] * 100, color=colores)
        
        plt.xlabel('País', fontsize=12)
        plt.ylabel('Impuesto (%)', fontsize=12)
        plt.title('Porcentaje de Impuestos por País (Azul: Origen | Naranja: Destino)', fontsize=14)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval + 0.2, f"{yval:.1f}%", ha='center', va='bottom')
            
        plt.show()


class Juguete:
    def __init__(self, id_juguete=None, nombre=None, id_fabrica=None, stock=None, precio_venta=None, costo_produccion=None, id_pais_destino=None, impuesto_destino=None): 
        self.id_juguete = id_juguete
        self.nombre = nombre
        self.id_fabrica = id_fabrica
        self.stock = stock
        self.precio_venta = precio_venta
        self.costo_produccion = costo_produccion
        self.id_pais_destino = id_pais_destino
        self.impuesto_destino = impuesto_destino

    def mostrar_juguete(self, df_juguete, df_fabrica=None, df_pais=None):
        imprimir_titulo("INVENTARIO DE JUGUETES (tb_juguete)")
        if df_juguete.empty:
            print("⚠️ No hay juguetes registrados.")
            return
        
        if df_fabrica is not None and df_pais is not None and not df_fabrica.empty and not df_pais.empty:
            df_merged = df_juguete.merge(df_fabrica[['id_fabrica', 'nombre']], on='id_fabrica', how='left', suffixes=('', '_fab'))
            df_merged = df_merged.merge(df_pais[['id_pais', 'nombre']], left_on='id_pais_destino', right_on='id_pais', how='left', suffixes=('', '_pais'))
            df_merged = df_merged.rename(columns={'nombre_fab': 'fabrica_origen', 'nombre_pais': 'pais_destino'})
            
            renombrar = {
                'id_juguete': 'ID',
                'nombre': 'Nombre Juguete',
                'fabrica_origen': 'Fábrica de Origen',
                'stock': 'Stock (Uds)',
                'precio_venta': 'Precio Venta',
                'costo_produccion': 'Costo Prod.',
                'pais_destino': 'País Destino',
                'impuesto_destino': 'Imp. Destino'
            }
            cols = ['id_juguete', 'nombre', 'fabrica_origen', 'stock', 'precio_venta', 'costo_produccion', 'pais_destino', 'impuesto_destino']
            print(formatear_tabla(df_merged[cols], columnas_moneda=['precio_venta', 'costo_produccion'], columnas_porcentaje=['impuesto_destino'], columnas_renombradas=renombrar))
        else:
            print(df_juguete.to_string(index=False))

    def agregar_juguete(self, df_juguete, df_fabrica, df_pais, file_path):
        imprimir_titulo("REGISTRAR NUEVO JUGUETE")
        if df_fabrica.empty or df_pais.empty:
            print("❌ Error: Debe registrar al menos una fábrica y un país de destino antes de crear un juguete.")
            return df_juguete
        
        try:
            nuevo_id = 1 if df_juguete.empty else int(df_juguete['id_juguete'].max()) + 1
            print(f"📌 ID asignado automáticamente: {nuevo_id}\n")

            nombre = input("▶ Nombre del juguete: ").strip()
            
            print("\n🏭 Fábricas disponibles:")
            print(formatear_tabla(df_fabrica[['id_fabrica', 'nombre']], columnas_renombradas={'id_fabrica': 'ID Fábrica', 'nombre': 'Nombre Fábrica'}))
            
            id_fabrica = int(input("\n▶ Ingrese el ID de la Fábrica de origen (FK): "))
            if id_fabrica not in df_fabrica['id_fabrica'].values:
                print("❌ Error: El ID de la fábrica no existe.")
                return df_juguete

            stock = int(input("▶ Stock disponible (unidades): "))
            precio_venta = float(input("▶ Precio de venta ($): "))
            costo_produccion = float(input("▶ Costo de producción ($): "))

            df_destinos = df_pais[df_pais['tipo'].str.lower() == 'destino']
            print("\n🌎 Países de destino disponibles:")
            print(formatear_tabla(df_destinos[['id_pais', 'nombre', 'impuesto']], columnas_porcentaje=['impuesto'], columnas_renombradas={'id_pais': 'ID País', 'nombre': 'Nombre País', 'impuesto': 'Impuesto'}))
            
            id_pais_destino = int(input("\n▶ Ingrese el ID del País de destino (FK): "))
            pais_match = df_pais[df_pais['id_pais'] == id_pais_destino]
            if pais_match.empty:
                print("❌ Error: El ID del país de destino no existe.")
                return df_juguete
            
            impuesto_destino = float(pais_match['impuesto'].values[0])

            nuevo_registro = pd.DataFrame([{
                'id_juguete': nuevo_id,
                'nombre': nombre,
                'id_fabrica': id_fabrica,
                'stock': stock,
                'precio_venta': precio_venta,
                'costo_produccion': costo_produccion,
                'id_pais_destino': id_pais_destino,
                'impuesto_destino': impuesto_destino
            }])

            df_juguete = pd.concat([df_juguete, nuevo_registro], ignore_index=True)
            guardar_hoja(df_juguete, "tb_juguete", file_path)
            print(f"✓ Juguete '{nombre}' (ID: {nuevo_id}) agregado exitosamente.")
            return df_juguete
        except ValueError:
            print("❌ Error: Ingrese valores numéricos válidos.")
            return df_juguete

    def calcular_ingresos_juguetes(self, df_juguete):
        imprimir_titulo("INGRESOS TOTALES POR JUGUETE (Precio Venta x Stock)")
        if df_juguete.empty:
            print("⚠️ No hay juguetes registrados.")
            return
        df_temp = df_juguete.copy()
        df_temp['ingreso_total'] = df_temp['precio_venta'] * df_temp['stock']
        renombrar = {
            'id_juguete': 'ID',
            'nombre': 'Nombre Juguete',
            'stock': 'Stock (Uds)',
            'precio_venta': 'Precio Venta',
            'ingreso_total': 'Ingresos Totales'
        }
        print(formatear_tabla(df_temp[['id_juguete', 'nombre', 'stock', 'precio_venta', 'ingreso_total']], columnas_moneda=['precio_venta', 'ingreso_total'], columnas_renombradas=renombrar))

    def calcular_impuestos_juguete(self, df_juguete, df_pais=None):
        imprimir_titulo("CÁLCULO DE IMPUESTOS DE DESTINO POR PRODUCTO")
        if df_juguete.empty:
            print("⚠️ No hay juguetes registrados.")
            return
        df_temp = df_juguete.copy()
        df_temp['impuesto_unitario'] = df_temp['costo_produccion'] * df_temp['impuesto_destino']
        df_temp['impuesto_total'] = df_temp['impuesto_unitario'] * df_temp['stock']
        
        if df_pais is not None and not df_pais.empty:
            df_temp = df_temp.merge(df_pais[['id_pais', 'nombre']], left_on='id_pais_destino', right_on='id_pais', how='left', suffixes=('', '_pais'))
            df_temp = df_temp.rename(columns={'nombre_pais': 'pais_destino'})
            renombrar = {
                'id_juguete': 'ID',
                'nombre': 'Nombre Juguete',
                'stock': 'Stock',
                'pais_destino': 'País Destino',
                'impuesto_destino': '% Imp. Destino',
                'impuesto_unitario': 'Imp. Unitario',
                'impuesto_total': 'Impuesto Total'
            }
            cols = ['id_juguete', 'nombre', 'stock', 'pais_destino', 'impuesto_destino', 'impuesto_unitario', 'impuesto_total']
            print(formatear_tabla(df_temp[cols], columnas_moneda=['impuesto_unitario', 'impuesto_total'], columnas_porcentaje=['impuesto_destino'], columnas_renombradas=renombrar))
        else:
            print(df_temp.to_string(index=False))

    def generar_grafico_valorizacion_inventario(self, df_juguete):
        """GRÁFICO 3: Valorización del Inventario en Almacén"""
        if df_juguete.empty:
            print("⚠️ No hay datos de juguetes para graficar.")
            return
        df_temp = df_juguete.copy()
        df_temp['valor_almacen'] = df_temp['stock'] * df_temp['costo_produccion']
        df_temp = df_temp.sort_values(by='valor_almacen', ascending=True)

        plt.figure(figsize=(10, 6))
        bars = plt.barh(df_temp['nombre'], df_temp['valor_almacen'], color='#9467bd')
        plt.xlabel('Valorización Total en Costos ($)', fontsize=12)
        plt.ylabel('Juguete / Producto', fontsize=12)
        plt.title('Valorización de Inventario en Almacén por Producto ($)', fontsize=14)
        plt.tight_layout()

        for bar in bars:
            xval = bar.get_width()
            plt.text(xval + 1000, bar.get_y() + bar.get_height()/2, f"${xval:,.2f}", ha='left', va='center', fontsize=9)

        plt.show()

    def generar_grafico_destinos_exportacion(self, df_juguete, df_pais):
        """GRÁFICO 4: Distribución de Stock por País de Destino"""
        if df_juguete.empty or df_pais.empty:
            print("⚠️ Datos insuficientes para generar el gráfico.")
            return
        df_merged = df_juguete.merge(df_pais[['id_pais', 'nombre']], left_on='id_pais_destino', right_on='id_pais', how='left')
        df_destinos = df_merged.groupby('nombre_y')['stock'].sum().reset_index()

        plt.figure(figsize=(8, 8))
        colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#c2c2f0']
        plt.pie(df_destinos['stock'], labels=df_destinos['nombre_y'], autopct='%1.1f%%', startangle=140, colors=colors, wedgeprops={'edgecolor': 'white', 'linewidth': 2})
        plt.title('Distribución Porcentual del Stock por País de Destino', fontsize=14)
        plt.tight_layout()
        plt.show()


class Fabrica:
    def __init__(self, id_fabrica=None, nombre=None, id_pais_origen=None, capacidad=None, impuesto_origen=None):
        self.id_fabrica = id_fabrica
        self.nombre = nombre
        self.id_pais_origen = id_pais_origen
        self.capacidad = capacidad
        self.impuesto_origen = impuesto_origen
        self.lista_juguetes = []

    def mostrar_fabrica(self, df_fabrica, df_pais=None):
        imprimir_titulo("DATOS DE LAS FÁBRICAS (tb_fabrica)")
        if df_fabrica.empty:
            print("⚠️ No hay fábricas registradas.")
            return
        if df_pais is not None and not df_pais.empty:
            df_merged = df_fabrica.merge(df_pais[['id_pais', 'nombre']], left_on='id_pais_origen', right_on='id_pais', how='left', suffixes=('', '_pais'))
            df_merged = df_merged.rename(columns={'nombre_pais': 'pais_origen'})
            renombrar = {
                'id_fabrica': 'ID Fábrica',
                'nombre': 'Nombre Fábrica',
                'pais_origen': 'País Origen',
                'capacidad': 'Capacidad Anual (Uds)',
                'impuesto_origen': 'Imp. Origen'
            }
            cols = ['id_fabrica', 'nombre', 'pais_origen', 'capacidad', 'impuesto_origen']
            print(formatear_tabla(df_merged[cols], columnas_porcentaje=['impuesto_origen'], columnas_renombradas=renombrar))
        else:
            print(df_fabrica.to_string(index=False))

    def agregar_fabrica(self, df_fabrica, df_pais, file_path):
        imprimir_titulo("REGISTRAR NUEVA FÁBRICA")
        df_origen = df_pais[df_pais['tipo'].str.lower() == 'origen'] if not df_pais.empty else pd.DataFrame()
        if df_origen.empty:
            print("❌ Error: Debe existir al menos un País registrado con tipo 'origen' antes de agregar una fábrica.")
            return df_fabrica
        
        try:
            nuevo_id = 101 if df_fabrica.empty else int(df_fabrica['id_fabrica'].max()) + 1
            print(f"📌 ID asignado automáticamente: {nuevo_id}\n")

            nombre = input("▶ Nombre de la fábrica: ").strip()
            print("\n🌎 Países de origen disponibles:")
            print(formatear_tabla(df_origen[['id_pais', 'nombre', 'impuesto']], columnas_porcentaje=['impuesto'], columnas_renombradas={'id_pais': 'ID País', 'nombre': 'Nombre País', 'impuesto': 'Impuesto'}))
            
            id_pais_origen = int(input("\n▶ Ingrese el ID del País de origen (FK): "))
            pais_match = df_pais[df_pais['id_pais'] == id_pais_origen]
            if pais_match.empty:
                print("❌ Error: El ID del país de origen no existe.")
                return df_fabrica

            impuesto_origen = float(pais_match['impuesto'].values[0])
            capacidad = int(input("▶ Capacidad de producción anual (unidades): "))

            nuevo_registro = pd.DataFrame([{
                'id_fabrica': nuevo_id,
                'nombre': nombre,
                'id_pais_origen': id_pais_origen,
                'capacidad': capacidad,
                'impuesto_origen': impuesto_origen
            }])

            df_fabrica = pd.concat([df_fabrica, nuevo_registro], ignore_index=True)
            guardar_hoja(df_fabrica, "tb_fabrica", file_path)
            print(f"✓ Fábrica '{nombre}' (ID: {nuevo_id}) agregada exitosamente.")
            return df_fabrica
        except ValueError:
            print("❌ Error: Ingrese datos numéricos válidos.")
            return df_fabrica

    def calcular_ingresos_total(self, df_fabrica, df_juguete):
        imprimir_titulo("TOTAL DE INGRESOS POR FÁBRICA (JOIN tb_fabrica - tb_juguete)")
        if df_fabrica.empty or df_juguete.empty:
            print("⚠️ Datos insuficientes de fábricas o juguetes.")
            return

        ingresos = []
        for _, fab in df_fabrica.iterrows():
            f_id = fab['id_fabrica']
            f_nombre = fab['nombre']
            juguetes_fab = df_juguete[df_juguete['id_fabrica'] == f_id]
            
            if not juguetes_fab.empty:
                total = (juguetes_fab['precio_venta'] * juguetes_fab['stock']).sum()
                cant_juguetes = len(juguetes_fab)
            else:
                total = 0.0
                cant_juguetes = 0
                
            ingresos.append({
                'id_fabrica': f_id,
                'nombre_fabrica': f_nombre,
                'variedad_juguetes': cant_juguetes,
                'ingresos_totales': round(total, 2)
            })
        
        df_res = pd.DataFrame(ingresos)
        renombrar = {
            'id_fabrica': 'ID Fábrica',
            'nombre_fabrica': 'Nombre Fábrica',
            'variedad_juguetes': 'Productos Distintos',
            'ingresos_totales': 'Ingresos Generados'
        }
        print(formatear_tabla(df_res, columnas_moneda=['ingresos_totales'], columnas_renombradas=renombrar))

    def calcular_impuestos_fabrica(self, df_fabrica, df_pais=None):
        imprimir_titulo("IMPUESTOS DE ORIGEN POR FÁBRICA")
        if df_fabrica.empty:
            print("⚠️ No hay fábricas registradas.")
            return
        df_temp = df_fabrica.copy()
        df_temp['impuesto_total_origen'] = df_temp['capacidad'] * df_temp['impuesto_origen']
        
        if df_pais is not None and not df_pais.empty:
            df_temp = df_temp.merge(df_pais[['id_pais', 'nombre']], left_on='id_pais_origen', right_on='id_pais', how='left', suffixes=('', '_pais'))
            df_temp = df_temp.rename(columns={'nombre_pais': 'pais_origen'})
            renombrar = {
                'id_fabrica': 'ID Fábrica',
                'nombre': 'Nombre Fábrica',
                'pais_origen': 'País Origen',
                'capacidad': 'Capacidad Anual',
                'impuesto_origen': '% Imp. Origen',
                'impuesto_total_origen': 'Impuesto Estimado Total'
            }
            cols = ['id_fabrica', 'nombre', 'pais_origen', 'capacidad', 'impuesto_origen', 'impuesto_total_origen']
            print(formatear_tabla(df_temp[cols], columnas_moneda=['impuesto_total_origen'], columnas_porcentaje=['impuesto_origen'], columnas_renombradas=renombrar))
        else:
            print(df_temp.to_string(index=False))

    def calcular_margen_ganancia(self, df_fabrica, df_juguete):
        imprimir_titulo("ANÁLISIS DE MARGEN DE GANANCIA POR PRODUCTO")
        if df_juguete.empty:
            print("⚠️ No hay juguetes registrados.")
            return
        df_temp = df_juguete.copy()
        df_temp['ingreso_total'] = df_temp['precio_venta'] * df_temp['stock']
        df_temp['costo_total'] = df_temp['costo_produccion'] * df_temp['stock']
        df_temp['ganancia_bruta'] = df_temp['ingreso_total'] - df_temp['costo_total']
        df_temp['margen_pct'] = np.where(df_temp['ingreso_total'] > 0, (df_temp['ganancia_bruta'] / df_temp['ingreso_total']) * 100, 0)
        
        renombrar = {
            'id_juguete': 'ID',
            'nombre': 'Producto',
            'stock': 'Stock',
            'ingreso_total': 'Ventas Totales',
            'costo_total': 'Costos Totales',
            'ganancia_bruta': 'Ganancia Bruta',
            'margen_pct': 'Margen Ganancia'
        }
        cols = ['id_juguete', 'nombre', 'stock', 'ingreso_total', 'costo_total', 'ganancia_bruta', 'margen_pct']
        print(formatear_tabla(df_temp[cols], columnas_moneda=['ingreso_total', 'costo_total', 'ganancia_bruta'], columnas_porcentaje=['margen_pct'], columnas_renombradas=renombrar))

    def agregar_juguete_fabrica(self, id_fabrica, juguete_obj, df_fabrica):
        if not df_fabrica[df_fabrica['id_fabrica'] == id_fabrica].empty:
            self.lista_juguetes.append({'id_fabrica': id_fabrica, 'juguete': juguete_obj})
            print(f"✓ El juguete '{juguete_obj.nombre}' fue asignado a la memoria de la fábrica ID {id_fabrica}.")
        else:
            print(f"❌ Error: La fábrica con ID {id_fabrica} no existe.")

    def mostrar_juguetes_fabrica(self, df_fabrica):
        imprimir_titulo("JUGUETES ASIGNADOS EN MEMORIA A UNA FÁBRICA")
        if df_fabrica.empty:
            print("⚠️ No hay fábricas en el sistema.")
            return
        try:
            id_fabrica = int(input("▶ Ingrese el ID de la fábrica a consultar: "))
            fab_match = df_fabrica[df_fabrica['id_fabrica'] == id_fabrica]
            if not fab_match.empty:
                nombre_fab = fab_match['nombre'].values[0]
                juguetes_asociados = [item['juguete'] for item in self.lista_juguetes if item['id_fabrica'] == id_fabrica]
                if not juguetes_asociados:
                    print(f"\n⚠️ No hay juguetes agregados en memoria para la fábrica '{nombre_fab}' (ID: {id_fabrica}).")
                else:
                    print(f"\n📋 Juguetes en memoria para '{nombre_fab}' (ID: {id_fabrica}):")
                    for i, jg in enumerate(juguetes_asociados, start=1):
                        print(f"  {i}. {jg.nombre:<35} | Stock: {jg.stock:<6} | Precio: ${jg.precio_venta:<6.2f} | Destino ID: {jg.id_pais_destino}")
            else:
                print(f"❌ Error: La fábrica con ID {id_fabrica} no existe.")
        except ValueError:
            print("❌ Error: ID no válido.")

    def generar_reporte_impuestos_fabrica(self, df_fabrica):
        """GRÁFICO 2: Impuestos de Origen por Fábrica"""
        if df_fabrica.empty:
            print("⚠️ No hay fábricas registradas para graficar.")
            return
        plt.figure(figsize=(10, 6))
        bars = plt.bar(df_fabrica['nombre'], df_fabrica['impuesto_origen'] * 100, color='#2ca02c')
        plt.title('Impuestos de Origen por Fábrica (%)', fontsize=14)
        plt.xlabel('Fábrica', fontsize=12)
        plt.ylabel('Impuesto (%)', fontsize=12)
        plt.xticks(rotation=30, ha='right')  
        plt.tight_layout()
        
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval + 0.2, f"{yval:.1f}%", ha='center', va='bottom')
            
        plt.show()

    def generar_grafico_capacidad_vs_stock(self, df_fabrica, df_juguete):
        """GRÁFICO 5: Capacidad de Producción vs. Stock Actual por Fábrica"""
        if df_fabrica.empty or df_juguete.empty:
            print("⚠️ Datos insuficientes para generar el gráfico.")
            return
        
        stock_fab = df_juguete.groupby('id_fabrica')['stock'].sum().reset_index()
        df_merged = df_fabrica.merge(stock_fab, on='id_fabrica', how='left').fillna(0)

        x = np.arange(len(df_merged['nombre']))
        width = 0.35

        plt.figure(figsize=(10, 6))
        plt.bar(x - width/2, df_merged['capacidad'], width, label='Capacidad Anual', color='#1f77b4')
        plt.bar(x + width/2, df_merged['stock'], width, label='Stock Producido', color='#ff7f0e')

        plt.xlabel('Fábrica', fontsize=12)
        plt.ylabel('Unidades', fontsize=12)
        plt.title('Comparativo: Capacidad de Producción Anual vs. Stock Actual', fontsize=14)
        plt.xticks(x, df_merged['nombre'], rotation=30, ha='right')
        plt.legend()
        plt.tight_layout()
        plt.show()


def menu_principal():
    imprimir_titulo("SISTEMA DE GESTIÓN LOGÍSTICA DE JUGUETES (POO & BD)")
    print(" 1 ➔ Operaciones con PAÍS (tb_pais)")
    print(" 2 ➔ Operaciones con JUGUETES (tb_juguete)")
    print(" 3 ➔ Operaciones con FÁBRICA (tb_fabrica)")
    print(" 4 ➔ Salir del sistema")
    print("-" * 80)

def menu_pais():
    imprimir_titulo("MÓDULO DE GESTIÓN DE PAÍSES")
    print(" 1 ➔ Mostrar lista de países")
    print(" 2 ➔ Registrar nuevo país")
    print(" 3 ➔ 📊 Reporte Gráfico: Porcentaje de Impuestos por País")
    print(" 4 ➔ Regresar al menú principal")
    print("-" * 80)

def menu_juguete():
    imprimir_titulo("MÓDULO DE GESTIÓN DE JUGUETES")
    print(" 1 ➔ Mostrar inventario de juguetes")
    print(" 2 ➔ Registrar nuevo juguete")
    print(" 3 ➔ Calcular ingresos por juguete") 
    print(" 4 ➔ Calcular impuestos de destino por juguete")
    print(" 5 ➔ 📊 Reporte Gráfico: Valorización de Inventario en Almacén ($)")
    print(" 6 ➔ 📊 Reporte Gráfico: Distribución de Exportación por País Destino")
    print(" 7 ➔ Regresar al menú principal")
    print("-" * 80)

def menu_fabrica(): 
    imprimir_titulo("MÓDULO DE GESTIÓN DE FÁBRICAS")
    print(" 1 ➔ Mostrar información de fábricas")
    print(" 2 ➔ Registrar nueva fábrica")
    print(" 3 ➔ Calcular ingresos totales agrupados por fábrica") 
    print(" 4 ➔ Calcular impuestos de origen por fábrica") 
    print(" 5 ➔ Calcular margen de ganancia de productos")  
    print(" 6 ➔ 📊 Reporte Gráfico: Impuestos de Origen por Fábrica")
    print(" 7 ➔ 📊 Reporte Gráfico: Capacidad de Producción vs Stock Actual")
    print(" 8 ➔ Mostrar juguetes asignados en memoria") 
    print(" 9 ➔ Asignar juguete a fábrica en memoria")
    print(" 10 ➔ Regresar al menú principal")
    print("-" * 80)


def main():  
    df_pais, df_fabrica, df_juguete = cargar_datos(FILE_PATH)

    fabrica = Fabrica()
    pais = Pais()
    juguete = Juguete()

    while True:
        menu_principal()
        op = input("▶ Seleccione una opción: ").strip()
        match op:
            case "1":
                while True:
                    menu_pais()
                    opcion = input("▶ Seleccione una opción: ").strip()
                    if opcion == "1":
                        pais.mostrar_pais(df_pais)
                    elif opcion == "2":
                        df_pais = pais.agregar_pais(df_pais, FILE_PATH)
                    elif opcion == "3":
                        pais.generar_reporte_impuesto_paises(df_pais)
                    elif opcion == "4":
                        break
                    else:
                        print("❌ Opción no válida.")

            case "2":
                while True:
                    menu_juguete()
                    opcion = input("▶ Seleccione una opción: ").strip()
                    if opcion == "1":
                        juguete.mostrar_juguete(df_juguete, df_fabrica, df_pais)
                    elif opcion == "2":
                        df_juguete = juguete.agregar_juguete(df_juguete, df_fabrica, df_pais, FILE_PATH)
                    elif opcion == "3":
                        juguete.calcular_ingresos_juguetes(df_juguete)
                    elif opcion == "4":
                        juguete.calcular_impuestos_juguete(df_juguete, df_pais)
                    elif opcion == "5":
                        juguete.generar_grafico_valorizacion_inventario(df_juguete)
                    elif opcion == "6":
                        juguete.generar_grafico_destinos_exportacion(df_juguete, df_pais)
                    elif opcion == "7":
                        break
                    else:
                        print("❌ Opción no válida.")

            case "3":
                while True:
                    menu_fabrica() 
                    opcion = input("▶ Seleccione una opción: ").strip()
                    if opcion == "1":
                        fabrica.mostrar_fabrica(df_fabrica, df_pais)
                    elif opcion == "2":
                        df_fabrica = fabrica.agregar_fabrica(df_fabrica, df_pais, FILE_PATH)
                    elif opcion == "3":
                        fabrica.calcular_ingresos_total(df_fabrica, df_juguete)
                    elif opcion == "4":
                        fabrica.calcular_impuestos_fabrica(df_fabrica, df_pais)
                    elif opcion == "5":
                        fabrica.calcular_margen_ganancia(df_fabrica, df_juguete)
                    elif opcion == "6":
                        fabrica.generar_reporte_impuestos_fabrica(df_fabrica)
                    elif opcion == "7":
                        fabrica.generar_grafico_capacidad_vs_stock(df_fabrica, df_juguete)
                    elif opcion == "8":
                        fabrica.mostrar_juguetes_fabrica(df_fabrica) 
                    elif opcion == "9":
                        if df_fabrica.empty:
                            print("⚠️ No hay fábricas disponibles.")
                            continue
                        print("\n🏭 Fábricas registradas:")
                        print(formatear_tabla(df_fabrica[['id_fabrica', 'nombre']], columnas_renombradas={'id_fabrica': 'ID Fábrica', 'nombre': 'Nombre Fábrica'}))
                        try:
                            id_fab = int(input("\n▶ Ingrese el ID de la fábrica: "))
                            nom_juguete = input("▶ Nombre del juguete: ").strip()
                            stock_j = int(input("▶ Stock: "))
                            precio_j = float(input("▶ Precio de venta ($): "))
                            costo_j = float(input("▶ Costo de producción ($): "))
                            id_dest = int(input("▶ ID del país de destino (FK): "))
                            imp_dest = float(input("▶ Impuesto de destino (ej. 0.05): "))
                            
                            nuevo_j = Juguete(None, nom_juguete, id_fab, stock_j, precio_j, costo_j, id_dest, imp_dest)
                            fabrica.agregar_juguete_fabrica(id_fab, nuevo_j, df_fabrica)
                        except ValueError:
                            print("❌ Error: Datos numéricos inválidos.")
                    elif opcion == "10":
                        break
                    else:
                        print("❌ Opción no válida.")

            case "4":
                print("\n👋 ¡Gracias por usar el Sistema de Gestión Logística! Saliendo...\n") 
                break 

            case _:
                print("❌ Opción no válida.")

if __name__ == "__main__":
    main()