import os
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import Tk, ttk, Frame, Label, Button, Entry, messagebox, filedialog, BooleanVar, Checkbutton, StringVar, Radiobutton
from tkcalendar import DateEntry
from database import Database
from datetime import datetime
import matplotlib.pyplot as plt 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from math import sqrt, ceil
import re
import numpy as np


class AppVentas:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestión de Ventas")
        self.root.geometry("1500x800")
        self.db = Database()
        self._crear_interfaz()

    def _crear_interfaz(self):
        # Frame para los controles
        frame_controles = Frame(self.root)
        frame_controles.pack(pady=10)

        # Selector de fecha
        self.label_fecha = Label(frame_controles, text="Fecha de Carga:")
        self.label_fecha.pack(side="left", padx=10)
        self.cal_fecha = DateEntry(frame_controles, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.cal_fecha.set_date(datetime.now())  # Establecer la fecha actual por defecto
        self.cal_fecha.pack(side="left", padx=10)

        # Botón para cargar un archivo de Excel
        self.btn_cargar = Button(frame_controles, text="Cargar Archivo", command=self.procesar_archivo)
        self.btn_cargar.pack(side="left", padx=10)

        # Campo de búsqueda
        self.label_busqueda = Label(frame_controles, text="Buscar:")
        self.label_busqueda.pack(side="left", padx=10)
        self.entry_busqueda = Entry(frame_controles, width=15)
        self.entry_busqueda.pack(side="left", padx=10)
        self.btn_buscar = Button(frame_controles, text="Buscar", command=self.buscar_datos)
        self.btn_buscar.pack(side="left", padx=10)

        # Campo para ingresar el nombre del producto
        self.label_producto = Label(frame_controles, text="Producto:")
        self.label_producto.pack(side="left", padx=10)
        self.entry_producto = Entry(frame_controles, width=15)
        self.entry_producto.pack(side="left", padx=10)

        # Botón para abrir la ventana de selección de rango de fechas y generar el gráfico
        self.btn_grafico = Button(frame_controles, text="Generar Gráfico de Ventas", command=self.abrir_ventana_grafico)
        self.btn_grafico.pack(side="left", padx=10)

        self.btn_comparar_meses = Button(frame_controles, text="Comparar Meses", command=self.generar_grafico_comparativo_meses)
        self.btn_comparar_meses.pack(side="left", padx=10)

        # Botón para cargar inventario y calcular pedidos
        self.btn_cargar_inventario = Button(frame_controles, text="Cargar Inventario y Calcular Pedidos", command=self.cargar_inventario_y_calcular_pedidos)
        self.btn_cargar_inventario.pack(side="left", padx=10)

        # Tabla para mostrar los datos
        self.frame_tabla = Frame(self.root)
        self.frame_tabla.pack(fill="both", expand=True, padx=10, pady=10)

        # Configurar la tabla
        self.columnas = ("ID", "Código", "Nombre", "Cantidad", "Fecha Carga")
        self.tabla = ttk.Treeview(self.frame_tabla, columns=self.columnas, show="headings")
        for col in self.columnas:
            self.tabla.heading(col, text=col)
        self.tabla.pack(fill="both", expand=True)

        # Barra de desplazamiento
        self.scrollbar = ttk.Scrollbar(self.frame_tabla, orient="vertical", command=self.tabla.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tabla.configure(yscrollcommand=self.scrollbar.set)

        # Boton de cambio de prioridad
        self.btn_prioridades = Button(frame_controles, text="Gestionar Prioridades", command=self.gestionar_prioridades,)
        self.btn_prioridades.pack(side=tk.LEFT, padx=5)

        # Actualizar la tabla al iniciar
        self.actualizar_tabla()

    def seleccionar_carpeta_mes(self):
        """Selecciona una carpeta de mes y devuelve su ruta"""
        carpeta = filedialog.askdirectory(
            initialdir=os.path.join(os.getcwd(), 'ventas'),
            title="Selecciona la carpeta del mes (ej: 2023-11)"
        )
        return carpeta if carpeta else None

    def procesar_archivo(self):
        # Seleccionar carpeta del mes en lugar de archivo individual
        carpeta_mes = self.seleccionar_carpeta_mes()
        if not carpeta_mes:
            return

        try:
            # Extraer año y mes del nombre de la carpeta (ej: "ventas/2023-11" -> 2023-11)
            nombre_carpeta = os.path.basename(carpeta_mes)
            if not re.match(r'\d{4}-\d{2}', nombre_carpeta):
                raise ValueError("Formato de carpeta incorrecto. Debe ser AAAA-MM")

            año, mes = nombre_carpeta.split('-')
            
            # Procesar cada archivo Excel en la carpeta
            archivos_procesados = 0
            for archivo in os.listdir(carpeta_mes):
                if archivo.endswith('.xlsx'):
                    # Extraer día del nombre del archivo (ej: "ventas_15.xlsx" -> 15)
                    match = re.search(r'(\d{1,2})\.xlsx$', archivo)
                    if not match:
                        continue
                    
                    dia = match.group(1).zfill(2)
                    fecha_carga = f"{año}-{mes}-{dia}"
                    ruta_archivo = os.path.join(carpeta_mes, archivo)

                    # Procesar el archivo Excel
                    df = pd.read_excel(ruta_archivo)
                    
                    # Verificar columnas requeridas
                    if not all(col in df.columns for col in ['Codigo', 'Nombre', 'Cantidad']):
                        continue

                    # Insertar datos en la base de datos
                    for _, row in df.iterrows():
                        self.db.cursor.execute('''
                        INSERT INTO ventas (codigo, nombre, cantidad, fecha_carga)
                        VALUES (?, ?, ?, ?)
                        ''', (row['Codigo'], row['Nombre'], row['Cantidad'], fecha_carga))
                    
                    archivos_procesados += 1

            self.db.conn.commit()
            messagebox.showinfo("Éxito", 
                f"Procesados {archivos_procesados} archivos de {nombre_carpeta}")
            self.actualizar_tabla()

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar archivos: {str(e)}")

    def actualizar_tabla(self):
        # Limpiar la tabla actual
        for row in self.tabla.get_children():
            self.tabla.delete(row)

        # Obtener la fecha seleccionada
        fecha_seleccionada = self.cal_fecha.get_date().strftime("%Y-%m-%d")

        # Obtener los datos de la base de datos para la fecha seleccionada
        self.db.cursor.execute('''
        SELECT * FROM ventas WHERE fecha_carga = ?
        ''', (fecha_seleccionada,))
        rows = self.db.cursor.fetchall()

        # Insertar los datos en la tabla
        for row in rows:
            self.tabla.insert("", "end", values=row)

    def buscar_datos(self):
        # Obtener el valor de búsqueda
        busqueda = self.entry_busqueda.get()

        # Limpiar la tabla actual
        for row in self.tabla.get_children():
            self.tabla.delete(row)

        # Obtener la fecha seleccionada
        fecha_seleccionada = self.cal_fecha.get_date().strftime("%Y-%m-%d")

        # Buscar en la base de datos para la fecha seleccionada
        self.db.cursor.execute('''
        SELECT * FROM ventas 
        WHERE (codigo LIKE ? OR nombre LIKE ? OR cantidad LIKE ?) AND fecha_carga = ?
        ''', (f"%{busqueda}%", f"%{busqueda}%", f"%{busqueda}%", fecha_seleccionada))
        rows = self.db.cursor.fetchall()

        # Insertar los resultados en la tabla
        for row in rows:
            self.tabla.insert("", "end", values=row)

    def abrir_ventana_grafico(self):
        # Crear una nueva ventana para seleccionar el rango de fechas
        ventana_rango_fechas = tk.Toplevel(self.root)
        ventana_rango_fechas.title("Seleccionar Rango de Fechas")
        ventana_rango_fechas.geometry("400x200")

        # Selector de fecha inicial
        label_fecha_inicio = Label(ventana_rango_fechas, text="Fecha Inicial:")
        label_fecha_inicio.pack(pady=5)
        cal_fecha_inicio = DateEntry(ventana_rango_fechas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        cal_fecha_inicio.pack(pady=5)

        # Selector de fecha final
        label_fecha_fin = Label(ventana_rango_fechas, text="Fecha Final:")
        label_fecha_fin.pack(pady=5)
        cal_fecha_fin = DateEntry(ventana_rango_fechas, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        cal_fecha_fin.pack(pady=5)

        # Función para generar el gráfico con el rango de fechas seleccionado
        def generar_grafico():
            fecha_inicio = cal_fecha_inicio.get_date().strftime("%Y-%m-%d")
            fecha_fin = cal_fecha_fin.get_date().strftime("%Y-%m-%d")

            if fecha_inicio > fecha_fin:
                messagebox.showerror("Error", "La fecha inicial no puede ser mayor que la fecha final.")
                return

            # Obtener el nombre del producto a analizar
            producto = self.entry_producto.get()

            if not producto:
                messagebox.showerror("Error", "Por favor, ingresa el nombre del producto.")
                return

            # Consultar las ventas del producto en el rango de fechas seleccionado
            self.db.cursor.execute('''
            SELECT fecha_carga, SUM(cantidad) as total_cantidad
            FROM ventas
            WHERE codigo = ? AND fecha_carga BETWEEN ? AND ?
            GROUP BY fecha_carga
            ORDER BY fecha_carga
            ''', (producto, fecha_inicio, fecha_fin))
            ventas_rango = self.db.cursor.fetchall()

            if not ventas_rango:
                messagebox.showinfo("Información", f"No hay datos de ventas para el producto: {producto} en el rango de fechas seleccionado.")
                return

            # Separar las fechas y las cantidades
            fechas = [venta[0] for venta in ventas_rango]
            cantidades = [venta[1] for venta in ventas_rango]

            # Crear una nueva ventana para mostrar el gráfico
            ventana_grafico = Tk()
            ventana_grafico.title(f"Ventas de {producto} ({fecha_inicio} a {fecha_fin})")
            ventana_grafico.geometry("800x600")

            # Crear una figura de matplotlib
            fig, ax = plt.subplots()
            ax.plot(fechas, cantidades, marker='o', linestyle='-', color='b')
            ax.set_title(f"Ventas de {producto} ({fecha_inicio} a {fecha_fin})")
            ax.set_xlabel("Fecha")
            ax.set_ylabel("Cantidad Vendida")
            ax.grid(True)
            plt.xticks(rotation=50)  # Rotar las etiquetas del eje X para mejor visualización

            # Integrar el gráfico en la ventana de Tkinter
            canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

            # Cerrar la ventana de selección de fechas
            ventana_rango_fechas.destroy()

            ventana_grafico.mainloop()

        # Botón para generar el gráfico
        btn_generar_grafico = Button(ventana_rango_fechas, text="Generar Gráfico", command=generar_grafico)
        btn_generar_grafico.pack(pady=10)

        ventana_rango_fechas.mainloop()

    def generar_grafico_comparativo_meses(self):
        # Crear ventana para selección de meses (usando Toplevel en lugar de Tk)
        ventana_seleccion_meses = tk.Toplevel(self.root)
        ventana_seleccion_meses.title("Selección de Meses para Comparación")
        ventana_seleccion_meses.geometry("400x400")
        
        # Obtener el nombre del producto a analizar
        producto = self.entry_producto.get()
        
        if not producto:
            messagebox.showerror("Error", "Por favor, ingresa el nombre del producto.")
            ventana_seleccion_meses.destroy()
            return
        
        # Frame para los checkboxes de meses
        frame_meses = tk.Frame(ventana_seleccion_meses)
        frame_meses.pack(pady=10)
        
        # Obtener todos los meses disponibles para el producto
        self.db.cursor.execute('''
        SELECT DISTINCT strftime('%Y-%m', fecha_carga) as mes
        FROM ventas
        WHERE codigo = ?
        ORDER BY mes
        ''', (producto,))
        meses_disponibles = [row[0] for row in self.db.cursor.fetchall()]
        
        if not meses_disponibles:
            messagebox.showinfo("Información", f"No hay datos de ventas para el producto: {producto}.")
            ventana_seleccion_meses.destroy()
            return
        
        # Variables para los checkboxes (como atributo de la ventana)
        ventana_seleccion_meses.vars_meses = {mes: tk.BooleanVar(value=False) for mes in meses_disponibles}

        # Crear checkboxes para cada mes disponible
        label = tk.Label(frame_meses, text="Selecciona los meses a comparar:")
        label.pack(pady=5)

        for mes in meses_disponibles:
            cb = tk.Checkbutton(
                frame_meses, 
                text=mes,
                variable=ventana_seleccion_meses.vars_meses[mes],
                onvalue=True,
                offvalue=False,
                anchor='w'
            )
            cb.pack(fill='x', padx=5, pady=2)
        
        # Frame para el año de referencia
        frame_anio = tk.Frame(ventana_seleccion_meses)
        frame_anio.pack(pady=10)
        tk.Label(frame_anio, text="Año de referencia (opcional):").pack()
        entry_anio = tk.Entry(frame_anio)
        entry_anio.pack()
        
        def generar_comparacion():
            # Obtener meses seleccionados (CORRECCIÓN IMPORTANTE)
            meses_seleccionados = [
                mes for mes in meses_disponibles 
                if ventana_seleccion_meses.vars_meses[mes].get()  # Acceso correcto a las variables
            ]
            
            # print("Meses seleccionados:", meses_seleccionados)  # Debug
            
            if len(meses_seleccionados) < 2:
                messagebox.showerror("Error", "Selecciona al menos 2 meses para comparar.")
                return
            
            # Resto del código permanece igual...
            # [Aquí iría el resto de tu función generar_comparacion()]
            
            # Obtener año de referencia si se especificó
            anio_referencia = entry_anio.get().strip()
            
            # Consultar datos para cada mes seleccionado
            datos_meses = {}
            
            for mes in meses_seleccionados:
                # Construir rango de fechas (todo el mes)
                fecha_inicio = f"{mes}-01"
                
                # Calcular fecha fin (versión simplificada)
                if anio_referencia:
                    mes_num = mes.split('-')[1]
                    fecha_inicio = f"{anio_referencia}-{mes_num}-01"

                def es_bisiesto(year):
                    """Determina si un año es bisiesto"""
                    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)
                
                year, month = map(int, fecha_inicio.split('-')[:2])
                
                # Determinar último día del mes
                if month == 2:  # Febrero
                    last_day = 29 if es_bisiesto(year) else 28
                elif month in [4, 6, 9, 11]:  # Abril, Junio, Septiembre, Noviembre
                    last_day = 30
                else:  # Resto de meses
                    last_day = 31
                
                fecha_fin = f"{year}-{month:02d}-{last_day:02d}"
                
                # Consultar datos
                self.db.cursor.execute('''
                SELECT fecha_carga, SUM(cantidad) as total_cantidad
                FROM ventas
                WHERE codigo = ? AND fecha_carga BETWEEN ? AND ?
                GROUP BY fecha_carga
                ORDER BY fecha_carga
                ''', (producto, fecha_inicio, fecha_fin))
                
                ventas_mes = self.db.cursor.fetchall()
                
                if ventas_mes:
                    # Normalizar fechas para que todas aparezcan como si fueran del mismo año
                    fechas = []
                    cantidades = []
                    
                    for venta in ventas_mes:
                        fecha = venta[0]
                        cantidad = venta[1]
                        
                        # Si hay año de referencia, cambiar el año en las fechas
                        if anio_referencia:
                            # Mantener mes y día original, cambiar año
                            fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
                            fecha_normalizada = fecha_obj.replace(year=int(anio_referencia)).strftime("%Y-%m-%d")
                        else:
                            # Usar fecha original
                            fecha_normalizada = fecha
                        
                        fechas.append(fecha_normalizada)
                        cantidades.append(cantidad)
                    
                    datos_meses[mes] = (fechas, cantidades)
            
            if not datos_meses:
                messagebox.showinfo("Información", "No hay datos para los meses seleccionados.")
                ventana_seleccion_meses.destroy()
                return
            
            # Crear ventana para el gráfico comparativo (usando Toplevel)
            ventana_grafico = tk.Toplevel(self.root)
            ventana_grafico.title(f"Comparativo de Ventas - {producto}")
            ventana_grafico.geometry("1100x700")
            
            # Crear figura de matplotlib
            fig, ax = plt.subplots(figsize=(12, 6))
            
            # Generar colores distintos para cada mes
            colors = plt.cm.tab10(np.linspace(0, 1, len(datos_meses)))
            
            # Graficar cada mes
            for (mes, color), (fechas, cantidades) in zip(zip(datos_meses.keys(), colors), datos_meses.values()):
                # Extraer días del mes (1-31)
                dias = [int(f.split('-')[2]) for f in fechas]
                
                ax.plot(dias, cantidades, 'o-', color=color, label=mes, markersize=8)
            
            # Configurar el gráfico
            titulo = f"Comparativo de Ventas - {producto}"
            if anio_referencia:
                titulo += f" (Año de referencia: {anio_referencia})"
            
            ax.set_title(f"Comparativo Mensual de Ventas - {producto}")
            ax.set_xlabel("Día del Mes")
            ax.set_ylabel("Cantidad Vendida")
            ax.legend(title="Meses")
            ax.grid(True, alpha=0.3)
            
            # Formatear el eje X para mostrar solo el día
            ax.set_xlim(0.5, 31.5)  # Margen adicional
            ax.set_xticks(range(1, 32))
            ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x)}'))
            
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            # Integrar el gráfico en Tkinter
            canvas = FigureCanvasTkAgg(fig, master=ventana_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            
            # Botón para guardar el gráfico
            def guardar_grafico():
                filepath = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    filetypes=[("PNG", "*.png"), ("JPEG", "*.jpg"), ("PDF", "*.pdf"), ("Todos", "*.*")],
                    title="Guardar gráfico como..."
                )
                if filepath:
                    fig.savefig(filepath, dpi=300, bbox_inches='tight')
                    messagebox.showinfo("Éxito", f"Gráfico guardado en:\n{filepath}")
            
            btn_guardar = tk.Button(ventana_grafico, text="Guardar Gráfico", command=guardar_grafico)
            btn_guardar.pack(pady=10)
            
            ventana_seleccion_meses.destroy()
        
        # Botón para generar la comparación
        btn_generar = tk.Button(ventana_seleccion_meses, text="Generar Comparación", command=generar_comparacion)
        btn_generar.pack(pady=20)

    def cargar_inventario_y_calcular_pedidos(self):
        try:
            # Cargar el archivo de Excel con el inventario actual
            archivo = filedialog.askopenfilename(
                initialdir=os.path.join(os.getcwd(), 'inventario'),
                title="Selecciona un archivo de Excel",
                filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
            )
            if not archivo:
                return

            # Leer el archivo de Excel
            df = pd.read_excel(archivo)

            # Verificar columnas requeridas
            columnas_requeridas = ['codigo', 'nombre', 'cantidad']
            if not all(col in df.columns for col in columnas_requeridas):
                messagebox.showerror("Error", "El archivo de Excel no tiene las columnas requeridas.")
                return

            # Crear ventana para recomendaciones
            ventana_recomendaciones = tk.Toplevel(self.root)
            ventana_recomendaciones.title("Recomendaciones de Pedidos")
            ventana_recomendaciones.geometry("1200x600")

            # Crear tabla de recomendaciones
            columnas = ("Producto", "Inventario Actual", "Demanda Promedio Diaria", 
                    "Punto de Reorden", "Cantidad a Pedir (EOQ)", "Cajas a Pedir")
            tabla_recomendaciones = ttk.Treeview(ventana_recomendaciones, columns=columnas, show="headings")
            for col in columnas:
                tabla_recomendaciones.heading(col, text=col)
            tabla_recomendaciones.pack(fill="both", expand=True)

            # Parámetros configurables
            TIEMPO_ENTREGA = 5
            COSTO_PEDIDO = 30
            COSTO_ALMACENAMIENTO = 69
            BUFFER_PRIORIDAD = {
                'alta': 1.2,   # 20% sobre el punto de reorden
                'media': 1.1,  # 10% sobre el punto de reorden
                'baja': 1.0    # Sin buffer
            }

            # Procesar cada producto
            for index, row in df.iterrows():
                codigo = row['codigo']
                nombre = row['nombre']
                inventario_actual = row['cantidad']

                # Obtener datos del producto
                self.db.cursor.execute('''
                SELECT cantidad_por_caja, prioridad FROM productos WHERE codigo = ?
                ''', (codigo,))
                resultado = self.db.cursor.fetchone()
                
                if not resultado:
                    continue  # Si no existe en productos, saltar
                    
                cantidad_por_caja, prioridad = resultado
                prioridad = prioridad.lower() if prioridad else 'baja'

                # Obtener ventas históricas
                self.db.cursor.execute('''
                SELECT SUM(cantidad) FROM ventas WHERE codigo = ?
                GROUP BY fecha_carga
                ''', (codigo,))
                ventas = [v[0] for v in self.db.cursor.fetchall()]
                
                if not ventas:
                    continue

                # Cálculos principales
                demanda_diaria = sum(ventas) / len(ventas)
                punto_reorden = demanda_diaria * TIEMPO_ENTREGA
                eoq = sqrt((2 * demanda_diaria * 365 * COSTO_PEDIDO) / COSTO_ALMACENAMIENTO)
                
                # Aplicar buffer según prioridad
                buffer = BUFFER_PRIORIDAD.get(prioridad, 1.0)
                punto_efectivo = punto_reorden * buffer

                # Calcular cajas a pedir
                cajas_a_pedir = 0
                if inventario_actual < punto_efectivo:
                    if inventario_actual < punto_reorden:
                        cantidad_necesaria = (punto_reorden - inventario_actual) + eoq
                    else:
                        cantidad_necesaria = eoq
                    
                    cajas_a_pedir = max(0, ceil(cantidad_necesaria / cantidad_por_caja) - 1)

                # Insertar en tabla
                tabla_recomendaciones.insert("", "end", values=(
                    nombre,
                    inventario_actual,
                    round(demanda_diaria, 2),
                    round(punto_reorden, 2),
                    round(eoq, 2),
                    cajas_a_pedir
                ))

            # Botón de exportación
            btn_exportar = tk.Button(
                ventana_recomendaciones,
                text="Generar Orden Automática",
                command=lambda: self.generar_orden_desde_plantilla(tabla_recomendaciones),
                bg="#4CAF50",
                fg="white"
            )
            btn_exportar.pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar inventario: {str(e)}")

    def generar_orden_desde_plantilla(self, tabla_recomendaciones):
        try:
            # 1. Obtener productos a pedir (cajas > 0)
            pedidos = []
            for item in tabla_recomendaciones.get_children():
                valores = tabla_recomendaciones.item(item)['values']
                if valores[5] > 0:  # Cajas a Pedir > 0
                    pedidos.append({
                        'producto': valores[0],  # Nombre
                        'cajas': valores[5]      # Cantidad
                    })

            if not pedidos:
                messagebox.showinfo("Info", "No hay productos para pedir")
                return

            # 2. Seleccionar plantilla
            plantilla_path = filedialog.askopenfilename(
                title="Seleccionar plantilla de pedido",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not plantilla_path:
                return

            # 3. Cargar plantilla
            wb = openpyxl.load_workbook(plantilla_path)
            
            # 4. Seleccionar hoja específica (ajusta el nombre)
            centroS = "Arequipa"
            hoja_pedidos = wb[centroS]  # Nombre exacto de tu hoja
            
            # 5. Configuración de columnas (ajustar según tu plantilla)
            COL_PRODUCTO = 2   # Columna B
            COL_CANTIDAD = 8   # Columna G
            FILA_INICIO = 4    # Fila donde empiezan los productos

            # 6. Rellenar datos
            productos_procesados = 0
            
            for fila in hoja_pedidos.iter_rows(min_row=FILA_INICIO):
                celda_producto = fila[COL_PRODUCTO - 1]
                if celda_producto.value:
                    # Buscar coincidencia (insensible a mayúsculas/espacios)
                    producto_plantilla = str(celda_producto.value).strip().lower()
                    
                    for pedido in pedidos:
                        if pedido['producto'].strip().lower() in producto_plantilla:
                            fila[COL_CANTIDAD - 1].value = pedido['cajas']
                            productos_procesados += 1
                            break

            # 7. Guardar
            if productos_procesados > 0:
                archivo_salida = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"Orden_Pedido_{centroS}{datetime.now().strftime('%Y-%m-%d')}",
                    filetypes=[("Excel files", "*.xlsx")]
                )
                
                if archivo_salida:
                    wb.save(archivo_salida)
                    messagebox.showinfo("Éxito", f"Orden generada con {productos_procesados} productos\nGuardada en:\n{archivo_salida}")
            else:
                messagebox.showwarning("Advertencia", "No se encontraron coincidencias con la plantilla")

        except Exception as e:
            messagebox.showerror("Error", f"Error al generar orden:\n{str(e)}")

    def gestionar_prioridades(self):
        try:
            # Crear ventana
            ventana_prioridades = tk.Toplevel(self.root)
            ventana_prioridades.title("Gestión de Prioridades")
            ventana_prioridades.geometry("800x600")
            
            # Frame para controles
            frame_controles = tk.Frame(ventana_prioridades)
            frame_controles.pack(pady=10)
            
            # Buscador
            tk.Label(frame_controles, text="Buscar:").pack(side=tk.LEFT)
            self.entry_buscar_prioridad = tk.Entry(frame_controles, width=30)
            self.entry_buscar_prioridad.pack(side=tk.LEFT, padx=5)
            btn_buscar = tk.Button(frame_controles, text="Buscar", command=self.actualizar_lista_prioridades)
            btn_buscar.pack(side=tk.LEFT)
            
            # Tabla de productos
            columnas = ("Código", "Nombre", "Prioridad")
            self.tabla_prioridades = ttk.Treeview(
                ventana_prioridades, 
                columns=columnas, 
                show="headings",
                selectmode="browse"
            )
            
            for col in columnas:
                self.tabla_prioridades.heading(col, text=col)
                self.tabla_prioridades.column(col, width=100, anchor="center")
            
            self.tabla_prioridades.pack(fill="both", expand=True, padx=10, pady=5)
            
            # Frame para cambiar prioridad
            frame_cambiar = tk.Frame(ventana_prioridades)
            frame_cambiar.pack(pady=10)
            
            tk.Label(frame_cambiar, text="Cambiar prioridad a:").pack(side=tk.LEFT)
            
            self.opcion_prioridad = tk.StringVar(value="baja")
            tk.Radiobutton(frame_cambiar, text="Baja", variable=self.opcion_prioridad, value="baja").pack(side=tk.LEFT, padx=5)
            tk.Radiobutton(frame_cambiar, text="Media", variable=self.opcion_prioridad, value="media").pack(side=tk.LEFT, padx=5)
            tk.Radiobutton(frame_cambiar, text="Alta", variable=self.opcion_prioridad, value="alta").pack(side=tk.LEFT, padx=5)
            
            btn_actualizar = tk.Button(
                frame_cambiar, 
                text="Aplicar Cambio", 
                command=self.actualizar_prioridad,
                bg="#4CAF50",
                fg="white"
            )
            btn_actualizar.pack(side=tk.LEFT, padx=10)
            
            # Cargar datos iniciales
            self.actualizar_lista_prioridades()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la interfaz: {str(e)}")

    def actualizar_lista_prioridades(self):
        try:
            # Limpiar tabla
            for item in self.tabla_prioridades.get_children():
                self.tabla_prioridades.delete(item)
                
            # Obtener término de búsqueda
            busqueda = self.entry_buscar_prioridad.get().strip()
            
            # Consultar productos
            query = '''
            SELECT codigo, nombre, prioridad 
            FROM productos
            WHERE codigo LIKE ? OR nombre LIKE ?
            ORDER BY codigo
            '''
            params = (f"%{busqueda}%", f"%{busqueda}%")
            
            productos = self.db.ejecutar_consulta(query, params)
            
            # Llenar tabla
            for prod in productos:
                self.tabla_prioridades.insert("", "end", values=prod)
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los productos: {str(e)}")

    def actualizar_prioridad(self):
        try:
            # Obtener producto seleccionado
            seleccion = self.tabla_prioridades.selection()
            if not seleccion:
                messagebox.showwarning("Advertencia", "Selecciona un producto primero")
                return
                
            # Obtener datos
            item = self.tabla_prioridades.item(seleccion[0])
            codigo = item['values'][0]
            nueva_prioridad = self.opcion_prioridad.get()
            
            # Actualizar en BD
            self.db.cursor.execute('''
            UPDATE productos 
            SET prioridad = ?
            WHERE codigo = ?
            ''', (nueva_prioridad, codigo))
            self.db.conn.commit()
            
            # Actualizar lista
            self.actualizar_lista_prioridades()
            messagebox.showinfo("Éxito", f"Prioridad actualizada para {item['values'][1]}")
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar: {str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = AppVentas(root)
    root.mainloop()