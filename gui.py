import os
import pandas as pd
from tkinter import Tk, ttk, Frame, Label, Button, Entry, messagebox, filedialog
from tkcalendar import DateEntry
from database import Database
from datetime import datetime
import matplotlib.pyplot as plt 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from math import sqrt, ceil
import re


class AppVentas:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestión de Ventas")
        self.root.geometry("1400x800")
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
        self.btn_cargar = Button(frame_controles, text="Cargar Archivo Excel", command=self.procesar_archivo)
        self.btn_cargar.pack(side="left", padx=10)

        # Campo de búsqueda
        self.label_busqueda = Label(frame_controles, text="Buscar:")
        self.label_busqueda.pack(side="left", padx=10)
        self.entry_busqueda = Entry(frame_controles, width=30)
        self.entry_busqueda.pack(side="left", padx=10)
        self.btn_buscar = Button(frame_controles, text="Buscar", command=self.buscar_datos)
        self.btn_buscar.pack(side="left", padx=10)

        # Campo para ingresar el nombre del producto
        self.label_producto = Label(frame_controles, text="Producto:")
        self.label_producto.pack(side="left", padx=10)
        self.entry_producto = Entry(frame_controles, width=30)
        self.entry_producto.pack(side="left", padx=10)

        # Botón para abrir la ventana de selección de rango de fechas y generar el gráfico
        self.btn_grafico = Button(frame_controles, text="Generar Gráfico de Ventas", command=self.abrir_ventana_grafico)
        self.btn_grafico.pack(side="left", padx=10)

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

        # Botón Analizar siguiente pedido
        #self.btn_analizar_pedidos = Button(frame_controles, text="Analizar Pedidos", command=self.analizar_pedidos)
        #self.btn_analizar_pedidos.pack(side="left", padx=10)

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
        ventana_rango_fechas = Tk()
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

            # Verificar que las columnas requeridas estén presentes
            columnas_requeridas = ['codigo', 'nombre', 'cantidad']
            if not all(col in df.columns for col in columnas_requeridas):
                messagebox.showerror("Error", "El archivo de Excel no tiene las columnas requeridas.")
                return

            # Crear una ventana para mostrar las recomendaciones
            ventana_recomendaciones = Tk()
            ventana_recomendaciones.title("Recomendaciones de Pedidos")
            ventana_recomendaciones.geometry("1200x600")

            # Crear una tabla para mostrar las recomendaciones
            columnas = ("Producto", "Inventario Actual", "Demanda Promedio Diaria", "Punto de Reorden", "Cantidad a Pedir (EOQ)", "Cajas a Pedir")
            tabla_recomendaciones = ttk.Treeview(ventana_recomendaciones, columns=columnas, show="headings")
            for col in columnas:
                tabla_recomendaciones.heading(col, text=col)
            tabla_recomendaciones.pack(fill="both", expand=True)

            # Parámetros para el cálculo de pedidos
            tiempo_entrega = 6.5  # Tiempo de entrega en días (ajustar según proveedor)
            costo_pedido = 30  # Costo fijo por hacer un pedido (ajustar según negocio)
            costo_almacenamiento = 59  # Costo de almacenamiento por unidad (ajustar según negocio)

            # Procesar cada producto en el inventario
            for index, row in df.iterrows():
                codigo = row['codigo']
                nombre = row['nombre']
                inventario_actual = row['cantidad']

                # Obtener la cantidad de unidades por caja desde la tabla productos
                self.db.cursor.execute('''
                SELECT cantidad_por_caja FROM productos WHERE codigo = ?
                ''', (codigo,))
                resultado = self.db.cursor.fetchone()
                if not resultado:
                    continue  # Si el producto no está en la tabla productos, saltar
                cantidad_por_caja = resultado[0]

                # Obtener las ventas históricas del producto
                self.db.cursor.execute('''
                SELECT fecha_carga, SUM(cantidad) as total_cantidad
                FROM ventas
                WHERE codigo = ?
                GROUP BY fecha_carga
                ORDER BY fecha_carga
                ''', (codigo,))
                ventas_producto = self.db.cursor.fetchall()

                if not ventas_producto:
                    continue

                # Calcular la demanda promedio diaria
                fechas = [venta[0] for venta in ventas_producto]
                cantidades = [venta[1] for venta in ventas_producto]
                demanda_promedio_diaria = sum(cantidades) / len(cantidades)

                # Calcular el punto de reorden
                punto_reorden = demanda_promedio_diaria * tiempo_entrega

                # Calcular la cantidad económica de pedido (EOQ)
                demanda_anual = demanda_promedio_diaria * 365  # Suponiendo 365 días al año
                eoq = sqrt((2 * demanda_anual * costo_pedido) / costo_almacenamiento)

                # Calcular el número de cajas a pedir (solo si el inventario es menor que el punto de reorden)
                if inventario_actual < punto_reorden:
                    cajas_a_pedir = (ceil((punto_reorden - inventario_actual + eoq) / cantidad_por_caja)-1)
                else:
                    cajas_a_pedir = 0

                # Insertar los resultados en la tabla
                tabla_recomendaciones.insert("", "end", values=(
                    nombre,
                    inventario_actual,
                    round(demanda_promedio_diaria, 2),
                    round(punto_reorden, 2),
                    round(eoq, 2),
                    cajas_a_pedir
                ))

            ventana_recomendaciones.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al cargar el inventario o calcular los pedidos: {e}")


    