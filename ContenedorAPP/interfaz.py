import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from operaciones import calcular_totales, calcular_contenedores, mostrar_resultados

class Interfaz:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Subir archivo")
        self.root.geometry("600x750")  # Ajustamos el tamaño de la ventana
        self.root.configure(bg="#fff")

        self.label = tk.Label(
            self.root, text="Por favor, suba su archivo:", bg="#fff",
            font=("Arial", 14)
        )
        self.label.pack(pady=10)

        # Campo de entrada para buscar
        self.entry_buscar = tk.Entry(self.root, font=("Arial", 12))
        self.entry_buscar.pack(pady=5)
        self.entry_buscar.bind('<KeyRelease>', self.filtrar_datos)

        # Listbox para mostrar 'Nombre'
        self.listbox_nombres = tk.Listbox(
            self.root, selectmode=tk.MULTIPLE, width=50, height=20, exportselection=0
        )
        self.listbox_nombres.pack(pady=10)

        self.boton_subir = tk.Button(
            self.root, text="Subir archivo", command=self.subir_archivo,
            bg="#4caf50", fg="white", font=("Arial", 12)
        )
        self.boton_subir.pack(pady=10)

        # Datos completos para el filtrado
        self.df_subido = pd.DataFrame()
        self.datos_filtrados = pd.DataFrame()

        # Botón para buscar datos
        self.boton_buscar = tk.Button(
            self.root, text="Buscar datos", command=self.buscar_datos,
            bg="#2196f3", fg="white", font=("Arial", 12)
        )
        self.boton_buscar.pack(pady=10)

        # Botón para procesar selección
        self.boton_procesar = tk.Button(
            self.root, text="Procesar selección", command=self.procesar_seleccion,
            bg="#2196f3", fg="white", font=("Arial", 12)
        )
        self.boton_procesar.pack(pady=10)

        # Variable para la capacidad del contenedor
        self.capacidad_contenedor = tk.DoubleVar()
        self.capacidad_contenedor.set(33.2)  # Valor por defecto para 20ft

    def subir_archivo(self):
        archivo_path = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if archivo_path:
            self.cargar_datos(archivo_path)

    def cargar_datos(self, archivo_path):
        try:
            df_subido = pd.read_excel(archivo_path, header=0)
            df_subido.dropna(how='all', inplace=True)
            df_subido.columns = df_subido.columns.str.strip()

            # Verificar si las columnas necesarias existen
            columnas_necesarias = ['Lote', 'Nombre', 'Cliente']
            for columna in columnas_necesarias:
                if columna not in df_subido.columns:
                    messagebox.showerror(
                        "Error",
                        f"La columna '{columna}' no se encontró en el archivo."
                    )
                    return

            # Crear la columna 'Grupo' basada en la columna 'Lote'
            df_subido['Grupo'] = df_subido['Lote'].astype(str).apply(
                lambda x: 'GCFOODS' if 'GZ' in x else 'NOEL'
            )

            # Solicitar al usuario que seleccione los grupos
            self.seleccionar_grupo(df_subido)

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar los datos: {e}")

    def seleccionar_grupo(self, df_subido):
        # Crear una nueva ventana para seleccionar los grupos
        grupo_window = tk.Toplevel(self.root)
        grupo_window.title("Seleccionar Grupo")

        tk.Label(
            grupo_window, text="Seleccione el grupo de datos que desea procesar:",
            font=("Arial", 12)
        ).pack(pady=10)

        # Variables para almacenar la selección
        self.grupo_gcfoods = tk.BooleanVar(value=True)
        self.grupo_noel = tk.BooleanVar(value=True)

        # Checkbuttons para seleccionar GCFOODS y NOEL
        tk.Checkbutton(
            grupo_window, text="GCFOODS", variable=self.grupo_gcfoods
        ).pack()
        tk.Checkbutton(
            grupo_window, text="NOEL", variable=self.grupo_noel
        ).pack()

        def confirmar_seleccion():
            if not self.grupo_gcfoods.get() and not self.grupo_noel.get():
                messagebox.showwarning(
                    "Advertencia", "Debe seleccionar al menos un grupo."
                )
                return
            grupo_window.destroy()
            self.mostrar_datos(df_subido)

        tk.Button(
            grupo_window, text="Confirmar", command=confirmar_seleccion
        ).pack(pady=10)

    def mostrar_datos(self, df_subido):
        # Filtrar df_subido según los grupos seleccionados
        grupos_seleccionados = []
        if self.grupo_gcfoods.get():
            grupos_seleccionados.append('GCFOODS')
        if self.grupo_noel.get():
            grupos_seleccionados.append('NOEL')

        df_subido = df_subido[df_subido['Grupo'].isin(grupos_seleccionados)]

        # Modificar 'Lote' para el grupo GCFOODS
        df_subido.loc[df_subido['Grupo'] == 'GCFOODS', 'Lote'] = (
            '20' + df_subido.loc[df_subido['Grupo'] == 'GCFOODS', 'Lote'].astype(str)
        )

        # Asegurar que las columnas clave sean de tipo string
        df_subido['Lote'] = df_subido['Lote'].astype(str)
        df_subido = df_subido.dropna(subset=['Nombre', 'Cliente'])
        df_subido['Nombre'] = df_subido['Nombre'].astype(str).str.strip()
        df_subido['Cliente'] = df_subido['Cliente'].astype(str).str.strip()

        # Almacenar los datos para el filtrado
        self.df_subido = df_subido.copy()
        self.datos_filtrados = pd.DataFrame(
            {'Nombre': self.df_subido['Nombre'].unique()}
        )

        # Mostrar todos los datos inicialmente
        self.actualizar_listbox(self.datos_filtrados['Nombre'].tolist())

    def actualizar_listbox(self, lista_nombres):
        self.listbox_nombres.delete(0, tk.END)
        for nombre in lista_nombres:
            self.listbox_nombres.insert(tk.END, nombre)

    def filtrar_datos(self, event):
        texto_busqueda = self.entry_buscar.get().lower()
        if texto_busqueda == '':
            nombres_filtrados = self.df_subido['Nombre'].unique()
        else:
            # Filtrar donde el texto coincida en 'Nombre' o 'Cliente'
            df_filtrado = self.df_subido[
                self.df_subido['Nombre'].str.lower().str.contains(texto_busqueda) |
                self.df_subido['Cliente'].str.lower().str.contains(texto_busqueda)
            ]
            nombres_filtrados = df_filtrado['Nombre'].unique()
        self.actualizar_listbox(nombres_filtrados)

    def procesar_seleccion(self):
        # Obtener los nombres seleccionados en el Listbox
        seleccionados = [
            self.listbox_nombres.get(i) for i in self.listbox_nombres.curselection()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un elemento para continuar."
            )
            return

        # Filtrar df_subido según los nombres seleccionados
        df_filtrado = self.df_subido[self.df_subido['Nombre'].isin(seleccionados)]

        # Antes de procesar el archivo, seleccionar el tipo de contenedor
        self.seleccionar_contenedor(df_filtrado)

    def seleccionar_contenedor(self, df_filtrado):
        # Crear una nueva ventana para seleccionar el tipo de contenedor
        contenedor_window = tk.Toplevel(self.root)
        contenedor_window.title("Seleccionar Tipo de Contenedor")

        tk.Label(
            contenedor_window, text="Seleccione el tipo de contenedor:", font=("Arial", 12)
        ).pack(pady=10)

        # Función para actualizar la capacidad según la selección
        def actualizar_capacidad():
            seleccion = opcion_seleccionada.get()
            if seleccion == '20ft':
                self.capacidad_contenedor.set(33.2)
            elif seleccion == '40ft':
                self.capacidad_contenedor.set(67.4)
            elif seleccion == '40 high':
                self.capacidad_contenedor.set(76.4)
            elif seleccion == 'Otro':
                # Pedir al usuario que ingrese la capacidad
                capacidad = simpledialog.askfloat(
                    "Capacidad del Contenedor",
                    "Ingrese la capacidad máxima del contenedor:",
                    minvalue=0.1
                )
                if capacidad:
                    self.capacidad_contenedor.set(capacidad)
                else:
                    self.capacidad_contenedor.set(0)  # Valor por defecto si cancela
            else:
                self.capacidad_contenedor.set(0)

        # Variable para la opción seleccionada
        opcion_seleccionada = tk.StringVar()
        opcion_seleccionada.set('20ft')  # Valor por defecto

        # Opciones de contenedores
        opciones = ['20ft', '40ft', '40 high', 'Otro']

        for opcion in opciones:
            tk.Radiobutton(
                contenedor_window,
                text=opcion,
                variable=opcion_seleccionada,
                value=opcion,
                command=actualizar_capacidad
            ).pack(anchor='w')

        def confirmar_seleccion():
            if self.capacidad_contenedor.get() <= 0:
                messagebox.showwarning(
                    "Advertencia", "Debe seleccionar un tipo de contenedor válido."
                )
                return
            contenedor_window.destroy()
            # Después de cerrar la ventana, procesar el archivo
            self.procesar_archivo(df_filtrado)

        tk.Button(
            contenedor_window, text="Confirmar", command=confirmar_seleccion
        ).pack(pady=10)

    def procesar_archivo(self, df_filtrado):
        try:
            columnas_necesarias = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'  # Agregamos 'Contador'
            ]

            # Intentar leer el archivo Portafolio con diferentes filas de encabezado
            df_portafolio = None
            for header_row in [0, 1, 2, 3, 4]:
                df_temp = pd.read_excel(
                    "Data_Base\PortafoliocompletointernacionalJulio2024.xlsx",
                    header=header_row
                )
                df_temp.dropna(how='all', inplace=True)
                df_temp.columns = df_temp.columns.str.strip()
                if any(col in df_temp.columns for col in columnas_necesarias):
                    df_portafolio = df_temp
                    break

            if df_portafolio is None:
                raise ValueError(
                    "No se encontraron las columnas necesarias en el portafolio."
                )

            # Mapear las columnas
            def find_closest_column(column_name, df_columns):
                for col in df_columns:
                    if column_name.lower() == col.lower():
                        return col
                for col in df_columns:
                    if column_name.lower() in col.lower():
                        return col
                return None

            columnas_mapeadas = {}
            for columna in columnas_necesarias:
                columna_encontrada = find_closest_column(columna, df_portafolio.columns)
                if columna_encontrada:
                    columnas_mapeadas[columna] = columna_encontrada
                else:
                    if columna != 'Contador':
                        raise ValueError(
                            f"No se encontró una columna similar a '{columna}' "
                            "en el portafolio."
                        )

            # Verificar y eliminar valores nulos en las columnas clave
            df_filtrado.dropna(
                subset=['Material', 'Texto breve material', 'Cliente', 'Nombre'],
                inplace=True
            )
            df_portafolio.dropna(
                subset=[
                    columnas_mapeadas['Material'],
                    columnas_mapeadas['Texto de mensaje'],
                    columnas_mapeadas['Cliente'],
                    columnas_mapeadas['Nombre']
                ],
                inplace=True
            )

            # Convertir las columnas clave a string y eliminar espacios en blanco
            claves_filtrado = ['Material', 'Texto breve material', 'Cliente', 'Nombre']
            claves_portafolio = [
                columnas_mapeadas['Material'],
                columnas_mapeadas['Texto de mensaje'],
                columnas_mapeadas['Cliente'],
                columnas_mapeadas['Nombre']
            ]

            for col in claves_filtrado:
                df_filtrado[col] = df_filtrado[col].astype(str).str.strip()
            for col in claves_portafolio:
                df_portafolio[col] = df_portafolio[col].astype(str).str.strip()

            # Realizar el merge
            df_resultado = pd.merge(
                df_filtrado,
                df_portafolio[list(columnas_mapeadas.values())],
                left_on=claves_filtrado,
                right_on=claves_portafolio,
                how='inner',
                suffixes=('_filtrado', '_portafolio')
            )

            if df_resultado.empty:
                messagebox.showerror(
                    "Error",
                    "No se encontraron coincidencias entre los archivos."
                )
                return


            # Columnas adicionales que queremos agregar del archivo subido
            columnas_adicionales = ['Doc.comer.', 'LibrUtiliz', 'Grupo', 'Lote']

            # Determinar el nombre correcto de las columnas adicionales
            columnas_adicionales_encontradas = []
            for col_adicional in columnas_adicionales:
                if f"{col_adicional}_filtrado" in df_resultado.columns:
                    columnas_adicionales_encontradas.append(f"{col_adicional}_filtrado")
                elif col_adicional in df_resultado.columns:
                    columnas_adicionales_encontradas.append(col_adicional)
                else:
                    messagebox.showwarning(
                        "Advertencia",
                        f"'{col_adicional}' no se encontró en los datos."
                    )

            # Filtrar las columnas necesarias
            columnas_para_df_final = list(columnas_mapeadas.values()) + \
                columnas_adicionales_encontradas
            df_final = df_resultado[columnas_para_df_final]

            # Renombrar las columnas adicionales
            columnas_renombradas = {}
            for col in columnas_adicionales_encontradas:
                if '_filtrado' in col:
                    columnas_renombradas[col] = col.replace('_filtrado', '')
            df_final.rename(columns=columnas_renombradas, inplace=True)

            # Eliminar comas y convertir a numérico
            for col in ['Bruto', 'Neto', 'Volumen', 'Importe']:
                df_final[col] = df_final[col].replace({',': ''}, regex=True)
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

            # Convertir 'LibrUtiliz' a numérico si no lo es
            if 'LibrUtiliz' in df_final.columns:
                df_final['LibrUtiliz'] = pd.to_numeric(
                    df_final['LibrUtiliz'], errors='coerce'
                ).fillna(0)

            # Calcular totales y contenedores
            totales = calcular_totales(df_final)
            total_neto = totales.get("Total peso Neto", 0)

            if total_neto == 0:
                raise ValueError(
                    "No se pudo calcular el Total peso Neto. Verifica los valores."
                )

            contenedores_pesos, mensajes_contenedores = calcular_contenedores(
                df_final, columnas_mapeadas, self.capacidad_contenedor.get()
            )

            mostrar_resultados(
                totales, contenedores_pesos, mensajes_contenedores,
                df_final, columnas_mapeadas, self.capacidad_contenedor.get()
            )

        except Exception as e:
            messagebox.showerror("Error", f"Error procesando el archivo: {e}")

    def buscar_datos(self):
        # Obtener los nombres seleccionados en el Listbox
        seleccionados = [
            self.listbox_nombres.get(i) for i in self.listbox_nombres.curselection()
        ]

        if not seleccionados:
            messagebox.showwarning(
                "Advertencia",
                "Debe seleccionar al menos un nombre en la lista para buscar."
            )
            return

        # Si hay múltiples nombres seleccionados, tomaremos el primero
        nombre_seleccionado = seleccionados[0]

        # Ventana para ingresar los datos de búsqueda
        input_window = tk.Toplevel(self.root)
        input_window.title("Buscar datos")

        tk.Label(
            input_window, text="Material:", font=("Arial", 12)
        ).grid(row=0, column=0, padx=10, pady=5)
        material_entry = tk.Entry(input_window, font=("Arial", 12))
        material_entry.grid(row=0, column=1, padx=10, pady=5)

        # Mostrar el nombre seleccionado
        tk.Label(
            input_window, text=f"Nombre seleccionado: {nombre_seleccionado}",
            font=("Arial", 12)
        ).grid(row=1, column=0, columnspan=2, padx=10, pady=5)

        tk.Button(
            input_window, text="Buscar",
            command=lambda: self.realizar_busqueda(
                material_entry.get(),
                nombre_seleccionado,
                input_window
            ),
            bg="#4caf50", fg="white", font=("Arial", 12)
        ).grid(row=2, column=0, columnspan=2, pady=10)

    def realizar_busqueda(self, material, nombre, input_window):
        # Mantener abierta la ventana de búsqueda
        # input_window.destroy()

        # Cargar el portafolio
        df_portafolio, columnas_mapeadas = self.cargar_portafolio()

        if df_portafolio is None:
            return  # Error ya mostrado en cargar_portafolio

        # Asegurarse de que las columnas son cadenas y eliminar espacios
        df_portafolio[columnas_mapeadas['Material']] = df_portafolio[
            columnas_mapeadas['Material']
        ].astype(str).str.strip()
        df_portafolio[columnas_mapeadas['Nombre']] = df_portafolio[
            columnas_mapeadas['Nombre']
        ].astype(str).str.strip()
        df_portafolio[columnas_mapeadas['Texto de mensaje']] = df_portafolio[
            columnas_mapeadas['Texto de mensaje']
        ].astype(str).str.strip()
        df_portafolio[columnas_mapeadas['Cliente']] = df_portafolio[
            columnas_mapeadas['Cliente']
        ].astype(str).str.strip()

        # Realizar la búsqueda
        mask = (
            (df_portafolio[columnas_mapeadas['Material']] == material.strip()) &
            (df_portafolio[columnas_mapeadas['Nombre']] == nombre.strip())
        )
        df_resultado = df_portafolio[mask]

        if df_resultado.empty:
            messagebox.showinfo(
                "Sin resultados",
                "No se encontraron coincidencias con los datos ingresados."
            )
        else:
            # Copiar los datos al portapapeles en formato Excel
            columnas_a_mostrar = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'
            ]
            self.copiar_al_portapapeles(df_resultado, columnas_mapeadas, columnas_a_mostrar)

    def copiar_al_portapapeles(self, df_resultado, columnas_mapeadas, columnas_a_copiar):
        # Obtener los datos en el orden de las columnas
        datos = df_resultado[[columnas_mapeadas[col] for col in columnas_a_copiar]].copy()

        # Formatear números con comas en lugar de puntos
        def format_number_with_comma(x):
            try:
                return ('{0:.3f}'.format(float(x))).replace('.', ',')
            except:
                return x

        for col in ['Bruto', 'Neto', 'Volumen', 'Importe']:
            mapped_col = columnas_mapeadas.get(col)
            if mapped_col in datos.columns:
                datos[mapped_col] = datos[mapped_col].apply(format_number_with_comma)

        # Convertir los datos a texto separado por tabulaciones
        datos_string = datos.to_csv(sep='\t', index=False, header=False)

        # Copiar al portapapeles
        self.root.clipboard_clear()
        self.root.clipboard_append(datos_string)

        messagebox.showinfo(
            "Datos copiados",
            "Los datos han sido copiados al portapapeles en formato Excel."
        )

    def cargar_portafolio(self):
        try:
            columnas_necesarias = [
                'Material', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen',
                'Importe', 'Cliente', 'Nombre', 'Contador'
            ]

            # Intentar leer el archivo con diferentes filas de encabezado
            df_portafolio = None
            for header_row in [0, 1, 2, 3, 4]:
                df_temp = pd.read_excel(
                    "Data_Base\PortafoliocompletointernacionalJulio2024.xlsx",
                    header=header_row
                )
                df_temp.dropna(how='all', inplace=True)
                df_temp.columns = df_temp.columns.str.strip()
                if any(col in df_temp.columns for col in columnas_necesarias):
                    df_portafolio = df_temp
                    break

            if df_portafolio is None:
                raise ValueError(
                    "No se encontraron las columnas necesarias en el portafolio."
                )

            # Mapear las columnas
            def find_closest_column(column_name, df_columns):
                for col in df_columns:
                    if column_name.lower() == col.lower():
                        return col
                for col in df_columns:
                    if column_name.lower() in col.lower():
                        return col
                return None

            columnas_mapeadas = {}
            for columna in columnas_necesarias:
                columna_encontrada = find_closest_column(columna, df_portafolio.columns)
                if columna_encontrada:
                    columnas_mapeadas[columna] = columna_encontrada
                else:
                    if columna != 'Contador':
                        raise ValueError(
                            f"No se encontró una columna similar a '{columna}' "
                            "en el portafolio."
                        )

            return df_portafolio, columnas_mapeadas
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el portafolio: {e}")
            return None, None

    def run(self):
        self.root.mainloop()