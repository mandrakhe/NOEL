# operaciones.py

import os
from tkinter import ttk
import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from openpyxl.styles import Font, PatternFill
from natsort import natsort_keygen
from openpyxl.utils import get_column_letter

def exportar_a_excel(contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max):
    try:
        # Obtener la ruta del escritorio dependiendo del sistema operativo
        if os.name == 'nt':  # Para Windows
            desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        else:  # Para macOS y Linux
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

        # Definir el nombre base del archivo
        base_filename = 'contenedores_exportados'
        version = 1
        filename = f'{base_filename}({version}).xlsx'
        filepath = os.path.join(desktop_path, filename)

        # Incrementar el número de versión si el archivo ya existe
        while os.path.exists(filepath):
            version += 1
            filename = f'{base_filename}({version}).xlsx'
            filepath = os.path.join(desktop_path, filename)

        # Crear la clave de ordenación natural
        natsort_key = natsort_keygen()

        # Guardar el archivo en el escritorio
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for i, (volumen_contenedor, mensajes) in enumerate(zip(contenedores_volumenes, mensajes_contenedores)):
                # Filtrar los datos del contenedor actual utilizando índices únicos
                df_contenedor = df_final.loc[mensajes].copy()

                # Asegurar que las columnas sean numéricas
                for col in ['Bruto', 'Neto', 'Volumen', 'Importe', 'LibrUtiliz', 'Contador']:
                    if col in df_contenedor.columns:
                        df_contenedor[col] = pd.to_numeric(df_contenedor[col], errors='coerce').fillna(0)

                # Asegurar que las columnas 'Lote' y 'Doc.comer.' sean de tipo string
                if 'Lote' in df_contenedor.columns:
                    df_contenedor['Lote'] = df_contenedor['Lote'].astype(str)
                if 'Doc.comer.' in df_contenedor.columns:
                    df_contenedor['Doc.comer.'] = df_contenedor['Doc.comer.'].astype(str)

                # Crear una clave de ordenación natural para 'Lote'
                if 'Lote' in df_contenedor.columns:
                    df_contenedor['Lote_sort_key'] = df_contenedor['Lote'].map(natsort_key)

                    # Ordenar df_contenedor por 'Doc.comer.' y luego por 'Lote' utilizando natsort
                    df_contenedor.sort_values(by=['Doc.comer.', 'Lote_sort_key'], inplace=True)

                    # Eliminar la columna auxiliar 'Lote_sort_key'
                    df_contenedor.drop(columns=['Lote_sort_key'], inplace=True)

                # Excluir 'Volumen_LibrUtiliz' y 'Color_Volumen' del export
                columns_to_export = [col for col in df_contenedor.columns if col not in ['Volumen_LibrUtiliz', 'Color_Volumen']]

                # Verificar que 'Nombre' esté incluido
                if 'Nombre' not in columns_to_export:
                    columns_to_export.append('Nombre')

                # Guardar los datos del contenedor en una hoja de Excel sin las columnas excluidas
                hoja_nombre = f'Contenedor_{i+1}'
                df_contenedor[columns_to_export].to_excel(writer, sheet_name=hoja_nombre, index=False)
                print(f"Exportando {len(df_contenedor)} filas al archivo Excel en la hoja '{hoja_nombre}'.")

                # Obtener el libro y la hoja de trabajo
                workbook = writer.book
                worksheet = writer.sheets[hoja_nombre]

                # Forzar el cálculo automático de fórmulas
                workbook.calcMode = 'auto'

                # ----------- MODIFICACIÓN PARA UBICAR LOS TOTALES EN O Y P -----------
                # Definir las columnas fijas para los totales
                total_label_col = 'O'
                total_value_col = 'P'

                # Crear una fuente en negrita
                bold_font = Font(bold=True)

                # Escribir los títulos de los totales
                total_names = ['Total peso Bruto', 'Total peso Neto', 'Total Volumen', 'Total Importe', 'Total LibrUtiliz']
                totals_start_row = df_contenedor.shape[0] + 2  # Espacio debajo de los datos
                for idx, total_name in enumerate(total_names):
                    cell_row = totals_start_row + idx
                    worksheet[f'{total_label_col}{cell_row}'] = total_name
                    worksheet[f'{total_label_col}{cell_row}'].font = bold_font

                # Correspondencia de columnas para las fórmulas
                bruto_col_idx = df_contenedor.columns.get_loc('Bruto') + 1
                neto_col_idx = df_contenedor.columns.get_loc('Neto') + 1
                volumen_col_idx = df_contenedor.columns.get_loc('Volumen') + 1
                importe_col_idx = df_contenedor.columns.get_loc('Importe') + 1
                contador_col_idx = df_contenedor.columns.get_loc('Contador') + 1
                librutiliz_col_idx = df_contenedor.columns.get_loc('LibrUtiliz') + 1

                bruto_col_letter = get_column_letter(bruto_col_idx)
                neto_col_letter = get_column_letter(neto_col_idx)
                volumen_col_letter = get_column_letter(volumen_col_idx)
                importe_col_letter = get_column_letter(importe_col_idx)
                contador_col_letter = get_column_letter(contador_col_idx)
                librutiliz_col_letter = get_column_letter(librutiliz_col_idx)

                # Definir el rango de las fórmulas desde la fila 2 hasta la última fila con datos
                start_row = 2
                end_row = df_contenedor.shape[0] + 1  # Ajustar al número real de filas

                # Escribir las fórmulas de los totales
                # Total peso Bruto: SUMPRODUCT((MINUSC(N2:N100)<>"x")*(Bruto * LibrUtiliz))
                worksheet[f'{total_value_col}{totals_start_row}'] = f"=SUMPRODUCT((LOWER($N${start_row}:$N${end_row})<>\"x\")*({bruto_col_letter}${start_row}:{bruto_col_letter}${end_row})*({librutiliz_col_letter}${start_row}:{librutiliz_col_letter}${end_row}))"

                # Total peso Neto: SUMPRODUCT((MINUSC(N2:N100)<>"x")*(Neto * LibrUtiliz))
                worksheet[f'{total_value_col}{totals_start_row + 1}'] = f"=SUMPRODUCT((LOWER($N${start_row}:$N${end_row})<>\"x\")*({neto_col_letter}${start_row}:{neto_col_letter}${end_row})*({librutiliz_col_letter}${start_row}:{librutiliz_col_letter}${end_row}))"

                # Total Volumen: SUMPRODUCT((MINUSC(N2:N100)<>"x")*(Volumen * LibrUtiliz))
                worksheet[f'{total_value_col}{totals_start_row + 2}'] = f"=SUMPRODUCT((LOWER($N${start_row}:$N${end_row})<>\"x\")*({volumen_col_letter}${start_row}:{volumen_col_letter}${end_row})*({librutiliz_col_letter}${start_row}:{librutiliz_col_letter}${end_row}))"

                # Total Importe: SUMPRODUCT((MINUSC(N2:N100)<>"x")*(Importe * Contador * LibrUtiliz))
                worksheet[f'{total_value_col}{totals_start_row + 3}'] = f"=SUMPRODUCT((LOWER($N${start_row}:$N${end_row})<>\"x\")*({importe_col_letter}${start_row}:{importe_col_letter}${end_row})*({contador_col_letter}${start_row}:{contador_col_letter}${end_row})*({librutiliz_col_letter}${start_row}:{librutiliz_col_letter}${end_row}))"

                # Total LibrUtiliz: SUMPRODUCT((MINUSC(N2:N100)<>"x")*(LibrUtiliz))
                worksheet[f'{total_value_col}{totals_start_row + 4}'] = f"=SUMPRODUCT((LOWER($N${start_row}:$N${end_row})<>\"x\")*({librutiliz_col_letter}${start_row}:{librutiliz_col_letter}${end_row}))"

                # Ajustar el ancho de la columna de los totales después de definir 'total_value_col'
                length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in worksheet[total_value_col])
                adjusted_width = length + 2  # Ajuste para agregar un pequeño margen
                worksheet.column_dimensions[total_value_col].width = adjusted_width

                # Aplicar el color amarillo a las celdas de 'Volumen' donde 'Color_Volumen' es 'yellow'
                fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                for idx, row in enumerate(df_contenedor.itertuples(), start=2):
                    # Verificar si 'Color_Volumen' es 'yellow' para la fila actual
                    if hasattr(row, 'Color_Volumen') and getattr(row, 'Color_Volumen') == 'yellow':
                        worksheet[f'{volumen_col_letter}{idx}'].fill = fill_yellow

        # Verificar que todas las filas hayan sido exportadas
        filas_exportadas = sum(len(mensajes) for mensajes in mensajes_contenedores)
        total_filas = df_final.shape[0]
        if filas_exportadas != total_filas:
            messagebox.showerror("Error de Exportación", f"Se han exportado {filas_exportadas} filas, pero el total es {total_filas}.")
            print(f"⚠️ Error de Exportación: Se han exportado {filas_exportadas} filas, pero el total es {total_filas}.")
        else:
            print("✅ Todas las filas han sido exportadas correctamente a Excel.")
            messagebox.showinfo("Exportación exitosa", f"Los contenedores se han exportado a '{filename}' en el escritorio.")
            print(f"Exportación completada: {filename}")

    except Exception as e:
        messagebox.showerror("Error de Exportación", f"Ocurrió un error al exportar: {e}")
        print(f"Error durante la exportación: {e}")

def calcular_totales(df):
    # Asegurar que las columnas sean numéricas
    for col in ['Bruto', 'Neto', 'Volumen', 'Importe', 'LibrUtiliz', 'Contador']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Calcular totales multiplicando por 'LibrUtiliz' donde corresponda
    total_bruto = round((df['Bruto'] * df['LibrUtiliz']).sum(), 2)
    total_neto = round((df['Neto'] * df['LibrUtiliz']).sum(), 2)
    total_volumen = round((df['Volumen'] * df['LibrUtiliz']).sum(), 2)
    total_importe = round((df['Importe'] * df['Contador'] * df['LibrUtiliz']).sum(), 2)
    total_librUtiliz = round(df['LibrUtiliz'].sum(), 2)

    return {
        'Total peso Bruto': total_bruto,
        'Total peso Neto': total_neto,
        'Total Volumen': total_volumen,
        'Total Importe': total_importe,
        'Total LibrUtiliz': total_librUtiliz
    }

def calcular_contenedores(df, columnas_mapeadas, capacidad_contenedor_max):
    # Ordenar el DataFrame por la columna 'Lote' de menor a mayor
    df['Lote'] = pd.to_numeric(df['Lote'], errors='coerce')  # Aseguramos que la columna Lote sea numérica
    df = df.sort_values(by='Lote')

    # Inicializar listas de contenedores y mensajes
    contenedores_volumenes = []
    mensajes_contenedores = []

    # Crear una nueva columna 'Volumen_LibrUtiliz'
    df['Volumen_LibrUtiliz'] = df['Volumen'] * df['LibrUtiliz']
    df['Color_Volumen'] = ''

    # Inicializar variables
    current_container = []
    cumulative_vol = 0
    total_rows = df.shape[0]
    processed_rows = 0

    for index, row in df.iterrows():
        processed_rows += 1
        vol_lib_util = row['Volumen_LibrUtiliz']
        mensaje = index  # Usar el índice único como identificador

        # Si el volumen del mensaje excede la capacidad, creamos un contenedor para él solo
        if vol_lib_util > capacidad_contenedor_max:
            # Marcar la fila para colorear
            df.at[index, 'Color_Volumen'] = 'yellow'
            # Incluir la fila en su propio contenedor
            contenedores_volumenes.append(vol_lib_util)
            mensajes_contenedores.append([mensaje])
            print(f"Fila {processed_rows}/{total_rows} asignada a Contenedor individual por exceso de volumen.")
            continue

        # Si el mensaje no cabe en el contenedor actual, llenamos el contenedor y empezamos uno nuevo
        if cumulative_vol + vol_lib_util > capacidad_contenedor_max:
            # Finalizar contenedor actual
            if current_container:
                contenedores_volumenes.append(cumulative_vol)
                mensajes_contenedores.append(current_container)
                print(f"Contenedor finalizado con volumen total: {cumulative_vol:.2f}")

            # Iniciar un nuevo contenedor con la fila actual
            current_container = [mensaje]
            cumulative_vol = vol_lib_util
            print(f"Fila {processed_rows}/{total_rows} iniciando un nuevo Contenedor. Volumen: {cumulative_vol:.2f}")
        else:
            # Agregar a contenedor actual
            current_container.append(mensaje)
            cumulative_vol += vol_lib_util
            print(f"Fila {processed_rows}/{total_rows} agregada al Contenedor actual. Volumen acumulado: {cumulative_vol:.2f}")

    # Si hay un contenedor actual al final, agregarlo
    if current_container:
        contenedores_volumenes.append(cumulative_vol)
        mensajes_contenedores.append(current_container)
        print(f"Último Contenedor finalizado con volumen total: {cumulative_vol:.2f}")

    print(f"Total de filas procesadas: {processed_rows}/{total_rows}")
    print(f"Total de contenedores creados: {len(contenedores_volumenes)}")

    # Validar que todas las filas hayan sido asignadas a algún contenedor
    filas_asignadas = sum(len(mensajes) for mensajes in mensajes_contenedores)
    if filas_asignadas != total_rows:
        print(f"⚠️ Error: Se han asignado {filas_asignadas} filas, pero el total es {total_rows}.")
    else:
        print("✅ Todas las filas han sido asignadas correctamente a contenedores.")

    return contenedores_volumenes, mensajes_contenedores

def mostrar_resultados(totales, contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max):
    root = tk.Tk()
    root.title("Resultados de Contenedores")
    root.geometry("900x600")  # Ajuste del tamaño de la ventana

    # Configurar estilo
    style = ttk.Style(root)
    style.theme_use("clam")  # Puedes cambiar el tema según preferencia

    # Definir estilos personalizados
    style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#4CAF50", foreground="white")
    style.configure("Treeview", font=("Arial", 11), rowheight=25, fieldbackground="#f0f0f0")
    style.map("Treeview", background=[("selected", "#ADD8E6")], foreground=[("selected", "black")])

    # Frame para Totales
    frame_totales = ttk.LabelFrame(root, text="Totales", padding=(20, 10))
    frame_totales.pack(fill='x', padx=20, pady=10)

    # Treeview para Totales
    columns_totales = ("Descripción", "Valor")
    tree_totales = ttk.Treeview(frame_totales, columns=columns_totales, show="headings", height=5)
    tree_totales.heading("Descripción", text="Descripción")
    tree_totales.heading("Valor", text="Valor")
    tree_totales.column("Descripción", anchor='w', width=200)
    tree_totales.column("Valor", anchor='center', width=200)

    # Insertar datos de totales
    for desc, valor in totales.items():
        tree_totales.insert("", "end", values=(desc, f"{valor:,.2f}"))

    # Scrollbar para Totales
    scrollbar_totales = ttk.Scrollbar(frame_totales, orient="vertical", command=tree_totales.yview)
    tree_totales.configure(yscroll=scrollbar_totales.set)
    scrollbar_totales.pack(side='right', fill='y')
    tree_totales.pack(fill='x')

    # Frame para Contenedores
    frame_contenedores = ttk.LabelFrame(root, text="Contenedores", padding=(20, 10))
    frame_contenedores.pack(fill='both', expand=True, padx=20, pady=10)

    # Treeview para Contenedores
    columns_contenedores = ("Número de Contenedor", "Peso Neto (unidades)")
    tree_contenedores = ttk.Treeview(frame_contenedores, columns=columns_contenedores, show="headings", height=9)
    tree_contenedores.heading("Número de Contenedor", text="Número de Contenedor")
    tree_contenedores.heading("Peso Neto (unidades)", text="Peso Neto (unidades)")

    tree_contenedores.column("Número de Contenedor", anchor='center', width=200)
    tree_contenedores.column("Peso Neto (unidades)", anchor='center', width=200)

    # Insertar datos de contenedores
    for i, contenedor in enumerate(contenedores_volumenes, start=1):
        tree_contenedores.insert("", "end", iid=i-1, values=(f"Contenedor {i}", f"{contenedor:,.2f}"))

    # Scrollbar para Contenedores
    scrollbar_contenedores = ttk.Scrollbar(frame_contenedores, orient="vertical", command=tree_contenedores.yview)
    tree_contenedores.configure(yscroll=scrollbar_contenedores.set)
    scrollbar_contenedores.pack(side='right', fill='y')
    tree_contenedores.pack(fill='both', expand=True)

    def on_select(event):
        selected_item = tree_contenedores.selection()
        if not selected_item:
            return

        contenedor_index = int(selected_item[0])

        # Crear una nueva ventana para mostrar los mensajes
        mensajes_ventana = tk.Toplevel(root)
        mensajes_ventana.title(f"Mensajes para Contenedor {contenedor_index + 1}")
        mensajes_ventana.geometry("800x400")  # Ajuste del tamaño de la ventana de mensajes

        # Configurar estilo para la nueva ventana
        style_mensajes = ttk.Style(mensajes_ventana)
        style_mensajes.theme_use("clam")
        style_mensajes.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#2196F3", foreground="white")
        style_mensajes.configure("Treeview", font=("Arial", 11), rowheight=25, fieldbackground="#f0f0f0")
        style_mensajes.map("Treeview", background=[("selected", "#ADD8E6")], foreground=[("selected", "black")])

        # Frame para los mensajes
        frame_mensajes = ttk.Frame(mensajes_ventana, padding=(20, 10))
        frame_mensajes.pack(fill='both', expand=True)

        # Treeview para Mensajes con todas las columnas
        columnas_mensajes = list(columnas_mapeadas.keys())  # Asegúrate de que esto contenga las columnas necesarias
        tree_mensajes = ttk.Treeview(frame_mensajes, columns=columnas_mensajes, show="headings", height=15)

        # Definir encabezados y columnas
        for col in columnas_mensajes:
            tree_mensajes.heading(col, text=col)
            tree_mensajes.column(col, anchor='center', width=150)

        # Insertar datos de mensajes
        for mensaje_idx in mensajes_contenedores[contenedor_index]:
            mensaje = df_final.loc[mensaje_idx]
            valores = [mensaje[col] for col in columnas_mensajes]
            tree_mensajes.insert("", "end", values=valores)

        # Scrollbar para Mensajes
        scrollbar_mensajes = ttk.Scrollbar(frame_mensajes, orient="vertical", command=tree_mensajes.yview)
        tree_mensajes.configure(yscroll=scrollbar_mensajes.set)
        scrollbar_mensajes.pack(side='right', fill='y')
        tree_mensajes.pack(fill='both', expand=True)

        def transferir_mensajes():
            seleccionados = list(tree_mensajes.selection())
            if not seleccionados:
                messagebox.showwarning("Sin Selección", "No se ha seleccionado ningún mensaje para transferir.")
                return

            destino_num = simpledialog.askinteger(
                "Transferir a Contenedor",
                f"Seleccione el número del contenedor de destino (1-{len(contenedores_volumenes)})",
                minvalue=1, maxvalue=len(contenedores_volumenes)
            )

            if destino_num is None or destino_num == contenedor_index + 1:
                return

            destino_index = destino_num - 1
            volumen_a_transferir = 0
            mensajes_a_transferir = []
            mensajes_no_transferidos = []

            for item in seleccionados:
                mensaje_seleccionado_idx = mensajes_contenedores[contenedor_index][int(item)]
                volumen_mensaje = df_final.loc[mensaje_seleccionado_idx, 'Volumen_LibrUtiliz']
                if contenedores_volumenes[destino_index] + volumen_mensaje <= capacidad_contenedor_max:
                    volumen_a_transferir += volumen_mensaje
                    mensajes_a_transferir.append(mensaje_seleccionado_idx)
                    contenedores_volumenes[destino_index] += volumen_mensaje
                else:
                    mensajes_no_transferidos.append(mensaje_seleccionado_idx)

            # Actualizar el contenedor de origen
            for mensaje in mensajes_a_transferir:
                mensajes_contenedores[contenedor_index].remove(mensaje)
                mensajes_contenedores[destino_index].append(mensaje)
                contenedores_volumenes[contenedor_index] -= df_final.loc[mensaje, 'Volumen_LibrUtiliz']

            # Mostrar advertencia si se alcanzó el límite
            if mensajes_no_transferidos:
                advertencia = "No se pudieron transferir los siguientes mensajes porque se alcanzó el límite de volumen del contenedor de destino:\n"
                advertencia += "\n".join([f"{df_final.loc[m, 'Texto de mensaje']} (Nombre: {df_final.loc[m, 'Nombre']})" for m in mensajes_no_transferidos])
                messagebox.showwarning("Advertencia", advertencia)

            # Actualizar la interfaz
            tree_mensajes.delete(*tree_mensajes.get_children())
            for mensaje_idx in mensajes_contenedores[contenedor_index]:
                mensaje = df_final.loc[mensaje_idx]
                valores = [mensaje[col] for col in columnas_mensajes]
                tree_mensajes.insert("", "end", values=valores)

            # Actualizar Treeview de contenedores
            tree_contenedores.item(selected_item, values=(f"Contenedor {contenedor_index + 1}", f"{contenedores_volumenes[contenedor_index]:,.2f}"))
            tree_contenedores.item(destino_index, values=(f"Contenedor {destino_num}", f"{contenedores_volumenes[destino_index]:,.2f}"))

            # Si un contenedor queda vacío, poner su volumen a 0
            if len(mensajes_contenedores[contenedor_index]) == 0:
                contenedores_volumenes[contenedor_index] = 0
                tree_contenedores.item(selected_item, values=(f"Contenedor {contenedor_index + 1}", f"{contenedores_volumenes[contenedor_index]:,.2f}"))

        # Botón para Transferir Mensajes
        boton_transferir = ttk.Button(mensajes_ventana, text="Transferir mensajes", command=transferir_mensajes)
        boton_transferir.pack(pady=10)

    # Bind del evento de selección
    tree_contenedores.bind("<<TreeviewSelect>>", on_select)

    # Botón para Exportar a Excel
    boton_exportar = ttk.Button(root, text="Exportar a Excel", command=lambda: exportar_a_excel(
        contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max
    ))
    boton_exportar.pack(pady=10)

    # Verificación de integridad antes de iniciar el loop
    filas_asignadas = sum(len(mensajes) for mensajes in mensajes_contenedores)
    total_filas = df_final.shape[0]
    if filas_asignadas != total_filas:
        messagebox.showerror("Error de Integridad", f"Se han asignado {filas_asignadas} filas, pero el total es {total_filas}.")
        print(f"⚠️ Error: Se han asignado {filas_asignadas} filas, pero el total es {total_filas}.")
    else:
        print("✅ Todas las filas han sido asignadas correctamente a contenedores.")

    root.mainloop()

def main_proceso(df_final, columnas_mapeadas, capacidad_contenedor_max):
    # Verificar que 'Nombre' y 'Texto de mensaje' estén presentes
    required_columns = ['Nombre', 'Texto de mensaje', 'Bruto', 'Neto', 'Volumen', 'Importe', 'LibrUtiliz', 'Contador']
    for col in required_columns:
        if col not in df_final.columns:
            messagebox.showerror("Error", f"La columna '{col}' no está presente en los datos.")
            print(f"❌ Error: La columna '{col}' no está presente en los datos.")
            return

    # Rellenar valores nulos en la columna 'Nombre' si es necesario
    df_final['Nombre'] = df_final['Nombre'].fillna('Sin Nombre')

    # Calcular contenedores
    contenedores_volumenes, mensajes_contenedores = calcular_contenedores(df_final, columnas_mapeadas, capacidad_contenedor_max)

    # Calcular totales
    totales = calcular_totales(df_final)

    # Mostrar resultados
    mostrar_resultados(
        totales, contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max
    )

def cargar_datos(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo)
        print(f"✅ Datos cargados exitosamente desde '{ruta_archivo}'.")
        return df
    except Exception as e:
        messagebox.showerror("Error de Carga", f"Ocurrió un error al cargar el archivo: {e}")
        print(f"❌ Error al cargar el archivo: {e}")
        return None

def main():
    # Crear la ventana principal para seleccionar el archivo de datos
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    messagebox.showinfo("Seleccionar Archivo", "Seleccione el archivo Excel que contiene los datos.")

    ruta_archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )

    if not ruta_archivo:
        messagebox.showwarning("Sin Selección", "No se ha seleccionado ningún archivo. El programa se cerrará.")
        print("⚠️ No se ha seleccionado ningún archivo. El programa se cerrará.")
        return

    df_final = cargar_datos(ruta_archivo)
    if df_final is None:
        return

    # Definir las columnas mapeadas si es necesario (ejemplo)
    columnas_mapeadas = {
        'Material': 'Material',
        'Texto de mensaje': 'Texto de mensaje',
        'Bruto': 'Bruto',
        'Neto': 'Neto',
        'Volumen': 'Volumen',
        'Importe': 'Importe',
        'Cliente': 'Cliente',
        'Nombre': 'Nombre',
        'Contador': 'Contador'
    }

    # Definir la capacidad máxima del contenedor (ejemplo)
    capacidad_contenedor_max = 1000  # Ajusta este valor según tus necesidades

    # Llamar a la función principal de procesamiento
    main_proceso(df_final, columnas_mapeadas, capacidad_contenedor_max)

if __name__ == "__main__":
    main()
