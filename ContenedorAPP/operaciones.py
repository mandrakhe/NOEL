# operaciones.py

import os
from tkinter import ttk
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from openpyxl.styles import Font, PatternFill
from natsort import natsort_keygen
from openpyxl.utils import get_column_letter

def exportar_a_excel(contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas):
    try:
        # Importar estilos para formatear celdas
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter

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
                # Filtrar los datos del contenedor actual
                df_contenedor = df_final[df_final[columnas_mapeadas['Texto de mensaje']].isin(mensajes)].copy()

                # Asegurar que las columnas sean numéricas
                df_contenedor['Bruto'] = pd.to_numeric(df_contenedor['Bruto'], errors='coerce').fillna(0)
                df_contenedor['Neto'] = pd.to_numeric(df_contenedor['Neto'], errors='coerce').fillna(0)
                df_contenedor['Volumen'] = pd.to_numeric(df_contenedor['Volumen'], errors='coerce').fillna(0)
                df_contenedor['Importe'] = pd.to_numeric(df_contenedor['Importe'], errors='coerce').fillna(0)
                df_contenedor['LibrUtiliz'] = pd.to_numeric(df_contenedor['LibrUtiliz'], errors='coerce').fillna(0)
                df_contenedor['Contador'] = pd.to_numeric(df_contenedor['Contador'], errors='coerce').fillna(0)

                # Asegurar que las columnas 'Lote' y 'Doc.comer.' sean de tipo string
                df_contenedor['Lote'] = df_contenedor['Lote'].astype(str)
                df_contenedor['Doc.comer.'] = df_contenedor['Doc.comer.'].astype(str)

                # Crear una clave de ordenación natural para 'Lote'
                df_contenedor['Lote_sort_key'] = df_contenedor['Lote'].map(natsort_key)

                # Ordenar df_contenedor por 'Doc.comer.' y luego por 'Lote' utilizando natsort
                df_contenedor.sort_values(by=['Doc.comer.', 'Lote_sort_key'], inplace=True)

                # Eliminar la columna auxiliar 'Lote_sort_key'
                df_contenedor.drop(columns=['Lote_sort_key'], inplace=True)

                # Excluir 'Volumen_LibrUtiliz' y 'Color_Volumen' del export
                columns_to_export = [col for col in df_contenedor.columns if col not in ['Volumen_LibrUtiliz', 'Color_Volumen']]

                # Guardar los datos del contenedor en una hoja de Excel sin las columnas excluidas
                hoja_nombre = f'Contenedor_{i+1}'
                df_contenedor[columns_to_export].to_excel(writer, sheet_name=hoja_nombre, index=False)

                # Obtener el libro y la hoja de trabajo
                workbook = writer.book
                worksheet = writer.sheets[hoja_nombre]

                # Forzar el cálculo automático de fórmulas
                workbook.calcMode = 'auto'

                # Ajustar el ancho de las columnas automáticamente
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
                    adjusted_width = length + 2
                    column_letter = get_column_letter(column_cells[0].column)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

                # ----------- MODIFICACIÓN PARA UBICAR LOS TOTALES EN Q Y R -----------
                # Definir las columnas fijas para los totales
                total_label_col = 'N'
                total_value_col = 'O'

                # Crear una fuente en negrita
                bold_font = Font(bold=True)

                # Escribir los títulos de los totales
                total_names = ['Total peso Bruto', 'Total peso Neto', 'Total Volumen', 'Total Importe', 'Total LibrUtiliz']
                totals_start_row = 2  # Comenzamos en la fila 2, debajo del encabezado
                for idx, total_name in enumerate(total_names):
                    cell_row = totals_start_row + idx
                    worksheet[f'{total_label_col}{cell_row}'] = total_name
                    worksheet[f'{total_label_col}{cell_row}'].font = bold_font

                # Correspondencia de columnas para las fórmulas
                bruto_col_letter = get_column_letter(df_contenedor.columns.get_loc('Bruto') + 1)
                neto_col_letter = get_column_letter(df_contenedor.columns.get_loc('Neto') + 1)
                volumen_col_letter = get_column_letter(df_contenedor.columns.get_loc('Volumen') + 1)
                importe_col_letter = get_column_letter(df_contenedor.columns.get_loc('Importe') + 1)
                contador_col_letter = get_column_letter(df_contenedor.columns.get_loc('Contador') + 1)
                librutiliz_col_letter = get_column_letter(df_contenedor.columns.get_loc('LibrUtiliz') + 1)

                # Definir el rango de las fórmulas desde la fila 2 hasta la última fila con datos
                start_row = 2
                end_row = df_contenedor.shape[0] + 1  # +1 porque Excel empieza en 1

                # Escribir las fórmulas de los totales
                # Total peso Bruto: SUMPRODUCT(Bruto * LibrUtiliz)
                worksheet[f'{total_value_col}{totals_start_row}'] = f"=SUMPRODUCT({bruto_col_letter}{start_row}:{bruto_col_letter}{end_row}, {librutiliz_col_letter}{start_row}:{librutiliz_col_letter}{end_row})"

                # Total peso Neto: SUMPRODUCT(Neto * LibrUtiliz)
                worksheet[f'{total_value_col}{totals_start_row + 1}'] = f"=SUMPRODUCT({neto_col_letter}{start_row}:{neto_col_letter}{end_row}, {librutiliz_col_letter}{start_row}:{librutiliz_col_letter}{end_row})"

                # Total Volumen: SUMPRODUCT(Volumen * LibrUtiliz)
                worksheet[f'{total_value_col}{totals_start_row + 2}'] = f"=SUMPRODUCT({volumen_col_letter}{start_row}:{volumen_col_letter}{end_row}, {librutiliz_col_letter}{start_row}:{librutiliz_col_letter}{end_row})"

                # Total Importe: SUMPRODUCT(Importe * Contador * LibrUtiliz)
                worksheet[f'{total_value_col}{totals_start_row + 3}'] = f"=SUMPRODUCT({importe_col_letter}{start_row}:{importe_col_letter}{end_row}, {contador_col_letter}{start_row}:{contador_col_letter}{end_row}, {librutiliz_col_letter}{start_row}:{librutiliz_col_letter}{end_row})"

                # Total LibrUtiliz: SUM(LibrUtiliz)
                worksheet[f'{total_value_col}{totals_start_row + 4}'] = f"=SUM({librutiliz_col_letter}{start_row}:{librutiliz_col_letter}{end_row})"
                # ---------------------------------------------------------------------------

                # Aplicar el color amarillo a las celdas de 'Volumen' donde 'Color_Volumen' es 'yellow'
                fill_yellow = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                for idx, row in enumerate(df_contenedor.itertuples(), start=2):
                    # Verificar si 'Color_Volumen' es 'yellow' para la fila actual
                    if hasattr(row, 'Color_Volumen') and getattr(row, 'Color_Volumen') == 'yellow':
                        worksheet[f'{volumen_col_letter}{idx}'].fill = fill_yellow

        # Mensaje de éxito
        messagebox.showinfo("Exportación exitosa", f"Los contenedores se han exportado a '{filename}' en el escritorio.")

    except Exception as e:
        messagebox.showerror("Error de Exportación", f"Ocurrió un error al exportar: {e}")

def calcular_totales(df):
    # Asegurar que las columnas sean numéricas
    df['Bruto'] = pd.to_numeric(df['Bruto'], errors='coerce').fillna(0)
    df['Neto'] = pd.to_numeric(df['Neto'], errors='coerce').fillna(0)
    df['Volumen'] = pd.to_numeric(df['Volumen'], errors='coerce').fillna(0)
    df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
    df['LibrUtiliz'] = pd.to_numeric(df['LibrUtiliz'], errors='coerce').fillna(0)
    df['Contador'] = pd.to_numeric(df['Contador'], errors='coerce').fillna(0)

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

def calcular_contenedores(df, columnas_mapeadas):
    capacidad_contenedor_min = 69
    capacidad_contenedor_max = 71.5
    contenedores_volumenes = []
    mensajes_contenedores = []

    # Crear una nueva columna 'Volumen_LibrUtiliz'
    df['Volumen_LibrUtiliz'] = df['Volumen'] * df['LibrUtiliz']
    df['Color_Volumen'] = ''

    # Inicializar variables
    current_container = []
    cumulative_vol = 0

    for index, row in df.iterrows():
        vol_lib_util = row['Volumen_LibrUtiliz']
        mensaje = row[columnas_mapeadas['Texto de mensaje']]

        if vol_lib_util > capacidad_contenedor_max:
            # Marcar la fila para colorear
            df.at[index, 'Color_Volumen'] = 'yellow'
            # Incluir la fila en su propio contenedor
            contenedores_volumenes.append(vol_lib_util)
            mensajes_contenedores.append([mensaje])
            continue

        elif capacidad_contenedor_min <= vol_lib_util <= capacidad_contenedor_max:
            # La fila forma un contenedor por sí sola
            contenedores_volumenes.append(vol_lib_util)
            mensajes_contenedores.append([mensaje])
            continue

        else:  # vol_lib_util < capacidad_contenedor_min
            if cumulative_vol + vol_lib_util <= capacidad_contenedor_max:
                # Agregar a contenedor actual
                current_container.append(mensaje)
                cumulative_vol += vol_lib_util
                if capacidad_contenedor_min <= cumulative_vol <= capacidad_contenedor_max:
                    # Finalizar contenedor actual
                    contenedores_volumenes.append(cumulative_vol)
                    mensajes_contenedores.append(current_container)
                    current_container = []
                    cumulative_vol = 0
            else:
                # No se puede agregar a contenedor actual
                if capacidad_contenedor_min <= cumulative_vol <= capacidad_contenedor_max:
                    # Finalizar contenedor actual
                    contenedores_volumenes.append(cumulative_vol)
                    mensajes_contenedores.append(current_container)
                    # Iniciar nuevo contenedor con la fila actual
                    current_container = [mensaje]
                    cumulative_vol = vol_lib_util
                else:
                    # Iniciar nuevo contenedor con la fila actual
                    current_container = [mensaje]
                    cumulative_vol = vol_lib_util

    # Si hay un contenedor actual al final, agregarlo
    if current_container:
        contenedores_volumenes.append(cumulative_vol)
        mensajes_contenedores.append(current_container)

    return contenedores_volumenes, mensajes_contenedores

def mostrar_resultados(totales, contenedores, mensajes_contenedores, df_final, columnas_mapeadas):
    root = tk.Tk()
    root.title("Resultados de Contenedores")
    root.geometry("800x600")  # Ajuste del tamaño de la ventana

    # Mostrar la tabla de totales
    frame_totales = tk.Frame(root)
    frame_totales.pack(pady=10)

    tk.Label(frame_totales, text="Descripción", font=("Arial", 12, "bold")).grid(row=0, column=0)
    tk.Label(frame_totales, text="Valor", font=("Arial", 12, "bold")).grid(row=0, column=1)

    for i, (desc, valor) in enumerate(totales.items()):
        tk.Label(frame_totales, text=desc, font=("Arial", 12)).grid(row=i+1, column=0)
        tk.Label(frame_totales, text=str(valor), font=("Arial", 12)).grid(row=i+1, column=1)

    # Crear la tabla de contenedores
    frame_contenedores = tk.Frame(root)
    frame_contenedores.pack(pady=20)

    tk.Label(frame_contenedores, text="Contenedores", font=("Arial", 14, "bold")).pack()

    tree = ttk.Treeview(frame_contenedores, columns=("Volumen"), show="headings", height=10)
    tree.heading("Volumen", text="Volumen Total")
    tree.column("Volumen", width=300)  # Ajuste del ancho de la columna
    tree.pack(expand=True, fill='both')  # Ajuste para expandir el Treeview

    for i, contenedor in enumerate(contenedores):
        tree.insert("", "end", iid=i, values=(f"Contenedor {i+1}: {contenedor:.2f} unidades de volumen",))

    def on_select(event):
        if not tree.selection():
            return  # Si no hay ningún elemento seleccionado, salir de la función.

        selected_item = tree.selection()[0]
        contenedor_index = int(selected_item)

        # Crear una nueva ventana para mostrar y mover los mensajes
        mensajes_ventana = tk.Toplevel(root)
        mensajes_ventana.title(f"Mensajes para Contenedor {contenedor_index + 1}")
        mensajes_ventana.geometry("600x400")  # Ajuste del tamaño de la ventana de mensajes

        mensajes_listbox = tk.Listbox(mensajes_ventana, selectmode=tk.MULTIPLE, width=50, height=10)
        mensajes_listbox.pack(pady=10)

        for mensaje in mensajes_contenedores[contenedor_index]:
            mensajes_listbox.insert(tk.END, mensaje)

        def transferir_mensajes():
            seleccionados = list(mensajes_listbox.curselection())
            if not seleccionados:
                return

            destino_index = tk.simpledialog.askinteger(
                "Transferir a Contenedor",
                f"Seleccione el número del contenedor de destino (1-{len(contenedores)})",
                minvalue=1, maxvalue=len(contenedores)
            )

            if destino_index is None or destino_index == contenedor_index + 1:
                return

            destino_index -= 1
            volumen_a_transferir = 0
            mensajes_a_transferir = []
            mensajes_no_transferidos = []

            for i in seleccionados:
                mensaje_seleccionado = mensajes_contenedores[contenedor_index][i]
                volumen_mensaje = df_final.loc[df_final[columnas_mapeadas['Texto de mensaje']] == mensaje_seleccionado, 'Volumen_LibrUtiliz'].values[0]
                if contenedores[destino_index] + volumen_mensaje <= 71.5:
                    volumen_a_transferir += volumen_mensaje
                    mensajes_a_transferir.append(mensaje_seleccionado)
                    contenedores[destino_index] += volumen_mensaje
                else:
                    mensajes_no_transferidos.append(mensaje_seleccionado)

            # Actualizar el contenedor de origen
            for mensaje in mensajes_a_transferir:
                mensajes_contenedores[contenedor_index].remove(mensaje)
                mensajes_contenedores[destino_index].append(mensaje)
                volumen_mensaje = df_final.loc[df_final[columnas_mapeadas['Texto de mensaje']] == mensaje, 'Volumen_LibrUtiliz'].values[0]
                contenedores[contenedor_index] -= volumen_mensaje

            if mensajes_no_transferidos:
                advertencia = f"No se pudieron transferir los siguientes mensajes porque se alcanzó el límite de volumen del contenedor de destino:\n{', '.join(map(str, mensajes_no_transferidos))}"
                messagebox.showwarning("Advertencia", advertencia)

            # Actualizar la interfaz
            mensajes_listbox.delete(0, tk.END)
            for mensaje in mensajes_contenedores[contenedor_index]:
                mensajes_listbox.insert(tk.END, mensaje)

            tree.item(selected_item, values=(f"Contenedor {contenedor_index + 1}: {contenedores[contenedor_index]:.2f} unidades de volumen",))
            tree.item(destino_index, values=(f"Contenedor {destino_index + 1}: {contenedores[destino_index]:.2f} unidades de volumen",))

            # Si un contenedor queda vacío, poner su volumen a 0
            if len(mensajes_contenedores[contenedor_index]) == 0:
                contenedores[contenedor_index] = 0

        tk.Button(mensajes_ventana, text="Transferir mensajes", command=transferir_mensajes).pack(pady=10)

    tree.bind("<<TreeviewSelect>>", on_select)

    tk.Button(root, text="Exportar a Excel", command=lambda: exportar_a_excel(contenedores, mensajes_contenedores, df_final, columnas_mapeadas)).pack(pady=10)

    root.mainloop()
