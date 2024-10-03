import os
from tkinter import ttk
import pandas as pd
import tkinter as tk
from tkinter import messagebox, simpledialog
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
                totals_start_row = df_contenedor.shape[0] + 2  # Espacio debajo de los datos
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
                end_row = 100

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



def mostrar_resultados(totales, contenedores, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max):
    root = tk.Tk()
    root.title("Resultados de Contenedores")
    root.geometry("1000x700")  # Ajuste del tamaño de la ventana

    # Mostrar la tabla de totales
    frame_totales = tk.Frame(root)
    frame_totales.pack(pady=10)

    tk.Label(frame_totales, text="Descripción", font=("Arial", 12, "bold")).grid(row=0, column=0, padx=10, pady=5)
    tk.Label(frame_totales, text="Valor", font=("Arial", 12, "bold")).grid(row=0, column=1, padx=10, pady=5)

    for i, (desc, valor) in enumerate(totales.items()):
        tk.Label(frame_totales, text=desc, font=("Arial", 12)).grid(row=i+1, column=0, sticky='w', padx=10, pady=2)
        tk.Label(frame_totales, text=str(valor), font=("Arial", 12)).grid(row=i+1, column=1, sticky='w', padx=10, pady=2)

    # Crear la tabla de contenedores
    frame_contenedores = tk.Frame(root)
    frame_contenedores.pack(pady=20, fill='both', expand=True)

    tk.Label(frame_contenedores, text="Contenedores", font=("Arial", 14, "bold")).pack()

    tree = ttk.Treeview(frame_contenedores, columns=("Volumen"), show="headings", height=20)
    tree.heading("Volumen", text="Volumen Total")
    tree.column("Volumen", width=600, anchor='center')  # Ajuste del ancho de la columna
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
        mensajes_ventana.geometry("800x500")  # Ajuste del tamaño de la ventana de mensajes

        mensajes_listbox = tk.Listbox(mensajes_ventana, selectmode=tk.MULTIPLE, width=100, height=25)
        mensajes_listbox.pack(pady=10, padx=10, fill='both', expand=True)

        # Obtener los mensajes (índices únicos) y mostrarlos con detalles
        for mensaje_idx in mensajes_contenedores[contenedor_index]:
            texto_mensaje = df_final.loc[mensaje_idx, 'Texto de mensaje']
            nombre = df_final.loc[mensaje_idx, 'Nombre']
            mensajes_listbox.insert(tk.END, f"{texto_mensaje} (Nombre: {nombre})")

        def transferir_mensajes():
            seleccionados = list(mensajes_listbox.curselection())
            if not seleccionados:
                messagebox.showwarning("Sin Selección", "No se ha seleccionado ningún mensaje para transferir.")
                return

            destino_index = simpledialog.askinteger(
                "Transferir a Contenedor",
                f"Seleccione el número del contenedor de destino (1-{len(contenedores)})",
                minvalue=1, maxvalue=len(contenedores)
            )

            if destino_index is None:
                return  # Cancelado por el usuario
            if destino_index == contenedor_index + 1:
                messagebox.showwarning("Selección Inválida", "El contenedor de destino no puede ser el mismo que el de origen.")
                return

            destino_index -= 1  # Ajustar al índice basado en 0
            volumen_a_transferir = 0
            mensajes_a_transferir = []
            mensajes_no_transferidos = []

            # Asumimos que cada mensaje tiene un volumen específico
            for i in seleccionados:
                mensaje_seleccionado_idx = mensajes_contenedores[contenedor_index][i]
                volumen_mensaje = df_final.loc[mensaje_seleccionado_idx, 'Volumen_LibrUtiliz']
                if contenedores[destino_index] + volumen_mensaje <= capacidad_contenedor_max:
                    volumen_a_transferir += volumen_mensaje
                    mensajes_a_transferir.append(mensaje_seleccionado_idx)
                    contenedores[destino_index] += volumen_mensaje
                else:
                    mensajes_no_transferidos.append(mensaje_seleccionado_idx)

            # Actualizar el contenedor de origen
            for mensaje in mensajes_a_transferir:
                mensajes_contenedores[contenedor_index].remove(mensaje)
                mensajes_contenedores[destino_index].append(mensaje)
                contenedores[contenedor_index] -= df_final.loc[mensaje, 'Volumen_LibrUtiliz']

            # Mostrar advertencia si se alcanzó el límite
            if mensajes_no_transferidos:
                advertencia = "No se pudieron transferir los siguientes mensajes porque se alcanzó el límite de volumen del contenedor de destino:\n"
                advertencia += "\n".join([f"{df_final.loc[m, 'Texto de mensaje']} (Nombre: {df_final.loc[m, 'Nombre']})" for m in mensajes_no_transferidos])
                messagebox.showwarning("Advertencia", advertencia)

            # Actualizar la interfaz
            mensajes_listbox.delete(0, tk.END)
            for mensaje_idx in mensajes_contenedores[contenedor_index]:
                texto_mensaje = df_final.loc[mensaje_idx, 'Texto de mensaje']
                nombre = df_final.loc[mensaje_idx, 'Nombre']
                mensajes_listbox.insert(tk.END, f"{texto_mensaje} (Nombre: {nombre})")

            tree.item(selected_item, values=(f"Contenedor {contenedor_index + 1}: {contenedores[contenedor_index]:.2f} unidades de volumen",))
            tree.item(destino_index, values=(f"Contenedor {destino_index + 1}: {contenedores[destino_index]:.2f} unidades de volumen",))

            # Si un contenedor queda vacío, poner su volumen a 0
            if len(mensajes_contenedores[contenedor_index]) == 0:
                contenedores[contenedor_index] = 0
                tree.item(selected_item, values=(f"Contenedor {contenedor_index + 1}: {contenedores[contenedor_index]:.2f} unidades de volumen",))

        tk.Button(mensajes_ventana, text="Transferir mensajes", command=transferir_mensajes).pack(pady=10)

    tree.bind("<<TreeviewSelect>>", on_select)

    tk.Button(root, text="Exportar a Excel", command=lambda: exportar_a_excel(
        contenedores, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max)).pack(pady=10)

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
    mostrar_resultados(totales, contenedores_volumenes, mensajes_contenedores, df_final, columnas_mapeadas, capacidad_contenedor_max)


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

    from tkinter import filedialog

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
        # 'columna_original': 'columna_mapeada',
        # Añade tus mapeos aquí si los hay
    }

    # Definir la capacidad máxima del contenedor (ejemplo)
    capacidad_contenedor_max = 1000  # Ajusta este valor según tus necesidades

    # Llamar a la función principal de procesamiento
    main_proceso(df_final, columnas_mapeadas, capacidad_contenedor_max)


if __name__ == "__main__":
    main()