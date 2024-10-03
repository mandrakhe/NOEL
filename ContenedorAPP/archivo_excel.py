import tkinter as tk
from tkinter import ttk, messagebox
from operaciones import exportar_a_excel

def mostrar_resultados(totales, contenedores, mensajes_contenedores, df_filtrado, columnas_mapeadas):
    root = tk.Tk()
    root.title("Resultados de Contenedores")
    root.geometry("800x600")  # Ajuste del tamaño de la ventana

    # Mostrar la tabla de totales
    frame_totales = tk.Frame(root)
    frame_totales.pack(pady=10)

    tk.Label(frame_totales, text="Descripción", font=(
        "Arial", 12, "bold")).grid(row=0, column=0)
    tk.Label(frame_totales, text="Valor", font=(
        "Arial", 12, "bold")).grid(row=0, column=1)

    for i, (desc, valor) in enumerate(totales.items()):
        tk.Label(frame_totales, text=desc, font=(
            "Arial", 12)).grid(row=i+1, column=0)
        tk.Label(frame_totales, text=str(valor), font=(
            "Arial", 12)).grid(row=i+1, column=1)

    # Crear la tabla de contenedores
    frame_contenedores = tk.Frame(root)
    frame_contenedores.pack(pady=20)

    tk.Label(frame_contenedores, text="Contenedores",
                font=("Arial", 14, "bold")).pack()

    tree = ttk.Treeview(frame_contenedores, columns=(
        "Peso"), show="headings", height=10)
    tree.heading("Peso", text="Peso Neto")
    tree.column("Peso", width=300)  # Ajuste del ancho de la columna
    tree.pack(expand=True, fill='both')  # Ajuste para expandir el Treeview

    for i, contenedor in enumerate(contenedores):
        tree.insert("", "end", iid=i, values=(
            f"Contenedor {i+1}: {contenedor:.2f} unidades de peso neto",))

    def on_select(event):
        if not tree.selection():
            # Si no hay ningún elemento seleccionado, salir de la función.
            return

        selected_item = tree.selection()[0]
        contenedor_index = int(selected_item)

        # Crear una nueva ventana para mostrar y mover los mensajes
        mensajes_ventana = tk.Toplevel(root)
        mensajes_ventana.title(f"Mensajes para Contenedor {
                            contenedor_index + 1}")
        # Ajuste del tamaño de la ventana de mensajes
        mensajes_ventana.geometry("600x400")

        mensajes_listbox = tk.Listbox(
            mensajes_ventana, selectmode=tk.MULTIPLE, width=50, height=10)
        mensajes_listbox.pack(pady=10)

        for mensaje in mensajes_contenedores[contenedor_index]:
            mensajes_listbox.insert(tk.END, mensaje)

        def transferir_mensajes():
            seleccionados = list(mensajes_listbox.curselection())
            if not seleccionados:
                return

            destino_index = tk.simpledialog.askinteger(
                "Transferir a Contenedor",
                f"Seleccione el número del contenedor de destino (1-{
                    len(contenedores)})",
                minvalue=1, maxvalue=len(contenedores)
            )

            if destino_index is None or destino_index == contenedor_index + 1:
                return

            destino_index -= 1
            peso_a_transferir = 0
            mensajes_a_transferir = []
            mensajes_no_transferidos = []

            # Asumimos que cada mensaje tiene un peso específico
            for i in seleccionados:
                mensaje_seleccionado = mensajes_contenedores[contenedor_index][i]
                peso_mensaje = df_filtrado.loc[df_filtrado[columnas_mapeadas['Texto de mensaje']]
                                            == mensaje_seleccionado, 'Neto'].values[0]
                if contenedores[destino_index] + peso_mensaje <= 71.5:
                    peso_a_transferir += peso_mensaje
                    mensajes_a_transferir.append(mensaje_seleccionado)
                    contenedores[destino_index] += peso_mensaje
                else:
                    mensajes_no_transferidos.append(mensaje_seleccionado)

            # Actualizar el contenedor de origen
            for mensaje in mensajes_a_transferir:
                mensajes_contenedores[contenedor_index].remove(mensaje)
                mensajes_contenedores[destino_index].append(mensaje)
                contenedores[contenedor_index] -= df_filtrado.loc[df_filtrado[columnas_mapeadas['Texto de mensaje']]
                                                                == mensaje, 'Neto'].values[0]

            # Mostrar advertencia si se alcanzó el límite
            if mensajes_no_transferidos:
                advertencia = f"No se pudieron transferir los siguientes mensajes porque se alcanzó el límite de peso del contenedor de destino:\n{
                    ', '.join(map(str, mensajes_no_transferidos))}"
                messagebox.showwarning("Advertencia", advertencia)

            # Actualizar la interfaz
            mensajes_listbox.delete(0, tk.END)
            for mensaje in mensajes_contenedores[contenedor_index]:
                mensajes_listbox.insert(tk.END, mensaje)

            tree.item(selected_item, values=(f"Contenedor {
                        contenedor_index + 1}: {contenedores[contenedor_index]:.2f} unidades de peso neto",))
            tree.item(destino_index, values=(f"Contenedor {
                        destino_index + 1}: {contenedores[destino_index]:.2f} unidades de peso neto",))

            # Si un contenedor queda vacío, poner su peso a 0
            if len(mensajes_contenedores[contenedor_index]) == 0:
                contenedores[contenedor_index] = 0

        tk.Button(mensajes_ventana, text="Transferir mensajes",
                    command=transferir_mensajes).pack(pady=10)

    tree.bind("<<TreeviewSelect>>", on_select)

    tk.Button(root, text="Exportar a Excel", command=lambda: exportar_a_excel(
        contenedores, mensajes_contenedores, df_filtrado, columnas_mapeadas)).pack(pady=10)

    root.mainloop()