import os
import shutil
import pandas as pd
from tkinter import PhotoImage, BOTH, CENTER, W, X, messagebox
from ttkbootstrap import ttk, Style
from tkinter import Tk, Toplevel
from logica.operaciones import obtener_preguntas, reemplazar_pregunta, guardar_preguntas

CARPETA_ARCHIVOS = "."
CARPETA_EDITADOS = "editados"
COLUMNAS_CURSO = ["SARGENTO 24-25", "OFICIAL 24-25", "CAMBIO ESCALA 24-25", "CABO 24-25", "GC+GJ 24-25"]

def iniciar_interfaz():
    root = Tk()
    app = ExcelApp(root)
    root.mainloop()

class ExcelApp:
    def __init__(self, root):
        self.root = root
        root.title("Buscador de Preguntas Disponibles")
        root.geometry("820x750")

        if os.path.exists("./logo.png"):
            logo = PhotoImage(file="./logo.png")
            root.iconphoto(False, logo)

        style = Style("cosmo")

        container = ttk.Frame(root, padding=40)
        container.pack(expand=True, fill=BOTH)

        title = ttk.Label(container, text="Buscador de Preguntas Excel", font=("Segoe UI", 20, "bold"), bootstyle="primary")
        title.pack(pady=(0, 30))

        form = ttk.Frame(container)
        form.pack(anchor=CENTER, fill=X, expand=True)

        def add_input(label_text, widget):
            group = ttk.Frame(form)
            group.pack(pady=8, fill=X, expand=True)
            ttk.Label(group, text=label_text, anchor=W).pack(anchor=W)
            widget.pack(fill=X, expand=True, ipady=5)

        self.archivo_combo = ttk.Combobox(form, values=self.get_archivos(), state="readonly")
        add_input("\U0001F4C2 Archivo de preguntas:", self.archivo_combo)

        self.temas_entry = ttk.Entry(form)
        add_input("\U0001F4D8 Tema(s) (separados por coma):", self.temas_entry)

        self.curso_entry = ttk.Entry(form)
        add_input("\U0001F393 Curso destino:", self.curso_entry)

        self.tema_curso_entry = ttk.Entry(form)
        add_input("\U0001F4DA Tema dentro del curso:", self.tema_curso_entry)

        self.num_preguntas_entry = ttk.Entry(form)
        add_input("\U0001F522 Número de preguntas:", self.num_preguntas_entry)

        self.boton_buscar = ttk.Button(container, text="\U0001F50D Buscar preguntas", bootstyle="success outline", command=self.buscar_preguntas)
        self.boton_buscar.pack(pady=30, ipadx=12, ipady=6)

    def get_archivos(self):
        return [f for f in os.listdir(CARPETA_ARCHIVOS) if f.endswith(".xlsx")]

    def buscar_preguntas(self):
        archivo = self.archivo_combo.get()
        temas = [t.strip() for t in self.temas_entry.get().split(',') if t.strip()]
        curso = self.curso_entry.get().strip()
        tema_curso = self.tema_curso_entry.get().strip()
        num_preguntas = self._try_parse_int(self.num_preguntas_entry.get())

        if not archivo or not temas or not curso or not tema_curso or not num_preguntas:
            messagebox.showwarning("Atención", "Por favor, completa todos los campos.")
            return

        ruta_archivo = os.path.join(CARPETA_ARCHIVOS, archivo)

        try:
            encabezados, seleccionadas, disponibles = obtener_preguntas(ruta_archivo, temas, COLUMNAS_CURSO, num_preguntas)
            if not seleccionadas:
                messagebox.showinfo("Sin resultados", "No se encontraron preguntas disponibles.")
                return

            self.mostrar_resultado(encabezados, seleccionadas, disponibles, archivo)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

    def mostrar_resultado(self, encabezados, seleccionadas, disponibles, archivo_original):
        self.resultado = Toplevel(self.root)
        self.resultado.title("Revisión de Preguntas")
        self.resultado.geometry("850x500")

        self.encabezados = encabezados
        self.preguntas = seleccionadas
        self.pool = disponibles
        self.archivo_original = archivo_original

        self.tree = ttk.Treeview(self.resultado, columns=("#", "Pregunta", "Respuesta", "Correcta"), show="headings")
        self.tree.heading("#", text="#")
        self.tree.heading("Pregunta", text="Pregunta")
        self.tree.heading("Respuesta", text="Opciones")
        self.tree.heading("Correcta", text="Correcta")
        self.tree.column("#", width=30, anchor=CENTER)
        self.tree.column("Pregunta", width=300)
        self.tree.column("Respuesta", width=350)
        self.tree.column("Correcta", width=150)

        for i, fila in enumerate(self.preguntas, start=1):
            self.tree.insert("", "end", iid=str(i - 1), values=(i, fila[1], fila[2], fila[6]))

        self.tree.pack(fill=BOTH, expand=True, padx=10, pady=10)

        boton_reemplazar = ttk.Button(self.resultado, text="Reemplazar seleccionada", command=self.reemplazar_pregunta)
        boton_guardar = ttk.Button(self.resultado, text="Guardar archivo", command=self.guardar_excel)
        boton_reemplazar.pack(pady=5)
        boton_guardar.pack(pady=5)

    def reemplazar_pregunta(self):
        item = self.tree.selection()
        if not item:
            messagebox.showinfo("Atención", "Selecciona una pregunta en la tabla para reemplazar.")
            return
        index = int(item[0])
        nueva = reemplazar_pregunta(self.pool, self.preguntas, index)
        if nueva:
            self.tree.item(item, values=(index + 1, nueva[1], nueva[2], nueva[6]))
        else:
            messagebox.showinfo("Sin reemplazos", "No hay más preguntas disponibles para reemplazar.")

    def guardar_excel(self):
        os.makedirs(CARPETA_EDITADOS, exist_ok=True)
        destino = os.path.join(CARPETA_EDITADOS, f"EDITADO_{self.archivo_original}")
        guardar_preguntas(self.encabezados, self.preguntas, destino)
        messagebox.showinfo("Éxito", f"Archivo guardado como {destino}")
        self.resultado.destroy()

    def _try_parse_int(self, value):
        try:
            return int(value)
        except:
            return None
            