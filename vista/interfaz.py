import os
import shutil
import pandas as pd
from tkinter import PhotoImage, BOTH, CENTER, W, X, messagebox
from ttkbootstrap import ttk, Style
from tkinter import Tk
from openpyxl import load_workbook

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
        root.geometry("600x750")
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
            df = pd.read_excel(ruta_archivo)

            # Filtrar preguntas libres por tema
            df_filtrado = df[df["TEMA"].isin(temas) & df[COLUMNAS_CURSO].isna().all(axis=1)]
            seleccionadas = df_filtrado.sample(n=min(num_preguntas, len(df_filtrado)))

            if seleccionadas.empty:
                messagebox.showinfo("Sin resultados", "No se encontraron preguntas disponibles con esos criterios.")
                return

            # Aquí iría la interfaz para mostrar preguntas y permitir descartar
            messagebox.showinfo("Preguntas encontradas", f"Se encontraron {len(seleccionadas)} preguntas libres.")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

    def _try_parse_int(self, value):
        try:
            return int(value)
        except:
            return None
