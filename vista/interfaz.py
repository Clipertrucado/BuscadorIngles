import os
from tkinter import PhotoImage, BOTH, CENTER, W, X, messagebox, Tk, Toplevel
from ttkbootstrap import ttk, Style
from logica.operaciones import (
    obtener_preguntas,
    reemplazar_pregunta,
    actualizar_archivo_original,
    cargar_datos,
    extraer_columnas_curso
)

CARPETA_ARCHIVOS = "Preguntas"

def iniciar_interfaz():
    os.makedirs(CARPETA_ARCHIVOS, exist_ok=True)
    root = Tk()
    app = ExcelApp(root)
    root.mainloop()

class ExcelApp:
    def __init__(self, root):
        self.root = root
        root.title("Buscador de Preguntas Disponibles")
        root.geometry("880x750")

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
        self.archivo_combo.bind("<<ComboboxSelected>>", self.actualizar_cursos)
        add_input("📂 Archivo de preguntas:", self.archivo_combo)

        self.temas_entry = ttk.Entry(form)
        add_input("📘 Tema(s) (separados por coma, dejar vacío para general):", self.temas_entry)

        self.curso_combo = ttk.Combobox(form, state="readonly")
        add_input("🎓 Curso destino:", self.curso_combo)

        self.tema_curso_entry = ttk.Entry(form)
        add_input("📚 Tema dentro del curso:", self.tema_curso_entry)

        self.num_preguntas_entry = ttk.Entry(form)
        add_input("🔢 Número de preguntas:", self.num_preguntas_entry)

        self.boton_buscar = ttk.Button(container, text="🔍 Buscar preguntas", bootstyle="success outline", command=self.buscar_preguntas)
        self.boton_buscar.pack(pady=30, ipadx=12, ipady=6)

    def get_archivos(self):
        return [f for f in os.listdir(CARPETA_ARCHIVOS) if f.endswith(".xlsx")]

    def actualizar_cursos(self, event):
        archivo = self.archivo_combo.get()
        if not archivo:
            return
        ruta = os.path.join(CARPETA_ARCHIVOS, archivo)
        _, _, encabezados = cargar_datos(ruta)
        cursos = extraer_columnas_curso(encabezados)
        self.curso_combo["values"] = cursos
        if cursos:
            self.curso_combo.current(0)

    def buscar_preguntas(self):
        archivo = self.archivo_combo.get()
        temas = [t.strip() for t in self.temas_entry.get().split(',') if t.strip()]
        curso = self.curso_combo.get().strip()
        tema_curso = self.tema_curso_entry.get().strip()
        num_preguntas = self._try_parse_int(self.num_preguntas_entry.get())

        # ✅ Permitir que 'temas' esté vacío
        if not archivo or not curso or not tema_curso or not num_preguntas:
            messagebox.showwarning("Atención", "Por favor, completa todos los campos.")
            return

        ruta_archivo = os.path.join(CARPETA_ARCHIVOS, archivo)

        try:
            encabezados, seleccionadas, disponibles, total_unicas = obtener_preguntas(
                ruta_archivo, temas, num_preguntas, curso
            )

            if total_unicas == 0:
                messagebox.showinfo("Sin resultados", "No se encontraron preguntas disponibles.")
                return

            if total_unicas < num_preguntas:
                messagebox.showwarning(
                    "Cantidad insuficiente",
                    f"Solo se encontraron {total_unicas} preguntas únicas para los temas seleccionados.\n"
                    f"Se mostrarán esas {total_unicas} preguntas disponibles."
                )

            self.mostrar_resultado(encabezados, seleccionadas, disponibles, archivo, curso, tema_curso)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

    def mostrar_resultado(self, encabezados, seleccionadas, disponibles, archivo_original, curso_destino, valor_asignacion):
        self.resultado = Toplevel(self.root)
        self.resultado.title("Revisión de Preguntas")
        self.resultado.geometry("1000x500")

        self.encabezados = encabezados
        self.preguntas = seleccionadas
        self.pool = disponibles
        self.archivo_original = archivo_original
        self.curso_destino = curso_destino
        self.valor_asignacion = valor_asignacion

        self.tree = ttk.Treeview(
            self.resultado,
            columns=("#", "Tema", "Pregunta", "Respuesta", "Correcta"),
            show="headings"
        )
        self.tree.heading("#", text="#")
        self.tree.heading("Tema", text="Tema")
        self.tree.heading("Pregunta", text="Pregunta")
        self.tree.heading("Respuesta", text="Opciones")
        self.tree.heading("Correcta", text="Correcta")

        self.tree.column("#", width=30, anchor=CENTER)
        self.tree.column("Tema", width=60, anchor=CENTER)
        self.tree.column("Pregunta", width=300)
        self.tree.column("Respuesta", width=350)
        self.tree.column("Correcta", width=150)

        for i, fila in enumerate(self.preguntas, start=1):
            self.tree.insert("", "end", iid=str(i - 1), values=(i, fila[0], fila[1], fila[2], fila[6]))

        self.tree.pack(fill=BOTH, expand=True, padx=10, pady=10)

        ttk.Button(self.resultado, text="Reemplazar seleccionada", command=self.reemplazar_pregunta).pack(pady=5)
        ttk.Button(self.resultado, text="Guardar archivo", command=self.guardar_excel).pack(pady=5)

    def reemplazar_pregunta(self):
        item = self.tree.selection()
        if not item:
            messagebox.showinfo("Atención", "Selecciona una pregunta en la tabla para reemplazar.")
            return
        index = int(item[0])
        nueva = reemplazar_pregunta(self.pool, self.preguntas, index)
        if nueva:
            self.tree.item(item, values=(index + 1, nueva[0], nueva[1], nueva[2], nueva[6]))
        else:
            messagebox.showinfo("Sin reemplazos", "No hay más preguntas disponibles para reemplazar.")

    def guardar_excel(self):
        ruta = os.path.join(CARPETA_ARCHIVOS, self.archivo_original)
        try:
            actualizar_archivo_original(ruta, self.preguntas, self.encabezados, self.curso_destino, self.valor_asignacion)
            messagebox.showinfo("Éxito", f"Archivo actualizado: {ruta}")
            self.resultado.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar el archivo: {e}")

    def _try_parse_int(self, value):
        try:
            return int(value)
        except:
            return None
