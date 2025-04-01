import tkinter as tk
from logica.operaciones import sumar

def iniciar_interfaz():
    ventana = tk.Tk()
    ventana.title("Mi App")

    entrada1 = tk.Entry(ventana)
    entrada2 = tk.Entry(ventana)
    resultado = tk.Label(ventana, text="Resultado")

    def calcular():
        a = int(entrada1.get())
        b = int(entrada2.get())
        res = sumar(a, b)
        resultado.config(text=f"Resultado: {res}")

    boton = tk.Button(ventana, text="Sumar", command=calcular)

    entrada1.pack()
    entrada2.pack()
    boton.pack()
    resultado.pack()

    ventana.mainloop()
