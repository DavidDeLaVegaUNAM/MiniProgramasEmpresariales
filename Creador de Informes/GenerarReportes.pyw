import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
from logica_GenerarReportes import ReportGenerator

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Reportes Comerciales")
        self.root.geometry("500x250")
        self.root.resizable(False,False)
        
        # Variables de control
        self.marca = tk.StringVar()
        self.fecha_inicio = tk.StringVar()
        self.fecha_fin = tk.StringVar()
        self.progreso = tk.DoubleVar()
        
        self.setup_ui()

    def setup_ui(self):
        """Configura la interfaz gráfica"""
        style = ttk.Style()
        style.configure("TLabel", padding=5, font=("Arial", 10))
        style.configure("TButton", padding=5, font=("Arial", 10))
        
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill="both")
        
        # Campos de entrada
        ttk.Label(main_frame, text="Marca:").grid(row=0, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.marca).grid(row=0, column=1, sticky="ew", pady=5)
        
        ttk.Label(main_frame, text="Fecha Inicio (AAAA/MM):").grid(row=1, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.fecha_inicio).grid(row=1, column=1, sticky="ew", pady=5)
        
        ttk.Label(main_frame, text="Fecha Fin (AAAA/MM):").grid(row=2, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.fecha_fin).grid(row=2, column=1, sticky="ew", pady=5)
        
        # Botones
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)
        
        self.btn_generar = ttk.Button(
            btn_frame, 
            text="Generar Reporte", 
            command=self.iniciar_generacion
        )
        self.btn_generar.pack(side="left", padx=5)
        
        # Barra de progreso
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            variable=self.progreso, 
            maximum=100
        )
        self.progress_bar.grid(row=4, column=0, columnspan=2, sticky="ew", pady=10)
        
        main_frame.columnconfigure(1, weight=1)

    def validar_fecha(self, fecha):
        """Valida el formato de fecha AAAA/MM"""
        try:
            datetime.strptime(fecha, "%Y/%m")
            return True
        except ValueError:
            return False

    def iniciar_generacion(self, event=None):
        if not self.validar_campos():
            return
        
        ruta_guardado = filedialog.askdirectory(
            title="Seleccione la carpeta para guardar el reporte"
        )
        
        if not ruta_guardado:  # Si el usuario cancela
            return
        
        self.btn_generar.config(state="disabled")
        thread = threading.Thread(
            target=self.ejecutar_generacion,
            args=(ruta_guardado,)
        )
        thread.start()

    def validar_campos(self):
        if not self.marca.get().strip():
            messagebox.showerror("Error", "Debe ingresar una marca")
            return False
        if not self.validar_fecha(self.fecha_inicio.get()):
            messagebox.showerror("Error", "Formato de fecha inicio inválido")
            return False
        if not self.validar_fecha(self.fecha_fin.get()):
            messagebox.showerror("Error", "Formato de fecha fin inválido")
            return False
        return True

    def actualizar_progreso(self, valor):
        self.progreso.set(valor)
        self.root.update_idletasks()

    def ejecutar_generacion(self, ruta_guardado):
        try:
            generador = ReportGenerator(
                marca=self.marca.get().strip(),
                fecha_inicio=self.fecha_inicio.get(),
                fecha_fin=self.fecha_fin.get()
            )
            
            archivo = generador.generar_reporte(ruta_guardado, self.actualizar_progreso)
            
            messagebox.showinfo(
                "Éxito", 
                f"Reporte generado exitosamente!\n\n"
                f"Archivo: {archivo}"
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar el reporte:\n{str(e)}")
        finally:
            self.btn_generar.config(state="normal")
            self.progreso.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()