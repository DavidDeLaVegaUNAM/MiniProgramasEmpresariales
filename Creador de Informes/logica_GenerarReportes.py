import xlsxwriter
from datetime import datetime
import os

class ReportGenerator:
    def __init__(self, marca, fecha_inicio, fecha_fin):
        self.marca = marca
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self.estilos = None

    def crear_formatos(self, workbook):
        """Crea los formatos de celda para el archivo Excel"""
        formatos = {
            "header": workbook.add_format({
                "bg_color": "#1F497D",
                "font_color": "white",
                "bold": True,
                "border": 1
            }),
            "moneda": workbook.add_format({
                "num_format": "#,##0.00",
                "border": 1
            }),
            "porcentaje": workbook.add_format({
                "num_format": "0.00%",
                "border": 1
            }),
            "texto": workbook.add_format({
                "border": 1
            })
        }
        return formatos

    def generar_reporte(self, ruta_guardado, progress_callback):
        """Genera el archivo Excel con todas las hojas y datos"""
        try:
            # Nombre del archivo con formato claro
            nombre_base = f"Reporte_{self.marca}_{self.fecha_inicio.replace('/', '-')}_a_{self.fecha_fin.replace('/', '-')}.xlsx"
            nombre_archivo = os.path.join(ruta_guardado, nombre_base)
            
            # Verificar si el archivo ya existe
            if os.path.exists(nombre_archivo):
                raise PermissionError(f"El archivo '{nombre_archivo}' ya existe. Por favor, cierra el archivo o elimínalo.")
            
            workbook = xlsxwriter.Workbook(nombre_archivo)
            self.estilos = self.crear_formatos(workbook)
            
            progress_callback(10)
            
            # Crear hojas de trabajo
            hojas = {
                "comparativa": workbook.add_worksheet("Comparativa"),
                "pdv": workbook.add_worksheet("PDV Día"),
                "zonas": workbook.add_worksheet("Zonas")
            }
            
            progress_callback(30)
            self.insertar_comparativa(hojas["comparativa"])
            progress_callback(50)
            
            self.insertar_pdv(hojas["pdv"])
            progress_callback(70)
            
            self.insertar_zonas(hojas["zonas"])
            progress_callback(90)
            
            workbook.close()
            progress_callback(100)
            
            return nombre_archivo
        
        except Exception as e:
            raise e

    def obtener_datos_ejemplo(self):
        """Devuelve datos de ejemplo para demostración"""
        return {
            "comparativa": [
                ["Ventas", 150000, 135000],
                ["Clientes", 450, 420],
                ["Margen", 0.35, 0.32]
            ],
            "pdv": {
                "PDV Norte": {"ventas": 45000, "visitas": 1200},
                "PDV Centro": {"ventas": 75000, "visitas": 1800}
            },
            "zonas": {
                "Norte": 45000,
                "Centro": 75000
            }
        }

    def insertar_comparativa(self, worksheet):
        """Inserta datos en la hoja Comparativa"""
        datos = self.obtener_datos_ejemplo()["comparativa"]
        
        # Encabezados
        worksheet.write(0, 0, "Métrica", self.estilos["header"])
        worksheet.write(0, 1, self.fecha_inicio, self.estilos["header"])
        worksheet.write(0, 2, self.fecha_fin, self.estilos["header"])
        worksheet.write(0, 3, "Variación", self.estilos["header"])
        
        # Datos
        for row, (metrica, actual, anterior) in enumerate(datos, 1):
            variacion = (actual - anterior) / anterior if anterior != 0 else 0
            worksheet.write(row, 0, metrica, self.estilos["texto"])
            worksheet.write(row, 1, actual, self.estilos["moneda"])
            worksheet.write(row, 2, anterior, self.estilos["moneda"])
            worksheet.write(row, 3, variacion, self.estilos["porcentaje"])
        
        # Formato de columnas
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 3, 15)

    def insertar_pdv(self, worksheet):
        """Inserta datos en la hoja PDV Día"""
        datos = self.obtener_datos_ejemplo()["pdv"]
        
        worksheet.write(0, 0, "PDV", self.estilos["header"])
        worksheet.write(0, 1, "Ventas", self.estilos["header"])
        worksheet.write(0, 2, "Visitas", self.estilos["header"])
        worksheet.write(0, 3, "Conversión", self.estilos["header"])
        
        for row, (pdv, datos_pdv) in enumerate(datos.items(), 1):
            conversion = datos_pdv["ventas"] / datos_pdv["visitas"] if datos_pdv["visitas"] > 0 else 0
            worksheet.write(row, 0, pdv, self.estilos["texto"])
            worksheet.write(row, 1, datos_pdv["ventas"], self.estilos["moneda"])
            worksheet.write(row, 2, datos_pdv["visitas"], self.estilos["texto"])
            worksheet.write(row, 3, conversion, self.estilos["porcentaje"])
        
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 3, 15)

    def insertar_zonas(self, worksheet):
        """Inserta datos en la hoja Zonas"""
        datos = self.obtener_datos_ejemplo()["zonas"]
        total = sum(datos.values())
        
        worksheet.write(0, 0, "Zona", self.estilos["header"])
        worksheet.write(0, 1, "Ventas", self.estilos["header"])
        worksheet.write(0, 2, "Participación", self.estilos["header"])
        
        for row, (zona, ventas) in enumerate(datos.items(), 1):
            participacion = ventas / total if total > 0 else 0
            worksheet.write(row, 0, zona, self.estilos["texto"])
            worksheet.write(row, 1, ventas, self.estilos["moneda"])
            worksheet.write(row, 2, participacion, self.estilos["porcentaje"])
        
        worksheet.set_column(0, 0, 25)
        worksheet.set_column(1, 2, 15)