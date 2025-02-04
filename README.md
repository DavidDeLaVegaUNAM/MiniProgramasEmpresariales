# Proyectos Actuales

- Creador de Informes en *Excel*. V. $\beta 0.1$

## Proyecto: Creador de Informes en *Excel* V. $\beta 0.1$

### Requisitos del sistema:
Para esta primer versión se ha creado Windows para , se requiere que la computadora tenga una versión de Python 3.10 en adelante.
En caso de no tenerlo, con el comando "python" ejecutado en el CMD se instalará en el sistema.

Las librerías usadas en este programa son `Tkinter` (para la interfaz) y `xlsxwriter` (para crear archivos en excel), para instalarlas en el sistema para que corran correctamente con los comandos :

```
pip install tkinter
pip install xlsxwriter
```

Cada uno de manera independiente en una terminal. 

### Ejecución del Programa
Tras ello ya es posible correr el programa. Para hacerlo se ejecuta el programa: `GenerarReportes.pyw` el cual abrirá una ventana emergente con el título **Generador de Reportes Comerciales**. 
Contiene los siguientes campos requeridos para generar una reporte de venta:

-Marca del Producto
-Fecha de Inicio (En formato AAAA/MM)
-Fecha Fin (En formato AAAA/MM)

Por fecha de inicio y fin se refiere a dos meses de años iguales o distintos que queremos comparar.

### Resultado
Para obtener un resultado se después de llenar los campos se da click en el botón `Generar Reporte`, se abrirá una ventana para permitir al usuario escoger dónde guardar su archivo. Tras ello vamos a obtener un reporte de excel nombrado
```
Reporte_[Marca del Producto]_[Fecha de Inicio]-[Fecha Fin].xlsx
```
En él, habrá 3 hojas de nombre `Comparativa`, `PDV Día` y `Zona`. 

-Comparativa: Está dada por las columnas Métrica, las dos fechas y la variación. Incluye datos de muestra para mostrar su uso. 
-PDV Día: Contiene los puntos de venta del producto, ventas, visitas al lugar y una relación entre estas.
-Zonas: Para las distintas regiones que presenta un negocio, las ventas de cada una y su participación total en la empresa.

