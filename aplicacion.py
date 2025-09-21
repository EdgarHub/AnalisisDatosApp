# Autor: Edgar Torres
# Backend: Python 3.9.7, pandas 1.5.1, matplotlib 3.9.4
# Frontend: QT5

# Librerias
import sys, os, pandas
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt
from PyQt5.uic import loadUi

# Clase
class Aplicacion(QMainWindow):
    # Constructor
    def __init__(self):
        #Constructor de la clase Heredada
        super(Aplicacion, self).__init__()
        # Llamar la interfaz desde el backend
        loadUi(r"C:\Users\soporte\Desktop\Python Analisis Ed\Aplicacion\interfaz.ui",self)
        # Vincular los metodos a los Widgets
        self.btnCancelar.clicked.connect(self.cancelar)
        self.btnCargar.clicked.connect(self.cargar_archivo)
        self.btnColor.clicked.connect(self.color_grafica)
        self.btnConsultar.clicked.connect(self.consultar)
        self.btnExportar.clicked.connect(self.exportar_datos)
        self.btnGraficar.clicked.connect(self.graficar)
        self.btnRegresar.clicked.connect(self.regresar)
        # Agregar un texto fantasma
        self.txtConsulta.setPlaceholderText("Escriba su consulta...")
        # Declarar DataFrames globales
        self.df = pandas.DataFrame()
        self.dfConsultado = pandas.DataFrame()


    
    def cargar_archivo(self):
         # Crear un cuadro de dialogo
        opciones = QFileDialog.Options()
        seleccion, _ = QFileDialog.getOpenFileName(self,
                                              "Seleccione un Archivo",
                                              "",
                                              "Archivos (*.xlsx *.csv *.json)",
                                              options=opciones)
        if seleccion:
                try:
                # Detectar el tipo de archivo
                    if seleccion.endswith('.xlsx'):
                        self.df = pandas.read_excel(seleccion)
                    elif seleccion.endswith('.csv'):
                        self.df = pandas.read_csv(seleccion)
                    elif seleccion.endswith('.json'):
                        self.df = pandas.read_json(seleccion)
                    # Llamar metodo mostrar_datos()
                    self.mostrar_datos(self.df, "Importacion exitosa")
                    self.dfConsultado = self.df
                    # Llamar metodo para mostrar estadisticas
                    self.mostrar_estadisticas(self.df)
                except Exception as error:
                    QMessageBox.warning(self, "Error", f"Se presento el error:{str(error)}")
        else:
            QMessageBox.warning(self, "Atencion", "Se cancelo la operacion")

    def regresar(self):
        self.mostrar_datos(self.df, "Mostrando todos los datos")
        self.mostrar_estadisticas(self.df)
        self.txtConsulta.clear()       

    def mostrar_datos(self, df, mensaje):
        # Mostrar los datos en la tabla
                # Limpiar tabla para quitar datos/basura/consulta anterior
                self.tablaDatos.clear()
                # Establecer cantidad de filas
                self.tablaDatos.setRowCount(len(df))
                # Establecer cantidad de Columnas
                self.tablaDatos.setColumnCount(len(df.columns))
                # Agregar el nombre de los Rotulos de las columnas
                self.tablaDatos.setHorizontalHeaderLabels(df.columns.astype(str))
                # Color especifico a los headers de filas y columnas
                self.tablaDatos.setStyleSheet("""QHeaderView::section{
                                                color:#FFFFFF;
                                                background-color:#05587C;
                                                }""")
                # Agregar informacion de las celdas repitiendo el ciclo x el numero de filas existentes en el archivo
                # range convierte un valor en un objeto iterable
                for fila in range(len(df)):
                    #Repetir el ciclo anidado por cada celda de cada fila
                    for columna in range(len(self.df.columns)):
                        # Recuperar el valor de la celda
                        # iloc = Acceder a una coordenada
                        valor = str(df.iloc[fila,columna])
                        # Enviar el valor de la celda a la tabla
                        self.tablaDatos.setItem(fila,columna,QTableWidgetItem(valor))
                # Aviso de cargado de datos
                QMessageBox.information(self, "Datos", f"{mensaje}: {len(df)} filas agregadas" )

    def mostrar_estadisticas(self, df):
        # DataFrame para valores numericos
        # Solo agregara a este DF los campos numericos (int float etc)
        dfNumerico = df.select_dtypes(include="number")
        # Validar que no haya quedado vacio el DF
        if dfNumerico.empty:
            self.tablaResumen.setColumnCount()
            self.tablaResumen.setRowCount()
            return 
        # Configurar las Estadisticas
        # T = Transpose
        estadisticas = dfNumerico.describe().T[["mean", "std", "min", "max"]]
        # Nombre Personalizado
        estadisticas.columns = ["MEDIA", "DESVIACION", "MINIMO", "MAXIMO"]
        #Establecer la cantidad de columnas
        self.tablaResumen.setColumnCount(len(estadisticas.columns))
        #Establecer la cantidad de filas
        self.tablaResumen.setRowCount(len(estadisticas.index))
        # Establecer nombres columnas
        self.tablaResumen.setHorizontalHeaderLabels(estadisticas.columns)
        # Establecer nombres Filas
        self.tablaResumen.setVerticalHeaderLabels(estadisticas.index)
        # Ciclos para llenado de tabla
        for fila in range(len(estadisticas.index)):
            # Ciclo para recorrer todas las celdas de una fila
            for columna in range(len(estadisticas.columns)):
                # iat = tomar el valor de tal posicion
                valor = QTableWidgetItem(f"{estadisticas.iat[fila,columna]:.2f}")
                valor.setTextAlignment(Qt.AlignCenter)
                self.tablaResumen.setItem(fila,columna,valor)
        # Columnas ajustables
        self.tablaResumen.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)


    def consultar(self):
        # Validar que la tabla contenga datos
        if self.tablaDatos.columnCount() == 0 and self.tablaDatos.rowCount() == 0:
            QMessageBox.warning(self, "Atencion", "Realiza una importacion de datos para su consulta")
            return
        
        """# Validar que la tabla contenga datos
        if len(self.df) == 0:
            QMessageBox.warning(self, "Atencion", "Realiza una importacion de datos para su consulta")
            return"""
        
        consulta = self.txtConsulta.toPlainText()
        if len(consulta) == 0:
            QMessageBox.warning(self, "Atencion", "Capture una consulta")
            return
        # Intentar Aplicar Consulta
        try:    
            self.dfConsultado = self.df.query(consulta)
            self.mostrar_datos(self.dfConsultado, "Consulta exitosa")
            self.mostrar_estadisticas(self.dfConsultado)
        except Exception as error:
            QMessageBox.warning(self, "Error", f"Error de consulta: {str(error)}")

    def exportar_datos(self):
        formato = self.cbFormato.currentText()
         # Intentar llevar a cabo la exportacion
        try:
            # Revisar que el DataFrame no este vacio
            if len(self.dfConsultado) == 0:
                QMessageBox.warning(self, "Atencion", "Realiza una consulta de datos para su exportacion")
                return
            
            # Configuracion para el cuadro de dialogo
            filtros = {
                "Excel": ("Archivos Excel (*.xlsx)", ".xlsx"),
                "JSON": ("Archivos JSON (*.json)", ".json"),
                "CSV": ("Archivos CSV (*.csv)", ".csv"),
                "HTML": ("Archivos HTML (*.html)", ".html")
            }

            filtro, extension = filtros.get(formato, (None, None))
            if filtro == None:
                QMessageBox.warning(self, "Atencion", "Formato invalido")
                return
            
            # Cuadro de dialogo
            # path, extension
            ruta, _ = QFileDialog.getSaveFileName(self,
                                                  "Guardar archivo",
                                                  f"Exportacion {extension}",
                                                  filtro)
            if not ruta:
                return
            
            # Corroborar el formato
            if formato == "Excel":
                # Convertir la QTABLEWIDGET a dataframe,
                dfResumen = self.convertir_tabla_dataframe(self.tablaResumen)
                # Exportar varias hojas
                with pandas.ExcelWriter(ruta, engine="openpyxl") as writer:
                    self.dfConsultado.to_excel(writer, sheet_name="Tabla", index=False)
                    dfResumen.to_excel(writer, sheet_name="Estadisticas", index=False)
                    QMessageBox.information(self, "Aviso",  "Archivo exportado con exito en Excel")
            elif formato == "JSON":
                self.dfConsultado.to_json(ruta, orient="records", force_ascii=False)
                QMessageBox.information(self, "Aviso",  "Archivo exportado con exito en JSON")
            elif formato == "CSV":
                self.dfConsultado.to_csv(ruta, index=False)
                QMessageBox.information(self, "Aviso",  "Archivo exportado con exito en CSV")
            elif formato == "HTML":
                self.dfConsultado.to_html(ruta, index=False)
                QMessageBox.information(self, "Aviso", "Archivo exportado con exito en HTML")

                

        except Exception as error:
                QMessageBox.warning(self,
                                     "Error",
                                     f"Se presento el error:{str(error)}")
    def convertir_tabla_dataframe(self, tabla):
        # Obtener filas y columnas de la tabla
        filas = tabla.rowCount()
        columnas = tabla.columnCount()

        # Obtener los encabezados de las columnas
        encabezados = ["MEDIA", "DESVIACION", "MINIMO", "MAXIMO"]
        
        # Lista para almacenar losd atos de cada fila
        datos =[]

        # Recorrer las filas de la tabla
        for fila in range(filas):
            # Crear una lista temporal para guardar los datos de cada fila
            fila_datos = []
            for celda in range(columnas):
                # obtener el Item que este guardado
                item = tabla.item(fila,celda)
                # Agregar (append) el valor a la fila
                # If item else "" = si no hay dato, guarda un str vacio
                fila_datos.append(item.text() if item else "")
            
            # Agregar la fila a la lista general (donde estan todas las filas)
            datos.append(fila_datos)
        # Devolver un Dataframe con datos
        return pandas.DataFrame(datos, columns=encabezados)


    def color_grafica(self):
        # Cuadro de dialogo para sleccionar color
        color = QColorDialog.getColor()
        # Verificar que se haya seleccionado algun color
        if color.isValid():
            # Guardar el color de manera global
            # Almacenando el color de manera Hexadecimal
            self.color_grafico = color.name()

    def graficar(self):
        # Comprobar que la tablaResumen tenga datos
        if self.tablaResumen.rowCount() == 0 or self.tablaResumen.columnCount() == 0:
            QMessageBox.warning(self, "Sin Datos",
                                "No hay datos por graficar")
            return
        
        # Verificar que el usuario haya seleccionado algun color
        # hasattr = metodo para conocer si el codigo contiene ciertos atributos
        if not hasattr(self, 'color_grafico') or not self.color_grafico:
            QMessageBox.warning(self, "Error en color",
                                "Seleccione un color para el grafico")
            return
        
        # Obtener las etiquetas de las columnas 
        etiquetas_columnas = ["MEDIA", "DESVIACION", "MINIMO", "MAXIMO"]

        # Mostrar un cuadro de dialogo para seleccionar la estadistica a graficar
        columna, ok_columna = QInputDialog.getItem(self, "Seleccione una opcion",
                                                  "Seleccione una estadistica a graficar",
                                                  etiquetas_columnas,
                                                  0,
                                                  False)
        
        # Si no se selecciono una opcion, entocnes salir del metodo
        if not ok_columna:
            return
        
        # Conocer el indice seleccionado de las opciones/lista
        indice_columna = etiquetas_columnas.index(columna)

        # Obtener los nombres de las filas (categorias)
        # Usando i como iterador, el numero de veces que retorne el metodo rowCount
        etiquetas_filas = [self.tablaResumen.verticalHeaderItem(i).text() for i in range (self.tablaResumen.rowCount())]

        # Crear un cuadro de dialogo personalizado
        dialogo = QDialog(self)
        dialogo.setWindowTitle("Seleccione las categorias a graficar")
        layout = QVBoxLayout(dialogo)

        # Opciones/Casillas de la lista
        checks = []
        for etiqueta in etiquetas_filas:
            check = QCheckBox(etiqueta) # Crear Casillas
            check.setChecked(False) # Poner/Quitar la Palomita de seleccion
            layout.addWidget(check)
            checks.append(check)

        # Agregar botones de aceptar y cancelar
        botones = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(botones)
        botones.accepted.connect(dialogo.accept)
        botones.rejected.connect(dialogo.reject)
        


        # Ejecutar el dialogo
        if dialogo.exec() != QDialog.Accepted:
            return
        
        #Generar una lista con los valores seleccionados
        etiquetas_seleccionadas = [
            check.text() # Texto de las casillas selecionada
            for check in checks 
            if check.isChecked()
        ]
        
        # Revisar que no se haya quedado sin selecciones
        if not etiquetas_seleccionadas:
            QMessageBox.warning(self, "Error", "Selecciona Casillas para graficar")
            return
        
        # Lista con los valores a graficar
        valores = []
        # Ciclo para recorrer las categorias/filas
        for fila in range(self.tablaResumen.rowCount()):
            # Recuperar el nombre de la etiqueta seleccionada
            etiqueta = self.tablaResumen.verticalHeaderItem(fila).text()
            # Comparar si la etiqueta pertenece al filtro creado
            if etiqueta in etiquetas_seleccionadas:
                # Guardar elemento
                elemento = self.tablaResumen.item(fila, indice_columna)
                # Calidar que contenga dato
                if elemento:
                    try:
                        # Si 'elemento' tiene un numero > a cero
                        valores.append(float(elemento.text()))
                    except ValueError:
                        valores.append(0.0)
        
        # Personalizar
        
        plt.figure(figsize= (10,6), facecolor="#8E8D8D", edgecolor="#FFFFFF")
        ax = plt.gca()
        ax.set_facecolor("#8E8D8D")
        # Color de los valores ejes X, Y o both
        plt.tick_params(axis="both", colors="#FFFFFF", labelsize=12)
        

        barras = plt.bar(etiquetas_seleccionadas, 
                         valores, 
                         color=self.color_grafico)
        
        # Agregar los valores a cada barra del Grafico
        for barra in barras:
            altura = barra.get_height()
            plt.text(
                barra.get_x() + barra.get_width()/2, # centro de la barra
                altura, #Posicionar ligeramente arriba de la barra
                f"{altura:.2f}", # Valor de la altura a str para que se pueda imprimir
                ha="center", # Alineacion Horizontal
                va="bottom", # Alineacion Vertical
                fontsize=11 # tamaño de la fuente para el valor
            )
        plt.title(f"{columna} Datos Estadisticos", color=self.color_grafico, backgroundcolor="#6F7A72", fontweight='bold')
        plt.xlabel("Valores", color="#710961", backgroundcolor="#6F7A72", fontweight='bold')
        plt.ylabel(columna.title(), color='#710961', backgroundcolor='#6F7A72', fontweight='bold')    
        # Mostrar Grafico
        plt.show()



    def cancelar(self):
        # Crear un cuadro de dialogo personalizado
        mensaje = QMessageBox()
        mensaje.setWindowTitle("Confirmación")
        mensaje.setText("¿Desea limpiar la interfaz?")
        mensaje.setIcon(QMessageBox.Question)
        boton_si = mensaje.addButton("Si", QMessageBox.YesRole)
        boton_no = mensaje.addButton("No", QMessageBox.NoRole)
        mensaje.setStyleSheet("background-color: #034765; color:#FFFFFF;")
        boton_si.setStyleSheet("color:#FFFFFF; background-color:#05587C;")
        boton_no.setStyleSheet("color:#FFFFFF; background-color:#05587C;")
        mensaje.setWindowIcon(QIcon(r"C:\Users\soporte\Desktop\Python Analisis Ed\Aplicacion\Alas.png"))

        # Ejecutar ventana
        mensaje.exec()
        # Verificar el boton Presionado
        if mensaje.clickedButton() == boton_si:
            #Limpiar las tablas
            self.tablaDatos.setRowCount(0)
            self.tablaDatos.setColumnCount(0)
            self.tablaResumen.setRowCount(0)
            self.tablaResumen.setColumnCount(0)
            # Limpiar Memoria del sistema
            self.df = pandas.DataFrame()
            self.dfConsultado = pandas.DataFrame()
            # Limpiar Caja de consultas
            self.txtConsulta.clear()


        

# Estructura de arranque
if __name__ == "__main__":
    # Instancia a nivel terminal
    app = QApplication(sys.argv)
    # Instancia a nivel grafico
    aplicacion = Aplicacion()
    # Muestro la interfaz
    aplicacion.show()
    # Ejecutar la app
    app.exec()

