#################################################################################
# Copyright (C) 2026 Binryan-void
#
# Este programa es software libre: puedes redistribuirlo y/o modificarlo
# bajo los terminos de la Licencia Publica General GNU publicada por la 
# Free Software Foundation, ya sea la version 3 de la Licencia o 
# (a tu eleccion) cualquier version posterior
#
# Este programa se distribuye con la esperanza de que sea util, pero 
# SIN GARANTIA ALGUNA; ni siquiera garantia implicita de 
# MERCANTILIDAD o APTITUD PARA UN PROPOSITO DETERMINADO.
# Consulte la Licencia Publica General GNU para obtener mas detalles.
#
# Deberias haber recibido una copia de la Licencia Publica General GNU 
# junto con este programa. Si no es asi, consulta <https://www.gnu.org/licenses/>.
##################################################################################
try:
    import sqlite3
    from datetime import date, datetime
    import os
    import pandas as pd
    import re
    import tkinter as tk
    from tkinter import messagebox, ttk
except ModuleNotFoundError:
    print("""ERROR, primero activa el entorno virtual
Si seguiste las instruccionoes puedes activarlo con:
source env/bin/activate
si usas windows usa:
call .\\env\\Scripts\\activate
en CMD o 
.\\env\\Scripts\\Activate.ps1
en Powershell""")

class Database:
    def __init__(self, carpeta):
        self.carpeta = carpeta
        self.conexion = None
        self.cursor = None
        self.conectar()
        self.tabla()


    def conectar(self):
        if not os.path.exists(self.carpeta):
            os.makedirs(self.carpeta)
        self.conexion = sqlite3.connect(os.path.join(self.carpeta, 'finanzas.db'))
        self.cursor = self.conexion.cursor()


    def tabla(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS finanzas (
            id INTEGER PRIMARY KEY,
            user TEXT,
            fecha TEXT,
            monto REAL,
            descripcion TEXT
            )
        ''')


    def insertar(self, fecha, dinero, descripcion):
        if (not isinstance(fecha, date) 
            or not isinstance(dinero, float) 
            or not isinstance(descripcion, str)
        ):
            messagebox.showerror("Error de Formato", "El formato de los datos es incorrecto")
        
        try:
            self.cursor.execute("INSERT INTO finanzas (fecha, monto, descripcion) VALUES (?, ?, ?)", (str(fecha), dinero, descripcion))
            self.conexion.commit()
            messagebox.showinfo("Exito", "Movimiento registrado exitosamente")
        except sqlite3.IntegrityError as e:
            messagebox.showerror("Error de Duplicado", f"Ocurrio un error de duplicado de los datos \nEl error es {e}")

    def ver_todos(self):
        self.cursor.execute("SELECT * FROM finanzas")
        filas = self.cursor.fetchall()
        return filas


    def ver_fechas(self, dia):
        self.cursor.execute("SELECT * FROM finanzas WHERE fecha = ?", (dia,))
        filas = self.cursor.fetchall()
        return filas


    def ver_meses(self, inicio, fin):
        self.cursor.execute("SELECT * FROM finanzas WHERE fecha >= ? AND fecha < ?", (str(inicio), str(fin)))
        filas = self.cursor.fetchall()
        return filas


    def mes_total(self, inicio, fin):
        self.cursor.execute("SELECT SUM(monto) FROM finanzas WHERE fecha >= ? AND fecha < ?", (str(inicio), str(fin)))
        total_res = self.cursor.fetchone()[0] or -0 
        total = total_res if total_res is not None else 0.0
        return total


    def excel(self):
        data = pd.read_sql_query("SELECT * FROM finanzas", self.conexion)
        for col in data.select_dtypes(include=['object']):
            data[col] = data[col].apply(lambda x: re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(x)) if x else x)
        data.to_excel('finanzas.xlsx', index=False, engine='openpyxl')
        messagebox.showinfo("Exito", "El archivo se ha creado exitosamente")


    def usuario(self, usuario):
        self.cursor.execute("""
            REPLACE INTO finanzas (id, user)
            VALUES (1, ?)
        """, (usuario,))
        self.conexion.commit()
        messagebox.showinfo("Exito", f"Se ha guardado el nombre de usuario \nTu usuario es {usuario}")


    def ver_usuario(self):
        self.cursor.execute("SELECT user FROM finanzas WHERE id = 1")
        usuario = self.cursor.fetchone()
        if usuario:
            return usuario[0]
        return ""


    def comp_usuario(self):
        self.cursor.execute("SELECT 1 FROM finanzas WHERE id = 1")
        resultado = self.cursor.fetchone()
        if resultado is None or resultado[0] is None or resultado[0] == '':
            return True
        else:
            return False


    def cerrar(self):
        self.conexion.close()


class Ventana:
    def __init__(self, base_datos): 
        self.db = base_datos
        self.root = self.obtener_root()
        self.seleccion = None
        self.continuar = tk.BooleanVar(value=False)


    @staticmethod
    def fechas(anio, mes):
        try:
            anio = int(anio) 
            mes = int(mes)
        except ValueError:
            messagebox.showerror("Error", "Ingresa solo numeros enteros \nVuelve a intentarlo")
            return None, None
        else:
            if mes > 12 or mes < 1:
                messagebox.showerror("Error", f"Ingresaste {mes} \nIngresa solo numeros entre 1 y 12")
            else:
                inicio = date(anio, mes, 1)
                if mes == 12:
                    fin = date(anio + 1, 1, 1)
                else:
                    fin = date(anio, mes + 1, 1)
                return inicio, fin


    @staticmethod
    def scrollable(root):
        frame = ttk.Frame(root)
        frame.pack(padx=10, pady=10, fill="both", expand=True)
        scrollbar = ttk.Scrollbar(frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        txt_scrollable = tk.Text(frame, wrap="word", yscrollcommand=scrollbar.set)
        txt_scrollable.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=txt_scrollable.yview)
        return txt_scrollable


    def obtener_root(self):
        root = tk.Tk()
        root.title("Control de Operaciones Financieras Basicas")
        root.geometry("800x500")
        return root


    def regresar(self):
        self.root.wait_variable(self.continuar)


    def limpiar_pantalla(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        self.root.config(menu="")


    def menu(self):
        self.esperar = tk.BooleanVar(value=False)
        self.seleccion = (None)
        self.limpiar_pantalla()

        def devolver_seleccion():
            self.seleccion = seleccion.get()
            self.esperar.set(True)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Menu Principal", menu=menu)

        bienvenida = tk.Label(self.root, text=f"BIENVENIDO A LA CALCULADORA DE FINANZAS {self.db.ver_usuario()}", font=("Arial", 15))
        bienvenida.pack(pady=10)
        pregunta = tk.Label(self.root, text="que quieres hacer", font=("Arial", 12))
        pregunta.pack(pady=15)

        opciones = [
            "registrar movimiento",
            "ver todos los movimientos",
            "ver movimientos por fecha",
            "ver movimientos por mes",
            "ver total de un mes",
            "pasar datos a excel",
            "cambiar usuario",
            "salir"
        ]
        seleccion = tk.StringVar(value="1")

        for indice, opcion in enumerate(opciones):
            if opcion != "salir":
                radio = tk.Radiobutton(
                    self.root,
                    text=opcion,
                    variable=seleccion,
                    value=str(indice + 1),
                )
            else:
                radio = tk.Radiobutton(
                    self.root,
                    text=opcion,
                    variable=seleccion,
                    value="0" 
                )
            radio.pack(anchor=tk.W, padx=20, pady=10)

        tk.Frame(self.root, height=10).pack()
        boton = tk.Button(self.root, text="Aceptar", command=devolver_seleccion)
        boton.pack(pady=10)
        self.root.wait_variable(self.esperar)
        return self.seleccion


    def movimiento(self):
        self.continuar.set(False)
        self.limpiar_pantalla()
        
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Registrar Movimiento", menu=menu)
        self.mov_monto()

        
    def mov_monto(self):
        def validar_monto():
            try:
                float(self.monto.get())
            except ValueError:
                messagebox.showerror("ERROR", "Ingresa unicamente numeros")
            else:
                self.mov_confirmacion()

        def volver():
            self.continuar.set(True)

        pedir_monto = tk.Label(self.root, text="Ingresa el monto que quieres registrar:", font=("Arial", 12))
        pedir_monto.pack(pady=10)
        self.monto = tk.Entry(self.root)
        self.monto.pack(pady=10)
        self.enviar = tk.Button(self.root, text="Enviar", command=validar_monto)
        self.enviar.pack(pady=10)
        self.volver_menu = tk.Button(self.root, text="Volver al Menu", command=volver)
        self.volver_menu.pack(side=tk.TOP, padx=20, pady=10)


    def mov_confirmacion(self):
        self.enviar.destroy()
        self.volver_menu.destroy()

        confirmacion = tk.Label(self.root, text=f"Ingresaste ${self.monto.get().strip()}", font=("Arial", 12))
        confirmacion.pack(pady=10)
        self.aceptar = tk.Button(self.root, text="Aceptar", command=self.mov_descripcion)
        self.aceptar.pack(side=tk.TOP, padx=50, pady=10)
        self.volver = tk.Button(self.root, text="Volver", command=self.mov_monto)
        self.volver.pack(side=tk.TOP, padx=50, pady=10)


    def mov_descripcion(self):
        self.aceptar.destroy()
        self.volver.destroy()

        pedir_descripcion = tk.Label(self.root, text="Ingresa una descripcion:", font=("Arial", 12))
        pedir_descripcion.pack(pady=10)
        self.descripcion = tk.Entry(self.root)
        self.descripcion.pack(pady=10)
        self.enviar = tk.Button(self.root, text="Enviar", command=self.mov_registrar)
        self.enviar.pack(pady=10)


    def mov_registrar(self):
        def volver():
            self.continuar.set(True)

        self.enviar.destroy()
        monto = float(self.monto.get().strip())
        descripcion = self.descripcion.get().strip()
        hoy = date.today()

        self.db.insertar(hoy, monto, descripcion)
        volver_menu = tk.Button(self.root, text="Continuar", command=volver)
        volver_menu.pack(side=tk.TOP, pady=20, padx=10)


    def ver_todo(self):
        self.limpiar_pantalla()
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver Todos Los Movimientos", menu=menu)

        self.continuar.set(False)
        def volver():
            self.continuar.set(True)
        
        filas = self.db.ver_todos()
        dato = Ventana.scrollable(self.root)
        for fila in filas:
            dato.insert("end", f"{fila}\n")
        dato.config(state="disabled")
        volver_menu = tk.Button(self.root, text="Continuar", command=volver)
        volver_menu.pack(side=tk.BOTTOM, pady=20, padx=10)


    def ver_fecha(self):
        self.limpiar_pantalla()
        self.continuar.set(False)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver Movimientos Por Fecha", menu=menu)

        pedir_fecha = tk.Label(self.root, text="Ingresa la fecha de la que quieres los movimientos", font=("Arial", 12))
        pedir_fecha.pack(pady=10)
        ejemplo = tk.Label(self.root, text="Ejemplo: 2026-06-04", font=("Arial", 12))
        ejemplo.pack(pady=10)
        self.fecha = tk.Entry(self.root)
        self.fecha.pack(pady=10)
        self.enviar = tk.Button(self.root, text="Enviar", command=self.datos_fecha)
        self.enviar.pack(pady=(10))


    def datos_fecha(self):
        def volver():
            self.continuar.set(True)
    
        self.enviar.destroy()
        try:
            fecha = self.fecha.get().strip()
            fecha_valida = datetime.strptime(fecha, "%Y-%m-%d").date()
            fecha_valida = str(fecha_valida)
        except ValueError:
            messagebox.showerror("Error", "formato o fecha invalida \nintenta nuevamente")
            self.ver_fecha()
        else:
            self.limpiar_pantalla()
            filas = self.db.ver_fechas(fecha_valida)
            dato = Ventana.scrollable(self.root)
            for fila in filas:
                dato.insert("end", f"{fila}\n")
            dato.config(state="disabled")
            volver_menu = tk.Button(self.root, text="Continuar", command=volver)
            volver_menu.pack(side=tk.TOP, pady=20, padx=10)


    def ver_mes(self):
        self.limpiar_pantalla()
        self.continuar.set(False)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver Movimientos Por Mes", menu=menu)

        anio_text = tk.Label(self.root, text="Ingresa el anio", font=("Arial", 12))
        anio_text.pack(pady=10)
        self.anio = tk.Entry(self.root)
        self.anio.pack(pady=10)
        mes_text = tk.Label(self.root, text="Ingresa el mes \nEjemplo: 05", font=("Arial", 12))
        mes_text.pack(pady=10)
        self.mes = tk.Entry(self.root)
        self.mes.pack(pady=10)

        self.enviar = tk.Button(self.root, text="Enviar", command=self.datos_mes)
        self.enviar.pack(pady=10)


    def datos_mes(self):
        def volver():
            self.continuar.set(True)

        anio = self.anio.get().strip()
        mes = self.mes.get().strip()

        self.limpiar_pantalla()
        inicio, fin = Ventana.fechas(anio, mes)
        if inicio is None:
            self.ver_mes()
        else:
            filas = self.db.ver_meses(inicio, fin)
            dato = Ventana.scrollable(self.root)
            for fila in filas:
                dato.insert("end", f"{fila}\n")
            dato.config(state="disable")
    
            volver_menu = tk.Button(self.root, text="Continuar", command=volver)
            volver_menu.pack(side=tk.TOP, pady=(50, 30), padx=(40, 5))


    def total_mes(self):
        self.limpiar_pantalla()
        self.continuar.set(False)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver Total De Un Mes", menu=menu)


        anio_text = tk.Label(self.root, text="Ingresa el anio", font=("Arial", 12))
        anio_text.pack(pady=10)
        self.anio = tk.Entry(self.root)
        self.anio.pack(pady=10)
        mes_text = tk.Label(self.root, text="Ingresa el mes \nEjemplo: 05", font=("Arial", 12))
        mes_text.pack(pady=10)
        self.mes = tk.Entry(self.root)
        self.mes.pack(pady=10)

        self.enviar = tk.Button(self.root, text="Enviar", command=self.datos_total)
        self.enviar.pack(pady=10)


    def datos_total(self):
        def volver():
            self.continuar.set(True)

        anio = self.anio.get().strip()
        mes = self.mes.get().strip()
        inicio,  fin = Ventana.fechas(anio, mes)
        if inicio is None:
            self.total_mes()
        else:
            total = self.db.mes_total(inicio, fin)
            dato = tk.Label(self.root, text=total, font=("Arial", 12))
            dato.pack(pady=10)
            volver_menu = tk.Button(self.root, text="Continuar", command=volver)
            volver_menu.pack(side=tk.TOP, pady=20, padx=10)


    def excel(self):
        self.limpiar_pantalla()
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Pasar Datos A Excel", menu=menu)

        self.continuar.set(False)
        def volver():
            self.continuar.set(True)

        def conectar_excel():
            self.db.excel()
            messagebox.showinfo("Exito", "Se ha creado el archivo de excel exitosamente")
            self.continuar.set(True)

        preguntar = tk.Label(self.root, text="Quieres Pasar Los Datos A Excel?", font=("Arial", 12))
        preguntar.pack(pady=10)
        volver_menu = tk.Button(self.root, text="Volver", command=volver)
        volver_menu.pack(side=tk.TOP, pady=10, padx=50)
        crear_excel = tk.Button(self.root, text="Crear Excel", command=conectar_excel)
        crear_excel.pack(side=tk.TOP, pady=10, padx=50)


    def cambiar_usuario(self):
        self.limpiar_pantalla()
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cambiar Usuario", menu=menu)

        self.continuar.set(False)
        def volver():
            self.continuar.set(True)

        usuario = tk.Label(self.root, text=f"tu usuario actual es {self.db.ver_usuario()}", font=("Arial", 12))
        usuario.pack(pady=10)
        preguntar = tk.Label(self.root, text="Quieres cambiarlo?", font=("Arial", 12))
        preguntar.pack(pady=10)
        self.si = tk.Button(self.root, text="Si", command=self.usr_cambiar)
        self.si.pack(side=tk.TOP, pady=20, padx=10)
        self.no = tk.Button(self.root, text="No", command=volver)
        self.no.pack(side=tk.TOP, pady=20, padx=10)


    def usr_cambiar(self):
        self.si.destroy()
        self.no.destroy()
        
        pedir_usr = tk.Label(self.root, text="Ingresa tu nuevo usuario", font=("Arial", 12))
        pedir_usr.pack(pady=10)
        self.user = tk.Entry(self.root)
        self.user.pack(pady=10)
        self.enviar = tk.Button(self.root, text="Enviar", command=self.usr_guardar)
        self.enviar.pack(pady=10)


    def usr_guardar(self):
        self.enviar.destroy()
        user = self.user.get().strip()
        self.db.usuario(user)
        self.continuar.set(True)


    def salir(self):
        self.db.cerrar()
        self.root.destroy()


    def iniciar(self):
        self.root.mainloop()
        

    def obtener_user(self):
        self.limpiar_pantalla()
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ingresar usuario", menu=menu)
        self.continuar.set(False)

        pedir_usuario = tk.Label(self.root, text="ingresa tu nombre de usuario", font=("Arial", 12))
        pedir_usuario.pack(pady=10)
        self.user = tk.Entry(self.root)
        self.user.pack(pady=10)
        self.aceptar = tk.Button(self.root, text="Aceptar", command=self.obt_usr_guardar)
        self.aceptar.pack(pady=10)


    def obt_usr_guardar(self):
        self.aceptar.destroy()
        user = self.user.get().strip()
        self.db.usuario(user)
        self.continuar.set(True)
        

def main(base, vn):
    if base.comp_usuario():
       vn.obtener_user()
       vn.regresar()

    while True:
        seleccion = vn.menu()
        if seleccion == "1":
            vn.movimiento()
            vn.regresar()
        elif seleccion == "2":
            vn.ver_todo()
            vn.regresar()
        elif seleccion == "3":
            vn.ver_fecha()
            vn.regresar()
        elif seleccion == "4":
            vn.ver_mes()
            vn.regresar()
        elif seleccion == "5":
            vn.total_mes()
            vn.regresar()
        elif seleccion == "6":
            vn.excel()
            vn.regresar()
        elif seleccion == "7":
            vn.cambiar_usuario()
            vn.regresar()
        elif seleccion == "0":
            vn.salir()
        else:
            messagebox.showerror("ERROR", "Ocurrio Un Error Inesperado")


if __name__ == "__main__":
    base_datos = Database('db')
    interfaz = Ventana(base_datos)
    base_datos.conectar()
    interfaz.root.after(100, lambda: main(base_datos, interfaz))
    interfaz.iniciar()
    base_datos.cerrar()
