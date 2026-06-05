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
import sys
try:
    import sqlite3
    from datetime import date
    import os
    import time
    import pandas as pd
    import re
except ModuleNotFoundError:
    print("""\033[31m[!] ERROR, \033[0m\033[1mprimero activa el entorno virtual
\033[0mSi seguiste las instruccionoes puedes activarlo con:
\033[4msource env/bin/activate
\033[0msi usas windows usa:
\033[4mcall .\\env\\Scripts\\activate
\033[0men CMD o 
\033[4m.\\env\\Scripts\\Activate.ps1
\033[0men Powershell""")
    sys.exit(1)

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
        or not isinstance(descripcion, str)):
            raise ValueError("\033[31mEl formato de los datos es incorrecto")
        
        try:
            self.cursor.execute("INSERT INTO finanzas (fecha, monto, descripcion) VALUES (?, ?, ?)", (str(fecha), dinero, descripcion))
            self.conexion.commit()
        except sqlite3.IntegrityError as e:
            print(e)
            raise ValueError("\033[31mYa existe un registro con esa fecha")


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


    def usuario(self, usuario):
        self.cursor.execute("""
            REPLACE INTO finanzas (id, user)
            VALUES (1, ?)
        """, (usuario,))
        self.conexion.commit()


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


class Operaciones:
    def __init__ (self, base):
        self.base = base


    @staticmethod
    def fechas():
        while True:
            try:
                anio = int(input("ingresa el anio \n--> "))
                mes = int(input("ingresa el mes \nejemplo: 03 \n--> "))
            except ValueError:
                limpiar()
                print("~" * 50)
                print("\033[31mERROR, \033[1mvuelve a intentarlo\n")
            else:
                if mes > 12 or mes < 1:
                    limpiar()
                    print(f"\033[31mERROR, \033[1mingresaste el mes {mes}")
                    print("\033[0mingresa solo numeros del 1 al 12")
                    time.sleep(1.2)
                    continue
                else:
                    break
        inicio = date(anio, mes, 1)
        if mes == 12:
            fin = date(anio + 1, 1, 1)
        else:
            fin = date(anio, mes + 1, 1)
        return inicio, fin


    def movimiento(self):
        while True:
            while True:
                try:
                    monto = float(input("ingresa el monto que quieres registrar: \n--> "))
                except ValueError:
                    print("\033[31mERROR, \033[1mingresa solo numeros")
                else:
                    break
            confirmacion = input(f"ingresaste ${monto}, es correcto? (s/n) ")
            print()
            if confirmacion == 's':
                break
            else:
                limpiar()
        fecha = date.today()
        descripcion = input("ingresa la descripcion del gasto: \n--> ")
        self.base.insertar(fecha, monto, descripcion)


    def ver_todo(self):
        filas = self.base.ver_todos()
        for fila in filas:
            print(fila)


    def ver_fecha(self):
        dia = input("ingresa la fecha de la que quieres conocer los movimientos \nejemplo: 2026-03-21 \n--> ")
        filas = self.base.ver_fechas(dia)
        for fila in filas:
            print(fila)


    def ver_mes(self):
        inicio, fin = Operaciones.fechas()
        filas = self.base.ver_meses(inicio, fin)
        for fila in filas:
            print(fila)


    def total_mes(self):
        inicio, fin = Operaciones.fechas()
        total = self.base.mes_total(inicio, fin)
        print(f"${total}")


    def camb_user(self):
        print(f"tu usuario actual es {self.base.ver_usuario()}")
        confirmacion = input("quieres cambiarlo? (s/n) ")
        if confirmacion == 's':
            usuario = input("ingresa tu nuevo usuario: ")
            self.base.usuario(usuario)
        print(f"nombre de usuario cambiado exitosamente a: {self.base.ver_usuario()}")


def limpiar():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')
def main(db, op):
    limpiar()
    if db.comp_usuario():
        usuario = input("\033[4mingresa tu nombre: ")
        db.usuario(usuario)
    while True:
        limpiar()
        print("\033[0m" + "~" * 50)
        print(f"\n\033[1mBIENVENIDO A COFB {db.ver_usuario()}")
        print("\n\033[0mque quieres hacer?")
        print("""\n
1 = registrar movimiento 
2 = ver todos los movimientos 
3 = ver movimientos por fecha 
4 = ver movimientos por mes 
5 = ver total de un mes 
6 = pasar datos a excel 
7 = cambiar usuario
0 = salir""")
        print()
        print("~" * 50)
        modo = input("\n--> ")
        time.sleep(0.4)
        limpiar()
        print("~" * 50, "\n")
        if modo == '1':
            op.movimiento()
            print("\033[32mExito, movimiento registrado\033[0m")
        elif modo == '2':
            op.ver_todo()
        elif modo == '3':
            op.ver_fecha()
        elif modo == '4':
            op.ver_mes()
        elif modo == '5':
            op.total_mes()
        elif modo == '6':
            db.excel()
        elif modo == '7':
            op.camb_user()
        elif modo == '0':
            break
        else:
            print("\033[31mError, \033[1mopcion no disponible")
        print()
        print("~" * 50)
        input("\npresiona Enter para continuar")


if __name__ == "__main__":
    base_datos = Database('db')
    operacion = Operaciones(base_datos)
    base_datos.conectar()
    main(base_datos, operacion)
    base_datos.cerrar()
