# <p align="center"> <img src="logo.png" width="200"> <br> **COFB** </p>

programa que ayuda a gestionar las finanzas del usuario registrando todos los movimintos financieros que realice

## Significado
El acronimo *C.O.F.B* significa **Control de Operaciones Financieras Basicas**

## Descripcion
Es un porgrama pensado para ayudar al usuario a gestionar su dinero guardando los movimientos economicos que realice el usuario guardandolos en una base de datos local y registrando la fecha, asi contando tambien con funciones para ver los moviientos realizados y pudiendo crear un archivo de excel para poder visualizar los datos de forma mas visual

## Funciones
* **SQLite** para la base de datos
* **OS** para crear la carpeta
* **datetime** para poder guardar la fecha
* **POO** codigo hecho utilizando la POO (programacion orientada a objetos)
* **Pandas** para crear el excel
* **instalacion:** incluye un script de instalacionpara automatizarla

## Tecnologias
* **Lenguaje:** Python
* **Sistema:** Arch Linux/Void Linux
* **Editor:** Neovim

## Como ejecutar
### Movil:
si te gustaria utilizar el programa en un dispositivo movil (Android o Iphone) sigue los siguientes pasos:
- instala el programa **Termux** en **Android** o **ISH Shell** en **Iphone**
    * puedes instalar **Termux** haciendo click [aquí](https://f-droid.org/repo/com.termux_1022.apk) **NOTA:** ***NO DESCARGAR DESDE LA PLAY STORE***
    * puedes descargar **ISH Shell** desde la **App Store**
* ejecuta ``pkg update && termux setup storage`` en **Termux** y ``apk update`` en **ISH Shell**
### descargar repositorio:
para descargar este repositorio deberás tener instalado git, puedes instalarlo asi:
* **Android:** ``pkg install git``
* **IOS:** ``apk add git``
* **MacOS:** ``xcode-select --install``
* **Windows:** abre terminal de administrador y ejecuta ``winget install --id Git.Git -e --silent``
* **Linux:** usa el gestor de paquetes de tu distribucion y descarga **git**
* *en el sistema en el que estés ejecuta estos comandos:*
``git clone https://github.com/Binryan-void/cofb.git``
``cd cofb``
### instalar dependencias:
para instalar lo necesario ejecuta los siguientes comandos dependiendo tu sistema
* **Android/IOS/Linux/MacOS:** ``sh install.sh``
* **Windows:** abre una terminal de administrador y ejecuta: ``.\install.bat``
### ejecutar:
para ejecutar el programa solo ejecuta esto:
* **Android/IOS/Linux/Macos:** ``. env/bin/activate && python3 cofb_cli.py``
* **Windows:** ``call .\env\Scripts\activate`` en **CMD** o ``.\env\Scripts\Activate.ps1`` en **PowerShell**
* **Windows:** ``python cofb_cli.py``
* **NOTA:** si en la versión de Android crea el archivo de excel paselo al almacenamiento interno con el siguiente comando: 
``mv finanzas.xlsx storage/shared``
* **IOS:** en iPhone lo puede ver desde su explorador de archivos
### extra:
hay una versión del programa GUI (interfaz grafica)
para ejecutar la version GUI haz lo siguiente:
* **NOTA:** esta version solo se puede ejecutar en computadora
*deberas hacer lo mismo que para ejecutar la version de terminal dependiendo tu sistema y para ejecutar usa el comando:* ``python3 cofb_gui.py`` para **Linux y MacOS** y ``python cofb_gui.py`` en **Windows**

## Instrucciones de uso
al ejecutar el programa por primera vez te pedira tu nombre o tu usuario, al ingresar el usuario y presionar la tecla *Enter* se guardara ese nombre
luego del primer uso cada vez que se inicie el programa se vera asi:
```
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BIENVENIDO A LA CALCULADORA DE FINANZAS usuario

que quieres hacer?

1 = registrar movimiento 
2 = ver todos los movimientos
3 = ver movimientos por fecha
4 = ver movimiinetos por mes 
5 = ver total de un mes 
6 = pasar datos a excel
7 = cambiar usuario 
0 = salir

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

-->
```

* **1-** si seleccionas la opcion 1 (registrar movimientos) cambiara lo que se muestra a una pantalla que te pidde que ingreses el monto a registrar, lo que se hace ahi es ingresar la cantidad de dinero la cual se quiere registrar
para representar una ganancia se pone el numero normal (EJEMPLO: 500) y para representar un gasto puedes utilizar el simbolo - (EJEMPLO: -500) o tambien lo puedes utilizar al reves si asi te resulta mas comodo, el programa no te pondra limitaciones por como lo uses
al ingresar la cantidad y presionar Enter aparecera un menu de confirmacion, escribe *s* si es correcto el monto que ingresaste y *n* si no y presiona Enter
luego de eso te pedira que ingreses una descripcion del gasto, esto seria para que en un futuro puedas saber cuanto gastaste y en que lo gastaste, ingresa la descripcion y presiona Enter (EJEMPLO: gasolina)
luego presionaras la tecla Enter para volver al menu inicial
* **2-** si seleccionas la opcion 2 (ver todos los movimientos) en la primer fila aparecera el nombre de usuario que tengas y en las fias siguientes la fecha del gasto en formato YYYY-MM-DD, luego estara la cantidad del gasto con el simbolo que se le haya puesto y seguido de eso la descripcion del gasto
luego presiona la tecla Enter para volver al menu inicial
* **3-** si seleccionas la opcion 3 (ver movimientos por fecha) seleccionaras la fecha del cual quieres conocer los movimientos registrados en formato YYYY-MM-DD (EJEMPLO: 2026-03-21) al ingresar la fecha te mostrara todos los movimientos registrados en esa fecha
* **4-** si seleccionas la opcion 4 (ver movimientos por mes) primero te pedira que ingreses el anio del mes del que quieres conocer los movimientos
al ingresar el anio (EJEMPLO: 2026) te pedira que ingreses el mes, y como ejemplo el programa pone 03, el mes debe de estar exactamente en ese formato, si es un numero de 1 solo dijito (ejemplo: abril que es el 4) se le debe agregar un 0 por delante (EJEMPLO: abril seria el 04) y te mostrara todos los movimientos de ese mes 
luego presiona Enter para volver al menu principal
* **5-** si seleccionas la opcion 5 (ver total de un mes) te pedira que ingreses el anio y el mes al igual que con la opcion 4, pero al escribirlo de dara el total de gastos y ganancias automaticamente de ese mes, para que no tengas que sumarlo todo manualmente
luego presiona Enter para volver al menu principal
* **6-** si seleccionas la opcion 6 (pasar datos a excel) aparecera algo de texto y se creara un archivo de excel con todos los datos que hayas ingresado en el programa
luego presiona Enter para volver al menu principal
* **7-** si seleccionas la opcion 7 (cambiar usuario) te mostrara tu nombre de usuario actual, con el cual te saluda el programa en el menu principal, si deseas cambiarlo presiona la tecla *s* y si no quieres cambiarlo escribe *n*
si escribes *s* te pedira tu nuevo nombre de usuairo, ahi escribiras el que quieres que sea tu nuevo nombre de usuario
depues al presionar Enter se mostrara tu nuevo nombre de usuario y un aviso que dice "para ver los cambios vuelva a iniciar el programa" por lo cual, antes de cerrar el programa se seguira viendo el nombre de usuario anterior
luego presiona Enter para volver al menu inicial
* **0-** al seleccionar la opcion **0** saldras del programa

## Desarrollado por *Bryan David Pérez Arana*
