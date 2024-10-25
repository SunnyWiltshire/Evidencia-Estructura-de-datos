import csv
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font 
import pandas as pd
import numpy as np
import sys
import sqlite3
from sqlite3 import Error
from tabulate import tabulate
import matplotlib.pyplot as plt



ruta = []
try:
    with sqlite3.connect('RentaBicicletas.db') as conn: 
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS UNIDAD \
            (Clave INTEGER NOT NULL PRIMARY KEY, \
            Rodada INTEGER NOT NULL, \
            Color TEXT NOT NULL);")
        cursor.execute("CREATE TABLE IF NOT EXISTS CLIENTES \
            (Clave INTEGER NOT NULL PRIMARY KEY, \
            Apellidos TEXT NOT NULL, \
            Nombres TEXT NOT NULL, \
            Telefono INTEGER NOT NULL);")
        cursor.execute("CREATE TABLE IF NOT EXISTS PRESTAMO \
            (Folio INTEGER NOT NULL PRIMARY KEY, \
            Fecha_Prestamo INTEGER NOT NULL, \
            Dias_Prestamo INTEGER NOT NULL, \
            Fecha_Retorno INTEGER NOT NULL, \
            Retorno INTEGER NOT NULL, \
            Clave_Cliente INTEGER NOT NULL, \
            Clave_Unidad INTEGER NOT NULL, \
            FOREIGN KEY (Clave_Cliente) REFERENCES CLIENTES(Clave), \
            FOREIGN KEY (Clave_Unidad) REFERENCES UNIDAD(Clave));")
        print("Base de datos y tablas creadas exitosamente.") 
except Error as e:
    print(e)
except Exception:
    print(f"Se produjo el siguiente error: {sys.exc_info()}")
    
def mostrar_ruta():
    print('\nRUTA: ')
    print(" > ".join(ruta))
    
#funcion que despliega el menu principal
def menu_principal():
    ruta.append('Menú Principal')
    while True:
        mostrar_ruta()
        print("\n--- MENÚ PRINCIPAL ---")
        print("1. Registro")
        print("2. Préstamo")
        print("3. Retorno")
        print("4. Informes")
        print("5. Salir\n")

        try:
            opcion = input("Elige una de las siguientes opciones: ")
            opcion = int(opcion)

            if opcion == 1:
                ruta.append("Registro")
                menu_registro()
                ruta.pop()
            elif opcion == 2:
                ruta.append("Prestamo")
                registrar_prestamo()
                ruta.pop()
            elif opcion == 3:
                ruta.append("Retorno")
                menu_retorno()
                ruta.pop()
            elif opcion == 4:
                ruta.append("Informes")
                menu_informes()
                ruta.pop()
            elif opcion == 5:
                confirmacion = input("¿Desea salir del programa? (S/N)").upper()
                if confirmacion == "S":
                    print("Saliendo del sistema...\n")
                    break
                elif confirmacion == "N":
                    return
                else:
                    print("Opción invalida, ingrese los valores de 'S' o 'N'.")
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')

#funcion que pregunta al usuario si desea cancelar la accion que estaba haciendo
def cancelar():
    while True:
        try:
            respuesta = int(input("\nHa ocurrido un error. ¿Deseas cancelar o intentar de nuevo? \n1: cancelar  \n2: intentar de nuevo \n"))
            if respuesta == 1:
                print("Operacion cancelada.")
                return True
            elif respuesta == 2:
                return False
            else:
                print("Opción no valida. Por favor, selecciona 1 para cancelar o 2 para intentar de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')
            
#funcion que despliega el sub menú de registro
def menu_registro():
    while True:
        mostrar_ruta()
        print("\n--- SUBMENÚ REGISTRO ---")
        print("1. Registrar una unidad")
        print("2. Registrar un cliente")
        print("3. Volver al menu principal\n")

        try:
            opcion = input("Elige una opción: ")
            opcion = int(opcion)

            if opcion == 1:
                ruta.append('Unidad')
                registro_Unidad()
                ruta.pop()
            elif opcion == 2:
                ruta.append('Cliente')
                registro_Cliente()
                ruta.pop()
            elif opcion == 3:
                break
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            if cancelar():
                break

## FUNCIONES PARA EL REGISTRO DE UNA UNIDAD

#funcion que permite registrar una unidad lista para un prestamo
def registro_Unidad():
    try:
        # Conectar a la base de datos
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            while True:
                mostrar_ruta()
                opcion = input("¿Deseas realizar un registro de unidad? (S/N): ").upper()

                if opcion == "S":
                    print("\n--- REGISTRO DE UNIDAD ---")
                    
                    # Obtener la clave más alta existente en la tabla UNIDAD y sumar 1
                    cursor.execute("SELECT IFNULL(MAX(Clave), 0) FROM UNIDAD")
                    clave = cursor.fetchone()[0] + 1
                    
                    while True:
                        entrada = input('Ingrese la rodada de la unidad (20, 26 o 29): ')
                        try:
                            rodada = int(entrada)
                            if rodada in [20, 26, 29]:
                                print("""\nTenemos disponibles los siguientes colores: \nRojo \nAzul \nAmarillo \nVerde \nRosa""")
                                color = input("Elige un color para la bicicleta: ").upper()
                                
                                if color in ["ROJO", "AZUL", "AMARILLO", "VERDE", "ROSA"]:
                                    # Insertar los datos en la tabla UNIDAD
                                    cursor.execute("INSERT INTO UNIDAD (Clave, Rodada, Color) VALUES (?, ?, ?)", 
                                                   (clave, rodada, color))
                                    conn.commit()  # Confirmar la transacción

                                    print(f"Unidad registrada con éxito. Clave: {clave}, Rodada: {rodada}, Color: {color}")
                                    return False
                                else:
                                    print("Color no válido. Por favor, elija entre Rojo, Azul, Amarillo, Verde, o Rosa.")
                                    if cancelar():
                                        return
                            else:
                                print("Por favor, ingrese un valor válido (20, 26 o 29).")
                                if cancelar():
                                    break

                        except ValueError:
                            print('Favor de ingresar un valor válido\n')
                            if cancelar():
                                break
                elif opcion == "N":
                    return False
                else:
                    print("Opción inválida. Debes ingresar 'S' o 'N'.")
                    if cancelar():
                        break
                    return
    except Error as e:
        print(f"Error de base de datos: {e}")
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()}")

## FUNCIONES PARA EL REGISTRO DE UN CLIENTE

#funcion que permite registrar un cliente          
def registro_Cliente():
    try:
        # Conectar a la base de datos
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            while True:
                mostrar_ruta()
                opcion = input("¿Deseas realizar un registro de cliente? (S/N): ").upper()

                if opcion == "S":
                    print("\n--- REGISTRO DE CLIENTE ---")
                    
                    # Obtener la clave más alta existente en la tabla CLIENTES y sumar 1
                    cursor.execute("SELECT IFNULL(MAX(Clave), 0) FROM CLIENTES")
                    clave_cliente = cursor.fetchone()[0] + 1
                    
                    # Captura de Apellidos
                    while True:
                        apellidos = input("Ingresa el apellido del cliente (max 40 caracteres): ")
                        if apellidos.replace(" ", "").isalpha() and len(apellidos) <= 40:
                            break
                        else:
                            print("Apellidos no válidos.")
                            if cancelar():
                                return
                    
                    # Captura de Nombre
                    while True:
                        nombre = input("Ingresa el nombre del cliente (max 40 caracteres): ")
                        if nombre.replace(" ", "").isalpha() and len(nombre) <= 40:
                            break
                        else:
                            print("Nombre no válido.")
                            if cancelar():
                                return
                    
                    # Captura de Teléfono
                    while True:
                        telefono = input("Ingrese el número de teléfono (10 dígitos): ")
                        if telefono.isdigit() and len(telefono) == 10:
                            break
                        else:
                            print("Teléfono no válido.")
                            if cancelar():
                                return
                    
                    # Insertar el cliente en la tabla CLIENTES
                    cursor.execute("INSERT INTO CLIENTES (Clave, Apellidos, Nombres, Telefono) VALUES (?, ?, ?, ?)", 
                                   (clave_cliente, apellidos, nombre, telefono))
                    conn.commit()  # Confirmar la transacción

                    print(f"Cliente registrado con éxito. Clave: {clave_cliente}, Nombre: {nombre} {apellidos}, Teléfono: {telefono}")
                    
                    # Salir del bucle después de registrar
                    break
                elif opcion == "N":
                    break  # Salir del bucle para regresar al menú
                else:
                    print("Opción inválida. Debes ingresar 'S' o 'N'.")
                    if cancelar():
                        break
                    return
    except Error as e:
        print(f"Error de base de datos: {e}")
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()}")

## FUNCIONES PARA EL REGISTRO DE UN PRÉSTAMO

## Apartado para registrar los préstamos
def registrar_prestamo():
    try:
        # Conectar a la base de datos
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            
            mostrar_ruta()
            while True:
                # Mostrar préstamos existentes (función opcional)
                tab_prestamos() 
                opcion = input("¿Deseas realizar un registro de préstamos? (S/N): ").upper()
                
                if opcion == "S":
                    print("\n--- REGISTRO DE PRÉSTAMO ---")
                    
                    fecha_actual = datetime.now().date()
                    
                    # Obtener el próximo folio de préstamo
                    cursor.execute("SELECT IFNULL(MAX(Folio), 0) FROM PRESTAMO")
                    folio = cursor.fetchone()[0] + 1

                    # Captura de la clave de la unidad
                    while True:
                        Clave_unidad = input("Clave de la unidad: ")
                        cursor.execute("SELECT Clave FROM UNIDAD WHERE Clave=?", (Clave_unidad,))
                        if cursor.fetchone():
                            Clave_unidad = int(Clave_unidad)
                            break
                        print("La clave de la unidad no es válida.")
                        if cancelar(): return

                    # Captura de la clave del cliente
                    while True:
                        Clave_cliente = input("Clave del cliente: ")
                        cursor.execute("SELECT Clave FROM CLIENTES WHERE Clave=?", (Clave_cliente,))
                        if cursor.fetchone():
                            Clave_cliente = int(Clave_cliente)
                            break
                        print("La clave del cliente no es válida.")
                        if cancelar(): return

                    # Elección de la fecha del préstamo
                    while True:
                        eleccion_de_fecha = input("¿Deseas que la fecha sea la del día de hoy?\n1. Sí\n2. No\nElige una opción: ")
                        if eleccion_de_fecha.isdigit():
                            eleccion_de_fecha = int(eleccion_de_fecha)
                            if eleccion_de_fecha == 1:
                                fecha_prestamo = fecha_actual
                                break
                            elif eleccion_de_fecha == 2:
                                while True:
                                    fecha_a_elegir = input("Indica la fecha del préstamo (MM/DD/AAAA): ")
                                    try:
                                        fecha_prestamo = datetime.strptime(fecha_a_elegir, "%m/%d/%Y").date()
                                        if fecha_prestamo >= fecha_actual:
                                            break
                                        print("La fecha no puede ser anterior a la actual.")
                                    except ValueError:
                                        print("Formato de fecha incorrecto, intenta de nuevo.")
                                break
                        print("Opción inválida, intenta de nuevo.")
                        if cancelar(): return

                    # Cantidad de días del préstamo
                    while True:
                        Cantidad_de_dias = input("¿Cuántos días de préstamo solicitas?: ")
                        if Cantidad_de_dias.isdigit() and int(Cantidad_de_dias) > 0:
                            Cantidad_de_dias = int(Cantidad_de_dias)
                            fecha_de_retorno = fecha_prestamo + timedelta(days=Cantidad_de_dias)
                            print(f"La fecha en la que se debe regresar la unidad es el: {fecha_de_retorno.strftime('%m/%d/%Y')}")
                            break
                        print("La cantidad de días debe ser un número mayor a 0.")
                        if cancelar(): return

                    # Registro del préstamo en la base de datos
                    cursor.execute(
                        "INSERT INTO PRESTAMO (Folio, Fecha_Prestamo, Dias_Prestamo, Fecha_Retorno, Retorno, Clave_Cliente, Clave_Unidad) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (folio, fecha_prestamo, Cantidad_de_dias, fecha_de_retorno, False, Clave_cliente, Clave_unidad)
                    )
                    conn.commit()  # Confirmar la transacción

                    print(f"Préstamo registrado exitosamente. Folio: {folio}, Cliente: {Clave_cliente}, Unidad: {Clave_unidad}, Fecha de Préstamo: {fecha_prestamo}")

                    # Salir del bucle después de registrar
                    break

                elif opcion == 'N':
                    break  # Salir del bucle para regresar al menú

                else: 
                    print("Favor de indicar un valor correcto (S/N)")
                    if cancelar(): return
    except Error as e:
        print(f"Error de base de datos: {e}")
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()}")

## Impresión tabular que muestra los clientes y unidades al momento de realizar un préstamo
def tab_prestamos():
    try:
        # Conexión a la base de datos
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            # Obtener los datos de los clientes
            cursor.execute("SELECT Clave, Nombres, Apellidos FROM CLIENTES")
            clientes = cursor.fetchall()

            # Obtener los datos de las unidades
            cursor.execute("SELECT Clave, Rodada, Color FROM UNIDAD")
            unidades = cursor.fetchall()

            # Generar tablas separadas para clientes y unidades
            print("\n--- Reporte de Clientes ---")
            headers_clientes = ["ID Cliente", "Nombre", "Apellido"]
            print(tabulate(clientes, headers=headers_clientes, tablefmt="rounded_outline"))

            print("\n--- Reporte de Unidades ---")
            headers_unidades = ["ID Unidad", "Rodada", "Color"]
            print(tabulate(unidades, headers=headers_unidades, tablefmt="rounded_outline"))

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")

## MENU DE RETORNO        
#Función que despliega menú para hacer el retorno de la unidad
def menu_retorno():
    mostrar_ruta()

    # Verificar si hay préstamos en la base de datos
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            # Obtener préstamos donde Retorno es 0 (no devuelto)
            cursor.execute("SELECT Folio, Clave_Cliente, Clave_Unidad FROM PRESTAMO WHERE Retorno = 0")
            prestamos = cursor.fetchall()

            if prestamos:
                print("\n--- SUBMENÚ RETORNO ---")
                while True:
                    opcion = input("¿Deseas retornar una unidad? (S/N): ").upper()

                    if opcion == "S":
                        while True:
                            print(f"{'Folio':^8}{'Clave del Cliente': <20}{'Clave de la Unidad': <22}")
                            print("=" * 50)

                            # Mostrar los préstamos pendientes de retorno
                            for folio, clave_cliente, clave_unidad in prestamos:
                                print(f"{folio:^8}{clave_cliente: <20}{clave_unidad: <20}")

                            print("=" * 50)

                            numdefolio = input("\nIngrese el número de folio de su préstamo: \n")
                            try:
                                numdefolio = int(numdefolio)

                                # Verificar si el folio ingresado existe en los préstamos no retornados
                                if any(f[0] == numdefolio for f in prestamos):
                                    today = datetime.now().date()

                                    # Actualizar el campo 'Retorno' a 1 (True) en la base de datos
                                    cursor.execute("UPDATE PRESTAMO SET Retorno = 1 WHERE Folio = ?", (numdefolio,))
                                    conn.commit()  # Confirmar los cambios en la base de datos
                                    print("\nRetornó su unidad exitosamente el día", today.strftime('%m/%d/%Y'), "\n")
                                    break
                                else:
                                    print("El número de folio no existe, inténtalo de nuevo.")
                                    if cancelar():
                                        break
                            except ValueError:
                                print("Por favor, ingrese un número entero.")
                                if cancelar():
                                    break
                        break
                    elif opcion == "N":
                        print("Volviendo al menú principal.")
                        break
                    else:
                        print("Opción inválida. Por favor, elige 'S' o 'N'.")
            else:
                print("No hay ningún préstamo realizado.")
    except sqlite3.Error as e:
        print(f"Error al conectar con la base de datos: {e}")


## MENU INFORMES
def menu_informes():
    mostrar_ruta()
    while True:
        print("\n--- MENÚ INFORMES ---")
        print("1. Reportes")
        print("2. Análisis")
        print("3. Volver al menú\n")

        try:
            opcion = input("Elige una de las siguientes opciones: ")
            opcion = int(opcion)

            if opcion == 1:
                ruta.append("Reportes")
                submenu_reportes()
                ruta.pop()
            elif opcion == 2:
                ruta.append("Análisis")
                submenu_analisis()
                ruta.pop()
            elif opcion == 3:
                return False
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')

def import_clientes():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conexion:
            cursor = conexion.cursor()
            # Consulta para obtener los datos de los clientes
            cursor.execute("SELECT * FROM CLIENTES")
            clientes = cursor.fetchall()
        clientes_dict={}
        if clientes:
            for cliente in clientes:
                clientes_dict[cliente[0]] = cliente[1],cliente[2],cliente[3]
            return clientes_dict
        else:
            print('No hay clientes registrados')
    except sqlite3.Error as e:
        print(e)
    except Exception:
        print('Algo ha salido mal...')

def import_unidades():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conexion:
            cursor = conexion.cursor()
            # Consulta para obtener los datos de las unidades
            cursor.execute("SELECT * FROM UNIDAD")
            unidades = cursor.fetchall()

        unidades_dict={}
        if unidades:
            for unidad in unidades:
                unidades_dict[unidad[0]] = unidad[1],unidad[2]
            return unidades_dict
        else:
            print('No hay unidades registradas')
    except sqlite3.Error as e:
        print(e)
    except Exception:
        print('Algo ha salido mal...')

def import_prestamos():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conexion:
            cursor = conexion.cursor()
            # Consulta para obtener los datos de los préstamos
            cursor.execute("SELECT * FROM PRESTAMO")
            prestamos = cursor.fetchall()

        prestamos_dict={}
        while True:
            if prestamos:
                for prestamo in prestamos: 
                    prestamos_dict[prestamo[0]] = prestamo[1],prestamo[2],prestamo[3],prestamo[4],prestamo[5],prestamo[6] 
                else: 
                    return prestamos_dict
            else:
                print('No hay prestamos registrados')
    except sqlite3.Error as e:
        print(e)
    except Exception:
        print('Algo ha salido mal...')


## MENU DE REPORTES
def submenu_reportes():

  while True:
    mostrar_ruta()
    print("\n--- SUBMENÚ REPORTES ---")
    print("1. Clientes.")
    print("2. Listado de unidades.")
    print("3. Retrasos.")
    print("4. Préstamos por retornar.")
    print("5. Préstamos por periodo.")
    print("6. Salir al menú principal\n")

    try:
        reporte_opcion = int(input("Elige alguna de las opciones mencionadas: "))
        if reporte_opcion == 1:
            ruta.append('Clientes')
            exportar_clientes()
            ruta.pop()
        elif reporte_opcion == 2:
            ruta.append('Unidades')
            exportar_unidades()
            ruta.pop()
        elif reporte_opcion == 3:
            ruta.append('Retrasos')
            retrasos(prestamos, clientes)
            ruta.pop()
        elif reporte_opcion == 4:
            ruta.append('Prestamos por retornar')
            reporte_prestamos_por_retornar(prestamos)
            ruta.pop()
        elif reporte_opcion == 5:
            ruta.append('Prestamos por periodo')
            prestamos_por_periodo(prestamos)
            ruta.pop()
        elif reporte_opcion == 6:
            return False    
        else:
            print("Ingresa una opción válida")
    except Exception as error_name:
        print(f"Ha ocurrido un error: {error_name}")
        if cancelar():
            break

## SUBMENU REPORTES CLIENTES
def exportar_clientes():
    # Conectar a la base de datos SQLite
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta para obtener los clientes
    cursor.execute("SELECT * FROM CLIENTES")
    clientes = cursor.fetchall()

    if clientes:  # Si hay clientes en la base de datos
        # Mostrar el reporte tabular con la librería 'tabulate'
        headers = ["Clave", "Apellidos", "Nombres", "Teléfono"]
        print(tabulate(clientes, headers, tablefmt="rounded_outline"))

        while True:
            try:
                export_opcion = int(input("Elige una opción de exportación: \n1. CSV\n2. Excel\n3. Ambos\n4. Salir al submenú\n"))

                # Exportar en formato CSV
                if export_opcion == 1:
                    export_csv_clientes()
                
                # Exportar en formato Excel
                elif export_opcion == 2:
                    export_excel_clientes()
                
                # Exportar en ambos formatos, CSV y Excel
                elif export_opcion == 3:
                    export_csv_clientes()
                    export_excel_clientes()
                
                # Salir al submenú
                elif export_opcion == 4:
                    break
                
                else:
                    print("Elige una opción válida.")
                    if cancelar():
                        break
            except ValueError:
                print("Error: Debes ingresar un número entero que sea válido.")
                if cancelar():
                    break
            except Exception as name_error:
                print(f"Ha ocurrido un error inesperado: {name_error}")
                if cancelar():
                    break
    else:
        print("No hay clientes para exportar.")
    
    # Cerrar la conexión con la base de datos
    conexion.close()
    
def exportar_unidades():
    # Conectar a la base de datos SQLite
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta para obtener las unidades
    cursor.execute("SELECT * FROM UNIDAD")
    unidades = cursor.fetchall()

    if unidades:  # Si hay unidades en la base de datos
        # Mostrar el reporte tabular con la librería 'tabulate'
        headers = ["Clave", "Rodada", "Color"]
        print(tabulate(unidades, headers, tablefmt="rounded_outline"))

        while True:
            try:
                export_opcion = int(input("Elige una opción de exportación: \n1. CSV\n2. Excel\n3. Ambos\n4. Salir al submenú\n"))

                # Exportar en formato CSV
                if export_opcion == 1:
                    export_csv_unidades()
                
                # Exportar en formato Excel
                elif export_opcion == 2:
                    export_excel_unidades()
                
                # Exportar en ambos formatos, CSV y Excel
                elif export_opcion == 3:
                    export_csv_unidades()
                    export_excel_unidades()
                
                # Salir al submenú
                elif export_opcion == 4:
                    break
                
                else:
                    print("Elige una opción válida.")
                    if cancelar():
                        break
            except ValueError:
                print("Error: Debes ingresar un número entero que sea válido.")
                if cancelar():
                    break
            except Exception as name_error:
                print(f"Ha ocurrido un error inesperado: {name_error}")
                if cancelar():
                    break
    else:
        print("No hay unidades para exportar.")
    
    # Cerrar la conexión con la base de datos
    conexion.close()

def export_excel_unidades(name_excel="Unidades.xlsx"):
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT Clave, Rodada, Color FROM UNIDAD")
            unidades = cursor.fetchall()
            
            libro = openpyxl.Workbook()
            hoja = libro.active
            hoja.title = "Unidades"
            
            # Encabezados
            hoja["A1"].value = "Clave"
            hoja["B1"].value = "Rodada"
            hoja["C1"].value = "Color"

            # Estilos en los encabezados
            hoja["A1"].font = Font(bold=True)
            hoja["B1"].font = Font(bold=True)
            hoja["C1"].font = Font(bold=True)

            for i, (Clave, Rodada, Color) in enumerate(unidades, start=2):
                hoja.cell(row=i, column=1).value = Clave
                hoja.cell(row=i, column=2).value = Rodada
                hoja.cell(row=i, column=3).value = Color
                
            ajustar_ancho_columnas(hoja)
            libro.save(name_excel)
            print(f"Datos exportados con éxito en {name_excel}")
    except sqlite3.Error as e:
        print(f"Error al exportar a Excel: {e}")
        raise

def tab_unidades_disponibles(unidades, prestamos):
    print("-----UNIDADES DISPONIBLES-----")
    print(f"{'Clave':^8}{'Rodada': <10}{'Color'}")
    print("=" * 30)
    
    # Crear un conjunto de unidades utilizadas, asegurando que sean del mismo tipo que las claves de 'unidades'
    unidades_utilizadas = {str(datos['Clave_unidad']) for datos in prestamos.values()}  # Convertimos a string
    
    # Iterar sobre las unidades y mostrar solo las que no están utilizadas
    for clave, datos in unidades.items():
        if str(clave) not in unidades_utilizadas:  # Convertimos clave a string para comparar
            print(f"{clave:^8}{datos[0]: <10}{datos[1]}")
    
    print("=" * 30)

def export_csv_unidades():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM UNIDAD")
            unidades = cursor.fetchall()

            with open("Unidades_bicicletas.csv", "w", encoding="latin1", newline="") as archivocsv_unidades:
                grabador = csv.writer(archivocsv_unidades)
                grabador.writerow(("Clave", "Rodada", "Color"))
                grabador.writerows(unidades)

            print("Datos exportados con éxito en Unidades_bicicletas.csv")

    except sqlite3.Error as e:
        print(f"Error al exportar a CSV: {e}")

## Impresión tabular que muestra los clientes
def tab_clientes():
    try:
        # Conectar a la base de datos
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            # Consultar los datos de los clientes
            cursor.execute("SELECT Clave, Apellidos, Nombres, Telefono FROM CLIENTES")
            clientes = cursor.fetchall()

            # Verificar si hay datos
            if clientes:
                # Encabezados de la tabla
                headers = ["Clave", "Apellidos", "Nombres", "Teléfono"]

                # Generar la tabla usando tabulate
                tabla = tabulate(clientes, headers, tablefmt="rounded_outline")

                # Imprimir la tabla formateada
                print(tabla)
            else:
                print("No hay clientes registrados en la base de datos.")
    
    except sqlite3.Error as e:
        print(f"Error al consultar la base de datos: {e}")


## Exporta los clientes en formato csv
def export_csv_clientes():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM CLIENTES")
            clientes = cursor.fetchall()

            with open("Clientes_bicicletas.csv", "w", encoding="latin1", newline="") as archivocsv_clientes:
                grabador = csv.writer(archivocsv_clientes)
                grabador.writerow(("Clave", "Apellidos", "Nombres", "Teléfono"))
                grabador.writerows(clientes)

            print("Datos exportados con éxito en Clientes_bicicletas.csv")

    except sqlite3.Error as e:
        print(f"Error al exportar a CSV: {e}")

## funcion para ajustar el ancho de las columnas en excel
def ajustar_ancho_columnas(hoja):
    for column_cells in hoja.columns:
        max_length = 0
        column = column_cells[0].column_letter  
        for cell in column_cells:
            try:  
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)  
        hoja.column_dimensions[column].width = adjusted_width

def export_excel_clientes(name_excel="Clientes.xlsx"):
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT Clave, Apellidos, Nombres, Telefono FROM CLIENTES")
            clientes = cursor.fetchall()

            libro = openpyxl.Workbook()
            hoja = libro.active
            hoja.title = "Clientes"

            # Encabezados
            hoja["A1"].value = "Clave"
            hoja["B1"].value = "Apellidos"
            hoja["C1"].value = "Nombres"
            hoja["D1"].value = "Teléfono"

            # Estilos en los encabezados
            hoja["A1"].font = Font(bold=True)
            hoja["B1"].font = Font(bold=True)
            hoja["C1"].font = Font(bold=True)
            hoja["D1"].font = Font(bold=True)

            # Insertar datos
            for i, (clave, apellidos, nombres, telefono) in enumerate(clientes, start=2):
                hoja.cell(row=i, column=1).value = clave
                hoja.cell(row=i, column=2).value = apellidos
                hoja.cell(row=i, column=3).value = nombres
                hoja.cell(row=i, column=4).value = telefono

            # Ajustar ancho de columnas
            ajustar_ancho_columnas(hoja)

            libro.save(name_excel)
            print(f"Datos exportados con éxito en {name_excel}")

    except sqlite3.Error as e:
        print(f"Error al exportar a Excel: {e}")

## SUBMENU LISTADO DE UNIDADES
def listado_unidades():
    while True:
        print("\n--- LISTADO DE UNIDADES ---")
        print("1. Completo")
        print("2. Por rodada")
        print("3. Por color")
        print("4. Volver al menú de listado de unidades\n")

        try:
            opcion = input("Elige una de las siguientes opciones: ")
            opcion = int(opcion)

            if opcion == 1:
                analisis_completo()
            elif opcion == 2:
                analisis_rodada()
            elif opcion == 3:
                analisis_color()
            elif opcion == 4:
                return False
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')

## SUBMENU RETRASOS
def retrasos(prestamos,clientes):
    mostrar_ruta()
    if prestamos:
        print(f"{'Folio':^8}{'Nombre del Cliente': <30}{'Fecha Préstamo': <20}{'Fecha Retorno': <20}{'Días de Retraso'}")
        print("=" * 100)

        fecha_actual = datetime.now().date()
        for folio, datos in prestamos.items():
            fecha_retorno = datetime.strptime(datos["Fecha_retorno"], "%m/%d/%Y").date()
            if fecha_actual > fecha_retorno:
                cliente_clave = int(datos['Clave_cliente'])
                cliente = clientes.get(cliente_clave)
                if cliente:
                    nombre, apellido = cliente[1], cliente[0]
                    dias_retraso = (fecha_actual - fecha_retorno).days
                    print(f"{folio:^8}{nombre + ' ' + apellido: <30}{datos['Fecha_prestamo']: <20}{datos['Fecha_retorno']: <20}{dias_retraso}")
                else:
                    print("Ha ocurrido un problema")

        print("=" * 100)
        exportar_opcion = int(input("Elige una opción para exportar: \n1. CSV\n2. Excel\n3. Ambos\n 4.No deseo exportarlo "))
        if exportar_opcion == 1:
            export_csv_retrasos(prestamos, fecha_actual)
        elif exportar_opcion == 2:
            export_excel_retrasos(prestamos, fecha_actual)
        elif exportar_opcion == 3:
            export_csv_retrasos(prestamos, fecha_actual)
            export_excel_retrasos(prestamos, fecha_actual)
        elif exportar_opcion == 4:
            return False
        else:
            print("Elige una opción válida (1, 2, 3 o 4).")
            if cancelar():
                return False
    else:
        print("No existen préstamos")
        
def export_csv_retrasos(prestamos, clientes):
    fecha_actual = datetime.now().date()
    with open("retrasos.csv", "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)
        grabador.writerow(("Folio", "Nombre del Cliente", "Fecha Préstamo", "Fecha Retorno", "Días de Retraso"))
        
        for folio, datos in prestamos.items():
            fecha_retorno = datetime.strptime(datos["Fecha_retorno"], "%m/%d/%Y").date()
            if fecha_actual > fecha_retorno:
                cliente_clave = int(datos['Clave_cliente'])
                cliente = clientes.get(cliente_clave)
                if cliente:
                    nombre, apellido = cliente[1], cliente[0]
                    dias_retraso = (fecha_actual - fecha_retorno).days
                    grabador.writerow((folio, nombre + ' ' + apellido, datos['Fecha_prestamo'], datos['Fecha_retorno'], dias_retraso))
    print("Datos exportados con éxito en retrasos.csv")

def export_excel_retrasos(prestamos, clientes, retrasos_excel="retrasos.xlsx"):
    fecha_actual = datetime.now().date()
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Retrasos"

    hoja["A1"].value = "Folio"
    hoja["B1"].value = "Nombre del Cliente"
    hoja["C1"].value = "Fecha Préstamo"
    hoja["D1"].value = "Fecha Retorno"
    hoja["E1"].value = "Días de Retraso"
    
    hoja["A1"].font = Font(bold=True)
    hoja["B1"].font = Font(bold=True)
    hoja["C1"].font = Font(bold=True)
    hoja["D1"].font = Font(bold=True)
    hoja["E1"].font = Font(bold=True)

    i = 2
    for folio, datos in prestamos.items():
        fecha_retorno = datetime.strptime(datos["Fecha_retorno"], "%m/%d/%Y").date()
        if fecha_actual > fecha_retorno:
            cliente_clave = int(datos['Clave_cliente'])
            cliente = clientes.get(cliente_clave)
            if cliente:
                nombre, apellido = cliente[1], cliente[0]
                dias_retraso = (fecha_actual - fecha_retorno).days
                hoja.cell(row=i, column=1).value = folio
                hoja.cell(row=i, column=2).value = nombre + ' ' + apellido
                hoja.cell(row=i, column=3).value = datos['Fecha_prestamo']
                hoja.cell(row=i, column=4).value = datos['Fecha_retorno']
                hoja.cell(row=i, column=5).value = dias_retraso
                i += 1

    ajustar_ancho_columnas(hoja)
    libro.save(retrasos_excel)
    print(f"Datos exportados con éxito en {retrasos_excel}")

## SUBMENU ANÁLISIS COMPLETO
def analisis_completo():
    print('lolol')

## SUBMENU ANÁLISIS POR RODADA
def analisis_rodada():
    print('lolol')

## SUBMENU ANÁLISIS POR COLOR
def analisis_color():
    print('lolol')

## SUBMENU REPORTES PRÉSTAMOS POR RETORNAR
def reporte_prestamos_por_retornar():
    # Conectar a la base de datos SQLite
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Preguntar por las fechas de inicio y final para el filtro de préstamos
    while True:
        try:
            fecha_inicial = input("\nIngresa la fecha inicial (MM/DD/AAAA): ")
            fecha_inicial = datetime.strptime(fecha_inicial, "%m/%d/%Y").date()
            break
        except ValueError:
            print("Formato de fecha incorrecto, intenta de nuevo.")
            if cancelar():
                return

    while True:
        try:
            fecha_final = input("Ingresa la fecha final (MM/DD/AAAA): ")
            fecha_final = datetime.strptime(fecha_final, "%m/%d/%Y").date()
            if fecha_final >= fecha_inicial:
                break
            else:
                print("La fecha final debe ser posterior o igual a la fecha inicial.")
                if cancelar():
                    return
        except ValueError:
            print("Formato de fecha incorrecto, intenta de nuevo.")
            if cancelar():
                return

    # Consulta SQL para obtener los préstamos que no han sido retornados en el rango de fechas
    query = '''
        SELECT folio, clave_cliente, clave_unidad, fecha_prestamo, fecha_retorno
        FROM prestamos
        WHERE retorno = 0
        AND fecha_retorno BETWEEN ? AND ?
    '''
    cursor.execute(query, (fecha_inicial, fecha_final))
    prestamos_por_retornar = cursor.fetchall()

    if prestamos_por_retornar:
        # Mostrar el reporte tabular con la librería 'tabulate'
        headers = ["Folio", "Clave del Cliente", "Clave de la Unidad", "Fecha Préstamo", "Fecha Retorno"]
        print(tabulate(prestamos_por_retornar, headers, tablefmt="rounded_outline"))

        # Opción de exportar
        while True:
            try:
                export_opcion = int(input("\nElige una opción de exportación: \n1. CSV\n2. Excel\n3. Ambos\n4. No deseo exportarlo\n"))
                if export_opcion == 1:
                    export_csv_prestamos_retornar(fecha_inicial, fecha_final)
                elif export_opcion == 2:
                    export_excel_prestamos_retornar(fecha_inicial, fecha_final)
                elif export_opcion == 3:
                    export_csv_prestamos_retornar(fecha_inicial, fecha_final)
                    export_excel_prestamos_retornar(fecha_inicial, fecha_final)
                elif export_opcion == 4:
                    break
                else:
                    print("Elige una opción válida (1, 2, 3 o 4).")
            except ValueError:
                print("Error: Debes ingresar un número entero que sea válido.")
                if cancelar():
                    return
            except Exception as name_error:
                print(f"Ha ocurrido un error inesperado: {name_error}")
                if cancelar():
                    return
    else:
        print("No se encontró ningún préstamo que coincida con los criterios.")

    # Cerrar la conexión con la base de datos
    conexion.close()

## Exporta los préstamos por retornar en formato excel
def export_excel_prestamos_retornar(fecha_prestamo, fecha_de_retorno, name_excel="Prestamos_por_retornar.xlsx"):
    # Conectar a la base de datos
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta SQL para obtener los préstamos no retornados en el rango de fechas
    query = '''
        SELECT folio, clave_unidad, clave_cliente, fecha_prestamo, fecha_retorno
        FROM prestamos
        WHERE retorno = 0
        AND fecha_retorno BETWEEN ? AND ?
    '''
    cursor.execute(query, (fecha_prestamo, fecha_de_retorno))
    prestamos = cursor.fetchall()

    # Crear un nuevo archivo de Excel
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Préstamos"

    # Encabezados
    hoja["A1"].value = "Folio"
    hoja["B1"].value = "Clave de la unidad"
    hoja["C1"].value = "Clave del cliente"
    hoja["D1"].value = "Fecha préstamo"
    hoja["E1"].value = "Fecha de retorno"

    # Negritas para los encabezados
    for col in ["A1", "B1", "C1", "D1", "E1"]:
        hoja[col].font = Font(bold=True)

    # Insertar los datos obtenidos de la base de datos
    for i, prestamo in enumerate(prestamos, start=2):
        hoja.cell(row=i, column=1).value = prestamo[0]  # Folio
        hoja.cell(row=i, column=2).value = prestamo[1]  # Clave de la unidad
        hoja.cell(row=i, column=3).value = prestamo[2]  # Clave del cliente
        hoja.cell(row=i, column=4).value = prestamo[3]  # Fecha préstamo
        hoja.cell(row=i, column=5).value = prestamo[4]  # Fecha de retorno

    # Ajustar ancho de columnas (opcional)
    ajustar_ancho_columnas(hoja)

    # Guardar el archivo Excel
    libro.save(name_excel)
    print(f"Datos exportados con éxito en {name_excel}")

    # Cerrar la conexión a la base de datos
    conexion.close()

## Exporta los préstamos por retornar en formato csv
def export_csv_prestamos_retornar(fecha_prestamo, fecha_de_retorno, nombre_csv="Prestamos_por_retornar.csv"):
    # Conectar a la base de datos
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta SQL para obtener los préstamos no retornados en el rango de fechas
    query = '''
        SELECT folio, clave_unidad, clave_cliente, fecha_prestamo, fecha_retorno
        FROM prestamos
        WHERE retorno = 0
        AND fecha_retorno BETWEEN ? AND ?
    '''
    cursor.execute(query, (fecha_prestamo, fecha_de_retorno))
    prestamos = cursor.fetchall()

    # Abrir archivo CSV y escribir los datos
    with open(nombre_csv, "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)

        # Encabezados
        grabador.writerow(("Folio", "Clave de la unidad", "Clave del cliente", "Fecha préstamo", "Fecha de retorno"))

        # Verificar si hay préstamos para exportar
        if prestamos:
            grabador.writerows(prestamos)
            print(f"Datos exportados con éxito en {nombre_csv}")
        else:
            print("No hay préstamos que coincidan con los criterios especificados.")

    # Cerrar la conexión a la base de datos
    conexion.close()

## SUBMENU PRÉSTAMOS POR PERIODO
def prestamos_por_periodo(prestamos):
    if prestamos:
            while True:
                mostrar_ruta()
                try:
                    fecha_inicial = input("\nIngresa la fecha inicial del periodo (MM/DD/AAAA): ")
                    fecha_inicial = datetime.strptime(fecha_inicial, "%m/%d/%Y").date()
                    break
                except ValueError:
                    print("Formato de fecha incorrecto, intenta de nuevo.")
                    if cancelar():
                        break

            while True:
                try:
                    fecha_final = input("Ingresa la fecha final del periodo (MM/DD/AAAA): ")
                    fecha_final = datetime.strptime(fecha_final, "%m/%d/%Y").date()
                    if fecha_final >= fecha_inicial:
                        break
                    else:
                        print("La fecha final debe ser posterior o igual a la fecha inicial.")
                        if cancelar():
                            break
                except ValueError:
                    print("Formato de fecha incorrecto, intenta de nuevo.")
                    if cancelar():
                        break

            print(f"{'Folio':^8}{'Clave del Cliente': <20}{'Clave de la Unidad': <20}{'Fecha Préstamo': <20}{'Fecha Retorno'}")
            print("=" * 80)

            for folio, datos in prestamos.items():
                fecha_prestamo = datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date()
                if fecha_inicial <= fecha_prestamo <= fecha_final:
                    print(f"{folio:^8}{datos['Clave_cliente']: <20}{datos['Clave_unidad']: <20}{datos['Fecha_prestamo']: <20}{datos['Fecha_retorno']}")

            print("=" * 80)

            export_opcion = int(input("Elige una opción de exportación: \n1. CSV\n2. Excel\n3. Ambos\n"))
            if export_opcion == 1:
                export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            elif export_opcion == 2:
                export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            elif export_opcion == 3:
                export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
                export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            elif export_opcion == 4:
                return False
            else:
                print("Elige una opción válida (1, 2, 3 o 4).")
                if cancelar():
                    return False
    else: 
        print("No hay préstamos para realizar un reporte")

## Exporta los préstamos por periodo en formato excel
def export_excel_prestamos_por_periodo(prestamos, fecha_prestamo, fecha_de_retorno, name_excel="Prestamos_por_periodo.xlsx"):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "préstamos"

    hoja["A1"].value = "Folio"
    hoja["B1"].value = "Clave de la unidad"
    hoja["C1"].value = "Clave del cliente"
    hoja["D1"].value = "Fecha préstamo"
    hoja["E1"].value = "Fecha de retorno"
    
    hoja["A1"].font = Font(bold=True)
    hoja["B1"].font = Font(bold=True)
    hoja["C1"].font = Font(bold=True)
    hoja["D1"].font = Font(bold=True)
    hoja["E1"].font = Font(bold=True)

    # Poner negrita en los encabezados (fila 1)
    hoja["A1"].font = Font(bold=True)
    hoja["B1"].font = Font(bold=True)
    hoja["C1"].font = Font(bold=True)
    hoja["D1"].font = Font(bold=True)
    hoja["E1"].font = Font(bold=True)

    i = 2  # Iniciar en la fila 2 porque la fila 1 son los encabezados
    for folio, datos in prestamos.items():
        fecha_prestamo = datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date()
        if fecha_prestamo <= fecha_prestamo <= fecha_de_retorno:
            # Asignar los valores a las celdas
            hoja.cell(row=i, column=1).value = folio
            hoja.cell(row=i, column=2).value = datos["Clave_unidad"]
            hoja.cell(row=i, column=3).value = datos["Clave_cliente"]
            hoja.cell(row=i, column=4).value = datos["Fecha_prestamo"]
            hoja.cell(row=i, column=5).value = datos["Fecha_retorno"]

            # Calcular el ancho según la longitud del contenido de cada fila
            hoja.column_dimensions["A"].width = max(hoja.column_dimensions["A"].width or 0, len(str(folio)))
            hoja.column_dimensions["B"].width = max(hoja.column_dimensions["B"].width or 0, len(str(datos["Clave_unidad"])))
            hoja.column_dimensions["C"].width = max(hoja.column_dimensions["C"].width or 0, len(str(datos["Clave_cliente"])))
            hoja.column_dimensions["D"].width = max(hoja.column_dimensions["D"].width or 0, len(str(datos["Fecha_prestamo"])))
            hoja.column_dimensions["E"].width = max(hoja.column_dimensions["E"].width or 0, len(str(datos["Fecha_retorno"])))

            i += 1  # Incrementa la fila
    ajustar_ancho_columnas(hoja)
    # Guarda el archivo Excel
    libro.save(name_excel)
    
    print(f"Datos exportados con éxito en {name_excel}")

## Exporta los préstamos por periodo en formato csv
def export_csv_prestamos_por_periodo(prestamos, fecha_prestamo, fecha_de_retorno, nombre_csv="Prestamos_por_periodo.csv"):
    with open(nombre_csv, "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)
        grabador.writerow(("Folio", "Clave de la unidad", "Clave del cliente", "Fecha préstamo", "Fecha de retorno"))

        prestamos_filtrados = [
            (folio, datos["Clave_unidad"], datos["Clave_cliente"], datos["Fecha_prestamo"], datos["Fecha_retorno"])
            for folio, datos in prestamos.items()
            if fecha_prestamo <= datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date() <= fecha_de_retorno
        ]

        if prestamos_filtrados:
            grabador.writerows(prestamos_filtrados)
            print(f"Datos exportados con éxito en {nombre_csv}")
        else:
            print("No hay préstamos que coincidan con los criterios especificados.")
            if cancelar():
                return False

##Submenú analísis
def submenu_analisis():
    mostrar_ruta()
    while True:
        print("\n--- SUBMENÚ ANÁLISIS ---")
        print("1. Duración de los préstamos.")
        print("2. Ranking de clientes.")
        print("3. Preferencias de rentas.")
        print("4. Volver al menú\n")

        try:
            opcion = input("Elige una de las siguientes opciones: ")
            opcion = int(opcion)

            if opcion == 1:
                estadisticas_prestamos()
            elif opcion == 2:
                ranking_clientes()
            elif opcion == 3:
                preferencias_rentas()
            elif opcion == 4:
                return False
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')

#EV3
#Datos estadisitcos de la duracion de los prestamos
def estadisticas_prestamos():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            query = "SELECT Dias_Prestamo FROM PRESTAMO"
            df = pd.read_sql_query(query, conn)

            media = df['Dias_Prestamo'].mean()
            mediana = df['Dias_Prestamo'].median()
            moda = df['Dias_Prestamo'].mode().tolist()  # La moda puede tener múltiples valores
            minimo = df['Dias_Prestamo'].min()
            maximo = df['Dias_Prestamo'].max()
            desviacion_estandar = df['Dias_Prestamo'].std()
            cuartiles = df['Dias_Prestamo'].quantile([0.25, 0.5, 0.75])

            print(f"Media: {media}")
            print(f"Mediana: {mediana}")
            print(f"Moda: {moda}")
            print(f"Mínimo: {minimo}")
            print(f"Máximo: {maximo}")
            print(f"Desviación estándar: {desviacion_estandar}")
            print("Cuartiles:")
            print(cuartiles)

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")




## SUBMENÚ DURACIÓN DE LOS PRÉSTAMOS
def duracion_prestamos(prestamos):
    dias_prestamo = [prestamo['Cantidad_dias'] for prestamo in prestamos.values()]

    if len(dias_prestamo) == 0:
        print("No hay registros de préstamos para calcular estadísticas.")
        return

    df = pd.DataFrame(dias_prestamo, columns=['Días de préstamo'])

    media = df['Días de préstamo'].mean()
    mediana = df['Días de préstamo'].median()
    moda = df['Días de préstamo'].mode().tolist()  # Convertir a lista en caso de múltiples modas
    minimo = df['Días de préstamo'].min()
    maximo = df['Días de préstamo'].max()
    desviacion_estandar = df['Días de préstamo'].std()
    cuartiles = np.percentile(df['Días de préstamo'], [25, 50, 75])

    reporte = {
        "Media": media,
        "Mediana": mediana,
        "Moda": moda,
        "Mínimo": minimo,
        "Máximo": maximo,
        "Desviación estándar": desviacion_estandar,
        "Cuartiles (25%, 50%, 75%)": cuartiles
    }

    for clave, valor in reporte.items():
        print(f"{clave}: {valor}")


#cargar rentas
def cargar_rentas_csv():
    rentas = {}
    try:
        with open('rentas.csv', mode='r') as file:
            reader = csv.reader(file)
            next(reader)  # Saltar la fila de encabezados
            for row in reader:
                clave_cliente = int(row[0])
                cantidad_rentas = int(row[1])
                rentas[clave_cliente] = cantidad_rentas
        print("Rentas cargadas correctamente.")
    except FileNotFoundError:
        print("Archivo de rentas no encontrado. Se inicializará una lista vacía.")
    except Exception as e:
        print(f"Error al cargar rentas: {e}")
    return rentas
    
#guardar rentas
def guardar_rentas_csv(rentas):
    try:
        df = pd.DataFrame(list(rentas.items()), columns=['Clave_cliente', 'Cantidad_rentas'])
        df.to_csv('rentas.csv', index=False)
    except Exception as e:
        print(f"Ocurrió un error al guardar las rentas: {e}")
    
#EV3
#Ranking de clientes, cantidad de prestamos(rentas) que realizo cada cliente.
def ranking_clientes():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            query = """
                SELECT 
                    CLIENTES.Clave, 
                    CLIENTES.Nombres || ' ' || CLIENTES.Apellidos AS Nombre_Completo,
                    CLIENTES.Telefono,
                    COUNT(PRESTAMO.Clave_Cliente) AS Cantidad_Rentas
                FROM CLIENTES
                JOIN PRESTAMO ON CLIENTES.Clave = PRESTAMO.Clave_Cliente
                GROUP BY CLIENTES.Clave
                ORDER BY Cantidad_Rentas DESC;
            """
            
            df = pd.read_sql_query(query, conn)

            print("\n--- Ranking de Clientes ---")
            headers = ["Clave", "Nombre Completo", "Teléfono", "Cantidad de Rentas"]
            print(tabulate(df, headers=headers, tablefmt="rounded_outline", showindex=False))

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")



## SUBMENÚ RANKING CLIENTES
# Función para generar el ranking de clientes
def ranking_clientes_old(prestamos, clientes, rentas):
    ranking_data = {
        'Cantidad_rentas': [],
        'Clave_cliente': [],
        'Nombre_completo': [],
        'Teléfono': []
    }

    # Contar las rentas acumuladas por cada cliente
    for clave_cliente, cantidad_rentas in rentas.items():
        if clave_cliente in clientes:
            cliente = clientes[clave_cliente]
            apellidos, nombre, telefono = cliente  # Desempaquetar la tupla

            ranking_data['Cantidad_rentas'].append(cantidad_rentas)
            ranking_data['Clave_cliente'].append(clave_cliente)
            ranking_data['Nombre_completo'].append(f"{nombre} {apellidos}")
            ranking_data['Teléfono'].append(telefono)

    # df para ordenar los resultados
    df_ranking = pd.DataFrame(ranking_data)
    df_ranking.sort_values(by='Cantidad_rentas', ascending=False, inplace=True)

    # ranking con formato de tabla
    print("\n--- RANKING DE CLIENTES ---")
    print(f"{'Posición':<10} {'Clave Cliente':<15} {'Nombre Completo':<30} {'Teléfono':<15} {'Cantidad de Rentas':<20}")
    print("=" * 90)
    for i, row in enumerate(df_ranking.itertuples(index=False), 1):
        print(f"{i:<10} {row.Clave_cliente:<15} {row.Nombre_completo:<30} {row.Teléfono:<15} {row.Cantidad_rentas:<20}")
    
    guardar_rentas_csv(rentas)
    guardar_ranking_csv(df_ranking)


def guardar_ranking_csv(df_ranking):
    df_ranking.to_csv("Ranking_clientes.csv", index=False, encoding="latin1")
    print("Ranking de clientes exportado exitosamente en 'Ranking_clientes.csv'.")



## SUBMENÚ PREFERENCIAS RENTAS
def preferencias_rentas():
    print("Elige el reporte que deseas generar:")
    print("1. Cantidad de préstamos por rodada")
    print("2. Cantidad de préstamos por color")
    print("3. Por días de la semana")
    
    while True:
        opcion_pref = input("Ingresa una opción (1 o 2): ")
        if opcion_pref.isdigit():
            opcion_pref = int(opcion_pref)
            if opcion_pref == 1:
                rodada_tab_count()
                break
            elif opcion_pref == 2:
                colores_tab_count()
                break
            elif opcion_pref == 3:
                prestamos_por_dia_semana()
                break
            else:
                print("Opción inválida. Debes ingresar 1 o 2.")
        else:
            print("Entrada inválida. Por favor ingresa un número (1 o 2).")

def reporte_prestamos_por_rodada(conteo_rodadas):
    # Ordenar las rodadas por la cantidad de préstamos en orden descendente
    datos_ordenados = sorted(conteo_rodadas.items(), key=lambda x: x[1], reverse=True)

    # Imprimir el reporte en formato tabular
    print("\n--- REPORTE DE PRÉSTAMOS POR RODADA ---")
    print("{:<10} {:<20}".format("Rodada", "Cantidad de Préstamos"))
    print("-" * 30)
    for rodada, cantidad in datos_ordenados:
        print("{:<10} {:<20}".format(rodada, cantidad))
    exportar_conteo_rodada_excel(conteo_rodadas)




def export_conteo_rodada(conteo_rodadas, nombre_archivo="Conteo_Rodadas.csv"):
    # Exportar el conteo de rodadas a un archivo CSV
    with open(nombre_archivo, "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)
        grabador.writerow(("Rodada", "Cantidad de Préstamos"))
        for rodada, cantidad in conteo_rodadas.items():
            grabador.writerow((rodada, cantidad))
    print(f"Conteo de rodadas exportado exitosamente en '{nombre_archivo}'")


def cargar_conteo_rodadas(nombre_archivo="Conteo_Rodadas.csv"):
    conteo_rodadas = {}
    try:
        # Abrir el archivo CSV para leer el conteo de rodadas
        with open(nombre_archivo, "r", encoding="latin1", newline="") as archivo_csv:
            lector = csv.reader(archivo_csv)
            # Saltar la fila de encabezado
            next(lector)
            # Leer cada fila y actualizar el diccionario conteo_rodadas
            for fila in lector:
                if len(fila) == 2:  # Asegurar que la fila tiene exactamente 2 columnas
                    rodada, cantidad = fila
                    conteo_rodadas[int(rodada)] = int(cantidad)
    except FileNotFoundError:
        print(f"El archivo '{nombre_archivo}' no existe. Asegúrate de que el archivo se haya exportado previamente.")
    
    return conteo_rodadas
    

def exportar_conteo_rodada_excel(conteo_rodadas, nombre_archivo="Conteo_Rodadas.xlsx"):
    # Convertir el conteo de rodadas a un DataFrame de pandas
    df = pd.DataFrame(list(conteo_rodadas.items()), columns=["Rodada", "Cantidad de Préstamos"])
    
    # Exportar a un archivo de Excel
    df.to_excel(nombre_archivo, index=False)
    print(f"Conteo de rodadas exportado exitosamente en '{nombre_archivo}'")


# 1. Función para exportar los colores a un archivo CSV
def exportar_colores_csv(unidades, nombre_archivo="Colores.csv"):
    conteo_colores = {}
    for clave, (rodada, color) in unidades.items():
        if color in conteo_colores:
            conteo_colores[color] += 1
        else:
            conteo_colores[color] = 1
    
    try:
        with open(nombre_archivo, mode='w', newline='', encoding='latin1') as archivo_csv:
            escritor = csv.writer(archivo_csv)
            escritor.writerow(["Color", "Cantidad"])  # Encabezados
            for color, cantidad in conteo_colores.items():
                escritor.writerow([color, cantidad])
        print(f"Colores exportados exitosamente a {nombre_archivo}.")
    except Exception as e:
        print(f"Error al exportar colores: {e}")

# 2. Función para cargar los colores desde un archivo CSV
def cargar_colores_csv(nombre_archivo="Colores.csv"):
    unidades = {}
    try:
        with open(nombre_archivo, mode='r', newline='', encoding='latin1') as archivo_csv:
            lector = csv.reader(archivo_csv)
            next(lector)  # Saltar encabezado
            for fila in lector:
                if len(fila) == 2:
                    color, cantidad = fila
                    unidades[color] = int(cantidad)
        print(f"Colores cargados exitosamente desde {nombre_archivo}.")
    except FileNotFoundError:
        print(f"El archivo '{nombre_archivo}' no existe.")
    except Exception as e:
        print(f"Error al cargar colores: {e}")
    return unidades

# 3. Función para generar un reporte tabular de los colores
def reporte_colores_tabular_ordenado(unidades):
    conteo_colores = {}
    for clave, (rodada, color) in unidades.items():
        if color in conteo_colores:
            conteo_colores[color] += 1
        else:
            conteo_colores[color] = 1

    # Ordenar los colores alfabéticamente
    colores_ordenados = sorted(conteo_colores.items())

    # Imprimir reporte en formato tabular
    print("\n--- REPORTE DE COLORES (ORDENADO) ---")
    print(f"{'Color':<15} {'Cantidad':<10}")
    print("-" * 25)
    for color, cantidad in colores_ordenados:
        print(f"{color:<15} {cantidad:<10}")
    exportar_colores_excel(unidades)
    exportar_colores_csv(unidades)    
  
def exportar_colores_excel(unidades, nombre_archivo="Colores.xlsx"):
    try:
        conteo_colores = {}
        for clave, (rodada, color) in unidades.items():
            if color in conteo_colores:
                conteo_colores[color] += 1
            else:
                conteo_colores[color] = 1

        # Convertir los datos a un DataFrame de pandas
        df = pd.DataFrame(list(conteo_colores.items()), columns=["Color", "Cantidad"])
        
        # Exportar a un archivo de Excel
        df.to_excel(nombre_archivo, index=False)
        print(f"Colores exportados exitosamente a {nombre_archivo}.")
    except Exception as e:
        print(f"Error al exportar colores a Excel: {e}")
   
#EV3
#Reporte tabular de preferencias del cliente, cantidad de prestamos por color
#version sin grafica   
def colores_tab_count1():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            cursor.execute("SELECT Color, COUNT(*) as Cantidad FROM UNIDAD GROUP BY Color ORDER BY Cantidad DESC")
            colores = cursor.fetchall()

            print("\n--- Reporte de Colores de Unidades ---")
            headers_colores = ["Color", "Cantidad"]
            print(tabulate(colores, headers=headers_colores, tablefmt="rounded_outline"))


    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")

#Version final con grafica que corresponde a sus colores
#Grafica de pastel, colores
def colores_tab_count():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            # Consulta SQL para obtener la cantidad de préstamos por color
            cursor.execute("SELECT Color, COUNT(*) as Cantidad FROM UNIDAD GROUP BY Color ORDER BY Cantidad DESC")
            colores = cursor.fetchall()

            # Convertir los resultados a un DataFrame
            df_colores = pd.DataFrame(colores, columns=["Color", "Cantidad"])

            print("\n--- Reporte de Cantidad de Préstamos por Color ---")
            print(tabulate(df_colores, headers="keys", tablefmt="rounded_outline", showindex=False))

            # Colores personalizados para la gráfica de pastel
            color_map = {
                "ROJO": "red",
                "AZUL": "blue",
                "AMARILLO": "yellow",
                "VERDE": "green",
                "ROSA": "pink"
            }
            # Asignar los colores según el nombre
            colores_grafica = [color_map[color] for color in df_colores['Color']]

            # Gráfica de pastel
            plt.figure(figsize=(8, 6))
            plt.pie(df_colores['Cantidad'], labels=df_colores['Color'], autopct='%1.1f%%', startangle=140, colors=colores_grafica)
            plt.title('Cantidad de préstamos por color y su proporción')
            plt.axis('equal')  # Para que el pie sea un círculo
            plt.show()

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")




#EV3
#Reporte tabular de preferencias del cliente, cantidad de prestamos por rodada
def rodada_tab_count():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            cursor.execute("SELECT Rodada, COUNT(*) as Cantidad FROM UNIDAD GROUP BY Rodada ORDER BY Cantidad DESC")
            colores = cursor.fetchall()

            print("\n--- Reporte de Rodadas de Unidades ---")
            headers_colores = ["Rodada", "Cantidad"]
            print(tabulate(colores, headers=headers_colores, tablefmt="rounded_outline"))

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")


#EV3
#Reporte tabular por dia de la semana y una grafica de barras
def prestamos_por_dia_semana():
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()

            # Consulta SQL para obtener la cantidad de préstamos por día de la semana
            query = """
                SELECT 
                    strftime('%w', Fecha_Prestamo) AS Dia_Semana, 
                    COUNT(*) AS Cantidad 
                FROM 
                    PRESTAMO 
                GROUP BY 
                    Dia_Semana 
                ORDER BY 
                    Dia_Semana;
            """
            cursor.execute(query)
            dias = cursor.fetchall()

            # Convertir los resultados a un DataFrame
            df_dias = pd.DataFrame(dias, columns=["Dia_Semana", "Cantidad"])

            # Mapeo de números de días a nombres de días
            dia_map = {
                '0': 'Domingo',
                '1': 'Lunes',
                '2': 'Martes',
                '3': 'Miércoles',
                '4': 'Jueves',
                '5': 'Viernes',
                '6': 'Sábado'
            }

            # Asignar nombres a los días
            df_dias['Dia_Semana'] = df_dias['Dia_Semana'].map(dia_map)

            # Reordenar días de la semana
            df_dias = df_dias.set_index("Dia_Semana").reindex(
                ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'],
                fill_value=0
            ).reset_index()

            print("\n--- Reporte de Cantidad de Préstamos por Día de la Semana ---")
            print(tabulate(df_dias, headers="keys", tablefmt="rounded_outline", showindex=False))

            # Gráfica de barras
            plt.figure(figsize=(10, 6))
            plt.bar(df_dias['Dia_Semana'], df_dias['Cantidad'], color='skyblue')
            plt.xlabel('Día de la Semana')
            plt.ylabel('Cantidad de Préstamos')
            plt.title('Cantidad de Préstamos por Día de la Semana')
            plt.xticks(rotation=45)
            plt.grid(axis='y')
            plt.show()

    except sqlite3.Error as e:
        print(f"Error de base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")


# Inicio del programa
clientes = import_clientes()
unidades = import_unidades()
print(unidades)
prestamos = import_prestamos()
print(prestamos)
conteo_colores = cargar_colores_csv()
conteo_rodadas = cargar_conteo_rodadas()
rentas = cargar_rentas_csv()
print("===== BIENVENIDO A NUESTRA RENTA DE BICICLETAS =====")
menu_principal()