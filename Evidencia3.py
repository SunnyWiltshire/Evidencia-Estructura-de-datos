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
import datetime
from datetime import datetime
import sqlite3
import datetime 
import pandas as pd
from tabulate import tabulate
from openpyxl import load_workbook

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
            listado_unidades_reporte()
            ruta.pop()
        elif reporte_opcion == 3:
            ruta.append('Retrasos')
            reporte_retrasos()
            ruta.pop()
        elif reporte_opcion == 4:
            ruta.append('Prestamos por retornar')
            reporte_prestamos_por_retornar()
            ruta.pop()
        elif reporte_opcion == 5:
            ruta.append('Prestamos por periodo')
            prestamos_por_periodo()
            ruta.pop()
        elif reporte_opcion == 6:
            return False    
        else:
            print("Ingresa una opción válida")
    except Exception as error_name:
        print(f"Ha ocurrido un error: {error_name}")
        if cancelar():
            break


#TEST UPDATE
def actualizar_retorno():
    try:
        with sqlite3.connect("RentaBicicletas.db") as conn:
            cursor = conn.cursor()

            # Pedir al usuario la clave del préstamo a actualizar
            clave_prestamo = input("Ingresa la clave del préstamo a actualizar: ")

            # Verificar si el préstamo existe
            cursor.execute("SELECT * FROM PRESTAMO WHERE Folio = ?", (clave_prestamo,))
            if not cursor.fetchone():
                print("No existe un préstamo con esa clave.")
                return

            # Preguntar al usuario si el préstamo ha sido retornado
            retorno = input("¿El préstamo ha sido retornado? (S/N): ").strip().upper()
            if retorno == 'S':
                cursor.execute("UPDATE PRESTAMO SET Retorno = ? WHERE Folio = ?", (True, clave_prestamo))
                print("Préstamo marcado como retornado.")
            elif retorno == 'N':
                cursor.execute("UPDATE PRESTAMO SET Retorno = ? WHERE Folio = ?", (False, clave_prestamo))
                print("Préstamo marcado como no retornado.")
            else:
                print("Opción no válida. Por favor ingresa 'S' o 'N'.")

            conn.commit()  # Confirmar la transacción

    except sqlite3.Error as e:
        print(f"Error en la base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")


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
def listado_rodada(): 
    while True:
        opcion_rodada = input("""Escoge una de las rodadas disponibles: 20,26,29: """)
        
        if int(opcion_rodada) in [20,26,29]: 
            try:
                with sqlite3.connect("RentaBicicletas.db") as conn:
                    mi_cursor = conn.cursor()
                    criterios = {"RODADA": opcion_rodada}
                    mi_cursor.execute("""
                        SELECT * FROM UNIDAD WHERE Rodada = :RODADA 
                        ORDER BY Clave;
                    """, criterios)
                    unidades = mi_cursor.fetchall()

                    if unidades:
                        exportar_unidades_rodada(opcion_rodada)
                        break
                    else:
                        print("No se encontraron unidades con la rodada indicada.")
            except Error as e:
                print(e)
            except:
                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
        else: 
            print("Opción invalida")

def listado_color():  
    while True:
        opcion_color = input("""Escoge uno de los siguientes colores 
                              ROJO
                              AZUL
                              AMARILLO
                              VERDE
                              ROSA: """).upper()
        
        if opcion_color in ["ROJO", "AZUL", "AMARILLO", "VERDE", "ROSA"]: 
            try:
                with sqlite3.connect("RentaBicicletas.db") as conn:
                    mi_cursor = conn.cursor()
                    criterios = {"COLOR": opcion_color}
                    mi_cursor.execute("""
                        SELECT * FROM UNIDAD WHERE Color = :COLOR 
                        ORDER BY Clave;
                    """, criterios)
                    unidades = mi_cursor.fetchall()

                    if unidades:
                        exportar_unidades_color(opcion_color)
                        break
                    else:
                        print("No se encontraron unidades con el color indicado.")
            except Error as e:
                print(e)
            except:
                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
        else: 
            print("Opción invalida")

def export_csv_unidades_color(color): 
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM UNIDAD WHERE Color = :COLOR ORDER BY Clave" ,{"COLOR": color})
            unidades = cursor.fetchall()

            with open("Unidades_bicicletas_color.csv", "w", encoding="latin1", newline="") as archivocsv_unidades:
                grabador = csv.writer(archivocsv_unidades)
                grabador.writerow(("Clave", "Rodada", "Color"))
                grabador.writerows(unidades)

            print("Datos exportados con éxito en Unidades_bicicletas_rodada.csv")

    except sqlite3.Error as e:
        print(f"Error al exportar a CSV: {e}")

def export_excel_unidades_color(color, name_excel="Unidades_color.xlsx"): 
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT Clave, Rodada, Color FROM UNIDAD WHERE Color = :COLOR ORDER BY Clave" ,{"COLOR": color})
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

def export_csv_unidades_rodada(rodada):
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM UNIDAD WHERE Rodada = :RODADA ORDER BY Clave" ,{"RODADA": rodada})
            unidades = cursor.fetchall()

            with open("Unidades_bicicletas_rodada.csv", "w", encoding="latin1", newline="") as archivocsv_unidades:
                grabador = csv.writer(archivocsv_unidades)
                grabador.writerow(("Clave", "Rodada", "Color"))
                grabador.writerows(unidades)

            print("Datos exportados con éxito en Unidades_bicicletas_rodada.csv")

    except sqlite3.Error as e:
        print(f"Error al exportar a CSV: {e}")

def exportar_unidades_color(color):
    # Conectar a la base de datos SQLite
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta para obtener las unidades
    cursor.execute("SELECT * FROM UNIDAD WHERE Color = :COLOR ORDER BY Clave", {"COLOR": color})
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
                    export_csv_unidades_color(color)
                
                # Exportar en formato Excel
                elif export_opcion == 2:
                    export_excel_unidades_color(color)
                
                # Exportar en ambos formatos, CSV y Excel
                elif export_opcion == 3:
                    export_csv_unidades_color(color)
                    export_excel_unidades_color(color)
                
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

def exportar_unidades_rodada(rodada):
    # Conectar a la base de datos SQLite
    conexion = sqlite3.connect('RentaBicicletas.db')
    cursor = conexion.cursor()

    # Consulta para obtener las unidades
    cursor.execute("SELECT * FROM UNIDAD WHERE Rodada = :RODADA ORDER BY Clave", {"RODADA": rodada})
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
                    export_csv_unidades_rodada(rodada)
                
                # Exportar en formato Excel
                elif export_opcion == 2:
                    export_excel_unidades_rodada(rodada)
                
                # Exportar en ambos formatos, CSV y Excel
                elif export_opcion == 3:
                    export_csv_unidades_rodada(rodada)
                    export_excel_unidades_rodada(rodada)
                
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

def export_excel_unidades_rodada(rodada, name_excel="Unidades_rodada.xlsx"):
    try:
        with sqlite3.connect('RentaBicicletas.db') as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT Clave, Rodada, Color FROM UNIDAD WHERE Rodada = :RODADA ORDER BY Clave" ,{"RODADA": int(rodada)})
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

def listado_unidades_reporte():
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
                exportar_unidades()
            elif opcion == 2:
                listado_rodada()  
            elif opcion == 3:
                listado_color()
            elif opcion == 4:
                return False
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')

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


## SUBMENU REPORTES PRÉSTAMOS POR RETORNAR
def reporte_prestamos_por_retornar():
    try:
        # Solicitar fechas de inicio y fin del período
        fecha_inicio = input("Ingrese la fecha de inicio del período (mm/dd/aaaa): ")
        fecha_fin = input("Ingrese la fecha de fin del período (mm/dd/aaaa): ")

        # Convertir las fechas a formato datetime
        fecha_inicio = datetime.datetime.strptime(fecha_inicio, "%m/%d/%Y").date()
        fecha_fin = datetime.datetime.strptime(fecha_fin, "%m/%d/%Y").date()

        with sqlite3.connect("RentaBicicletas.db", 
                             detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            
            # Consulta SQL para obtener préstamos pendientes de retorno en el período
            mi_cursor.execute("""
                SELECT 
                    PRESTAMO.Clave_Unidad, 
                    UNIDAD.Rodada, 
                    PRESTAMO.Fecha_Prestamo, 
                    CLIENTES.Nombres || ' ' || CLIENTES.Apellidos AS Nombre_Completo, 
                    CLIENTES.Telefono
                FROM 
                    PRESTAMO
                JOIN 
                    CLIENTES ON PRESTAMO.Clave_Cliente = CLIENTES.Clave
                JOIN 
                    UNIDAD ON PRESTAMO.Clave_Unidad = UNIDAD.Clave
                WHERE 
                    PRESTAMO.Retorno = False
                    AND DATE(PRESTAMO.Fecha_Prestamo) BETWEEN ? AND ?
            """, (fecha_inicio, fecha_fin))

            registros = mi_cursor.fetchall()

            # Verificar si los registros no están vacíos
            if registros:
                # Formatear las fechas en "mes/día/año"
                for i in range(len(registros)):
                    # Convertir la cadena de fecha a un objeto datetime
                    fecha_prestamo = registros[i][2]  # Este es un string
                    if isinstance(fecha_prestamo, str):
                        fecha_prestamo = datetime.datetime.strptime(fecha_prestamo, "%Y-%m-%d")  # Ajusta el formato según sea necesario
                    fecha_formateada = fecha_prestamo.strftime("%m/%d/%Y")  # Convertir la fecha al formato deseado
                    registros[i] = (registros[i][0], registros[i][1], fecha_formateada, registros[i][3], registros[i][4])

                # Mostrar resultados en forma tabular
                headers = ["Clave de Unidad", "Rodada", "Fecha de Préstamo", "Nombre Completo", "Teléfono"]
                print("\n--- Reporte de Préstamos por Retornar ---")
                print(tabulate(registros, headers=headers, tablefmt="rounded_outline"))

                # Preguntar si se desea exportar
                export_option = input("\n¿Desea exportar el reporte? (1: CSV, 2: Excel, 3: Ambas, 4: No exportar): ")
                
                # Crear DataFrame para exportar
                df = pd.DataFrame(registros, columns=headers)
                
                # Exportar según la elección
                if export_option == "1":
                    df.to_csv("prestamos_por_retornar.csv", index=False)
                    print("Reporte exportado a 'prestamos_por_retornar.csv'.")
                elif export_option == "2":
                    # Exportar a Excel
                    df.to_excel("prestamos_por_retornar.xlsx", index=False)
                    print("Reporte exportado a 'prestamos_por_retornar.xlsx'.")

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("prestamos_por_retornar.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("prestamos_por_retornar.xlsx")
                    
                elif export_option == "3":
                    df.to_csv("prestamos_por_retornar.csv", index=False)
                    df.to_excel("prestamos_por_retornar.xlsx", index=False)

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("prestamos_por_retornar.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("prestamos_por_retornar.xlsx")

                    print("Reporte exportado a 'prestamos_por_retornar.csv' y 'prestamos_por_retornar.xlsx'.")
                elif export_option == "4":
                    print("No se exportó el reporte.")
                else:
                    print("Opción no válida. No se exportó el reporte.")
            else:
                print("\nNo hay préstamos pendientes de retorno en el período indicado.")

    except sqlite3.Error as e:
        print(f"Error en la base de datos: {e}")
    except ValueError:
        print("Formato de fecha incorrecto. Por favor, utiliza el formato mm/dd/aaaa.")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        print("Se ha cerrado la conexión")


#EV3
#PRESTAMOS POR PERIODO TAB
def prestamos_por_periodo():
    try:
        # Solicitar fechas de inicio y fin del período
        fecha_inicio = input("Ingrese la fecha de inicio del período (mm/dd/aaaa): ")
        fecha_fin = input("Ingrese la fecha de fin del período (mm/dd/aaaa): ")

        # Convertir las fechas a formato datetime
        fecha_inicio = datetime.datetime.strptime(fecha_inicio, "%m/%d/%Y").date()
        fecha_fin = datetime.datetime.strptime(fecha_fin, "%m/%d/%Y").date()

        with sqlite3.connect("RentaBicicletas.db", 
                             detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()

            # Consulta SQL para obtener datos en el período
            mi_cursor.execute("""
                SELECT 
                    PRESTAMO.Clave_Unidad, 
                    UNIDAD.Rodada, 
                    PRESTAMO.Fecha_Prestamo, 
                    CLIENTES.Nombres || ' ' || CLIENTES.Apellidos AS Nombre_Completo, 
                    CLIENTES.Telefono
                FROM 
                    PRESTAMO
                JOIN 
                    CLIENTES ON PRESTAMO.Clave_Cliente = CLIENTES.Clave
                JOIN 
                    UNIDAD ON PRESTAMO.Clave_Unidad = UNIDAD.Clave
                WHERE 
                    DATE(PRESTAMO.Fecha_Prestamo) BETWEEN ? AND ?
            """, (fecha_inicio, fecha_fin))

            registros = mi_cursor.fetchall()

            # Verificar si los registros no están vacíos
            if registros:
                # Formatear las fechas en "mes/día/año"
                for i in range(len(registros)):
                    fecha_prestamo = registros[i][2]  # Este es un string
                    if isinstance(fecha_prestamo, str):
                        fecha_prestamo = datetime.datetime.strptime(fecha_prestamo, "%Y-%m-%d")  # Ajusta el formato según sea necesario
                    fecha_formateada = fecha_prestamo.strftime("%m/%d/%Y")  # Convertir la fecha al formato deseado
                    registros[i] = (registros[i][0], registros[i][1], fecha_formateada, registros[i][3], registros[i][4])

                # Mostrar resultados en forma tabular
                headers = ["Clave de Unidad", "Rodada", "Fecha de Préstamo", "Nombre Completo", "Teléfono"]
                print("\n--- Reporte por Período ---")
                print(tabulate(registros, headers=headers, tablefmt="rounded_outline"))

                # Preguntar si se desea exportar
                export_option = input("\n¿Desea exportar el reporte? (1: CSV, 2: Excel, 3: Ambas, 4: No exportar): ")
                
                # Crear DataFrame para exportar
                df = pd.DataFrame(registros, columns=headers)
                
                # Exportar según la elección
                if export_option == "1":
                    df.to_csv("reporte_por_periodo.csv", index=False)
                    print("Reporte exportado a 'reporte_por_periodo.csv'.")
                elif export_option == "2":
                    # Exportar a Excel
                    df.to_excel("reporte_por_periodo.xlsx", index=False)

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("reporte_por_periodo.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("reporte_por_periodo.xlsx")
                    
                elif export_option == "3":
                    df.to_csv("reporte_por_periodo.csv", index=False)
                    df.to_excel("reporte_por_periodo.xlsx", index=False)

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("reporte_por_periodo.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("reporte_por_periodo.xlsx")

                    print("Reporte exportado a 'reporte_por_periodo.csv' y 'reporte_por_periodo.xlsx'.")
                elif export_option == "4":
                    print("No se exportó el reporte.")
                else:
                    print("Opción no válida. No se exportó el reporte.")
            else:
                print("\nNo hay registros en el período indicado.")

    except sqlite3.Error as e:
        print(f"Error en la base de datos: {e}")
    except ValueError:
        print("Formato de fecha incorrecto. Por favor, utiliza el formato mm/dd/aaaa.")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        print("Se ha cerrado la conexión")

#EV3
#Retrasos
def reporte_retrasos():
    try:
        with sqlite3.connect("RentaBicicletas.db", 
                             detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES) as conn:
            mi_cursor = conn.cursor()
            
            # Obtener la fecha actual
            today = datetime.datetime.now().date()

            # Consulta SQL para obtener préstamos con retraso
            mi_cursor.execute("""
                SELECT 
                    (JULIANDAY(?) - JULIANDAY(PRESTAMO.Fecha_Prestamo)) AS Dias_Retraso,
                    PRESTAMO.Fecha_Retorno,
                    PRESTAMO.Clave_Unidad,
                    UNIDAD.Rodada,
                    UNIDAD.Color,
                    CLIENTES.Nombres || ' ' || CLIENTES.Apellidos AS Nombre_Completo,
                    CLIENTES.Telefono
                FROM 
                    PRESTAMO
                JOIN 
                    CLIENTES ON PRESTAMO.Clave_Cliente = CLIENTES.Clave
                JOIN 
                    UNIDAD ON PRESTAMO.Clave_Unidad = UNIDAD.Clave
                WHERE 
                    PRESTAMO.Retorno = False
                    AND PRESTAMO.Fecha_Retorno < ?;
            """, (today, today))

            registros = mi_cursor.fetchall()

            # Verificar si hay registros de retrasos
            if registros:
                # Formatear las fechas en "mes/día/año"
                for i in range(len(registros)):
                    fecha_retorno = registros[i][1]  # Fecha de retorno
                    if isinstance(fecha_retorno, str):
                        fecha_retorno = datetime.datetime.strptime(fecha_retorno, "%Y-%m-%d")  # Ajusta el formato según sea necesario
                    fecha_formateada_retorno = fecha_retorno.strftime("%m/%d/%Y")  # Formato deseado

                    # Actualizar el registro con la fecha formateada
                    registros[i] = (registros[i][0], fecha_formateada_retorno, registros[i][2], registros[i][3], registros[i][4], registros[i][5], registros[i][6])

                # Mostrar resultados en forma tabular
                headers = ["Días de Retraso", "Fecha de Retorno", "Clave de Unidad", "Rodada", "Color", "Nombre Completo", "Teléfono"]
                print("\n--- Reporte de Retrasos ---")
                print(tabulate(registros, headers=headers, tablefmt="rounded_outline"))

                # Preguntar si se desea exportar
                export_option = input("\n¿Desea exportar el reporte? (1: CSV, 2: Excel, 3: Ambas, 4: No exportar): ")

                # Crear DataFrame para exportar
                df = pd.DataFrame(registros, columns=headers)
                
                # Exportar según la elección
                if export_option == "1":
                    df.to_csv("retrasos.csv", index=False)
                    print("Reporte exportado a 'retrasos.csv'.")
                elif export_option == "2":
                    df.to_excel("retrasos.xlsx", index=False)

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("retrasos.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("retrasos.xlsx")

                    print("Reporte exportado a 'retrasos.xlsx'.")
                elif export_option == "3":
                    df.to_csv("retrasos.csv", index=False)
                    df.to_excel("retrasos.xlsx", index=False)

                    # Ajustar el ancho de las columnas
                    workbook = load_workbook("retrasos.xlsx")
                    sheet = workbook.active
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = (max_length + 2)
                        sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    workbook.save("retrasos.xlsx")

                    print("Reporte exportado a 'retrasos.csv' y 'retrasos.xlsx'.")
                elif export_option == "4":
                    print("No se exportó el reporte.")
                else:
                    print("Opción no válida. No se exportó el reporte.")

            else:
                print("\nNo hay préstamos con retraso.")

    except sqlite3.Error as e:
        print(f"Error en la base de datos: {e}")
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        print("Se ha cerrado la conexión")

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



## SUBMENÚ PREFERENCIAS RENTAS
def preferencias_rentas():
    while True:
        print("Elige el reporte que deseas generar:")
        print("1. Cantidad de préstamos por rodada")
        print("2. Cantidad de préstamos por color")
        print("3. Por días de la semana")
        print("4. Volver al submenu")
        
        opcion_pref = input("Ingresa una de las opciónes mencionadas: ")
        
        if opcion_pref.isdigit():
            opcion_pref = int(opcion_pref)
            if opcion_pref == 1:
                rodada_tab_count()
            elif opcion_pref == 2:
                colores_tab_count()
            elif opcion_pref == 3:
                prestamos_por_dia_semana()
            elif opcion_pref == 4:
                break
            else:
                print("Opción inválida. Debes ingresar 1, 2, 3 o 4.")
        else:
            print("Entrada inválida. Por favor ingresa un número (1, 2, 3 o 4).")

   
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


#EV3
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

            print("\n--- Reporte de cantidad de préstamos por día de la semana ---")
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
prestamos = import_prestamos()
print("===== BIENVENIDO A NUESTRA RENTA DE BICICLETAS =====")
menu_principal()