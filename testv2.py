import csv
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font 

#funcion que despliega el menu principal
def menu_completo():
    while True:
        print("\n--- MENÚ PRINCIPAL ---")
        print("1. Registro")
        print("2. Prestamo")
        print("3. Retorno")
        print("4. Reportes")
        print("5. Salir\n")

        try:
            opcion = input("Elige una de las siguientes opciones: ")
            opcion = int(opcion)

            if opcion == 1:
                menu_registro()
            elif opcion == 2:
                registrar_prestamo()
            elif opcion == 3:
                menu_retorno()
            elif opcion == 4:
                menu_export()
            elif opcion == 5:
                confirmacion = input("¿Desea salir del programa? (S/N)").upper()
                if confirmacion == "S":
                    print("Saliendo del sistema...\n")
                    break
                elif confirmacion == "N":
                    menu_completo()
                else:
                    print("Opción invalida, ingrese los valores de 'S' o 'N'.")
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            print('Favor de ingresar un valor numerico')
            
#funcion que despliega el sub menú de registro
def menu_registro():

    while True:
        print("\n--- SUBMENÚ REGISTRO ---")
        print("1. Registrar una unidad")
        print("2. Registrar un cliente")
        print("3. Volver al menu principal\n")

        try:
            opcion = input("Elige una opción: ")
            opcion = int(opcion)

            if opcion == 1:
                registro_Unidad()
            elif opcion == 2:
                registro_Cliente()
            elif opcion == 3:
                return False
            else:
                print("Opción invalida, intentalo de nuevo.")
        except ValueError:
            if cancelar():
                break

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

#funcion que permite registrar una unidad lista para un prestamo
def registro_Unidad():
    while True:
        opcion = input("¿Deseas realizar un registro de unidad? (S/N): ").upper()

        if opcion == "S":
            print("\n--- REGISTRO DE UNIDAD ---")
            clave = max(unidades, default=0) + 1
            while True:
                entrada = input('Ingrese la rodada de la unidad (20, 26 o 29): ')
                try:
                    rodada = int(entrada)
                    if rodada in [20, 26, 29]:
                        unidades[clave] = rodada
                        print(f"Unidad registrada con exito. Clave: {clave}, Rodada: {rodada}")
                        return False
                    else:
                        print("Por favor, ingrese un valor valido (20, 26 o 29).")
                        if cancelar():
                            return
                except ValueError:
                    if cancelar():
                        return
        elif opcion == "N":
            # Regresar al menú registro si elige 'N'
            return False
        else:
            print("Opción inválida. Debes ingresar 'S' o 'N'.")
            return
            
#funcion que permite registrar un cliente listo para solicitar un prestamo           
def registro_Cliente():
    while True:
        opcion = input("¿Deseas realizar un registro de cliente? (S/N): ").upper()

        if opcion == "S":
            print("\n--- REGISTRO DE CLIENTE ---")
            clave_cliente = max(clientes, default=0) + 1
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
            # Registro del cliente en la base de datos
            clientes[clave_cliente] = (apellidos, nombre, telefono)
            print(f"Cliente registrado con éxito. Clave: {clave_cliente}, Nombre: {nombre} {apellidos}, Teléfono: {telefono}")
            # Llamada a función para exportar datos
            export_clientes_auto(clientes)
            # Salir del bucle después de exportar
            break
        elif opcion == "N":
            # Regresar al menú de registro si elige 'N'
            menu_registro()
            break  # Salir del bucle para regresar al menú
        else:
            print("Opción inválida. Debes ingresar 'S' o 'N'.")

def export_clientes_auto(clientes):
    with open("Clientes_bicicletas.csv", "w", encoding="latin1", newline="") as archivocsv_clientes:
        grabador = csv.writer(archivocsv_clientes)
        grabador.writerow(("Clave", "Apellidos", "Nombres", "Teléfono"))
        grabador.writerows([(clave, datos[0], datos[1], datos[2]) for clave, datos in clientes.items()])     

def registrar_prestamo():
    while True:
            opcion = input("¿Deseas realizar un registro de préstamos? (S/N): ").upper()
            
            if opcion == "S":
                print("\n--- REGISTRO DE PRÉSTAMO ---")
                fecha_actual = datetime.datetime.now().date()
                folio = max(prestamos, default=0) + 1

                # Captura de la clave de la unidad
                while True:
                    Clave_unidad = input("Clave de la unidad: ")
                    try:
                        Clave_unidad = int(Clave_unidad)
                        if Clave_unidad in unidades:
                            break
                        else:
                            print("La clave de la unidad no es válida.")
                            if cancelar():
                                return
                    except ValueError:
                        if cancelar():
                            return

                # Captura de la clave del cliente
                while True:
                    Clave_cliente = input("Clave del cliente: ")
                    try:
                        Clave_cliente = int(Clave_cliente)
                        if Clave_cliente in clientes:
                            break
                        else:
                            print("La clave del cliente no es válida.")
                            if cancelar():
                                return
                    except ValueError:
                        if cancelar():
                            return

                # Elección de la fecha del préstamo
                while True:
                    eleccion_de_fecha = input("¿Deseas que la fecha sea la del día de hoy?\n1. Sí\n2. No\nElige una opción: ")
                    try:
                        eleccion_de_fecha = int(eleccion_de_fecha)
                        if eleccion_de_fecha == 1:
                            fecha_prestamo = fecha_actual
                            break
                        elif eleccion_de_fecha == 2:
                            while True:
                                fecha_a_elegir = input("Indica la fecha del préstamo (MM/DD/AAAA): ")
                                try:
                                    fecha_prestamo = datetime.datetime.strptime(fecha_a_elegir, "%m/%d/%Y").date()
                                    if fecha_prestamo >= fecha_actual:
                                        break
                                    else:
                                        print("La fecha no puede ser anterior al día de hoy.")
                                except ValueError:
                                    print("Formato de fecha incorrecto, intenta de nuevo.")
                            break
                        else:
                            print("Opción inválida, intenta de nuevo.")
                    except ValueError:
                        print("Entrada inválida, elige 1 o 2.")

                # Registro del préstamo
                prestamos[folio] = {
                    'Clave_cliente': Clave_cliente,
                    'Clave_unidad': Clave_unidad,
                    'Fecha_prestamo': fecha_prestamo.strftime("%m/%d/%Y"),
                    'Fecha_retorno': None  # Se completará cuando se registre el retorno
                }

                print(f"Préstamo registrado exitosamente. Folio: {folio}, Cliente: {Clave_cliente}, Unidad: {Clave_unidad}, Fecha de Préstamo: {fecha_prestamo}")

            elif opcion == "N":
                # Regresar al menú si elige 'N'
                break
            else:
                print("Opción inválida. Debes ingresar 'S' o 'N'.")
        
#Función que despliega menú para hacer el retorno de la unidad (Carlos)
def menu_retorno():
    if prestamos:
        opcion = input("¿Deseas retornar una unidad? (S/N): ").upper()
        if opcion == "S":
            while True:
                numdefolio = input("\nIngrese el número de folio de su préstamo: \n")
                try:
                    numdefolio = int(numdefolio)
                    if numdefolio in prestamos:
                        today = datetime.now().date()
                        prestamos[numdefolio]["Retornado"] = True  #v2
                        print("Retornó su unidad exitosamente el día", today.strftime('%m/%d/%Y'), "\n")
                        break
                    else:
                        print("El número de folio no existe. Por favor, inténtalo de nuevo.")
                except ValueError:
                    print("Por favor, ingrese un número entero.")
        elif opcion == "N":
            # Regresar al menú completo si elige 'N'
            return False
        else:
            print("Opción inválida. Debes ingresar 'S' o 'N'.")
            registro_Unidad()
    else:
        print("No hay ningún prestamo realizado.")

## JESÚS        
def export_excel_clientes(clientes, name_excel="Clientes.xlsx"):
    libro = openpyxl.Workbook()

    hoja = libro.active
    hoja.title = "Clientes"

    hoja["A1"].value = "Clave"
    hoja["B1"].value = "Apellidos"
    hoja["C1"].value = "Nombres"
    hoja["D1"].value = "Teléfono"

    i = 2

    for clave, (apellidos, nombres, telefono) in clientes.items():
        hoja.cell(row=i, column=1).value = clave
        hoja.cell(row=i, column=2).value = apellidos
        hoja.cell(row=i, column=3).value = nombres
        hoja.cell(row=i, column=4).value = telefono
        i += 1

    libro.save(name_excel)
    print(f"Datos exportados con éxito en {name_excel}")

def tab_clientes(clientes):
    print(f"{'Clave':^8}{'Apellidos': <41}{'Nombres': <41}{'Teléfono'}")
    print("=" * 100)
    for clave, datos in clientes.items():
        print(f"{clave:^8}{datos[0]: <41}{datos[1]: <41}{datos[2]}")
    print("=" * 100)

def export_csv_clientes(clientes):
    with open("Clientes_bicicletas.csv", "w", encoding="latin1", newline="") as archivocsv_clientes:
        grabador = csv.writer(archivocsv_clientes)
        grabador.writerow(("Clave", "Apellidos", "Nombres", "Teléfono"))
        grabador.writerows([(clave, datos[0], datos[1], datos[2]) for clave, datos in clientes.items()])
    print("Datos exportados con éxito en Clientes_bicicletas.csv")
  
def cargar_clientes_csv(nombre_archivo="Clientes_bicicletas.csv"):
    clientes = {}
    try:
        with open(nombre_archivo, "r", encoding="latin1", newline="") as archivocsv_clientes:
            lector = csv.reader(archivocsv_clientes)
            next(lector)
            for fila in lector:
                clave, apellidos, nombres, telefono = fila
                clientes[int(clave)] = (apellidos, nombres, telefono)
    except FileNotFoundError:
        print("El archivo no existe. Se creará uno nuevo al exportar.")
    return clientes

def reporte_prestamos_por_retornar(prestamos):
    if prestamos:
        try:
            print("\n--- REPORTES DE PRÉSTAMOS POR RETORNAR ---")
            fecha_inicial = input("Ingrese la fecha inicial (MM/DD/AAAA): ")
            fecha_inicial = datetime.strptime(fecha_inicial, "%m/%d/%Y").date()
            fecha_final = input("Ingrese la fecha final (MM/DD/AAAA): ")
            fecha_final = datetime.strptime(fecha_final, "%m/%d/%Y").date() 

            if fecha_final < fecha_inicial:
                print("La fecha final no puede ser anterior a la fecha inicial.")
                return
            
            print(f"{'Folio':^8}{'Clave del Cliente': <20}{'Clave de la Unidad': <20}{'Fecha Préstamo': <20}{'Fecha Retorno'}")
            print("=" * 80)

            prestamos_mostrados = False
            for folio, datos in prestamos.items():
                fecha_prestamo = datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date()
                if fecha_inicial <= fecha_prestamo <= fecha_final:
                    print(f"{folio:^8}{datos['Clave_cliente']: <20}{datos['Clave_unidad']: <20}{datos['Fecha_prestamo']: <20}{datos['Fecha_retorno']}")
                    prestamos_mostrados = True

            if not prestamos_mostrados:
                print("No se encontraron préstamos en el rango de fechas especificado.")

            print("=" * 80)

            # Opciones de exportación
            export_opcion = input("Elige una opción de exportación:\n1. CSV\n2. Excel\n3. Ambos\n")
            if export_opcion == '1':
                export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            elif export_opcion == '2':
                export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            elif export_opcion == '3':
                export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
                export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            else:
                print("Elige una opción válida (1, 2 o 3).")

        except ValueError as e:
            print(f"Error en el formato de la fecha: {e}")
        except KeyError as e:
            print(f"Error con los datos: {e}")
    else:
        print("No se encontró ningún préstamo.")

def export_excel_prestamos_retornar(prestamos, fecha_inicial, fecha_final, name_excel="Prestamos_por_retornar.xlsx"):
    libro = openpyxl.Workbook()

    hoja = libro.active
    hoja.title = "Prestamos"

    hoja["A1"].value = "Folio"
    hoja["B1"].value = "Clave de la unidad"
    hoja["C1"].value = "Clave del cliente"
    hoja["D1"].value = "Fecha prestamo"
    hoja["E1"].value = "Fecha de retorno"

    i = 2
    for folio, datos in prestamos.items():
        if not datos["Retornado"]:
            fecha_retorno = datetime.strptime(datos["Fecha_retorno"], "%m/%d/%Y").date()
            if fecha_inicial <= fecha_retorno <= fecha_final:
                hoja.cell(row=i, column=1).value = folio
                hoja.cell(row=i, column=2).value = datos["Clave_unidad"]
                hoja.cell(row=i, column=3).value = datos["Clave_cliente"]
                hoja.cell(row=i, column=4).value = datos["Fecha_prestamo"]
                hoja.cell(row=i, column=5).value = datos["Fecha_retorno"]
                i += 1

    libro.save(name_excel)
    print(f"Datos exportados con éxito en {name_excel}")

def export_csv_prestamos_retornar(prestamos, fecha_inicial, fecha_final, nombre_csv="Prestamos_por_retornar.csv"):
    with open(nombre_csv, "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)

        grabador.writerow(("Folio", "Clave de la unidad", "Clave del cliente", "Fecha prestamo", "Fecha de retorno"))


        prestamos_filtrados = [
            (folio, datos["Clave_unidad"], datos["Clave_cliente"], datos["Fecha_prestamo"], datos["Fecha_retorno"])
            for folio, datos in prestamos.items()
            if not datos["Retornado"] and fecha_inicial <= datetime.strptime(datos["Fecha_retorno"], "%m/%d/%Y").date() <= fecha_final
        ]

        if prestamos_filtrados:
            grabador.writerows(prestamos_filtrados)
            print(f"Datos exportados con éxito en {nombre_csv}")
        else:
            print("No hay préstamos que coincidan con los criterios especificados.")

def exportar_clientes():
    while True:
        if clientes:
            tab_clientes(clientes)
            try:
                export_opcion = int(input("Elige una opción de exportación: \n1. CSV\n2. Excel\n3. Ambos\n4. Salir al submenú\n"))
                if export_opcion == 1:
                    export_csv_clientes(clientes)
                elif export_opcion == 2:
                    export_excel_clientes(clientes)
                elif export_opcion == 3:
                    export_csv_clientes(clientes)
                    export_excel_clientes(clientes)
                elif export_opcion == 4:
                    break
                else:
                    print("Elige una opcion valida")
            except ValueError:
                    print("Error: Debes ingresar un número entero que sea válido.")
            except Exception as name_error:
                    print(f"Ha ocurrido un error inesperado: {name_error}")
            else:
                print("No hay clientes para exportar")
                break

def opciones_report():

  while True:
    print("\n--- SUBMENÚ REPORTES ---")
    print("1. Clientes")
    print("2. Prestamos por retornar")
    print("3. Prestamos por periodo")
    print("4. Salir al menú principal\n")

    try:
      reporte_opcion = int(input("Elige alguna de las opciones mencionadas: "))
      if reporte_opcion == 1:
        exportar_clientes()
      elif reporte_opcion == 2:
        reporte_prestamos_por_retornar(prestamos)
      elif reporte_opcion == 3:
        prestamos_por_periodo()
      elif reporte_opcion == 4:
        return False
      else:
        print("Ingresa una opción válida")
    except Exception as error_name:
        print("Ha ocurrido un error")

def export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final, name_excel="Prestamos_por_periodo.xlsx"):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.title = "Prestamos"

    hoja["A1"].value = "Folio"
    hoja["B1"].value = "Clave de la unidad"
    hoja["C1"].value = "Clave del cliente"
    hoja["D1"].value = "Fecha préstamo"
    hoja["E1"].value = "Fecha de retorno"

    # Poner negrita en los encabezados (fila 1)
    hoja["A1"].font = Font(bold=True)
    hoja["B1"].font = Font(bold=True)
    hoja["C1"].font = Font(bold=True)
    hoja["D1"].font = Font(bold=True)
    hoja["E1"].font = Font(bold=True)

    i = 2  # Iniciar en la fila 2 porque la fila 1 son los encabezados
    for folio, datos in prestamos.items():
        fecha_prestamo = datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date()
        if fecha_inicial <= fecha_prestamo <= fecha_final:
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

    # Guarda el archivo Excel
    libro.save(name_excel)
    
    print(f"Datos exportados con éxito en {name_excel}")
    
def export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final, nombre_csv="Prestamos_por_periodo.csv"):
    with open(nombre_csv, "w", encoding="latin1", newline="") as archivo_csv:
        grabador = csv.writer(archivo_csv)
        grabador.writerow(("Folio", "Clave de la unidad", "Clave del cliente", "Fecha préstamo", "Fecha de retorno"))

        prestamos_filtrados = [
            (folio, datos["Clave_unidad"], datos["Clave_cliente"], datos["Fecha_prestamo"], datos["Fecha_retorno"])
            for folio, datos in prestamos.items()
            if fecha_inicial <= datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date() <= fecha_final
        ]

        if prestamos_filtrados:
            grabador.writerows(prestamos_filtrados)
            print(f"Datos exportados con éxito en {nombre_csv}")
        else:
            print("No hay préstamos que coincidan con los criterios especificados.")

def prestamos_por_periodo():
    if prestamos:
        while True:
            fecha_inicial = input("Indica la fecha inicial del período (MM/DD/AAAA): ")
            try:
                fecha_inicial = datetime.datetime.strptime(fecha_inicial, "%m/%d/%Y").date()
                break
            except ValueError:
                print("Formato de fecha incorrecto, intenta de nuevo.")

        while True:
            fecha_final = input("Indica la fecha final del período (MM/DD/AAAA): ")
            try:
                fecha_final = datetime.datetime.strptime(fecha_final, "%m/%d/%Y").date()
                if fecha_final >= fecha_inicial:
                    break
                else:
                    print("La fecha final debe ser posterior o igual a la fecha inicial.")
            except ValueError:
                print("Formato de fecha incorrecto, intenta de nuevo.")

        # Mostrar el reporte
        print(f"{'Folio':^8}{'Clave del Cliente': <20}{'Clave de la Unidad': <20}{'Fecha Préstamo': <20}{'Fecha Retorno'}")
        print("=" * 80)

        for folio, datos in prestamos.items():
            fecha_prestamo = datetime.datetime.strptime(datos['Fecha_prestamo'], "%m/%d/%Y").date()
            if fecha_inicial <= fecha_prestamo <= fecha_final:
                print(f"{folio:^8}{datos['Clave_cliente']: <20}{datos['Clave_unidad']: <20}{datos['Fecha_prestamo']: <20}{datos['Fecha_retorno']}")

        print("=" * 80)

        # Opciones de exportación
        export_opcion = input("Elige una opción de exportación:\n1. CSV\n2. Excel\n3. Ambos\n")
        if export_opcion == "1":
            export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
        elif export_opcion == "2":
            export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
        elif export_opcion == "3":
            export_csv_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
            export_excel_prestamos_por_periodo(prestamos, fecha_inicial, fecha_final)
        else:
            print("Elige una opción válida (1, 2 o 3).")
    else:
        print("No se encontró ningún préstamo.")
        opciones_report()
        
def menu_export():
  opciones_report()
  
  
# Inicio del programa
clientes = cargar_clientes_csv()
unidades = {}
clientes = {}
prestamos = {}
print("===== BIENVENIDO A NUESTRA RENTA DE BICICLETAS =====")
menu_completo()