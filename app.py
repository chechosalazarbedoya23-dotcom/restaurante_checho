# -*- coding: utf-8 -*-
"""
SISTEMA DE RESERVAS DE RESTAURANTE  (programa para la terminal)

Menú:
  1. Registrar cliente
  2. Ver mesas
  3. Crear reserva
  4. Exportar información a Excel
  5. Cancelar reserva
  6. Ver reservas
  7. Ver clientes
  8. Salir

Requisito para exportar a Excel:   pip install openpyxl
Los datos se guardan solos en 'datos_restaurante.json' (misma carpeta del
programa), así que no se pierden al cerrar.
"""

import json
import os
from datetime import datetime, timedelta

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ----------------------------------------------------------------
# CONFIGURACIÓN
# ----------------------------------------------------------------
CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_DATOS = os.path.join(CARPETA, "datos_restaurante.json")
ARCHIVO_EXCEL = os.path.join(CARPETA, "reservas_restaurante.xlsx")

FORMATO_FECHA = "%d/%m/%Y"
FORMATO_HORA = "%H:%M"
DURACION_MINUTOS = 120        # cada reserva ocupa la mesa 2 horas
HORA_APERTURA = "11:00"
ULTIMA_HORA_RESERVA = "21:00"  # 21:00 + 2 horas = cierre a las 23:00
HORA_CIERRE = "23:00"

# Mesas: número, comensales, ubicación y valor de la reserva (en pesos)
mesas = [
    {"numero": 1, "capacidad": 2,  "ubicacion": "Terraza",         "valor": 20000},
    {"numero": 2, "capacidad": 2,  "ubicacion": "Terraza",         "valor": 20000},
    {"numero": 3, "capacidad": 4,  "ubicacion": "Salón principal", "valor": 30000},
    {"numero": 4, "capacidad": 4,  "ubicacion": "Salón principal", "valor": 30000},
    {"numero": 5, "capacidad": 4,  "ubicacion": "Ventana",         "valor": 35000},
    {"numero": 6, "capacidad": 6,  "ubicacion": "Salón principal", "valor": 50000},
    {"numero": 7, "capacidad": 6,  "ubicacion": "Patio",           "valor": 55000},
    {"numero": 8, "capacidad": 8,  "ubicacion": "Zona VIP",        "valor": 90000},
    {"numero": 9, "capacidad": 10, "ubicacion": "Zona VIP",        "valor": 120000},
]

# Datos que se guardan en el archivo JSON
clientes = []   # identificacion, nombre, apellido, telefono, otro_telefono, correo, fecha_registro
reservas = []   # codigo, identificacion, a_nombre_de, mesa, personas, fecha, hora, valor,
                # observaciones, estado, fecha_creacion, fecha_cancelacion
contador_reservas = 0


# ----------------------------------------------------------------
# GUARDAR Y CARGAR DATOS
# ----------------------------------------------------------------
def guardar_datos():
    datos = {
        "clientes": clientes,
        "reservas": reservas,
        "contador_reservas": contador_reservas,
    }
    with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)


def cargar_datos():
    global contador_reservas
    if not os.path.exists(ARCHIVO_DATOS):
        return
    try:
        with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
            datos = json.load(f)
        clientes.extend(datos.get("clientes", []))
        reservas.extend(datos.get("reservas", []))
        contador_reservas = datos.get("contador_reservas", len(reservas))
    except (json.JSONDecodeError, OSError):
        print("Aviso: no se pudo leer el archivo de datos. Se inicia vacío.")


# ----------------------------------------------------------------
# FUNCIONES DE APOYO (presentación)
# ----------------------------------------------------------------
def titulo(texto):
    print("\n" + "=" * 60)
    print(f"  {texto}")
    print("=" * 60)


def pausar():
    input("\nPresiona Enter para volver al menú...")


def dinero(valor):
    """20000 -> $20.000"""
    return "$" + f"{valor:,.0f}".replace(",", ".")


def nombre_completo(cliente):
    return f"{cliente['nombre']} {cliente['apellido']}"


# ----------------------------------------------------------------
# FUNCIONES DE APOYO (lectura y validación de datos)
# ----------------------------------------------------------------
def leer_entero(mensaje, minimo, maximo):
    while True:
        try:
            valor = int(input(mensaje).strip())
        except ValueError:
            print("Error: escribe un número entero.")
            continue
        if valor < minimo or valor > maximo:
            print(f"Error: el número debe estar entre {minimo} y {maximo}.")
            continue
        return valor


def leer_nombre(mensaje):
    """Solo letras (con tildes), espacios y guiones."""
    while True:
        valor = input(mensaje).strip()
        if valor == "":
            print("Este dato es obligatorio.")
        elif not valor.replace(" ", "").replace("-", "").isalpha():
            print("Solo se permiten letras.")
        else:
            return valor.title()


def leer_identificacion(mensaje):
    """Solo números, entre 5 y 12 dígitos."""
    while True:
        valor = input(mensaje).strip().replace(".", "").replace(" ", "")
        if not valor.isdigit() or not (5 <= len(valor) <= 12):
            print("Identificación inválida: solo números, entre 5 y 12 dígitos.")
        else:
            return valor


def leer_telefono(mensaje, obligatorio=True):
    """Solo números, entre 7 y 10 dígitos."""
    while True:
        valor = input(mensaje).strip().replace(" ", "")
        if valor == "" and not obligatorio:
            return ""
        if not valor.isdigit() or not (7 <= len(valor) <= 10):
            print("Teléfono inválido: solo números, entre 7 y 10 dígitos.")
        else:
            return valor


def leer_correo(mensaje):
    """Opcional. Si se escribe, debe parecer un correo (tener @ y un punto)."""
    while True:
        valor = input(mensaje).strip().lower()
        if valor == "":
            return ""
        if "@" in valor and "." in valor.split("@")[-1] and " " not in valor:
            return valor
        print("Correo inválido. Ejemplo: nombre@correo.com")


def leer_si_no(mensaje):
    while True:
        r = input(mensaje).strip().lower()
        if r in ("s", "si", "sí"):
            return True
        if r in ("n", "no"):
            return False
        print("Responde 's' (sí) o 'n' (no).")


def leer_fecha(mensaje, permitir_pasado=False, vacio_es_hoy=False):
    while True:
        texto = input(mensaje).strip()
        if texto == "" and vacio_es_hoy:
            return datetime.now().strftime(FORMATO_FECHA)
        try:
            fecha = datetime.strptime(texto, FORMATO_FECHA)
        except ValueError:
            print("Fecha inválida. Usa el formato DD/MM/AAAA, por ejemplo 15/11/2026.")
            continue
        if not permitir_pasado and fecha.date() < datetime.now().date():
            print("La fecha no puede ser anterior a hoy.")
            continue
        return fecha.strftime(FORMATO_FECHA)


def leer_hora(mensaje, fecha, validar_horario=True, vacio_es_ahora=False):
    while True:
        texto = input(mensaje).strip()
        if texto == "" and vacio_es_ahora:
            return datetime.now().strftime(FORMATO_HORA)
        try:
            hora = datetime.strptime(texto, FORMATO_HORA).strftime(FORMATO_HORA)
        except ValueError:
            print("Hora inválida. Usa el formato HH:MM en 24 horas, por ejemplo 19:30.")
            continue
        if validar_horario:
            if hora < HORA_APERTURA or hora > ULTIMA_HORA_RESERVA:
                print(f"El restaurante atiende de {HORA_APERTURA} a {HORA_CIERRE}. "
                      f"Las reservas pueden empezar entre las {HORA_APERTURA} y las {ULTIMA_HORA_RESERVA}.")
                continue
            if inicio_reserva(fecha, hora) <= datetime.now():
                print("Esa hora ya pasó. Elige una hora futura.")
                continue
        return hora


# ----------------------------------------------------------------
# FUNCIONES DE APOYO (clientes, mesas y reservas)
# ----------------------------------------------------------------
def buscar_cliente(identificacion):
    for c in clientes:
        if c["identificacion"] == identificacion:
            return c
    return None


def buscar_mesa(numero):
    for m in mesas:
        if m["numero"] == numero:
            return m
    return None


def inicio_reserva(fecha, hora):
    return datetime.strptime(f"{fecha} {hora}", f"{FORMATO_FECHA} {FORMATO_HORA}")


def hora_fin(hora):
    """Hora en que termina una reserva que empieza a 'hora'."""
    inicio = datetime.strptime(hora, FORMATO_HORA)
    return (inicio + timedelta(minutes=DURACION_MINUTOS)).strftime(FORMATO_HORA)


def reserva_de_mesa(numero_mesa, fecha, hora):
    """Si la mesa ya está reservada en esa fecha y hora (o se cruza con otra
    reserva activa), devuelve esa reserva. Si está libre, devuelve None."""
    nuevo = inicio_reserva(fecha, hora)
    for r in reservas:
        if r["mesa"] != numero_mesa or r["estado"] != "Activa":
            continue
        existente = inicio_reserva(r["fecha"], r["hora"])
        if abs((nuevo - existente).total_seconds()) < DURACION_MINUTOS * 60:
            return r
    return None


def texto_estado(numero_mesa, fecha, hora):
    r = reserva_de_mesa(numero_mesa, fecha, hora)
    if r is None:
        return "Libre"
    return f"RESERVADA ({r['a_nombre_de']}, {r['hora']} a {hora_fin(r['hora'])})"


# ----------------------------------------------------------------
# OPCIÓN 1: REGISTRAR CLIENTE
# ----------------------------------------------------------------
def registrar_cliente():
    titulo("REGISTRAR CLIENTE")

    identificacion = leer_identificacion("Identificación (cédula): ")
    existente = buscar_cliente(identificacion)
    if existente:
        print(f"\nEl cliente ya se encuentra registrado: "
              f"{nombre_completo(existente)} (identificación {identificacion}).")
        return

    nombre = leer_nombre("Nombre: ")
    apellido = leer_nombre("Apellido: ")
    telefono = leer_telefono("Teléfono: ")
    otro_telefono = leer_telefono("Otro teléfono (opcional, Enter para omitir): ", obligatorio=False)
    correo = leer_correo("Correo electrónico (opcional, Enter para omitir): ")

    clientes.append({
        "identificacion": identificacion,
        "nombre": nombre,
        "apellido": apellido,
        "telefono": telefono,
        "otro_telefono": otro_telefono,
        "correo": correo,
        "fecha_registro": datetime.now().strftime("%d/%m/%Y %H:%M"),
    })
    guardar_datos()
    print(f"\nCliente registrado con éxito: {nombre} {apellido} (identificación {identificacion}).")


# ----------------------------------------------------------------
# OPCIÓN 2: VER MESAS
# ----------------------------------------------------------------
def ver_mesas():
    titulo("VER MESAS")
    print("Consulta el estado de las mesas para una fecha y hora.")
    print("(Presiona Enter en ambas para usar la fecha y hora actuales)\n")
    fecha = leer_fecha("Fecha (DD/MM/AAAA): ", permitir_pasado=True, vacio_es_hoy=True)
    hora = leer_hora("Hora (HH:MM): ", fecha, validar_horario=False, vacio_es_ahora=True)

    print(f"\nEstado de las mesas el {fecha} a las {hora}")
    print("-" * 95)
    print(f"{'Mesa':<6}{'Comensales':<12}{'Ubicación':<18}{'Valor':<12}Estado")
    print("-" * 95)
    for m in mesas:
        print(f"{m['numero']:<6}{m['capacidad']:<12}{m['ubicacion']:<18}"
              f"{dinero(m['valor']):<12}{texto_estado(m['numero'], fecha, hora)}")
    print("-" * 95)
    print(f"Cada reserva ocupa la mesa {DURACION_MINUTOS // 60} horas.")


# ----------------------------------------------------------------
# OPCIÓN 3: CREAR RESERVA
# ----------------------------------------------------------------
def mostrar_comprobante(r):
    mesa = buscar_mesa(r["mesa"])
    cliente = buscar_cliente(r["identificacion"])
    print("\n" + "=" * 60)
    print("            COMPROBANTE DE RESERVA")
    print("=" * 60)
    print(f"Código de reserva : {r['codigo']}")
    print(f"Cliente           : {nombre_completo(cliente)}")
    print(f"Identificación    : {r['identificacion']}")
    print(f"Teléfono          : {cliente['telefono']}")
    print(f"A nombre de       : {r['a_nombre_de']}")
    print(f"Mesa              : {mesa['numero']} ({mesa['ubicacion']})")
    print(f"Comensales        : {r['personas']} (la mesa admite {mesa['capacidad']})")
    print(f"Fecha             : {r['fecha']}")
    print(f"Hora              : {r['hora']} a {hora_fin(r['hora'])}")
    print(f"Valor de reserva  : {dinero(r['valor'])}")
    print(f"Observaciones     : {r['observaciones'] or 'Ninguna'}")
    print(f"Estado            : {r['estado']}")
    print("=" * 60)
    print("Guarda tu código de reserva: lo necesitas para cancelar.")


def crear_reserva():
    global contador_reservas
    titulo("CREAR RESERVA")

    if len(clientes) == 0:
        print("Aún no hay clientes registrados. Primero usa la opción 1.")
        return

    # 1. Comprobar que el cliente esté registrado (por identificación)
    identificacion = leer_identificacion("Identificación del cliente: ")
    cliente = buscar_cliente(identificacion)
    if cliente is None:
        print("\nEse cliente no está registrado. Regístralo primero con la opción 1.")
        return
    print(f"Cliente encontrado: {nombre_completo(cliente)}")

    # 2. A nombre de quién queda la reserva
    if leer_si_no("¿La reserva queda a nombre del cliente? (s/n): "):
        a_nombre_de = nombre_completo(cliente)
    else:
        a_nombre_de = leer_nombre("Nombre de la persona a quien queda la reserva: ")

    # 3. Datos de la reserva
    capacidad_maxima = max(m["capacidad"] for m in mesas)
    personas = leer_entero(f"Cantidad de personas (1 a {capacidad_maxima}): ", 1, capacidad_maxima)
    fecha = leer_fecha("Fecha de la reserva (DD/MM/AAAA): ")
    hora = leer_hora(f"Hora de la reserva (HH:MM, entre {HORA_APERTURA} y {ULTIMA_HORA_RESERVA}): ", fecha)

    # 4. Mostrar mesas con capacidad suficiente y su estado
    aptas = [m for m in mesas if m["capacidad"] >= personas]
    libres = [m for m in aptas if reserva_de_mesa(m["numero"], fecha, hora) is None]

    print(f"\nMesas para {personas} persona(s) el {fecha} a las {hora}:")
    print("-" * 95)
    print(f"{'Mesa':<6}{'Comensales':<12}{'Ubicación':<18}{'Valor':<12}Estado")
    print("-" * 95)
    for m in aptas:
        print(f"{m['numero']:<6}{m['capacidad']:<12}{m['ubicacion']:<18}"
              f"{dinero(m['valor']):<12}{texto_estado(m['numero'], fecha, hora)}")
    print("-" * 95)

    if len(libres) == 0:
        print("\nNo hay mesas libres para esa fecha, hora y cantidad de personas.")
        return

    # 5. Elegir mesa (si está ocupada, avisa y pide otra)
    while True:
        numero_mesa = leer_entero("Número de mesa a reservar (0 para cancelar): ", 0, 999)
        if numero_mesa == 0:
            print("Reserva cancelada. No se guardó nada.")
            return
        mesa = buscar_mesa(numero_mesa)
        if mesa is None or mesa["capacidad"] < personas:
            print("Esa mesa no existe o no tiene capacidad suficiente. Elige otra de la lista.")
            continue
        ocupada = reserva_de_mesa(numero_mesa, fecha, hora)
        if ocupada:
            print(f"La mesa ya está en reserva (a nombre de {ocupada['a_nombre_de']}, "
                  f"de {ocupada['hora']} a {hora_fin(ocupada['hora'])}). Elige otra.")
            continue
        break

    observaciones = input("Observaciones (opcional, ej. cumpleaños, silla de bebé): ").strip()

    # 6. Confirmar
    print(f"\nResumen: mesa {mesa['numero']} ({mesa['ubicacion']}), {fecha} a las {hora}, "
          f"a nombre de {a_nombre_de}, valor {dinero(mesa['valor'])}.")
    if not leer_si_no("¿Confirmar la reserva? (s/n): "):
        print("Reserva no confirmada. No se guardó nada.")
        return

    contador_reservas += 1
    reserva = {
        "codigo": f"R{contador_reservas:03d}",
        "identificacion": identificacion,
        "a_nombre_de": a_nombre_de,
        "mesa": numero_mesa,
        "personas": personas,
        "fecha": fecha,
        "hora": hora,
        "valor": mesa["valor"],
        "observaciones": observaciones,
        "estado": "Activa",
        "fecha_creacion": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fecha_cancelacion": "",
    }
    reservas.append(reserva)
    guardar_datos()
    print("\nReserva creada con éxito.")
    mostrar_comprobante(reserva)


# ----------------------------------------------------------------
# OPCIÓN 4: EXPORTAR A EXCEL
# ----------------------------------------------------------------
def dar_estilo_hoja(ws, columnas_dinero=(), columna_estado=None):
    borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    relleno_titulo = PatternFill("solid", fgColor="1F3864")
    relleno_cancelada = PatternFill("solid", fgColor="F8D7DA")

    for fila in ws.iter_rows():
        for celda in fila:
            celda.border = borde
            celda.alignment = Alignment(vertical="center")
            if celda.row == 1:
                celda.font = Font(name="Arial", bold=True, color="FFFFFF")
                celda.fill = relleno_titulo
                celda.alignment = Alignment(horizontal="center", vertical="center")
            else:
                celda.font = Font(name="Arial")
                if celda.column in columnas_dinero:
                    celda.number_format = '"$"#,##0'
        if columna_estado and fila[0].row > 1 and fila[columna_estado - 1].value == "Cancelada":
            for celda in fila:
                celda.fill = relleno_cancelada

    for i, columna in enumerate(ws.columns, start=1):
        ancho = max(len(str(c.value)) if c.value is not None else 0 for c in columna)
        ws.column_dimensions[get_column_letter(i)].width = min(max(ancho + 3, 12), 45)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def exportar_excel():
    titulo("EXPORTAR A EXCEL")

    if not OPENPYXL_OK:
        print("Falta la librería openpyxl. Instálala con:  pip install openpyxl")
        return
    if len(reservas) == 0:
        print("Aún no hay reservas para exportar.")
        return

    wb = Workbook()

    # Hoja 1: reservas
    ws = wb.active
    ws.title = "Reservas"
    ws.append(["Código", "Fecha", "Hora", "Hora fin", "Mesa", "Ubicación", "Comensales",
               "Valor", "A nombre de", "Cliente", "Identificación", "Teléfono",
               "Observaciones", "Estado", "Creada el"])
    ordenadas = sorted(reservas, key=lambda x: inicio_reserva(x["fecha"], x["hora"]))
    for r in ordenadas:
        c = buscar_cliente(r["identificacion"])
        m = buscar_mesa(r["mesa"])
        ws.append([
            r["codigo"],
            datetime.strptime(r["fecha"], FORMATO_FECHA).date(),
            r["hora"],
            hora_fin(r["hora"]),
            r["mesa"],
            m["ubicacion"],
            r["personas"],
            r["valor"],
            r["a_nombre_de"],
            nombre_completo(c),
            r["identificacion"],
            c["telefono"],
            r["observaciones"],
            r["estado"],
            r["fecha_creacion"],
        ])
    dar_estilo_hoja(ws, columnas_dinero=(8,), columna_estado=14)
    for fila in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        fila[0].number_format = "dd/mm/yyyy"

    # Hoja 2: clientes
    ws2 = wb.create_sheet("Clientes")
    ws2.append(["Identificación", "Nombre", "Apellido", "Teléfono", "Otro teléfono",
                "Correo", "Registrado el"])
    for c in clientes:
        ws2.append([c["identificacion"], c["nombre"], c["apellido"], c["telefono"],
                    c["otro_telefono"] or "No registrado", c["correo"] or "No registrado",
                    c["fecha_registro"]])
    dar_estilo_hoja(ws2)

    try:
        wb.save(ARCHIVO_EXCEL)
    except PermissionError:
        print("No se pudo guardar. Si el archivo está abierto en Excel, ciérralo e intenta de nuevo.")
        return
    activas = sum(1 for r in reservas if r["estado"] == "Activa")
    print(f"Archivo creado: {ARCHIVO_EXCEL}")
    print(f"Reservas exportadas: {len(reservas)} ({activas} activas).")


# ----------------------------------------------------------------
# OPCIÓN 5: CANCELAR RESERVA
# ----------------------------------------------------------------
def cancelar_reserva():
    titulo("CANCELAR RESERVA")

    identificacion = leer_identificacion("Identificación del cliente: ")
    cliente = buscar_cliente(identificacion)
    if cliente is None:
        print("\nEse cliente no está registrado.")
        return

    activas = [r for r in reservas if r["identificacion"] == identificacion and r["estado"] == "Activa"]
    if len(activas) == 0:
        print(f"\n{nombre_completo(cliente)} no tiene reservas activas.")
        return

    print(f"\nReservas activas de {nombre_completo(cliente)}:")
    print("-" * 80)
    print(f"{'Código':<9}{'Fecha':<13}{'Hora':<8}{'Mesa':<7}{'Personas':<10}A nombre de")
    print("-" * 80)
    for r in activas:
        print(f"{r['codigo']:<9}{r['fecha']:<13}{r['hora']:<8}{r['mesa']:<7}{r['personas']:<10}{r['a_nombre_de']}")
    print("-" * 80)

    while True:
        codigo = input("Código de la reserva a cancelar (Enter para salir): ").strip().upper()
        if codigo == "":
            print("No se canceló ninguna reserva.")
            return
        seleccion = next((r for r in activas if r["codigo"] == codigo), None)
        if seleccion:
            break
        print("Ese código no corresponde a una reserva activa de este cliente.")

    if not leer_si_no(f"¿Seguro que deseas cancelar la reserva {codigo}? (s/n): "):
        print("No se canceló la reserva.")
        return

    seleccion["estado"] = "Cancelada"
    seleccion["fecha_cancelacion"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    guardar_datos()
    print(f"\nReserva {codigo} cancelada. La mesa {seleccion['mesa']} quedó libre "
          f"el {seleccion['fecha']} a las {seleccion['hora']}.")


# ----------------------------------------------------------------
# OPCIÓN 6: VER RESERVAS
# ----------------------------------------------------------------
def ver_reservas():
    titulo("RESERVAS")
    if len(reservas) == 0:
        print("Aún no hay reservas.")
        return

    print(f"{'Código':<9}{'Fecha':<13}{'Hora':<8}{'Mesa':<6}{'Pers.':<7}{'A nombre de':<24}{'Identificación':<16}Estado")
    print("-" * 95)
    for r in sorted(reservas, key=lambda x: inicio_reserva(x["fecha"], x["hora"])):
        print(f"{r['codigo']:<9}{r['fecha']:<13}{r['hora']:<8}{r['mesa']:<6}{r['personas']:<7}"
              f"{r['a_nombre_de']:<24}{r['identificacion']:<16}{r['estado']}")
    activas = sum(1 for r in reservas if r["estado"] == "Activa")
    print("-" * 95)
    print(f"Total: {len(reservas)} reserva(s), {activas} activa(s).")


# ----------------------------------------------------------------
# OPCIÓN 7: VER CLIENTES
# ----------------------------------------------------------------
def ver_clientes():
    titulo("CLIENTES REGISTRADOS")
    if len(clientes) == 0:
        print("Aún no hay clientes registrados.")
        return

    print(f"{'Identificación':<16}{'Nombre':<26}{'Teléfono':<13}{'Otro teléfono':<15}Correo")
    print("-" * 95)
    for c in clientes:
        print(f"{c['identificacion']:<16}{nombre_completo(c):<26}{c['telefono']:<13}"
              f"{(c['otro_telefono'] or '-'):<15}{c['correo'] or '-'}")
    print("-" * 95)
    print(f"Total: {len(clientes)} cliente(s).")


# ----------------------------------------------------------------
# MENÚ PRINCIPAL
# ----------------------------------------------------------------
def mostrar_menu():
    print("\n" + "=" * 40)
    print("        RESTAURANTE - RESERVAS")
    print("=" * 40)
    print("1. Registrar cliente")
    print("2. Ver mesas")
    print("3. Crear reserva")
    print("4. Exportar información a Excel")
    print("5. Cancelar reserva")
    print("6. Ver reservas")
    print("7. Ver clientes")
    print("8. Salir")


def main():
    cargar_datos()
    acciones = {
        "1": registrar_cliente,
        "2": ver_mesas,
        "3": crear_reserva,
        "4": exportar_excel,
        "5": cancelar_reserva,
        "6": ver_reservas,
        "7": ver_clientes,
    }
    try:
        while True:
            mostrar_menu()
            opcion = input("Elige una opción: ").strip()
            if opcion == "8":
                print("¡Hasta pronto!")
                break
            if opcion in acciones:
                acciones[opcion]()
                pausar()
            else:
                print("Opción no válida. Elige un número del 1 al 8.")
    except (KeyboardInterrupt, EOFError):
        print("\nPrograma cerrado. ¡Hasta pronto!")


if __name__ == "__main__":
    main()