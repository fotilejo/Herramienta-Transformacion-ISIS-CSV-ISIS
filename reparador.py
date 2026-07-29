# -*- coding: utf-8 -*-
"""
REPARADOR CAT — menu unico.

Envuelve los 3 pasos de siempre (exportar_cat.py, generar_iso.py,
aplicar_cambios.py) en un solo lugar, sin cambiarles la logica: son los
mismos scripts ya probados, esto es solo una forma mas simple de llamarlos
y de saber en todo momento que archivo hay que tocar.

Uso: doble clic en REPARADOR_CAT.bat. No hace falta correr nada mas a mano.
"""

import glob
import os
import subprocess
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR  = os.path.join(BASE_DIR, "Salidas_CSV")
ISO_DIR  = os.path.join(BASE_DIR, "Salidas_ISO")
BAK_DIR  = os.path.join(BASE_DIR, "Respaldos")

ISIS_BASE_DEFAULT = r"C:\Isis\cat\data\cat"

# Uso normal: sin argumentos, el paso 1 exporta siempre ISIS_BASE_DEFAULT.
# Uso puntual (exportar un clon, una carpeta de prueba, etc.):
#   python reparador.py --base "C:\Isis\catprueba\catprueba"
BASE_OVERRIDE = None
if "--base" in sys.argv:
    BASE_OVERRIDE = sys.argv[sys.argv.index("--base") + 1]


def pausa():
    input("\nPresione Enter para volver al menu... ")


def correr(script, argumentos=None):
    """Corre uno de los scripts existentes tal cual, sin tocarles nada."""
    cmd = [sys.executable, os.path.join(BASE_DIR, script)] + (argumentos or [])
    return subprocess.run(cmd, cwd=BASE_DIR).returncode


def mas_reciente(patron):
    encontrados = sorted(glob.glob(patron), key=os.path.getmtime)
    return encontrados[-1] if encontrados else None


def abrir(ruta):
    """Abre el archivo con el programa asociado (Excel, en el caso del .xlsx)."""
    try:
        os.startfile(ruta)  # solo Windows, que es donde corre esto
    except Exception:
        pass  # si no se puede abrir solo, igual se imprimio la ruta


def pista_de_estado():
    """Una linea que sugiere en que paso quedaste la ultima vez, por si te
    perdiste. Es solo una ayuda, no hace falta que sea exacta."""
    editable = mas_reciente(os.path.join(CSV_DIR, "*_editable.xlsx"))
    cambios  = mas_reciente(os.path.join(ISO_DIR, "*_cambios.iso"))
    if not editable:
        return None
    t_editable = os.path.getmtime(editable)
    t_cambios  = os.path.getmtime(cambios) if cambios else 0
    if t_editable > t_cambios:
        return "  (Ultimo archivo preparado: %s -- si ya lo editaste, segui con el paso 2)" % os.path.basename(editable)
    return "  (Ya generaste el ISO de cambios: %s -- si todavia no lo aplicaste, segui con el paso 3)" % os.path.basename(cambios)


def paso_1_preparar():
    print("\n" + "=" * 70)
    print("  PASO 1 -- Preparar una correccion (ISIS -> Excel)")
    print("=" * 70)

    # Sin preguntas: siempre exporta la base CAT real, todas las columnas.
    # Ya no hace falta elegir campos a mano -- el .xlsx sale con las celdas
    # bloqueadas en Texto, asi que pedir todo de entrada es seguro.
    #
    # Para exportar otra base (un clon, una carpeta de prueba), correr
    # reparador.py con --base "<ruta sin .mst>" desde la linea de comandos:
    # ese caso puntual no es el uso normal, no hace falta un menu para el.
    if BASE_OVERRIDE:
        argumentos = ["--base", BASE_OVERRIDE, "--editar", ""]
        nombre = os.path.basename(BASE_OVERRIDE)
        print("\n  Base: %s (indicada por --base)" % BASE_OVERRIDE)
    else:
        argumentos = ["--editar", ""]
        nombre = "cat"
        print("\n  Base: %s" % ISIS_BASE_DEFAULT)

    codigo = correr("exportar_cat.py", argumentos)
    if codigo != 0:
        print("\n  Algo fallo al exportar. Revisa el mensaje de arriba.")
        pausa()
        return

    fecha = datetime.now().strftime("%Y%m%d")
    ruta_editable = os.path.join(CSV_DIR, "%s_%s_editable.xlsx" % (nombre, fecha))
    ruta_base     = os.path.join(CSV_DIR, "%s_%s_base.csv" % (nombre, fecha))

    print("\n" + "-" * 70)
    if os.path.exists(ruta_editable):
        print("  Listo. El archivo para EDITAR es este:")
        print("    %s" % ruta_editable)
        print("\n  (El otro archivo que se genero, %s, es la linea base:"
              % os.path.basename(ruta_base))
        print("   NO se toca, es solo para que la herramienta compare despues.)")
        print("\n  Abriendolo en Excel...")
        abrir(ruta_editable)
    else:
        print("  No encontre el archivo esperado (%s)." % os.path.basename(ruta_editable))
        print("  Revisa Salidas_CSV a mano.")
    print("-" * 70)
    print("\n  Corregi lo que haga falta, GUARDA y CERRA el Excel, y volve a este")
    print("  menu para el paso 2.")
    pausa()


def paso_2_generar():
    print("\n" + "=" * 70)
    print("  PASO 2 -- Generar el archivo de cambios")
    print("=" * 70)
    print("\n  Esto compara lo que editaste contra la linea base y arma el ISO")
    print("  con solo los registros que cambiaron.\n")
    correr("generar_iso.py")
    print("\n  Leé el informe (Salidas_ISO, el .txt mas reciente) antes de seguir.")
    print("  Tiene que decir exactamente lo que corregiste, ni un registro de mas.")
    pausa()


def paso_3_aplicar():
    print("\n" + "=" * 70)
    print("  PASO 3 -- Aplicar los cambios a la base")
    print("=" * 70)
    print()
    correr("aplicar_cambios.py")
    pausa()


def main():
    while True:
        os.system("cls")
        print("=" * 70)
        print("  REPARADOR CAT")
        print("=" * 70)
        pista = pista_de_estado()
        if pista:
            print(pista)
        print("""
  Que queres hacer?

    1. Preparar una correccion nueva      (exportar de ISIS a Excel)
    2. Ya edite el Excel -> generar los cambios
    3. Aplicar los cambios a la base
    4. Salir
""")
        opcion = input("  Opcion: ").strip()
        if opcion == "1":
            paso_1_preparar()
        elif opcion == "2":
            paso_2_generar()
        elif opcion == "3":
            paso_3_aplicar()
        elif opcion == "4":
            break
        else:
            print("\n  Opcion invalida.")
            pausa()


if __name__ == "__main__":
    main()
