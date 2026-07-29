# -*- coding: utf-8 -*-
"""
Reexporta la base clonada y la compara contra la base real. Todo en un paso.

Es lo que hace falta para cerrar la prueba de fidelidad (fase A de
PROCEDIMIENTO.md): despues de importar el ISO en el clon, hay que sacarle un
CSV y ver si coincide con el original.

Se encarga de:
  1. Ubicar el .mst del clon (lo busca solo; si no lo encuentra, lo pregunta).
  2. Correr exportar_cat.py contra esa base.
  3. Correr verificar.py contra el CSV de la base real.

No toca ninguna base: solo lee.

Uso: doble clic en 03_Verificar_Clon.bat  (o python verificar_clon.py)
"""

import glob
import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR  = os.path.join(BASE_DIR, "Salidas_CSV")

NOMBRE_CLON = "catprueba"

# Lugares donde suele quedar una base de ABCD, en orden de probabilidad.
RAICES = [
    r"C:\ABCD", r"C:\abcd", r"D:\ABCD", r"D:\abcd",
    r"C:\Program Files\ABCD", r"C:\Program Files (x86)\ABCD",
    r"C:\Isis", r"C:\inetpub\wwwroot\ABCD",
]


def buscar_mst(nombre):
    """Devuelve la ruta al .mst del clon, o None."""
    objetivo = nombre.lower() + ".mst"

    # 1) rutas tipicas de ABCD, directo
    for raiz in RAICES:
        directa = os.path.join(raiz, "www", "bases", nombre, "data", nombre + ".mst")
        if os.path.exists(directa):
            return directa

    # 2) recorrido acotado de esas mismas raices
    for raiz in RAICES:
        if not os.path.isdir(raiz):
            continue
        for carpeta, _, archivos in os.walk(raiz):
            for a in archivos:
                if a.lower() == objetivo:
                    return os.path.join(carpeta, a)
    return None


def correr(titulo, argumentos):
    print("\n" + "=" * 70)
    print(titulo)
    print("=" * 70)
    r = subprocess.run([sys.executable] + argumentos, cwd=BASE_DIR)
    return r.returncode


def csv_mas_reciente(prefijo):
    encontrados = sorted(glob.glob(os.path.join(CSV_DIR, "%s_*_base.csv" % prefijo)))
    return encontrados[-1] if encontrados else None


def main():
    print("Verificacion de la base clonada")
    print("-" * 70)

    nombre = NOMBRE_CLON
    if len(sys.argv) > 1:
        nombre = sys.argv[1]

    print("Buscando %s.mst ..." % nombre)
    mst = buscar_mst(nombre)

    if not mst:
        print("\nNo lo encontre solo.")
        print("Abri el Explorador, ubica el archivo %s.mst, copia su ruta" % nombre)
        print("y pegala aca (con boton derecho -> Pegar). Enter vacio para salir.")
        pegada = input("\nRuta al .mst: ").strip().strip('"')
        if not pegada:
            return 1
        if not os.path.exists(pegada):
            print("\nEsa ruta no existe: %s" % pegada)
            return 1
        mst = pegada

    base_isis = mst[:-4] if mst.lower().endswith(".mst") else mst
    print("  encontrada: %s" % mst)

    original = csv_mas_reciente("cat")
    if not original:
        print("\nERROR: no hay ningun cat_*_base.csv en Salidas_CSV.")
        print("Corre antes 01_Exportar_CAT.bat.")
        return 1
    print("  CSV de la base real: %s" % os.path.basename(original))

    if correr("PASO 1 de 2  -  Exportando el clon a CSV", ["exportar_cat.py", "--base", base_isis]):
        print("\nLa exportacion del clon fallo. No se puede comparar.")
        return 1

    clon = csv_mas_reciente(os.path.basename(base_isis))
    if not clon:
        print("\nERROR: la exportacion no dejo ningun CSV del clon en Salidas_CSV.")
        return 1

    correr("PASO 2 de 2  -  Comparando el clon contra la base real",
           ["verificar.py", original, clon, "--por-orden"])

    print("\n" + "=" * 70)
    print("Como leer el resultado")
    print("=" * 70)
    print("  'el ciclo cierra'                -> listo, se puede empezar a editar.")
    print("  '*** CAMPOS REESCRITOS POR ABCD' -> ABCD reasigna v9 o pisa v4 al")
    print("                                      importar. Define como aplicar los")
    print("                                      cambios en produccion.")
    print("  diferencias en otros campos      -> es un defecto del conversor.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
