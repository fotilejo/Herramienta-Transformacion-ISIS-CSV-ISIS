# -*- coding: utf-8 -*-
"""
Aplica el ISO de cambios sobre una base ISIS, conservando los MFN.

Paso 3 del ciclo:  CAT (ISIS) -> CSV -> edicion manual -> CAT (ISIS)

Hace todo solo:
  1. Copia de seguridad del .mst y el .xrf.
  2. Aplica los cambios con mx (updatf), escribiendo cada registro en su MFN.
  3. Reexporta la base y verifica que quedo lo que se esperaba.
  4. Si algo no cierra, avisa y deja la copia de seguridad a mano.

No renumera nada: los registros que no se tocaron quedan intactos.

Uso: doble clic en 04_Aplicar_Cambios.bat
"""

import csv
import glob
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_DIR  = os.path.join(BASE_DIR, "Salidas_CSV")
ISO_DIR  = os.path.join(BASE_DIR, "Salidas_ISO")
BAK_DIR  = os.path.join(BASE_DIR, "Respaldos")
LOG_PATH = os.path.join(BASE_DIR, "log_aplicar_cambios.txt")


class _Tee:
    """Duplica todo lo que se imprime en pantalla hacia un archivo de registro.

    Sirve para no perder el detalle de lo que paso si se cierra la ventana
    antes de copiar la salida: queda guardado en log_aplicar_cambios.txt,
    que se puede revisar despues sin depender de la consola.
    """
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            s.write(data)
            s.flush()

    def flush(self):
        for s in self.streams:
            s.flush()

MX_EXE = r"C:\Isis\mx.exe"

# Base sobre la que se aplican los cambios. Poner la COPIA LOCAL, nunca la del
# servidor: el flujo es aplicar aca, verificar, y recien despues copiar los
# archivos al servidor a mano.
BASE_DESTINO = r"C:\Isis\cat\data\cat"

# Extensiones que forman una base ISIS. Se respaldan las dos primeras (son los
# datos); las demas son indices y se regeneran.
EXT_DATOS = [".mst", ".xrf"]


def ultimo(patron):
    encontrados = sorted(glob.glob(patron))
    return encontrados[-1] if encontrados else None


def leer_informe(ruta):
    """Extrae del informe .txt los cambios que se pidieron aplicar:
    {(mfn, columna): valor_esperado}.

    Verificar contra esto (en vez de contra el CSV editable) evita el falso
    positivo de comparar un archivo angosto de 2 columnas contra el CSV
    completo reexportado: ahi todo lo que no esta en el angosto aparece
    como "diferencia" aunque no se haya tocado.
    """
    esperado = {}
    mfn = col = None
    with open(ruta, encoding="utf-8") as f:
        for linea in f:
            if linea.startswith("MFN "):
                mfn = linea.strip().split()[1]
                col = None
            elif linea.startswith("   ") and not linea.startswith("      "):
                col = linea.strip()
            elif linea.strip().startswith("ahora :") and mfn and col:
                valor = linea.split("ahora :", 1)[1].strip()
                if valor == "(vacio)":
                    valor = ""
                esperado[(mfn, col)] = valor
    return esperado


def leer_valores_csv(ruta):
    """Devuelve {(mfn, columna): valor} para poder buscar celdas puntuales."""
    valores = {}
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f, delimiter=";")
        encabezados = next(r)
        for fila in r:
            mfn = fila[0].strip().zfill(6) if fila[0].strip().isdigit() else fila[0].strip()
            for col, celda in zip(encabezados[1:], fila[1:]):
                valores[(mfn, col)] = celda
    return valores


def respaldar(base):
    os.makedirs(BAK_DIR, exist_ok=True)
    sello = datetime.now().strftime("%Y%m%d_%H%M")
    destino = os.path.join(BAK_DIR, "%s_%s" % (os.path.basename(base), sello))
    os.makedirs(destino, exist_ok=True)
    for ext in EXT_DATOS:
        origen = base + ext
        if os.path.exists(origen):
            shutil.copy2(origen, destino)
    return destino


def main():
    log_file = open(LOG_PATH, "a", encoding="utf-8")
    log_file.write("\n\n" + "#" * 70 + "\n")
    log_file.write("# Corrida: %s\n" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log_file.write("#" * 70 + "\n")
    sys.stdout = _Tee(sys.stdout, log_file)

    print("=" * 70)
    print("  APLICAR CAMBIOS SOBRE LA BASE ISIS")
    print("=" * 70)

    iso = ultimo(os.path.join(ISO_DIR, "*_cambios.iso"))
    if not iso:
        print("\nNo hay ningun ISO de cambios en Salidas_ISO.")
        print("Corre antes 02_Generar_ISO.bat.")
        return 1

    informe = iso.replace(".iso", ".txt")

    # --forzar-limpieza-401-110: uso excepcional. v401 y v110 no son campos
    # de la FDT de CAT; mx los VA SUMANDO en vez de reemplazarlos, mande lo
    # que se mande (probado: ni omitirlos ni mandar el valor correcto evita
    # el duplicado). La unica forma de dejar una sola ocurrencia es borrar
    # con 'D401D110' antes de escribir.
    #
    # Esto NO se puede dejar prendido siempre: v110 en particular a veces
    # tiene datos reales (por ejemplo una URL, comprobado en el MFN 046015)
    # y no un sello de sistema, asi que borrarlo en TODOS los registros de
    # una tanda cualquiera podria destruir contenido real. Solo es seguro
    # cuando el ISO de esa corrida trae UNICAMENTE los pocos MFN ya
    # revisados a mano que de verdad tienen el duplicado.
    limpiar_401_110 = "--forzar-limpieza-401-110" in sys.argv

    print("\n  Sobre que base se aplican los cambios?")
    print("    1. La copia local de CAT   (%s)" % BASE_DESTINO)
    print("    2. La base de prueba catprueba")
    print("    3. Otra (se pide la ruta)")
    opcion = input("\n  Opcion [1/2/3]: ").strip()

    if opcion == "2":
        from verificar_clon import buscar_mst
        mst = buscar_mst("catprueba")
        if not mst:
            mst = input("  No la encontre. Pegar la ruta a catprueba.mst: ").strip().strip('"')
        base = mst[:-4] if mst.lower().endswith(".mst") else mst
    elif opcion == "3":
        r = input("  Ruta al .mst: ").strip().strip('"')
        base = r[:-4] if r.lower().endswith(".mst") else r
    else:
        base = BASE_DESTINO

    if not os.path.exists(base + ".mst"):
        print("\nNo encontre la base en %s.mst" % base)
        print("Editar BASE_DESTINO en aplicar_cambios.py.")
        return 1
    if not os.path.exists(MX_EXE):
        print("\nNo encontre mx.exe en %s" % MX_EXE)
        return 1

    print("\n  ISO de cambios : %s" % os.path.basename(iso))
    print("  Base destino   : %s.mst" % base)
    if os.path.exists(informe):
        with open(informe, encoding="utf-8") as f:
            for linea in f:
                if linea.startswith("Registros modificados"):
                    print("  %s" % linea.strip())
                    break
        print("\n  El detalle de los cambios esta en:")
        print("    %s" % informe)

    print("\n  Se hara una copia de seguridad antes de tocar nada.")
    if input("\n  Escribir SI para continuar: ").strip().upper() != "SI":
        print("\n  Cancelado. No se toco nada.")
        return 0

    print("\n[1/3] Copia de seguridad...")
    bak = respaldar(base)
    print("      guardada en %s" % bak)

    print("\n[2/3] Aplicando los cambios...")
    # proc: '=' + v900 fija el MFN del registro entrante; D900 borra el campo
    # auxiliar para que no quede en la base. updatf escribe sobre la base
    # existente, en la posicion indicada, sin renumerar nada.
    directiva = "'D900D401D110'" if limpiar_401_110 else "'D900'"
    if limpiar_401_110:
        print("      (--forzar-limpieza-401-110: se borran v401 y v110 antes de escribir,")
        print("       SOLO en los MFN de este ISO puntual)")
    cmd = [MX_EXE, "iso=" + iso, "proc='=',v900," + directiva,
           "updatf=" + base, "-all", "now"]
    print("      " + " ".join(cmd))
    r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    salida = r.stdout.decode("cp1252", "replace")
    if salida.strip():
        print("      " + salida.strip()[:600].replace("\n", "\n      "))
    if r.returncode != 0:
        print("\n  mx devolvio un error (codigo %d)." % r.returncode)
        print("  La base NO se modifico, o quedo a medias.")
        print("  Restaurar copiando los archivos de:")
        print("    %s" % bak)
        return 1

    print("\n[3/3] Verificando...")
    cambios_esperados = leer_informe(informe) if os.path.exists(informe) else {}
    if not cambios_esperados:
        print("\n  No se pudo leer el informe de cambios (%s)." % informe)
        print("  No se puede confirmar el resultado, pero la base ya fue tocada.")
        return 1

    # --editar "" evita que exportar_cat.py se quede esperando que elijamos
    # columnas por teclado: no nos interesa el CSV angosto aca, solo el
    # completo, que se escribe siempre.
    r = subprocess.run([sys.executable, "exportar_cat.py", "--base", base, "--editar", ""],
                       cwd=BASE_DIR)
    if r.returncode != 0:
        print("\n  No se pudo reexportar para verificar.")
        return 1

    obtenido = ultimo(os.path.join(CSV_DIR, "%s_*_base.csv" % os.path.basename(base)))
    if not obtenido:
        print("\n  No encontre el CSV reexportado para comparar.")
        return 1

    valores = leer_valores_csv(obtenido)

    # Se compara pedido por pedido contra el informe: solo los campos que
    # realmente se mandaron a cambiar. Esto no se confunde si el editable
    # que se uso para generar el ISO era una proyeccion angosta (2 o 3
    # columnas) en vez del CSV completo.
    ok, mal = [], []
    for (mfn, col), esperado_val in sorted(cambios_esperados.items()):
        real_val = valores.get((mfn, col))
        if real_val is None:
            mal.append((mfn, col, esperado_val, "(columna no encontrada)"))
        elif real_val == esperado_val:
            ok.append((mfn, col))
        else:
            mal.append((mfn, col, esperado_val, real_val))

    print("\n  Cambios pedidos   : %d" % len(cambios_esperados))
    print("  Confirmados en la base : %d" % len(ok))
    for mfn, col in ok:
        print("    MFN %s  %-25s OK" % (mfn, col))

    print("\n" + "=" * 70)
    if not mal:
        print("  LISTO. Los %d cambios quedaron en la base, tal como se pidio." % len(ok))
        print("  El resto de los registros no se toco (mismo mecanismo: updatf")
        print("  solo escribe en el MFN indicado, nunca reconstruye la base entera).")
        print("\n  Ya se puede copiar %s.mst y %s.xrf al servidor"
              % (os.path.basename(base), os.path.basename(base)))
        print("  y reconstruir los indices ahi.")
    else:
        print("  HAY %d CAMBIOS QUE NO QUEDARON COMO SE ESPERABA:" % len(mal))
        for mfn, col, esperado_val, real_val in mal:
            print("    MFN %s  %s" % (mfn, col))
            print("      se esperaba : %s" % (esperado_val or "(vacio)"))
            print("      quedo       : %s" % (real_val or "(vacio)"))
        print("\n  No copiar nada al servidor todavia.")
        print("  Para volver atras, restaurar los archivos de:")
        print("    %s" % bak)
    print("=" * 70)
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())
