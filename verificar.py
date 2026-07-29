# -*- coding: utf-8 -*-
"""
Verificador del ciclo completo.

Paso 3:  se aplico el ISO sobre la base, se reexporto con exportar_cat.py,
y ahora se compara lo que quedo contra lo que se esperaba.

Es el paso que convierte "creo que funciono" en "esta probado". Mientras esta
comparacion no de limpia, no se toca la base del servidor.

Uso:
  python verificar.py esperado.csv obtenido.csv
      Compara emparejando por MFN. Es el caso normal: actualizacion
      selectiva sobre la base existente, donde los MFN se conservan.

  python verificar.py esperado.csv obtenido.csv --por-orden
      Compara emparejando por posicion. Se usa en la prueba de fidelidad,
      cuando el ISO se importo en un clon VACIO: ahi CISIS asigna MFN nuevos
      y correlativos, asi que los numeros no coinciden pero el orden si.

Los campos v9 (numero de control) y v4 (fecha de registro) se informan aparte:
son los que ABCD podria reescribir solo al importar, y saber si lo hace es
justamente uno de los objetivos de la prueba.
"""

import csv
import os
import sys
from collections import Counter

csv.field_size_limit(10 ** 7)

SEP_COLUMNAS    = ";"
SEP_OCURRENCIAS = "<|>"

# Campos bajo sospecha de ser reescritos por ABCD al importar.
SOSPECHOSOS = {
    "9": "v9 numero de control (lo asigna ABCD solo)",
    "4": "v4 fecha de registro (ABCD podria ponerle la fecha de hoy)",
}


def leer(ruta):
    if not os.path.exists(ruta):
        sys.exit("ERROR: no se encontro %s" % ruta)
    filas, columnas = [], None
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f, delimiter=SEP_COLUMNAS)
        encabezados = next(r)
        columnas = [(c.split("_")[0][1:], c) for c in encabezados[1:]]
        for fila in r:
            campos = {}
            for (tag, _), celda in zip(columnas, fila[1:]):
                if celda:
                    campos[tag] = celda.split(SEP_OCURRENCIAS)
            filas.append((fila[0], campos))
    return filas, dict(columnas)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 2:
        sys.exit(__doc__)
    por_orden = "--por-orden" in sys.argv

    esperado, columnas = leer(args[0])
    obtenido, _        = leer(args[1])

    print("Esperado : %s   (%d registros)" % (os.path.basename(args[0]), len(esperado)))
    print("Obtenido : %s   (%d registros)" % (os.path.basename(args[1]), len(obtenido)))
    print("Empareja : %s\n" % ("por posicion" if por_orden else "por MFN"))

    if len(esperado) != len(obtenido):
        print("AVISO: la cantidad de registros no coincide (%+d)."
              % (len(obtenido) - len(esperado)))

    # --- emparejar ---
    if por_orden:
        pares = list(zip(esperado, obtenido))
        sobran = abs(len(esperado) - len(obtenido))
        if sobran:
            print("       Se comparan los primeros %d.\n" % len(pares))
    else:
        dic_obt = dict(obtenido)
        pares = [((m, c), (m, dic_obt[m])) for m, c in esperado if m in dic_obt]
        perdidos = [m for m, _ in esperado if m not in dic_obt]
        if perdidos:
            print("AVISO: %d MFN del esperado no aparecen en el obtenido: %s\n"
                  % (len(perdidos), ", ".join(perdidos[:10])))

    # --- comparar ---
    difs_por_campo = Counter()
    sospechosos_vistos = Counter()
    ejemplos = []
    registros_con_dif = 0

    for (mfn_e, campos_e), (mfn_o, campos_o) in pares:
        distintos = []
        for tag in sorted(set(campos_e) | set(campos_o), key=int):
            a, b = campos_e.get(tag, []), campos_o.get(tag, [])
            if a != b:
                distintos.append((tag, a, b))
                difs_por_campo[tag] += 1
                if tag in SOSPECHOSOS:
                    sospechosos_vistos[tag] += 1
        if distintos:
            registros_con_dif += 1
            if len(ejemplos) < 15:
                ejemplos.append((mfn_e, mfn_o, distintos))

    # --- informe ---
    print("Registros comparados      : %d" % len(pares))
    print("Registros con diferencias : %d" % registros_con_dif)

    if sospechosos_vistos:
        print("\n*** CAMPOS REESCRITOS POR ABCD ***")
        for tag, n in sospechosos_vistos.items():
            print("  %-12s cambio en %d registros  -> %s"
                  % ("v" + tag, n, SOSPECHOSOS[tag]))
        print("  Esto responde una de las preguntas abiertas del proyecto.")
        print("  Ver la seccion 9 de ESPECIFICACION_CSV.md.")

    otros = {t: n for t, n in difs_por_campo.items() if t not in SOSPECHOSOS}
    if otros:
        print("\nDiferencias por campo (sin contar los sospechosos):")
        for tag, n in sorted(otros.items(), key=lambda x: -x[1]):
            print("  %-34s %d registros" % (columnas.get(tag, "v" + tag), n))

    if ejemplos:
        print("\nEjemplos:")
        for mfn_e, mfn_o, distintos in ejemplos:
            ref = "MFN %s" % mfn_e if mfn_e == mfn_o else "MFN %s -> %s" % (mfn_e, mfn_o)
            print("  %s" % ref)
            for tag, a, b in distintos[:4]:
                print("     %s" % columnas.get(tag, "v" + tag))
                print("        esperado : %s" % (SEP_OCURRENCIAS.join(a) or "(vacio)")[:100])
                print("        obtenido : %s" % (SEP_OCURRENCIAS.join(b) or "(vacio)")[:100])

    print()
    if registros_con_dif == 0:
        print("RESULTADO: el ciclo cierra. Lo que quedo en la base es lo que se esperaba.")
    else:
        print("RESULTADO: hay %d registros con diferencias. Revisar antes de tocar el "
              "servidor." % registros_con_dif)
        sys.exit(1)


if __name__ == "__main__":
    main()
