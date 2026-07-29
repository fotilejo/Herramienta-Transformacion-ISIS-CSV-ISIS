# -*- coding: utf-8 -*-
"""
Exportador fiel de la base CAT (ISIS) a CSV.

Paso 1 del ciclo:  CAT (ISIS) -> CSV -> edicion manual -> CAT (ISIS)

Lee el ISO 2709 que produce mx, no un PFT. La diferencia importa:

  - Un PFT obliga a enumerar las etiquetas a mano. La FDT de CAT declara 53,
    pero en los datos hay 66: v21, v60, v76, v78, v80, v95, v102, v103, v110,
    v130, v131, v327 y v401 existen y no estan declaradas. Con PFT se perdian
    2.450 ocurrencias en cada corrida.
  - Hay 77 registros donde mx, consultado por PFT, devuelve vacio para campos
    que el ISO si trae. El caso testigo es el MFN 050356: tiene v10='true' y
    dos v80, con el directorio empezando por 080,080,010 y sin v3. El ISO los
    muestra; 'if p(v10)' no.

El ISO vuelca el directorio completo, asi que no depende de ninguna lista.

Uso:
  python exportar_cat.py                       exporta desde la base de ISIS_BASE
  python exportar_cat.py --base C:\\ruta\\clon   exporta desde otra base (sin el .mst)
  python exportar_cat.py --iso arch.iso        parsea un ISO ya generado (debug)

Salida: dos archivos en Salidas_CSV
  cat_<fecha>_base.csv        linea base intocable, para el diff posterior
  cat_<fecha>_editable.xlsx   el que se edita (Excel, celdas bloqueadas en Texto)
"""

import csv
import os
import re
import subprocess
import sys
from collections import Counter
from datetime import datetime

from openpyxl import Workbook

# =========================================================
# CONFIGURACION
# =========================================================
MX_EXE    = r"C:\Isis\mx.exe"
ISIS_BASE = r"C:\Isis\cat\data\cat"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR  = os.path.join(BASE_DIR, "Salidas_CSV")
TMP_DIR  = os.path.join(BASE_DIR, "temp")

SEP_COLUMNAS    = ";"
SEP_OCURRENCIAS = "<|>"
# Caracteres de control que aparecen DENTRO del contenido de los campos.
# En CAT hay 684 ocurrencias: 598 con un CR suelto pegado al final del texto
# ('Quimica\r'), 79 con CR LF y 7 con LF suelto. Es basura de carga que conviene
# limpiar, pero el exportador no la toca: la representa con tokens distintos
# para que el viaje de vuelta reponga exactamente los mismos bytes.
CONTROLES = [("\r\n", "<NL>"), ("\r", "<CR>"), ("\n", "<LF>")]

# Estructura del ISO que escribe CISIS. Verificado contra un ISO real de 43 MB.
ANCHO_LINEA = 80          # cada registro se parte en lineas de 80 + CRLF
FIN_LINEA   = b"\r\n"
FT          = b"#"        # fin de campo   (CISIS no usa 0x1E)
RT          = b"#"        # fin de registro (CISIS no usa 0x1D)

# El lider de CISIS tiene un patron fijo que sirve para resincronizar:
#   5 digitos (largo) | 00000 | 00 | 5 digitos (base) | 000 | 4500
PATRON_LIDER = re.compile(rb"(?=(\d{5}0000000\d{5}0004500))")

# mx escribe el ISO en modo texto: un salto de linea guardado DENTRO de un campo
# (que en el .mst es CR LF) sale al archivo como CR CR LF, porque el LF se
# expande a CR LF. Los CRLF de envoltura de linea salen iguales. Para poder
# distinguirlos, antes de desenvolver se reemplaza la secuencia CR CR LF por una
# marca de 2 bytes: asi conserva el largo que declara el lider y todos los CRLF
# que quedan en el archivo son, con certeza, envoltura.
MARCA_SALTO = b"\x01\x02"

# Restos de CP850 (ISIS de la epoca DOS). Son los 5 bytes que cp1252 deja sin
# definir, por eso se detectan sin ambiguedad. Los 5 tienen equivalente cp1252.
CP850_REMEDIOS = {0x81: "ü", 0x8D: "ì", 0x8F: "Å", 0x90: "É", 0x9D: "Ø"}

# Nombres de las 53 etiquetas declaradas en cat.fdt, en el orden de la FDT.
# Las que aparecen en los datos y no estan aca se agregan solas, al final.
NOMBRES_FDT = [
    ("3",   "v3_Fecha_de_ingreso"),            ("4",   "v4_Fecha_de_registro"),
    ("2",   "v2_Fuente_del_registro"),         ("5",   "v5_Nivel_bibliografico"),
    ("6",   "v6_Nivel_de_referencia"),         ("7",   "v7_Tipo_de_documento"),
    ("8",   "v8_Soporte_no_convencional"),     ("9",   "v9_Numero_de_control"),
    ("10",  "v10_ISBN_Monografia"),            ("11",  "v11_ISBN_Serie_o_coleccion"),
    ("12",  "v12_Codigo_de_doc"),              ("14",  "v14_Codigo_identif_de_doc_esp"),
    ("15",  "v15_ISSN"),                       ("18",  "v18_Proy_progr_u_otro_encuadre"),
    ("19",  "v19_Otros_codigos"),              ("20",  "v20_Titulo"),
    ("22",  "v22_Autor_personal"),             ("23",  "v23_Autor_institucional"),
    ("25",  "v25_Colaborador"),                ("26",  "v26_Afiliacion_del_autor"),
    ("29",  "v29_Reunion_Responsable"),        ("30",  "v30_Reunion_Nombre_y_nro"),
    ("31",  "v31_Reunion_Lugar_pais"),         ("32",  "v32_Reunion_Fecha"),
    ("33",  "v33_Titulo_de_serie_o_coleccion"),("39",  "v39_Autor_institucional_ser_col"),
    ("40",  "v40_Idioma"),                     ("42",  "v42_Edicion"),
    ("43",  "v43_Lugar_de_edicion"),           ("44",  "v44_Pais_de_edicion"),
    ("45",  "v45_Editor"),                     ("46",  "v46_Fecha_de_edicion"),
    ("50",  "v50_Colacion"),                   ("53",  "v53_Tesis_Inst_grado_fecha"),
    ("54",  "v54_Notas"),                      ("55",  "v55_Unidad_Tecnica"),
    ("56",  "v56_Resumen"),                    ("65",  "v65_Descriptores"),
    ("75",  "v75_Signatura_topografica"),      ("77",  "v77_Inventario"),
    ("81",  "v81_Tipo_res_cientif"),           ("82",  "v82_Version"),
    ("84",  "v84_Formato"),                    ("85",  "v85_Disponibilidad"),
    ("86",  "v86_Nivel_acceso"),               ("88",  "v88_Licencia"),
    ("90",  "v90_Exportable_SNRD"),            ("100", "v100_Codigo_bd"),
    ("120", "v120_Enviado_a_BVSDE"),           ("140", "v140_URL"),
    ("150", "v150_Direccion"),                 ("200", "v200_Catalogador"),
    ("210", "v210_Revisado"),
]
# Prefijo de los archivos de salida. Cambia si se exporta desde otra base.
ETIQUETA = [None]

ORDEN_FDT   = [t for t, _ in NOMBRES_FDT]
NOMBRE_TAG  = dict(NOMBRES_FDT)


# =========================================================
# 1. VOLCADO DESDE ISIS
# =========================================================
def correr_mx(isis_base, args, salida):
    proc = subprocess.run([MX_EXE, isis_base] + args + ["now"],
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if proc.returncode != 0:
        sys.exit("ERROR: mx devolvio codigo %d\n%s"
                 % (proc.returncode, proc.stderr.decode("cp1252", "replace")))
    if salida:
        return proc.stdout
    return None


def volcar(isis_base):
    """Genera el ISO y la lista de MFN. Devuelve (bytes_iso, [mfn, ...]).

    Hacen falta dos pasadas porque el ISO 2709 no tiene lugar para el numero de
    registro. Las dos recorren la base en el mismo orden y saltean los borrados
    logicos, asi que las listas se corresponden posicion a posicion. Si las
    cantidades no coinciden el proceso se cancela.
    """
    for ruta, que in ((MX_EXE, "mx.exe"), (isis_base + ".mst", "la base")):
        if not os.path.exists(ruta):
            sys.exit("ERROR: no se encontro %s en %s" % (que, ruta))

    os.makedirs(TMP_DIR, exist_ok=True)
    ruta_iso = os.path.join(TMP_DIR, "volcado_cat.iso")

    print("Volcando %s ..." % isis_base)
    correr_mx(isis_base, ["iso=" + ruta_iso], salida=False)
    with open(ruta_iso, "rb") as f:
        datos = f.read()
    print("  ISO: %.1f MB en %s" % (len(datos) / 1048576, ruta_iso))

    crudo = correr_mx(isis_base, ["pft=mfn/", "lw=40"], salida=True)
    mfns  = [l.strip().decode("ascii", "replace")
             for l in crudo.split(b"\n") if l.strip().isdigit()]
    print("  MFN listados: %d" % len(mfns))
    return datos, mfns


# =========================================================
# 2. CODIFICACION
# =========================================================
def remediar_cp850(datos):
    """Corrige los restos de CP850. Devuelve (bytes, [lineas afectadas])."""
    presentes = [b for b in CP850_REMEDIOS if bytes([b]) in datos]
    if not presentes:
        return datos, []
    afectados = [l.decode("cp850", "replace")[:90]
                 for l in datos.split(FIN_LINEA)
                 if any(bytes([b]) in l for b in presentes)]
    for b, ch in CP850_REMEDIOS.items():
        datos = datos.replace(bytes([b]), ch.encode("cp1252"))
    return datos, afectados


# =========================================================
# 3. PARSEO DEL ISO
# =========================================================
def desenvolver(datos, ini, largo):
    """Reconstruye un registro salteando los CRLF de envoltura de las lineas."""
    buf, pos, resta = bytearray(), ini, largo
    while resta > 0:
        n = min(ANCHO_LINEA, resta)
        buf += datos[pos:pos + n]
        pos += n
        resta -= n
        if datos[pos:pos + 2] == FIN_LINEA:
            pos += 2
    return bytes(buf)


def _registro_sano(reg, largo):
    """Un registro bien desenvuelto cierra en RT y tiene el directorio entero."""
    try:
        if len(reg) != largo or reg[-1:] != RT:
            return False
        base  = int(reg[12:17])
        direc = reg[24:base - 1]
        if base > largo or len(direc) % 12 or not direc:
            return False
        fin = 0
        for k in range(0, len(direc), 12):
            e = direc[k:k + 12]
            if not e.isdigit():
                return False
            fin = max(fin, int(e[7:12]) + int(e[3:7]))
        # largo = base + area de datos + 1 byte del terminador de registro
        return base + fin == largo - 1
    except (ValueError, IndexError):
        return False


def _proporcion_valida(datos, muestra=400):
    """Fraccion de registros que cierran bien, sobre una muestra."""
    ok = tot = 0
    for m in PATRON_LIDER.finditer(datos):
        try:
            ini   = m.start()
            largo = int(datos[ini:ini + 5])
            if _registro_sano(desenvolver(datos, ini, largo), largo):
                ok += 1
        except (ValueError, IndexError):
            pass
        tot += 1
        if tot >= muestra:
            break
    return ok / tot if tot else 0.0


def parsear_iso(datos):
    """Devuelve (lista de {tag: [ocurrencias]}, descartados).

    Se localizan los registros por el patron del lider en vez de recorrer el
    archivo de corrido. mx escribe el ISO en modo texto y le agrega un \\r a los
    campos que tienen un salto de linea adentro, lo que desincroniza la lectura
    secuencial; buscar el patron permite retomar en el registro siguiente.
    """
    if MARCA_SALTO in datos:
        sys.exit("ERROR: el ISO ya contiene la marca interna %r. Elegir otra en "
                 "MARCA_SALTO." % MARCA_SALTO)

    # El arreglo del modo texto solo corresponde a los ISO que escribe mx. Los
    # que genera generar_iso.py salen en binario y no lo necesitan: aplicarselo
    # les comeria un byte por cada CR seguido de un corte de linea. Se decide
    # probando: si sin el arreglo casi todos los registros cierran bien, no va.
    if _proporcion_valida(datos) < 0.99:
        corregido = datos.replace(b"\r\r\n", MARCA_SALTO)
        if _proporcion_valida(corregido) > _proporcion_valida(datos):
            datos = corregido

    # Se recorre en secuencia y se retoma con el patron del lider cuando un
    # registro no cierra. Cada tramo ilegible cuenta como UN registro, para que
    # la correspondencia posicional con la lista de MFN no se desplace.
    inicios = [m.start() for m in PATRON_LIDER.finditer(datos)]
    registros, descartados = [], []
    for idx, ini in enumerate(inicios):
        try:
            largo = int(datos[ini:ini + 5])
            reg   = desenvolver(datos, ini, largo)
            if not _registro_sano(reg, largo):
                descartados.append(idx)
                registros.append(None)
                continue
            base  = int(reg[12:17])
            direc = reg[24:base - 1]
            campos = {}
            for k in range(0, len(direc), 12):
                e   = direc[k:k + 12]
                tag = str(int(e[0:3]))
                ln  = int(e[3:7])
                st  = int(e[7:12])
                if ln <= 1:
                    continue          # ocurrencia sin contenido: en ISIS no existe
                val = reg[base + st: base + st + ln - 1]
                val = val.replace(MARCA_SALTO, b"\r\n").decode("cp1252", "replace")
                for crudo, token in CONTROLES:
                    val = val.replace(crudo, token)
                campos.setdefault(tag, []).append(val)
            registros.append(campos)
        except (ValueError, IndexError):
            descartados.append(idx)
            registros.append(None)
    return registros, descartados


# =========================================================
# 4. COLUMNAS Y ESCRITURA
# =========================================================
def armar_columnas(registros):
    """Orden de la FDT primero; lo que aparezca de mas, al final por numero.

    v900 se excluye siempre: no es un dato de catalogacion, es un artefacto
    de esta misma herramienta (el carril donde generar_iso.py guarda el MFN
    dentro del ISO). Si esta base ya paso antes por una carga nuestra, ese
    v900 puede haber quedado escrito adentro de los datos reales -- y si se
    lo deja pasar como columna editable, la proxima vuelta por generar_iso.py
    termina escribiendo DOS v900 en el mismo registro (el que se inyecta de
    nuevo mas el que quedo de la carga anterior). Eso es lo que produce
    'fatal: recupdat/mfn' en mx: el proc no sabe cual de los dos usar.
    """
    presentes = {t for r in registros for t in r}
    presentes.discard("900")
    extras    = sorted(presentes - set(ORDEN_FDT), key=int)
    tags      = ORDEN_FDT + extras
    columnas  = [(t, NOMBRE_TAG.get(t, "v%s_no_declarado" % t)) for t in tags]
    return columnas, extras


def escribir_csv(ruta, mfns, registros, columnas, solo=None):
    """Si 'solo' trae una lista de tags, escribe unicamente esas columnas.

    Sirve para editar sin que la planilla rompa el resto: Excel reformatea las
    fechas ('2002-09-24' -> '24/9/2002') y los numeros con ceros a la izquierda
    apenas los ve. Si esas columnas no estan en el archivo, no las puede tocar.
    """
    if solo:
        columnas = [(t, c) for t, c in columnas if t in solo]
    with open(ruta, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=SEP_COLUMNAS,
                       quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
        w.writerow(["MFN"] + [c for _, c in columnas])
        for mfn, campos in zip(mfns, registros):
            w.writerow([mfn] + [SEP_OCURRENCIAS.join(campos.get(t, []))
                                for t, _ in columnas])



# El formato XLSX (XML por debajo) no acepta ciertos caracteres de control,
# aunque cp1252 los pueda representar y el .mst los tenga guardados. Son
# bytes de basura de carga vieja (no CR/LF, esos ya se pasan a <NL>/<CR>/<LF>
# mas arriba) que aparecieron por primera vez al exportar la base completa.
# Se sacan al escribir el xlsx -- si no, openpyxl directamente rechaza la
# celda y frena toda la exportacion -- y se listan para poder revisarlos.
ILEGALES_XLSX = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
LIMPIADOS_XLSX = []


def escribir_xlsx(ruta, mfns, registros, columnas, solo=None):
    """Como escribir_csv, pero en un .xlsx con TODAS las celdas fijadas al
    formato 'Texto' (number_format='@') desde que se crean.

    Con el formato bloqueado asi, Excel no reinterpreta nada -- ni fechas ni
    numeros con ceros a la izquierda -- sin importar la forma del contenido.
    Es lo que permite editar el archivo completo (53 columnas) en vez de tener
    que exportar solo 1 o 2 campos por vez: en un .csv comun, Excel decide el
    tipo de cada columna con solo mirarla, apenas se abre y se guarda, y eso
    reformatea columnas enteras que ni se tocaron.

    data_type se fuerza a 's' (string) ademas del number_format, porque si el
    contenido de una celda empieza con '=', '+', '-' o '@', openpyxl lo puede
    guardar como formula en vez de como texto.
    """
    if solo:
        columnas = [(t, c) for t, c in columnas if t in solo]

    LIMPIADOS_XLSX.clear()

    wb = Workbook()
    ws = wb.active
    ws.title = "editable"

    def poner(fila, col, valor, mfn=None, nombre_col=None):
        if isinstance(valor, str) and ILEGALES_XLSX.search(valor):
            limpio = ILEGALES_XLSX.sub("", valor)
            if mfn is not None:
                LIMPIADOS_XLSX.append((mfn, nombre_col, valor, limpio))
            valor = limpio
        celda = ws.cell(row=fila, column=col, value=valor)
        celda.number_format = "@"
        celda.data_type = "s"

    encabezados = ["MFN"] + [c for _, c in columnas]
    for col_idx, texto in enumerate(encabezados, start=1):
        poner(1, col_idx, texto)

    for fila_idx, (mfn, campos) in enumerate(zip(mfns, registros), start=2):
        poner(fila_idx, 1, mfn)
        for col_idx, (t, nombre_col) in enumerate(columnas, start=2):
            poner(fila_idx, col_idx, SEP_OCURRENCIAS.join(campos.get(t, [])),
                  mfn, nombre_col)

    ws.freeze_panes = "B2"
    wb.save(ruta)


def validar(registros):
    problemas = []
    for i, campos in enumerate(registros):
        for tag, ocurrencias in campos.items():
            for oc in ocurrencias:
                if SEP_OCURRENCIAS in oc:
                    problemas.append("registro %d v%s contiene %s"
                                     % (i + 1, tag, SEP_OCURRENCIAS))
                if "\r" in oc or "\n" in oc:
                    problemas.append("registro %d v%s conserva un caracter de control "
                                     "sin codificar" % (i + 1, tag))
    if problemas:
        print("\nERROR: %d ocurrencias son incompatibles con el formato." % len(problemas))
        for p in problemas[:20]:
            print("  " + p)
        sys.exit("\nExportacion cancelada. Ver la seccion 3 de ESPECIFICACION_CSV.md.")


# =========================================================
# MAIN
# =========================================================
def main():
    if "--iso" in sys.argv:
        ruta = sys.argv[sys.argv.index("--iso") + 1]
        print("Leyendo ISO existente: %s" % ruta)
        datos = open(ruta, "rb").read()
        mfns  = None
    else:
        base  = sys.argv[sys.argv.index("--base") + 1] if "--base" in sys.argv else ISIS_BASE
        datos, mfns = volcar(base)
        if base != ISIS_BASE:
            # Exportando desde otra base (un clon de prueba): se le pone su
            # nombre al archivo para no pisar el CSV de la base real.
            ETIQUETA[0] = os.path.basename(base)

    datos, cp850 = remediar_cp850(datos)
    if cp850:
        print("\n  AVISO: %d lineas traen caracteres de CP850 (ISIS de DOS) y se "
              "corrigieron." % len(cp850))
        print("  Al reimportar quedaran en cp1252. Registros a revisar en ABCD:")
        for l in cp850[:12]:
            print("    " + l)
        print()

    print("Parseando el ISO...")
    registros, descartados = parsear_iso(datos)
    print("  posiciones de registro halladas: %d   ilegibles: %d"
          % (len(registros), len(descartados)))

    if mfns is None:
        mfns = ["%06d" % (i + 1) for i in range(len(registros))]
        print("  (sin lista de MFN: se numeran correlativos)")
    elif len(mfns) == len(registros) - 1:
        # el patron del lider puede dar un falso positivo de mas
        sobra = [i for i in descartados]
        if sobra:
            quitar = sobra[-1]
            registros.pop(quitar); descartados.remove(quitar)
            descartados = [d - 1 if d > quitar else d for d in descartados]
    if mfns is not None and len(mfns) != len(registros):
        sys.exit("\nERROR: mx devolvio %d MFN pero el ISO trae %d registros.\n"
                 "Las dos pasadas no se corresponden y no se puede asignar el MFN."
                 % (len(mfns), len(registros)))

    if descartados:
        print("\n  AVISO: %d registros no se pudieron leer del ISO y quedan FUERA del CSV."
              % len(descartados))
        print("  Son registros con caracteres de control (CR/LF) guardados dentro de un")
        print("  campo: mx escribe el ISO en modo texto y esos bytes rompen la estructura.")
        print("  MFN afectados: %s" % ", ".join(mfns[i] for i in descartados))
        print("  Conviene limpiarlos en ABCD; hasta entonces no se pueden editar con esta")
        print("  herramienta. Tampoco se sobrescriben, porque nunca entran al ISO de salida.\n")
        mfns      = [m for i, m in enumerate(mfns) if i not in set(descartados)]
        registros = [r for r in registros if r is not None]

    print("Validando...")
    validar(registros)

    columnas, extras = armar_columnas(registros)

    os.makedirs(OUT_DIR, exist_ok=True)
    fecha    = datetime.now().strftime("%Y%m%d")
    nombre   = ETIQUETA[0] or "cat"
    ruta_b   = os.path.join(OUT_DIR, "%s_%s_base.csv" % (nombre, fecha))
    ruta_e   = os.path.join(OUT_DIR, "%s_%s_editable.xlsx" % (nombre, fecha))
    solo = None
    if "--editar" not in sys.argv and sys.stdin.isatty():
        # El editable ahora es un .xlsx con las celdas bloqueadas en formato
        # Texto, asi que Excel no reformatea nada: se puede pedir todo sin
        # riesgo. Esto solo sirve para achicar el archivo si se prefiere.
        print("\n  Que campos vas a corregir? (opcional, para achicar el archivo)")
        print("    22 autores    65 descriptores    23 autor institucional")
        print("    26 afiliacion 45 editor          85 disponibilidad")
        print("\n  Se puede poner mas de uno separado por coma. Ejemplo: 22,26")
        print("  Enter sin nada = las 53 columnas completas (ahora es seguro: el")
        print("  archivo se abre en Excel con las celdas bloqueadas en Texto).")
        r = input("\n  Campos: ").strip()
        if r:
            sys.argv += ["--editar", r]

    if "--editar" in sys.argv:
        pedidos = sys.argv[sys.argv.index("--editar") + 1]
        solo = [t.strip().lstrip("vV") for t in pedidos.split(",") if t.strip()]
        validos = {t for t, _ in columnas}
        faltan  = [t for t in solo if t not in validos]
        if faltan:
            sys.exit("ERROR: estas columnas no existen: %s"
                     % ", ".join("v" + t for t in faltan))

    escribir_csv(ruta_b, mfns, registros, columnas)             # completo, intacto
    escribir_xlsx(ruta_e, mfns, registros, columnas, solo)      # el que se edita

    if LIMPIADOS_XLSX:
        print("\n  AVISO: %d celdas tenian un caracter de control que Excel no acepta "
              "(no CR/LF -- esos van como <NL>/<CR>/<LF> -- sino basura de carga vieja)."
              % len(LIMPIADOS_XLSX))
        print("  Se quitaron solo en el .xlsx para poder abrirlo. Registros a revisar:")
        vistos = set()
        for mfn, col, _, _ in LIMPIADOS_XLSX:
            if (mfn, col) not in vistos:
                vistos.add((mfn, col))
                print("    MFN %s  %s" % (mfn, col))
                if len(vistos) >= 12:
                    break
        if len(vistos) < len(set((m, c) for m, c, _, _ in LIMPIADOS_XLSX)):
            print("    ... y otros")
        print()

    total_oc = sum(len(v) for r in registros for v in r.values())
    multi    = sum(1 for r in registros for v in r.values() if len(v) > 1)
    nums     = sorted(int(m) for m in mfns)
    con_ctrl = sum(1 for r in registros for v in r.values() for oc in v
                   if any(tok in oc for _, tok in CONTROLES))

    print("\n--- Resumen ---")
    print("  Registros exportados      : %d" % len(registros))
    print("  MFN minimo / maximo       : %d / %d" % (nums[0], nums[-1]))
    print("  MFN faltantes (borrados)  : %d" % (nums[-1] - len(nums)))
    print("  Ocurrencias totales       : %d" % total_oc)
    print("  Campos con >1 ocurrencia  : %d" % multi)
    print("  Ocurrencias con control   : %d  (tokens %s)"
          % (con_ctrl, " ".join(t for _, t in CONTROLES)))
    print("  Columnas                  : %d" % (len(columnas) + 1))
    if extras:
        cuenta = Counter()
        for r in registros:
            for t in r:
                if t in extras:
                    cuenta[t] += len(r[t])
        print("  Etiquetas fuera de la FDT : %s"
              % ", ".join("v%s (%d)" % (t, cuenta[t]) for t in extras))

    print("\n  Linea base : %s   (NO EDITAR)" % ruta_b)
    print("  Editable   : %s" % ruta_e)
    print("\nListo.")


if __name__ == "__main__":
    main()
