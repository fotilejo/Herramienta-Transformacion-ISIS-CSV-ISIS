# Especificación del CSV fiel de CAT

> Formato de intercambio para el ciclo **CAT (ISIS) → CSV/Excel → edición manual → CAT (ISIS)**.
> Herramienta nueva e independiente del Gestor de Migración existente, que no se modifica.
>
> **Estado:** cerrada. Probada de punta a punta, incluyendo la actualización selectiva
> sobre una base con registros borrados. No queda nada pendiente de determinar.
> **Fecha:** 2026-07-27, actualizado 2026-07-28.

---

## 0. Origen de los datos: el ISO, no un PFT

`exportar_cat.py` lee el **ISO 2709 que produce `mx`**, no un volcado con formato PFT.
El cambio se hizo después de auditar el export contra un ISO real, y corrigió dos pérdidas:

- **Trece etiquetas fuera de la FDT.** La FDT declara 53; en los datos hay 66. Existen y no
  están declaradas: `v21`, `v60`, `v76`, `v78`, `v80`, `v95`, `v102`, `v103`, `v110`,
  `v130`, `v131`, `v327` y `v401` — unas 2.450 ocurrencias. `v110` es metadata de edición
  que escribe ABCD (`^iabcd^t201709291346^x1506703594`); `v76` y `v130` traen datos reales
  de sede. Un PFT obliga a enumerar etiquetas a mano y todo lo no enumerado se pierde.
- **77 registros donde el PFT no ve campos que el ISO sí trae.** Caso testigo, MFN 050356:
  tiene `v10='true'` y dos `v80`, con el directorio empezando por `080,080,010` y sin `v3`.
  `if p(v10)` devuelve vacío; el ISO lo muestra.

El ISO vuelca el directorio completo, así que el exportador no depende de ninguna lista.
Las columnas se arman solas: primero las 53 de la FDT en su orden, después lo que aparezca,
por número, con el nombre `vNN_no_declarado`.

Como el ISO 2709 no tiene lugar para el número de registro, se corren **dos pasadas de
`mx`**: una para el ISO y otra para la lista de MFN. Recorren la base en el mismo orden y
saltean los borrados lógicos, así que se corresponden posición a posición. Si las
cantidades no coinciden, el proceso se cancela.

### Registros que no se pueden leer

`mx` escribe el ISO en **modo texto**: los bytes de control guardados dentro de un campo se
expanden y rompen la estructura declarada en el lider. En CAT eso deja **15 registros
ilegibles**, que quedan fuera del CSV. El exportador los lista por MFN al terminar.

No se sobrescriben nunca, porque tampoco entran al ISO de salida. Pero **no se pueden
editar con esta herramienta** hasta que se limpien los caracteres de control en ABCD.

---

## 1. Propósito y alcance

Permitir editar el catálogo CAT a mano en una planilla y devolver los cambios a la base
de ABCD sin pérdida de información.

**Dentro del alcance:** únicamente la base **CAT**.
**Fuera del alcance:** NORMA y KARDEX, que no se tocan. El pipeline actual hacia
`Catalogo.ttl` y los espejos JSON sigue funcionando sin cambios.

La copia maestra de CAT vive en el servidor ABCD. La copia local es de trabajo.

---

## 2. Contexto verificado

Medido sobre la exportación real de CAT y sobre `cat.fdt`:

| Dato | Valor |
|---|---|
| Registros activos | 49.592 |
| MFN máximo | 51.468 |
| MFN faltantes (borrados lógicos) | 1.876 |
| MFN duplicados | 0 |
| Campos declarados en la FDT | 53 |
| Campos que exporta el PFT actual | 49 |
| Campos que se pierden hoy | 4 (`v9`, `v25`, `v120`, `v210`) |
| Ocurrencias de campo totales | ~1.487.296 |
| Máximo de ocurrencias en un campo | 160 (`v77` Inventario) |
| Codificación del `.mst` | **cp1252** |
| Codificación de `cat.fdt` | cp850 (sólo referencia) |

**Consecuencia de los huecos de MFN:** está prohibido reconstruir la base completa. Un
`.mst` regenerado desde 49.592 registros renumeraría todo a partir del MFN 104 y rompería
signaturas, referencias internas de ABCD y las URIs del TTL. Los cambios se aplican
**siempre de forma selectiva, por MFN**.

---

## 3. Forma del archivo

**Una fila por registro, una columna por campo.** 49.592 filas de datos + 1 de encabezado.

Se descartó la forma larga (una fila por ocurrencia) porque son ~1,49 millones de filas y
Excel y LibreOffice cortan en 1.048.576.

Se generan **dos archivos por corrida**, con formatos distintos a propósito:

| Archivo | Formato | Se edita |
|---|---|---|
| `*_base.csv` | CSV, UTF-8 con BOM | **No.** Es la línea base intocable para el diff. |
| `*_editable.xlsx` | Excel, celdas en formato Texto | **Sí.** Es el que se abre y corrige. |

**Por qué el editable es `.xlsx` y no `.csv`.** La primera versión de la herramienta usaba
CSV para los dos archivos. Se abandonó después de un incidente real: al abrir y guardar el
CSV completo en Excel, este reformateó solo columnas de fecha (`2002-09-24` → `24/9/2002`)
y números con ceros a la izquierda, en decenas de miles de registros que nadie había
tocado — mezclado con 2 ediciones genuinas, indistinguible sin comparar contra la línea
base. La solución no fue restringir qué se puede editar, sino sacarle a Excel el poder de
decidir el formato: `exportar_cat.py` escribe el `.xlsx` con **cada celda fijada al formato
"Texto" desde que se crea** (`number_format='@'`, más `data_type='s'` forzado para que un
valor que empiece con `=`, `+`, `-` o `@` no se tome como fórmula). Con el formato
bloqueado así, Excel no reinterpreta nada sin importar la forma del contenido, y las 53
columnas quedan editables sin riesgo.

| Parámetro | Valor |
|---|---|
| Separador de columnas (CSV) | `;` (Excel en configuración regional español) |
| Comillas (CSV) | RFC 4180: `"` sólo si hace falta, duplicadas para escapar |
| Codificación (CSV) | **UTF-8 con BOM** |
| Fin de línea (CSV) | CRLF |
| Separador de ocurrencias | `<\|>` (en ambos formatos) |
| CR LF dentro de un campo | `<NL>` |
| CR suelto dentro de un campo | `<CR>` |
| LF suelto dentro de un campo | `<LF>` |

**Sobre la codificación:** el archivo se edita en UTF-8 (o como texto de Excel), pero la
base es cp1252. La conversión ocurre al escribir de vuelta, con validación explícita —
ver §7.

**Sobre el separador de ocurrencias:** se verificó contra los 31 MB de la exportación real
que la secuencia `<|>` **no aparece nunca** en los datos de CAT. Candidatos descartados por
aparecer en los datos: `|` (5 líneas), `~` (7), `§` (509), `^^` (7). Alternativas también
limpias, por si alguna vez hiciera falta cambiarlo: `@@`, `¶`, `‡`, `#%#`.

**No se usa `//` como separador.** Es lo que hace la exportación actual y es ambiguo: el
texto catalogado puede contener `//` legítimamente.

**Saltos de línea dentro del contenido.** 18 ocurrencias de CAT tienen un CRLF guardado
adentro del campo, repartidas en `v10` (2), `v18` (1), `v23` (7), `v33` (1), `v53` (2),
`v54` (4) y `v65` (1). Se representan con el token `<NL>` en vez de un salto real, para
que cada registro ocupe **una sola fila** del CSV: así la planilla se edita, se ordena y
se compara sin filas partidas. Verificado: `<NL>` no aparece en los datos. Al escribir de
vuelta se restituye el CRLF original.

---

## 4. Las 54 columnas

Columna 1 = `MFN`. Las 53 restantes siguen el orden de la FDT, que es el orden de la
planilla de carga de ABCD.

| # | Tag | Nombre en la FDT | Repetible | Encabezado en el CSV |
|---|---|---|---|---|
| 0 | — | Número de registro ISIS | no | `MFN` |
| 1 | v3 | Fecha de ingreso | no | `v3_Fecha_de_ingreso` |
| 2 | v4 | Fecha de registro | no | `v4_Fecha_de_registro` |
| 3 | v2 | Fuente del registro | sí | `v2_Fuente_del_registro` |
| 4 | v5 | Nivel bibliográfico | no | `v5_Nivel_bibliografico` |
| 5 | v6 | Nivel de referencia | no | `v6_Nivel_de_referencia` |
| 6 | v7 | Tipo de documento | sí | `v7_Tipo_de_documento` |
| 7 | v8 | Soporte no convencional | sí | `v8_Soporte_no_convencional` |
| 8 | v9 | Número de control | no | `v9_Numero_de_control` |
| 9 | v10 | ISBN (Monografía) | sí | `v10_ISBN_Monografia` |
| 10 | v11 | ISBN (Serie o colección) | sí | `v11_ISBN_Serie_o_coleccion` |
| 11 | v12 | Código de doc. | sí | `v12_Codigo_de_doc` |
| 12 | v14 | Código identif. de doc. esp. | sí | `v14_Codigo_identif_de_doc_esp` |
| 13 | v15 | ISSN | sí | `v15_ISSN` |
| 14 | v18 | Proy., progr. u otro encuadre | sí | `v18_Proy_progr_u_otro_encuadre` |
| 15 | v19 | Otros códigos | sí | `v19_Otros_codigos` |
| 16 | v20 | Título | sí | `v20_Titulo` |
| 17 | v22 | Autor personal | sí | `v22_Autor_personal` |
| 18 | v23 | Autor institucional | sí | `v23_Autor_institucional` |
| 19 | v25 | Colaborador | sí | `v25_Colaborador` |
| 20 | v26 | Afiliación del autor | sí | `v26_Afiliacion_del_autor` |
| 21 | v29 | Reunión: Responsable | sí | `v29_Reunion_Responsable` |
| 22 | v30 | Reunión: Nombre y nro | sí | `v30_Reunion_Nombre_y_nro` |
| 23 | v31 | Reunión: Lugar-país | sí | `v31_Reunion_Lugar_pais` |
| 24 | v32 | Reunión: Fecha | sí | `v32_Reunion_Fecha` |
| 25 | v33 | Título de serie o colección | sí | `v33_Titulo_de_serie_o_coleccion` |
| 26 | v39 | Autor institucional (ser-col) | sí | `v39_Autor_institucional_ser_col` |
| 27 | v40 | Idioma | sí | `v40_Idioma` |
| 28 | v42 | Edición | no | `v42_Edicion` |
| 29 | v43 | Lugar de edición | sí | `v43_Lugar_de_edicion` |
| 30 | v44 | País de edición | sí | `v44_Pais_de_edicion` |
| 31 | v45 | Editor | sí | `v45_Editor` |
| 32 | v46 | Fecha de edición | sí | `v46_Fecha_de_edicion` |
| 33 | v50 | Colación | sí | `v50_Colacion` |
| 34 | v53 | Tesis (Inst.-grado-fecha) | no | `v53_Tesis_Inst_grado_fecha` |
| 35 | v54 | Notas | sí | `v54_Notas` |
| 36 | v55 | Unidad Técnica | sí | `v55_Unidad_Tecnica` |
| 37 | v56 | Resumen | no | `v56_Resumen` |
| 38 | v65 | Descriptores | sí | `v65_Descriptores` |
| 39 | v75 | Signatura topográfica | sí | `v75_Signatura_topografica` |
| 40 | v77 | Inventario | sí | `v77_Inventario` |
| 41 | v81 | Tipo_res_cientif | sí | `v81_Tipo_res_cientif` |
| 42 | v82 | Versión | sí | `v82_Version` |
| 43 | v84 | Formato | sí | `v84_Formato` |
| 44 | v85 | Disponibilidad | sí | `v85_Disponibilidad` |
| 45 | v86 | Nivel_acceso | sí | `v86_Nivel_acceso` |
| 46 | v88 | Licencia | sí | `v88_Licencia` |
| 47 | v90 | Exportable_SNRD | no | `v90_Exportable_SNRD` |
| 48 | v100 | Código bd | sí | `v100_Codigo_bd` |
| 49 | v120 | Enviado a BVSDE | sí | `v120_Enviado_a_BVSDE` |
| 50 | v140 | URL | sí | `v140_URL` |
| 51 | v150 | Dirección | sí | `v150_Direccion` |
| 52 | v200 | Catalogador | no | `v200_Catalogador` |
| 53 | v210 | Revisado | no | `v210_Revisado` |

El número de tag va en el encabezado para que el mapeo de vuelta sea explícito y no
dependa del orden ni del nombre.

---

## 5. Campos con tratamiento especial

### `v9` — Número de control · NO EDITAR

Único campo con tipo `AI` en la FDT (todos los demás son `X`), lo que indica asignación
automática por el sistema. ABCD lo genera solo.

- Se exporta para no perderlo, pero es **de sólo lectura**.
- El escritor **rechaza** cualquier fila cuyo `v9` difiera de la línea base.
- **Pendiente de verificar:** si ABCD regenera `v9` al importar. Ver §9.

### `v210` — Revisado · NO PISAR

Campo de estado del flujo de catalogación, en uso permanente. La herramienta nunca lo
escribe por iniciativa propia; sólo lo devuelve tal como salió, salvo edición explícita.

### `v4` — Fecha de registro · BAJO SOSPECHA

Es el candidato natural a que ABCD lo actualice automáticamente al modificar un registro.
Si lo hace, cada corrida de reparación pisaría la fecha real de catalogación.
**Pendiente de verificar.** Ver §9.

### `v77` — Inventario · CELDAS LARGAS

Llega a 160 ocurrencias en un mismo registro. En forma ancha eso es una celda con 160
valores separados por `<|>`. Es correcto pero incómodo; no es un campo destinado a
edición manual.

### `v20` — Título

Único campo con subcampos declarados en la FDT. Ver la nota sobre subcampos en §6.

### `v900` — NUNCA ES UNA COLUMNA

No es un campo de catalogación: es el mecanismo interno de esta herramienta para llevar el
MFN metido dentro del ISO 2709 (que no tiene lugar propio para el número de registro).
`generar_iso.py` lo agrega al escribir y lo saca con `proc=...,D900` al aplicar.

Si una base ya pasó antes por una carga de esta herramienta, ese campo puede haber quedado
escrito dentro de los datos reales. `armar_columnas()` en `exportar_cat.py` lo descarta
siempre, así nunca aparece como columna editable — si apareciera, la próxima vuelta por
`generar_iso.py` escribiría el tag `900` dos veces en el mismo registro (el que se inyecta
de nuevo más el viejo) y `mx` no sabría cuál usar como MFN destino. Esto pasó una vez en
las pruebas del 2026-07-28 (`fatal: recupdat/mfn`, ver §9) y quedó resuelto en dos lugares:
el exportador ya no expone la columna, y `armar_registro()` en `generar_iso.py` descarta
cualquier `900` que igual llegue a `campos`, por las dudas.

### `v401` — ABCD LO ESCRIBE SOLO AL IMPORTAR

Hallazgo del 2026-07-28: al importar un ISO por la interfaz de ABCD, cada registro queda
con `v401='true'`, campo que en el original está vacío. No es un campo de los que se usan
en la catalogación (control, colaborador, BVSDE, revisado); parece un marcador interno de
"esto entró por ISO". No bloquea nada, pero conviene saber que ABCD sí escribe algo por su
cuenta al importar — a diferencia de `v9` y `v4`, que confirmadamente no toca (ver §9).

### El `;` de ABCD · CONVENCIÓN DE PANTALLA, NO DATO

En el formulario de carga de ABCD las ocurrencias se escriben en líneas separadas y, al
reabrir el registro, se muestran unidas por `; `. **Ese `;` no existe en el `.mst`.**

Verificado sobre los datos reales:

| Campo | Celdas no vacías | Con `;` | Con ocurrencias reales |
|---|---|---|---|
| `v22` Autor personal | 36.848 | 29 | 17.325 |
| `v65` Descriptores | 47.427 | 1 | 40.786 |
| `v23` Autor institucional | 17.546 | 18 | 2.621 |

Si el `;` fuera dato, la proporción sería la inversa. ABCD divide en ocurrencias ISIS
genuinas al guardar. Por lo tanto el separador de ocurrencias del CSV sigue siendo `<|>`
y no hay conversión que hacer.

**Riesgo asociado — memoria muscular:** escribir `Autor 1; Autor 2` en una celda del CSV
**no** crea dos ocurrencias. Crea una sola ocurrencia con un `;` literal adentro. El
escritor emite un **aviso** (no bloqueante) por cada celda editada que contenga `;`, para
que se confirme si es texto legítimo o un intento de separar ocurrencias.

**Dato de calidad detectado:** las 29 celdas de `v22` con `;` son erratas de tipeo —
`;` donde correspondía `,` en el patrón `Apellido, Nombre` (p. ej. MFN 015768
`O'Farrell; Ernesto`, MFN 019260 `Gass; Saul I.`). Candidatas a corrección.

---

## 6. Reglas de ida (ISIS → CSV)

1. **Se exportan todos los campos de la FDT**, presentes o no en el registro.
2. **Las ocurrencias se leen del `.mst`, nunca se infieren de la FDT.**
   La FDT declara `v3`, `v4`, `v5` y `v42` como no repetibles, pero en los datos reales hay
   14 registros con ocurrencias múltiples en esos campos. ISIS no valida la FDT: el `.mst`
   guarda lo que sea. El exportador debe respetar el dato, no la declaración.
3. **Los subcampos `^x` viajan como texto literal, opacos.**
   CDS/ISIS no usa el separador `0x1F` del estándar: guarda `^n`, `^s`, `^l`, `^p` como
   caracteres normales dentro del contenido. No se parsean, no se normalizan, no se
   reordenan. La FDT sólo declara subcampos en `v20`, pero los datos los usan en `v23`,
   `v26`, `v43` y otros — otra razón para tratarlos como texto y no como estructura.
4. **Cero normalizaciones.** No se colapsan espacios, no se cambia capitalización, no se
   quitan `#`, no se resuelven entidades HTML, no se recortan bordes. Fiel es fiel.
5. **Celda vacía = campo ausente.** En ISIS un campo presente y vacío no existe, así que
   no hay ambigüedad que resolver.
6. **Los registros borrados lógicamente no se exportan.** Sus MFN quedan como huecos y la
   herramienta jamás intenta rellenarlos.
7. **Se guarda una copia pristina** del CSV recién exportado, como línea base para el diff.
   Nombre sugerido: `cat_YYYYMMDD_base.csv` junto a `cat_YYYYMMDD_editable.csv`.

---

## 7. Reglas de vuelta (CSV → ISIS)

1. **Diff contra la línea base.** Sólo se procesan los MFN cuyo contenido cambió. Un
   registro no modificado no se toca ni se reescribe.
2. **El MFN es la llave y es inmutable.** No se crean registros nuevos, no se borran, no
   se renumera nada. Si aparece un MFN que no está en la línea base, se rechaza el archivo.
3. **Desdoblado por `<|>`.** Cada fragmento vuelve a ser una ocurrencia independiente del
   campo. Fragmentos vacíos entre separadores se descartan.
4. **Validación de codificación, obligatoria y bloqueante.** Antes de escribir nada se
   verifica que todo carácter del CSV sea representable en cp1252. Si no lo es, el proceso
   **aborta** e informa MFN, columna y carácter.
   Casos típicos que introduce Excel: comillas tipográficas (`" "` `' '`), guion largo
   (`—`), puntos suspensivos (`…`), espacio fino y espacio duro.
5. **Se rechaza `v9` modificado.**
5bis. **Aviso por `;` en celdas editadas.** No bloquea, pero se lista para revisión: puede
   ser texto legítimo o un intento de separar ocurrencias al estilo ABCD. Ver §5.
6. **Informe previo obligatorio.** Antes de tocar la base se emite el detalle de qué MFN
   cambian, en qué campos, y el valor viejo contra el nuevo. Nada se escribe sin que ese
   informe se haya generado.
7. **Se aplica primero sobre la copia local**, nunca directo al servidor.

---

## 8. Validación: prueba de identidad ida y vuelta

Antes de confiar en la herramienta, y sin ninguna edición de por medio:

```
CAT local → CSV → (sin editar) → aplicar a base scratch → re-exportar CSV → comparar
```

Los dos CSV deben ser **idénticos byte a byte**. Cualquier diferencia es un defecto del
conversor y hay que resolverlo antes de editar nada real.

Esta prueba corre sobre los 49.592 registros y no toca el servidor.

---

## 9. Resultados de las pruebas (cerradas)

### Fase A — fidelidad de la carga completa (2026-07-27, repetida 2026-07-28)

2.000 registros importados en un clon vacío (`catprueba`, después `catpruebas`) y
reexportados. Comparación campo por campo contra el original.

| Pregunta | Respuesta |
|---|---|
| ¿ABCD regenera `v9` al importar? | **No.** Conservan su número de control. |
| ¿ABCD pisa `v4` (fecha de registro)? | **No.** Las fechas quedaron intactas. |
| ¿Sobreviven los subcampos `^n^s^l^p`? | **Sí**, literales. |
| ¿Sobreviven los acentos cp1252? | **Sí** (`Préstamo`, `Tecnología`). |
| ¿ABCD agrega campos propios al importar? | **Sí: `v401='true'`** en todos. Ver nota en §5. |
| ¿Sobrevive `v900` con el MFN? | **Sí**, es la llave de actualización — ver nota en §5. |
| ¿Llegan las etiquetas fuera de la FDT? | **Sí** (`v110` en los registros que lo traían). |
| ¿Sobreviven los caracteres de control (CR/LF) embebidos? | **No** — se limpian al escribir el ISO porque WXIS no puede importarlos. Es un cambio intencional, avisado en el informe. |

### Fase B — actualización selectiva conservando MFN, con borrados reales (2026-07-28)

Sobre una base con 1.997 registros activos y 3 borrados lógicos (MFN 100, 101 y 102,
borrados con la función normal de ABCD, para simular los huecos de la base real):

```
mx.exe iso=cambios.iso proc='=',v900,'D900' updatf=<base> -all now
```

- Corrigió un registro sin tocar ningún otro, verificado campo por campo.
- Los 3 borrados **siguieron borrados**, antes y después — confirmado con `mx` y
  visualmente en la interfaz de ABCD ("registro borrado").
- Cantidad de registros sin cambios: 1.997, con los mismos 3 huecos.

**Incidente durante la prueba:** la primera corrida falló con `fatal: recupdat/mfn` en
cualquier registro, incluso lejos de los huecos borrados. La causa **no fueron los
borrados** (sospecha inicial descartada) sino el campo `v900` duplicado — ver la nota en
§5. Resuelto en `exportar_cat.py` y `generar_iso.py`.

### Conclusión

El ciclo completo — exportar, editar en Excel sin restricciones, generar el ISO,
aplicar por MFN conservando huecos — está probado de punta a punta sobre una base con
las mismas condiciones que la real (borrados incluidos). No queda ningún punto abierto
para empezar la Fase C. Ver `PROCEDIMIENTO.md`.

---

## 10. Flujo operativo

Ver `PROCEDIMIENTO.md` — es la guía paso a paso, actualizada, para aplicar esto sobre la
base CAT real (Fase C). Este documento describe el **formato**; ese describe los **pasos**.
