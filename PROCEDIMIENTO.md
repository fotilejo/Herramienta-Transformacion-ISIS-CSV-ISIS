# Procedimiento — guía única y definitiva

> Ciclo **CAT (ISIS) → Excel → edición manual → CAT (ISIS)**.
> Herramienta independiente: no toca el Gestor de Migración ni las bases NORMA y KARDEX.
>
> **Estado: en uso real.** Las Fases A y B están cerradas (ver resumen abajo y el detalle
> en `ESPECIFICACION_CSV.md` §9). La Fase C ya se probó con una corrección real (unificar
> las variantes de "Martínez Krahmer, Daniel Osvaldo", 120 cambios en 85 registros) sobre
> un clon completo de los 49.592 registros, verificada campo por campo, celda por celda,
> contra la base real: la única diferencia sin explicar que quedó son 7 celdas de v401/v110
> duplicadas (ver "Efecto conocido" en §2) — ningún campo de catalogación afectado, ningún
> MFN de más ni de menos. Este documento es **el flujo a seguir de ahora en más** para
> cualquier corrección.

---

## 1. Un solo archivo para arrancar

**`REPARADOR_CAT.bat`.** Es el único que hace falta abrir. Muestra un menú con 3
opciones (preparar, generar cambios, aplicar) y va diciendo en cada paso qué archivo
corregir o qué esperar — no hace falta acordarse de nombres de scripts ni de cuál CSV
tocar, lo indica en pantalla y además abre solo el archivo que hay que editar.

Los pasos 1, 2 y 3 del menú son exactamente los mismos de la Fase C (§3 más abajo); el
menú simplemente los llama por vos, en orden, sin que haga falta recordar cuál va primero.

### El resto de los archivos de la carpeta

No hace falta abrirlos nunca directamente — quedaron documentados acá solo por si algún
día hace falta mirar el detalle:

| Archivo | Qué es |
|---|---|
| `exportar_cat.py`, `generar_iso.py`, `aplicar_cambios.py` | El código que corre el menú por detrás. |
| `verificar.py`, `verificar_clon.py` | Usados internamente por `aplicar_cambios.py` (no se corren solos). |
| `log_aplicar_cambios.txt` | Registro acumulado de cada corrida del paso 3, por si hay que revisar qué pasó. |

Los `.bat` numerados del 01 al 06 (las formas viejas de llamar cada paso por separado, y
las herramientas usadas solo durante las pruebas) ya se borraron — `REPARADOR_CAT.bat` no
dependía de ellos.

### Carpetas

| Carpeta | Contenido |
|---|---|
| `Salidas_CSV` | Los archivos que arma el paso 1 del menú (línea base + editable). |
| `Salidas_ISO` | Los `.iso` y sus informes `.txt` que arma el paso 2 del menú. |
| `Respaldos` | Copias de seguridad automáticas del paso 3 del menú, una carpeta por corrida. |
| `temp` | Archivos de trabajo intermedios. Se pueden borrar sin problema, se regeneran solos. |

**Antes de empezar con la base real:** conviene vaciar `Salidas_CSV`, `Salidas_ISO` y
`Respaldos` (los archivos de las pruebas de estos días), para que lo único que haya ahí
sea la corrida real. Nada de esto se borra solo — hay que hacerlo a mano desde el
Explorador de Windows.

---

## 2. Qué quedó probado (resumen — detalle en ESPECIFICACION_CSV.md §9)

- **Fidelidad ISIS ↔ Excel/CSV:** exportar y reimportar sin editar nada da cero
  diferencias, sobre 2.000 registros y sobre la base completa.
- **El editable es un `.xlsx`** con las 53 columnas y las celdas bloqueadas en formato
  Texto: se puede corregir cualquier campo, a mano o con ayuda de IA, sin que Excel
  reformatee fechas ni números por su cuenta.
- **La actualización selectiva conserva el MFN** y no toca nada más — probado sobre una
  base con registros borrados de verdad (huecos), que siguieron borrados después.
- **Nunca se recarga la base entera.** Eso solo se hizo en las pruebas (bases de
  descarte, sin huecos reales). Sobre CAT real, con sus 1.876 huecos, recargar todo
  correría los MFN y rompería las referencias de préstamos. Por eso la Fase C sólo usa
  actualización selectiva (§3, pasos 6-7).
- **Comportamientos de ABCD confirmados:** no regenera `v9` (número de control), no pisa
  `v4` (fecha de registro), agrega `v401='true'` a lo que importa (inofensivo), y limpia
  saltos de línea crudos guardados dentro de un campo (necesario, WXIS no los acepta).
- **Efecto conocido (v401 y v110):** estos dos campos no están en la FDT de CAT. `mx` los
  va sumando en vez de reemplazarlos, mande la herramienta lo que mande (probado: omitir
  el campo, mandar el valor viejo, mandar el valor correcto y hasta un borrado explícito
  `D401`/`D110` en el `proc` — ninguno evitó que, en algunos registros puntuales, quede
  duplicado como `true<|>true` o el sello de edición repetido dos veces). Pasa solo cuando
  el registro ya traía ese campo cargado de antes; no es predecible cuáles lo van a hacer.
  No afecta ningún campo de catalogación, ni el MFN, ni la búsqueda. **No intentar
  "limpiarlo" borrando v110 en bloque:** en algunos registros v110 no es un sello de
  sistema, tiene datos reales (ej. una URL, MFN 046015). Si aparece, es cosmético y se deja
  así.

---

## 3. Fase C — Aplicar correcciones sobre la base CAT real

Catalogador único, sin concurrencia. Requiere una ventana corta de congelamiento
(pasos 4 a 10): mientras tanto no se cataloga en ABCD, porque cualquier cambio hecho ahí
durante la ventana se pierde al copiar los archivos de vuelta.

### Antes de arrancar

Vaciar `Salidas_CSV`, `Salidas_ISO` y `Respaldos` si todavía tienen archivos de las
pruebas (ver §1). No es obligatorio, pero evita confundir un archivo viejo con uno de la
corrida real.

### ¿Hace falta clonar antes de tocar la base real?

**No siempre.** El mecanismo ya está probado a fondo, con una corrección real, a escala
completa. Para una corrección chica y bien entendida (unos pocos registros, un campo
puntual) se puede aplicar directo sobre la copia local real (paso 7, opción 1).

Para un lote grande o algo nuevo que no se probó antes, sigue siendo buena idea clonar
la base completa primero y aplicar ahí (ver más abajo, "Clonar la base completa antes de
aplicar"), simplemente por tranquilidad — no porque el mecanismo lo necesite.

### Pasos

1. **Dejar de catalogar en ABCD.**

2. **Backup de `cat.mst` y `cat.xrf` del servidor**, aparte de todo lo demás — esta es la
   copia de seguridad "de verdad", antes de tocar nada.

3. **Copiar `cat.mst` y `cat.xrf` del servidor a la PC local**, a la ruta configurada en
   `exportar_cat.py` (`C:\Isis\cat\data\cat`, salvo que se haya cambiado).

4. **Abrir `REPARADOR_CAT.bat` y elegir la opción 1 (Preparar una corrección).**
   No pregunta nada más: exporta siempre la base CAT real completa (las 53 columnas). Ya
   no hace falta elegir campos a mano — el `.xlsx` sale con las celdas bloqueadas en
   Texto, así que pedir todo de entrada es seguro.

   El programa abre solo, en Excel, el archivo que hay que editar, y lo dice en pantalla.
   El otro archivo que genera (la línea base) no se toca — es solo para que la
   herramienta compare después.

   Para exportar otra base en vez de CAT real (un clon, una carpeta de prueba), correr
   `python reparador.py --base "<ruta sin .mst>"` desde la línea de comandos — es un caso
   puntual, no hace falta un menú para eso.

5. **Editar ese Excel.** Corregir solo lo que corresponda. Recordatorios:
   - Las ocurrencias van separadas por `<|>`, no por `;` (ver ESPECIFICACION_CSV.md §5).
   - `v9` (número de control) no se edita — el escritor lo rechaza si cambia.
   - `v210` (Revisado) no se toca salvo que sea a propósito.
   - Guardar y cerrar Excel antes de seguir (si el archivo queda abierto, el siguiente
     paso puede fallar al intentar leerlo).

6. **Volver al menú de `REPARADOR_CAT.bat` y elegir la opción 2 (Generar los cambios).**

   **Leer el informe (se abre o se indica la ruta) antes de seguir.** Tiene que listar
   exactamente lo que se quiso corregir, MFN por MFN, valor viejo y nuevo. Si aparece un
   MFN que no se tocó a propósito, o si el número de "registros modificados" es mucho más
   alto de lo esperado, **parar ahí** — es la señal de que algo se coló sin querer
   (típicamente Excel reformateando de más, aunque ya no debería pasar con el `.xlsx`).
   Excepción: puede aparecer alguna línea de `v401_no_declarado` o `v110_no_declarado` en
   un registro que sí se editó a propósito por otro campo — es el efecto conocido de §2,
   no una señal de alarma.

7. **Elegir la opción 3 del menú (Aplicar los cambios), opción 1** dentro de esa
   (la copia local de CAT). Hace, en orden:
   - Copia de seguridad de la copia local (además del backup del paso 2).
   - Aplica los cambios por MFN, sin tocar ningún otro registro.
   - Reexporta y confirma, campo por campo, que cada cambio pedido quedó como se esperaba.

   Si dice **LISTO**: seguir al paso 8.
   Si dice **HAY CAMBIOS QUE NO QUEDARON COMO SE ESPERABA**: no seguir. Restaurar los
   archivos desde la carpeta de `Respaldos` que indica el mensaje, y avisar para revisar
   qué pasó antes de reintentar.

8. **Copiar `cat.mst` y `cat.xrf` (ya verificados) al servidor.**

9. **Reconstruir los índices en el servidor**, desde la interfaz de ABCD. Sin esto las
   búsquedas de ABCD no reflejan el cambio, aunque el archivo ya esté corregido.

10. **Reanudar la catalogación.**

---

## 4. Clonar la base completa antes de aplicar (opcional)

Forma de ensayar un lote de cambios contra una copia real completa (mismo MFN, mismos
huecos que la base real) sin arriesgar nada, y de paso poder revisarlo en la pantalla de
ABCD antes de tocar la base real. Así se probó la corrección de Martínez Krahmer.

**No alcanza con crear la base y listo** (importar un ISO adentro renumera todo y pierde
los huecos, ver §2). Los pasos que sí preservan todo:

1. **Crear la base clon desde la interfaz de ABCD** (Central → Administración de bases →
   Crear base de datos, tomando CAT como modelo). Va a quedar registrada en ABCD, pero
   vacía.
2. **En el servidor, reemplazar el `.mst` y el `.xrf`** que esa base recién creada generó
   por copias de los `cat.mst`/`cat.xrf` reales, **renombradas con el nombre del clon**
   (por ejemplo, si el clon se llama `catprueba`: quedan `catprueba.mst` y
   `catprueba.xrf`, con el contenido real adentro).
3. **Reconstruir los índices** de esa base en ABCD, para que la búsqueda funcione.
4. **Copiar esos dos archivos del servidor a la PC local**, a una carpeta propia (por
   ejemplo `C:\Isis\catprueba\`).
5. Seguir los pasos 4 a 7 de §3 normalmente, pero en el paso 7 elegir **opción 3 (Otra)**
   y apuntar a esta ruta local en vez de a la copia real de CAT.
6. Si el resultado es el esperado, aplicar lo mismo sobre la copia local real (paso 7,
   opción 1) y seguir con §3 desde el paso 8.

---

## 5. Si algo sale mal

- **Antes del paso 7:** no se tocó ningún archivo real todavía. Se puede cortar en
  cualquier momento sin consecuencias.
- **Durante o después del paso 7, si algo no cierra:** restaurar `cat.mst` y `cat.xrf`
  copiando los de la carpeta más reciente dentro de `Respaldos` (o los del paso 2, si
  hace falta ir más atrás) a `C:\Isis\cat\data\cat`. El servidor no se tocó todavía en
  ese punto, así que no hay nada que revertir ahí.
- **Si ya se copiaron los archivos al servidor (paso 8) y después se encuentra un
  problema:** restaurar con el backup del servidor hecho en el paso 2, y volver a
  reconstruir los índices ahí.
