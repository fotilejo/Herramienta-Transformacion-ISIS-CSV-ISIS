# Paso a paso simple

1. Ejecuto `REPARADOR_CAT.bat`.
2. Selecciono **"1. Preparar una corrección nueva (exportar de ISIS a Excel)"**. Sin
   preguntar nada más, exporta la base CAT real completa (todas las columnas).
3. Se generan unos Excel. Al que dice **"editable"** le hago los cambios que quiera, de
   la forma que mejor me convenga. Guardo y cierro. Cierro todo si quiero — retomo
   cuando esté listo, no hace falta hacerlo de una sentada.
4. Cuando termino de editar, vuelvo a abrir `REPARADOR_CAT.bat` y selecciono
   **"2. Ya edité el Excel -> generar los cambios"**. Se arma el ISO con los cambios.
   Reviso el informe.
5. Selecciono **"3. Aplicar los cambios a la base"** → opción **"3. Otra"** → pego la
   ruta del **clon** (por ejemplo `C:\Isis\catprueba\catprueba`). Tiene que decir LISTO.
6. (Opcional) Reviso el resultado en la pantalla de ABCD, copiando los archivos del clon
   al servidor y reconstruyendo su índice.
7. Si el clon quedó bien, repito el paso 5, pero eligiendo esta vez la opción
   **"1. La copia local de CAT"**, para aplicarlo a CAT real. Tiene que decir LISTO.
8. Copio `cat.mst` y `cat.xrf` (ya modificados) al servidor de ABCD, reemplazando los que
   están ahí.
9. Reconstruyo el índice de CAT en ABCD.
10. Cambios aplicados. Retomo la catalogación.

### Exportar otra base (el clon, una carpeta de prueba)

El paso 2 solo exporta la base CAT real — es lo que se usa siempre para preparar una
corrección. Si alguna vez hace falta exportar *otra* base (por ejemplo para revisar el
clon después de aplicarle cambios), se hace desde la línea de comandos, no desde el
menú:

```
python reparador.py --base "C:\Isis\catprueba\catprueba"
```

Y ahí sí, opción 1 del menú exporta esa ruta en vez de CAT real.

---

### Por qué el paso 6 va primero al clon y no directo a CAT real

No es obligatorio — el mecanismo (paso 8) ya está probado y no rompe nada. Pero probar
antes en el clon (paso 6-7) permite detectar un error de tipeo o un cambio mal hecho en
el Excel antes de que llegue a la base real. Cuanto más grande el lote de cambios, más
vale la pena. Para una corrección chica se puede saltear e ir directo al paso 8.

### Antes de un lote grande: refrescar el clon

Si pasó tiempo desde la última vez que se usó el clon (se siguió catalogando en ABCD
mientras tanto), conviene recrearlo primero: copiar `cat.mst`/`cat.xrf` reales sobre los
del clon en el servidor, reconstruir su índice, y traer esos dos archivos frescos a la
carpeta local del clon. Si el clon es reciente, se puede saltear.

### Importante

- No se toca `cat` real hasta el paso 7.
- Si en el informe del paso 5 aparece `v401_no_declarado` o `v110_no_declarado` en algún
  registro que sí edité por otro campo, no es una alarma: es un efecto de sistema ya
  conocido y documentado en `PROCEDIMIENTO.md` §2, sin impacto en la catalogación.
- Este archivo es la versión corta. El detalle completo está en `PROCEDIMIENTO.md`.
