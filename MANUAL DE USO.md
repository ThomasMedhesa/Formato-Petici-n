# Manual de uso — Formato de Petición

Guía breve para cualquier persona que necesite preparar un archivo de petición de electricidad.
No hace falta saber nada de programación ni de Excel: solo copiar una plantilla, pegar los datos y pulsar un botón.

---

## 1. ¿Qué hace esta herramienta?

Recibe tu Excel con los datos de los clientes y te devuelve **otro Excel ya maquetado y listo para enviar**, con:

- Los datos de cada cliente separados y bien delimitados.
- Una fila en blanco entre cliente y cliente.
- Los cálculos de consumo y las sumas por cliente ya escritos.
- Una fila final de **TOTAL** con la suma de todos los clientes.

**Importante:** tu archivo original **no se modifica nunca**. La herramienta siempre crea un archivo nuevo.

---

## 2. Antes de empezar: los dos archivos que necesitas

| Archivo | Para qué sirve |
|---|---|
| `FormatearPeticion.exe` | El programa. Solo hay que abrirlo, no se instala. |
| `EJEMPLO INICIAL.xlsx` | La plantilla donde copias y pegas tus datos. **No lo rellenes encima**: cópialo primero con otro nombre. |

> Los dos archivos deben estar en la misma carpeta.

---

## 3. Paso a paso

### Paso 1 — Abre el programa
Haz doble clic en **`FormatearPeticion.exe`**. Se abrirá una ventana pequeña con tres campos.

### Paso 2 — Prepara tu Excel de trabajo
1. Abre `EJEMPLO INICIAL.xlsx`.
2. **Guarda una copia** con el nombre de tu petición (por ejemplo `Peticion_Enero.xlsx`). Trabaja siempre en la copia.
3. Borra los datos de ejemplo que vienen a partir de la **fila 7**.
4. Pega o escribe tus datos a partir de la **fila 7**.

### Paso 3 — Rellena los datos
Respeta la estructura de la plantilla (ver el punto 5 de este manual). Cada cliente ocupa **varias filas seguidas**:

- **La primera fila del cliente** lleva todos sus datos (nombre, CIF, código, dirección, tarifa, potencias…).
- **Las filas siguientes** llevan solo el periodo y los consumos (`Periodo`, `E1`…`E6`).

### Paso 4 — Formatea
En la ventana del programa:

1. En **"Archivo Excel de entrada"**, pulsa **Examinar...** y elige tu archivo (`Peticion_Enero.xlsx`).
2. En **"Carpeta del resultado"**, pulsa **Examinar...** y elige dónde quieres guardar el resultado.
3. Pulsa **"Formatear archivo"**.

Tarda unos segundos. Cuando aparezca el aviso **"Formato completado"**, ya está listo.

### Paso 5 — Abre el resultado
Se ha creado un archivo nuevo llamado **`Peticion_Enero_peticion.xlsx`** en la carpeta que elegiste. Ábrelo, revísalo y envíalo.

Si ese archivo ya existe, el programa te preguntará si quieres reemplazarlo.

---

## 4. Qué hace la herramienta por ti

| Automatismo | Resultado |
|---|---|
| Detecta dónde acaba un cliente y empieza otro | Un bloque de filas limpio por cliente |
| Coloca los datos del cliente solo en la primera fila de su bloque | Evita repeticiones y errores de copiado |
| Escribe la fórmula del consumo total anual | Columna **X** |
| Escribe la fórmula de la suma de cada cliente | Columna **X**, primera fila de cada bloque |
| Añade la fila final **TOTAL** | Suma de todos los clientes |
| Copia colores, bordes y anchos de la plantilla oficial | Mismo aspecto que los formatos anteriores |
| Ajusta la altura de las filas y los saltos de página | Todo alineado, listo para imprimir |

Tus fórmulas del archivo original se **adaptan automáticamente** a las nuevas filas, así que no hay que corregirlas a mano.

---

## 5. Cómo debe ser tu Excel de entrada

### Reglas obligatorias

1. **Una sola hoja**, y debe llamarse exactamente **`PETICION OFERTAS`**.
2. **Los datos empiezan en la fila 7.** Las filas 1 a 6 son cabeceras: no las borres ni las modifiques.
3. **Los datos solo pueden estar entre las columnas B y AG.** La columna A y las columnas siguientes a AG deben estar **vacías**.
4. **No dejes filas totalmente vacías en mitad de los datos.** Si hay un hueco, el programa lo considera el final de la tabla y da error.
5. El archivo debe ser **`.xlsx`** (no `.xls` ni `.csv`).
6. Un mismo cliente debe estar **en filas consecutivas**, sin clientes intercalados.

### Estructura de cada cliente

```
Fila 1 del cliente  →  Cliente | CIF | Código | Dirección | C.P. | Población |
                        Provincia | CUPS | Tarifa | P1..P6 | Periodo | E1..E6 |
Fila 2 del cliente  →  (vacío)   | Periodo | E1..E6
Fila 3 del cliente  →  (vacío)   | Periodo | E1..E6
Siguiente cliente   →  vuelve a empezar por la fila 1 del cliente
```

> La herramienta agrupa las filas por **Cliente + CIF + Código**. Si esos tres datos son idénticos, las filas se consideran del mismo cliente.

### Referencia de columnas

| Columna | Contenido |
|---|---|
| B | Cliente |
| C | C.I.F. |
| D | Código |
| E | Dirección |
| F | C.P. |
| G | Población |
| H | Provincia |
| I | C.U.P.S. |
| J | Tarifa |
| K – P | Potencia Contratada (P1 … P6) |
| Q | Periodo |
| R – W | Consumo Energía Activa (E1 … E6) |
| X | Consumo Total (kWh/año) — **se calcula solo** |
| Y | Inicio contrato |
| Z – AE | Precio neto ofertado (P1 … P6) |
| AF | FEE |
| AG | Desvíos |

Las columnas **X** (consumo total) y **W** del total no hace falta rellenarlas: el programa escribe las fórmulas por ti.

---

## 6. Mensajes de error y cómo solucionarlos

| Mensaje | Causa | Solución |
|---|---|---|
| *No se encontraron datos a partir de la fila 7.* | No hay datos, o están en una fila anterior. | Comprueba que los datos empiezan en la fila 7. |
| *El archivo de entrada debe contener una sola hoja.* | Hay más de una hoja en el archivo. | Borra las hojas sobrantes; solo debe quedar `PETICION OFERTAS`. |
| *Ambos archivos deben contener la hoja «PETICION OFERTAS».* | El nombre de la hoja no es exacto. | Renómbrala exactamente a `PETICION OFERTAS` (sin tildes ni espacios al final). |
| *Se encontraron datos fuera de las columnas B:AG.* | Hay algo escrito en la columna A, en AH o posterior, o en una fila anterior a la 7. | Borra el contenido sobrante. El aviso indica la celda exacta. |
| *El archivo de salida debe tener extensión .xlsx.* | El archivo no es `.xlsx`. | Usa un Excel `.xlsx` (no `.xls`). |
| *Ya existe el archivo de salida.* | Ya había un resultado con ese nombre. | Responde **Sí** para reemplazarlo, o cambia la carpeta de destino. |
| *No se pudo adaptar la fórmula de …* | Una fórmula del archivo original no se puede mover de fila. | Simplifica o corrige esa fórmula en el archivo de entrada. |
| *El archivo no es un Excel válido.* | Archivo dañado o con contraseña. | Vuelve a guardarlo desde Excel sin contraseña. |

---

## 7. Preguntas frecuentes

**¿Puedo usar el archivo de ejemplo tal cual, sin tocar nada?**
Sí. `EJEMPLO INICIAL.xlsx` ya tiene datos de muestra: puedes probarlo para ver el resultado antes de usar datos reales.

**¿Se puede usar más de una vez?**
Sí, tantas veces como quieras.

**¿Qué pasa si meto 200 clientes?**
La herramienta insertará automáticamente la fila en blanco entre clientes y adjusts la maquetación. No hay límite práctico.

**¿Puedo abrir el archivo original mientras la herramienta trabaja?**
Sí. El programa crea una copia interna del archivo para trabajar; tu archivo queda libre. Aun así, te recomendamos cerrarlo antes por seguridad.

**Me he equivocado en un dato, ¿tengo que empezar de cero?**
No. Corrige el dato en tu archivo de trabajo y vuelve a pulsar **"Formatear archivo"**. Se generará un archivo nuevo.

**¿Se guardan mis datos en algún sitio?**
No. Todo el proceso ocurre en tu propio equipo. No se envía nada por Internet.

---

## 8. Consejo rápido

> Trabaja siempre sobre una **copia** de `EJEMPLO INICIAL.xlsx`, pega los datos de tus clientes a partir de la **fila 7** respetando el orden de las columnas, y comprueba que cada cliente va en filas **seguidas** sin huecos.
>
> Con eso, el archivo de salida queda correcto.
