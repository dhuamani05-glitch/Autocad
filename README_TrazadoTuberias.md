# modTrazadoTuberias — Trazado de tuberías de riego (VBA AutoCAD)

Complemento de `modDisenoAspersion` (distribución de aspersores). Traza la red
de tuberías que conecta **todos los aspersores** con un punto de **fuente**
(cabezal/válvula) usando **polilíneas**, con la **menor longitud total posible**.

## ¿Por qué es lo más eficiente posible?

Conectar N puntos con tubería, sin bucles (un árbol), gastando la menor
cantidad de tubería, es el problema del **Árbol de Expansión Mínima (MST)**.
El módulo lo resuelve con el algoritmo de **Prim**, que devuelve el árbol
**óptimo** (de longitud total mínima), no una aproximación.

Además dimensiona la red:

- Enraíza el árbol en la fuente y calcula el **caudal aguas abajo** de cada
  tramo (suma de los caudales de los aspersores que dependen de él).
- Elige el **diámetro comercial más pequeño** que respeta una **velocidad
  máxima** (por defecto 1.5 m/s) → la red más económica que cumple el criterio.
- Cada tramo se dibuja en una capa por diámetro (`RIEGO_TUB_16`, `_20`, …) y
  opcionalmente se rotula con diámetro y caudal.

## Uso

1. Ejecuta primero `modDisenoAspersion` para colocar los aspersores (o ten los
   bloques `ASPERSOR_RIEGO_*` en el dibujo).
2. Importa `modTrazadoTuberias.bas` en el Editor de VBA (Alt+F11 → Archivo →
   Importar), o pégalo en un módulo nuevo.
3. Ejecuta la macro **`TrazadoTuberias`**:
   - (Opcional) Escribe una **ZONA** para trazar solo esa válvula; vacío = todos.
   - Indica con el mouse el **punto de fuente** (cabezal/válvula).
   - Ingresa la **velocidad máxima** (Enter = 1.5 m/s).
   - Elige si **rotular** los tramos (S/N).

## Salida

- Polilíneas de tubería en capas `RIEGO_TUB_<diámetro>` (color por diámetro).
- Rótulos opcionales en `RIEGO_TUB_TXT`.
- Reporte: nº de tramos, longitud total, longitud por diámetro y caudal en la
  fuente. Usa **DATAEXTRACTION** filtrando por capa para la lista de materiales
  (metros por diámetro).

## Detalles técnicos

- **Detección de aspersores:** bloques cuyo nombre contiene `ASPERSOR`; lee los
  atributos `CAUDAL`, `NUM` y `ZONA`. Si no encuentra ninguno, permite
  **selección manual** (bloques, círculos o puntos).
- **Caudal por tramo:** acumulación de subárbol enraizado en la fuente.
- **Complejidad:** Prim O(n²), apto para cientos/miles de aspersores.
