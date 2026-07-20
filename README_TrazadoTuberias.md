# modTrazadoTuberias — Trazado de tuberías de riego (VBA AutoCAD)

Complemento de `modDisenoAspersion`. Traza la red de tuberías que conecta los
aspersores con un punto de **fuente** (cabezal/válvula) usando **polilíneas**,
y la dimensiona según las mejores prácticas de riego.

## Disposiciones (el usuario elige al ejecutar)

1. **Troncal única (recomendada)** — una sola línea principal continua recorre
   los aspersores en secuencia (**mínima ramificación → ideal para zanja de
   retroexcavadora**). El cabezal **no** entra de frente a un aspersor: se conecta
   con una **tee en el punto más cercano** de la troncal, y el caudal se reparte
   hacia los dos lados (favorece el 20%).
2. **Anillo (looped main)** — perímetro en bucle; caudal repartido en dos ramas
   → presión muy uniforme.
3. **Árbol (MST, Prim)** — mínima longitud total de tubería (menor material).
4. **Principal + laterales** (opcional) — troncal con laterales perpendiculares.

**Diámetro mínimo de lateral = 25 mm** (no se usan 16 ni 20 mm). Catálogo:
25, 32, 40, 50, 63, 75, 90, 110, 125, 160 mm.

## Criterio hidráulico (regla del 20%) con desnivel

En diseño de riego la variación de presión dentro del sector no debe superar
~20% de la presión nominal del emisor. El módulo:

- Calcula la **pérdida de carga Hazen-Williams** en cada tramo con el caudal
  acumulado aguas abajo.
- Incluye el **desnivel**: la variación de presión = fricción + diferencia de
  cota entre emisores. La cota sale de la **Z de los bloques** (si el dibujo es
  3D) y/o de una **pendiente de terreno** (%) y **azimut de subida** que se
  piden al ejecutar.
- Dimensiona el diámetro de cada tramo partiendo del **mínimo por velocidad** y
  **agranda el tramo más crítico** hasta que la variación ≤ % admisible.
- El reporte indica la variación obtenida, el aporte del desnivel y
  **CUMPLE / NO CUMPLE**. Si el desnivel por sí solo supera el %, avisa que
  ninguna tubería lo corrige (sectorizar o regular presión).

## Comparativo automático

Antes de dibujar, el módulo **evalúa las 3 disposiciones** (sin dibujarlas) y
muestra una tabla con: longitud total, variación de presión (m y %), índice de
costo (longitud × diámetro) y CUMPLE/FALLA. Marca la **recomendada** (menor
costo entre las que cumplen) y te deja elegir cuál dibujar (Enter = recomendada).

## Uso

1. Coloca los aspersores con `modDisenoAspersion` (bloques `ASPERSOR_RIEGO_*`).
2. Importa `modTrazadoTuberias.bas` en el editor VBA (Alt+F11 → Importar).
3. Ejecuta **`TrazadoTuberias`**:
   - Confirma los aspersores detectados.
   - Indica el **punto de fuente**.
   - Velocidad máx (1.5), presión nominal (m.c.a.), % admisible (20), C (150).
   - Elige **disposición** (1=Anillo, 2=Principal+laterales, 3=MST).
   - Rótulos S/N.

## Salida

- Polilíneas en capas `RIEGO_TUB_<diámetro>` (color por diámetro).
- Rótulos opcionales en `RIEGO_TUB_TXT`; marcador de fuente en `RIEGO_FUENTE`.
- Reporte: disposición, longitud total y por diámetro, caudal en la fuente y
  verificación del criterio de presión. Usa **DATAEXTRACTION** por capa para la
  lista de materiales (metros por diámetro).

## Notas técnicas

- **Detección de aspersores**: bloque cuyo nombre contiene `ASPERSOR`, **o** capa
  `RIEGO_ASPERSOR`, **o** con atributo `NUM`/`CAUDAL`. Fallback: selección manual
  o por círculos.
- **Filtro por zona**: solo se ofrece si hay varias `ZONA` distintas.
- El anillo usa un modelo de reparto de caudal equilibrado (dos vías) — es una
  aproximación de diseño, no un cálculo iterativo de redes malladas (Hardy-Cross).
- No considera desnivel (terreno plano). Puede añadirse leyendo la Z de los
  bloques si el dibujo es 3D.
