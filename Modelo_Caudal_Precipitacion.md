# Modelo matemático de caudal y precipitación (toberas R‑VAN y similares)

Derivado y **verificado numéricamente** contra los cuadros R‑VAN14 / R‑VAN18 /
R‑VAN24 (Rain Bird). Sirve para calcular el caudal de un aspersor con **cualquier
arco** (125°, 264°, 52°…) y **cualquier alcance reducido**, y su precipitación en
malla cuadrada y triangular — datos que las tablas solo dan para arcos fijos.

## Idea clave: precipitación ajustada (MPR)

En estas toberas la **precipitación es constante** (~16 mm/h cuadro, ~18 mm/h
triángulo) para todos los sectores, presiones y alcances. Eso es diseño
*Matched Precipitation Rate*. De ahí salen dos proporcionalidades exactas.

## 1) Caudal vs ARCO  (Q ∝ arco)

A igual presión y alcance, el caudal es proporcional al ángulo del sector:

```
Q(arco) = Q360 · (arco / 360)
```

Verificado (R‑VAN14, 3.1 bar, R=4.3, Q360=4.81):

| arco | calc = Q360·arco/360 | tabla |
|-----:|---------------------:|------:|
| 270° | 3.61 | 3.56 |
| 210° | 2.81 | 2.76 |
| 180° | 2.40 | 2.38 |
|  90° | 1.20 | 1.21 |

Para arcos NO tabulados (125°, 264°, 52°) se usa la misma proporción directa.

## 2) Caudal vs ALCANCE  (Q ∝ R²)  ← lo que faltaba

Al **reducir el radio nominal** (por presión o por el tornillo de reducción), el
caudal baja con el **cuadrado** del radio (porque la precipitación se mantiene):

```
Q(R) = Q_nom · (R / R_nom)²
```

Verificado tomando filas de distinto alcance del mismo modelo:

| modelo | Q₂/Q₁ (tabla) | (R₂/R₁)² |
|--------|--------------:|---------:|
| R‑VAN14 | 1.320 | 1.322 |
| R‑VAN18 | 1.278 | 1.260 |
| R‑VAN24 | 1.591 | 1.584 |

Es decir, si un aspersor se reduce al 80 % de su alcance (f = 0.8), su caudal
queda en **f² = 0.64 → 64 %** del nominal, no en 80 %.

## 3) Fórmula general (arco + alcance combinados)

```
Q = Q360_nom · (arco / 360) · (R_uso / R_nom)²
```

Con la constante de la familia `k = Q360_nom / R_nom²` (≈ 0.26 l/min·m⁻² para
R‑VAN, equivale a PR ≈ 16 mm/h) también se puede escribir:

```
Q = k · (arco / 360) · R_uso²
```

## 4) Precipitación: cuadro vs triángulo

La precipitación (mm/h) con marco a marco (separación S = R) es:

```
Cuadro     :  PR_c = 60 · Q / S²
Triángulo  :  PR_t = 60 · Q / (S² · √3/2) = PR_c · (2/√3) ≈ 1.155 · PR_c
```

- Pasar de **cuadro a triángulo**: × 2/√3 ≈ **×1.1547**
- Pasar de **triángulo a cuadro**: × √3/2 ≈ **×0.8660**

(El triángulo da más precipitación a igual separación porque empaqueta más
aspersores por área.) Q en l/min, S en m → PR en mm/h (1 l/m² = 1 mm).

Verificado (R‑VAN14 360°, R=4.3, Q=4.81): PR_c = 60·4.81/4.3² = **15.6 → 16**;
PR_t = 15.6·2/√3 = **18.0 → 18**. Coincide con la tabla.

Para un aspersor de arco parcial la precipitación **local** en el sector regado
es la misma que la del círculo completo (por eso la tabla repite 16/18 en todos
los sectores): el arco reduce el caudal y el área regada en la misma proporción.

## Resumen operativo

| Dato buscado | Fórmula |
|---|---|
| Caudal a otro arco | `Q = Q360·(arco/360)` |
| Caudal a radio reducido | `Q = Q_nom·(R/R_nom)²` |
| Caudal general | `Q = Q360_nom·(arco/360)·(R/R_nom)²` |
| Precip. cuadro | `PR_c = 60·Q/S²` |
| Precip. triángulo | `PR_t = 60·Q/(S²·√3/2) = 1.1547·PR_c` |
