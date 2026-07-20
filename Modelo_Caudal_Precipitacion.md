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

## 2) Caudal vs ALCANCE  (Q ∝ Rⁿ)  ← NO es universal

Al reducir el radio, el caudal baja con **Rⁿ**, pero el exponente **n depende de
la familia** de boquilla (lo verifiqué contra R‑VAN, HE‑VAN y VAN):

| familia | exponente n | por qué |
|---------|:-----------:|---------|
| **R‑VAN** | **≈ 2.0** | precipitación constante (matched) |
| **HE‑VAN** | ≈ 0.7 – 1.15 | precipitación baja con la presión |
| **VAN** (spray) | ≈ 0.76 – 0.84 | precipitación baja con la presión |

```
Q(R) = Q_nom · (R / R_nom)ⁿ
```

**Por eso el n = 2 solo vale para R‑VAN.** Para HE‑VAN o VAN, calcule n con dos
filas del catálogo de esa boquilla:

```
n = ln(Q₂/Q₁) / ln(R₂/R₁)
```

Ejemplo: si un R‑VAN se reduce al 80 % del alcance (f = 0.8), su caudal queda en
**f² = 0.64 → 64 %**; pero un HE‑VAN (n≈0.75) quedaría en **0.8^0.75 = 85 %**.
La forma robusta y válida para cualquier boquilla es **interpolar** el caudal
entre dos filas (presión, radio, caudal) del catálogo — función `CaudalInterp`.

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

## ¿Una fórmula por modelo, o una sola?

- **Arco y precipitación (cuadro/triángulo): una sola, universal.** La misma
  del R‑VAN encaja exacto con HE‑VAN y VAN (verificado). El factor
  triángulo/cuadro = 2/√3 se cumple en todas.
- **Radio → caudal: NO universal.** El exponente n cambia por familia. No se
  inventa una fórmula por modelo: se guardan **dos filas (presión, radio,
  caudal)** por boquilla en el catálogo y se **interpola** (`CaudalInterp`),
  o se calcula n con `ExponenteRadio`. Así vale para R‑VAN, HE‑VAN, VAN y
  cualquier tobera futura sin re‑derivar nada.

## Resumen operativo

| Dato buscado | Fórmula | Universal |
|---|---|:---:|
| Caudal a otro arco | `Q = Q360·(arco/360)` | ✅ |
| Caudal a radio reducido | `Q = Q_nom·(R/R_nom)ⁿ`  (n por familia) | ❌ |
| Exponente n | `n = ln(Q₂/Q₁)/ln(R₂/R₁)` (2 filas del catálogo) | ✅ método |
| Precip. cuadro | `PR_c = 60·Q/S²` | ✅ |
| Precip. triángulo | `PR_t = 1.1547·PR_c` | ✅ |
