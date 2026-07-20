Attribute VB_Name = "modCaudalAspersor"
Option Explicit
'==============================================================================
' MODELO DE CAUDAL Y PRECIPITACION PARA ASPERSORES (toberas MPR tipo R-VAN)
'------------------------------------------------------------------------------
' Verificado contra R-VAN14/18/24, HE-VAN 8/10/12/15 y VAN 4/6/8/10/12.
'
' UNIVERSAL (misma formula para todas las boquillas):
'   1) Caudal vs ARCO      :  Q = Q360 * (arco/360)
'   4) Precipitacion cuadro:  PR_c = 60*Q / S^2            (S = separacion, m)
'      Precipitacion triang:  PR_t = PR_c * 2/sqrt3  (~1.1547)
'
' NO universal (depende de la familia de boquilla):
'   2) Caudal vs ALCANCE   :  Q = Q_nom * (R/R_nom)^n
'      El exponente n NO es fijo:
'        - R-VAN  : n ~ 2.0  (precipitacion constante / matched)
'        - HE-VAN : n ~ 0.7 a 1.15
'        - VAN     : n ~ 0.76 a 0.84
'      Por eso NO se debe usar n=2 con HE-VAN ni VAN. Se calcula n de DOS
'      filas del catalogo con ExponenteRadio(), o se pasa a CaudalAspersor.
'
' Sirve para calcular el caudal de CUALQUIER arco (125, 264, 52 grados...) y de
' cualquier alcance, y su precipitacion en cuadro y triangulo.
' Ver "Modelo_Caudal_Precipitacion.md" para la derivacion y la verificacion.
'==============================================================================

Private Const SQRT3 As Double = 1.73205080756888   ' raiz de 3

'------------------------------------------------------------------------------
' CAUDAL efectivo (l/min) para un arco y un alcance de uso dados.
'   q360Nom : caudal nominal a 360 grados (l/min) del catalogo
'   rNom    : alcance nominal (m) al que corresponde q360Nom
'   arcoDeg : angulo del sector a usar (grados, 1..360) - cualquier valor
'   rUso    : alcance realmente usado (m); si = rNom no hay reduccion
'------------------------------------------------------------------------------
'   expRadio: exponente n de la ley Q ~ R^n. Por defecto 2 (R-VAN). Para
'             HE-VAN / VAN calcule n con ExponenteRadio() y paselo aqui.
Public Function CaudalAspersor(q360Nom As Double, rNom As Double, _
                               arcoDeg As Double, rUso As Double, _
                               Optional expRadio As Double = 2#) As Double
    If rNom <= 0# Or q360Nom <= 0# Then Exit Function
    Dim fArco As Double, fRad As Double
    fArco = arcoDeg / 360#
    If fArco < 0# Then fArco = 0#
    If fArco > 1# Then fArco = 1#
    fRad = rUso / rNom
    CaudalAspersor = q360Nom * fArco * (fRad ^ expRadio)
End Function

'------------------------------------------------------------------------------
' EXPONENTE n de la ley Q ~ R^n, a partir de DOS filas del catalogo de la
' MISMA boquilla (dos presiones): (r1,q1) y (r2,q2). Model-agnostic.
'   R-VAN ~ 2.0 ; HE-VAN ~ 0.7-1.15 ; VAN ~ 0.76-0.84
'------------------------------------------------------------------------------
Public Function ExponenteRadio(r1 As Double, q1 As Double, _
                               r2 As Double, q2 As Double) As Double
    If r1 <= 0# Or r2 <= 0# Or q1 <= 0# Or q2 <= 0# Or r1 = r2 Then
        ExponenteRadio = 2#            ' sin datos: por defecto matched (R-VAN)
        Exit Function
    End If
    ExponenteRadio = Log(q2 / q1) / Log(r2 / r1)   ' Log = logaritmo natural
End Function

'------------------------------------------------------------------------------
' CAUDAL por INTERPOLACION de dos filas del catalogo (lo mas exacto y valido
' para CUALQUIER boquilla): ajusta la ley Q ~ R^n con (r1,q1),(r2,q2) y evalua
' en rUso, luego aplica el arco.  q1,q2 son caudales a 360 grados.
'------------------------------------------------------------------------------
Public Function CaudalInterp(r1 As Double, q1 As Double, r2 As Double, q2 As Double, _
                             arcoDeg As Double, rUso As Double) As Double
    Dim n As Double: n = ExponenteRadio(r1, q1, r2, q2)
    CaudalInterp = CaudalAspersor(q1, r1, arcoDeg, rUso, n)
End Function

'------------------------------------------------------------------------------
' Caudal a 360 grados equivalente para un alcance dado (util para catalogos).
'------------------------------------------------------------------------------
Public Function Caudal360(q360Nom As Double, rNom As Double, rUso As Double) As Double
    Caudal360 = CaudalAspersor(q360Nom, rNom, 360#, rUso)
End Function

'------------------------------------------------------------------------------
' Constante de la familia k = Q360 / R^2  (l/min por m^2). PR_cuadro = 60*k.
'------------------------------------------------------------------------------
Public Function ConstanteK(q360Nom As Double, rNom As Double) As Double
    If rNom <= 0# Then Exit Function
    ConstanteK = q360Nom / (rNom * rNom)
End Function

'------------------------------------------------------------------------------
' PRECIPITACION (mm/h) en malla CUADRADA con separacion sepM (m).
'   Marco a marco (cabeza a cabeza) => sepM = alcance.
'   Q en l/min, sep en m -> mm/h (1 l/m^2 = 1 mm).
'------------------------------------------------------------------------------
Public Function PrecipCuadro(qLmin As Double, sepM As Double) As Double
    If sepM <= 0# Then Exit Function
    PrecipCuadro = 60# * qLmin / (sepM * sepM)
End Function

'------------------------------------------------------------------------------
' PRECIPITACION (mm/h) en malla TRIANGULAR (tresbolillo) con separacion sepM.
'   Area por aspersor = sep^2 * sqrt(3)/2.
'------------------------------------------------------------------------------
Public Function PrecipTriangulo(qLmin As Double, sepM As Double) As Double
    If sepM <= 0# Then Exit Function
    PrecipTriangulo = 60# * qLmin / (sepM * sepM * (SQRT3 / 2#))
End Function

'------------------------------------------------------------------------------
' Conversion directa entre precipitaciones (mismo espaciamiento):
'   cuadro -> triangulo : * 2/sqrt3 (~1.1547)
'   triangulo -> cuadro : * sqrt3/2 (~0.8660)
'------------------------------------------------------------------------------
Public Function CuadroATriangulo(prCuadro As Double) As Double
    CuadroATriangulo = prCuadro * (2# / SQRT3)
End Function

Public Function TrianguloACuadro(prTriangulo As Double) As Double
    TrianguloACuadro = prTriangulo * (SQRT3 / 2#)
End Function

'==============================================================================
' AUTO-VERIFICACION contra el cuadro R-VAN14 (ejecutar para comprobar).
'==============================================================================
Public Sub VerificarModeloRVAN14()
    ' Nominal 3.1 bar: R = 4.3 m, Q360 = 4.81 l/min
    Dim rNom As Double, q360 As Double
    rNom = 4.3: q360 = 4.81

    Dim s As String
    s = "VERIFICACION R-VAN14 (R=4.3 m, Q360=4.81 l/min, 3.1 bar)" & vbCrLf & _
        String(50, "-") & vbCrLf & _
        "Caudal por arco (calc vs tabla):" & vbCrLf & _
        "  270: " & Format(CaudalAspersor(q360, rNom, 270, rNom), "0.00") & "  (tabla 3.56)" & vbCrLf & _
        "  210: " & Format(CaudalAspersor(q360, rNom, 210, rNom), "0.00") & "  (tabla 2.76)" & vbCrLf & _
        "  180: " & Format(CaudalAspersor(q360, rNom, 180, rNom), "0.00") & "  (tabla 2.38)" & vbCrLf & _
        "   90: " & Format(CaudalAspersor(q360, rNom, 90, rNom), "0.00") & "  (tabla 1.21)" & vbCrLf & vbCrLf & _
        "Arco NO tabulado (ejemplos):" & vbCrLf & _
        "  125: " & Format(CaudalAspersor(q360, rNom, 125, rNom), "0.00") & " l/min" & vbCrLf & _
        "  264: " & Format(CaudalAspersor(q360, rNom, 264, rNom), "0.00") & " l/min" & vbCrLf & _
        "   52: " & Format(CaudalAspersor(q360, rNom, 52, rNom), "0.00") & " l/min" & vbCrLf & vbCrLf & _
        "Reduccion de alcance (360 grados):" & vbCrLf & _
        "  R=4.0 -> Q=" & Format(CaudalAspersor(q360, rNom, 360, 4#), "0.00") & "  (tabla 4.16)" & vbCrLf & _
        "  R=4.6 -> Q=" & Format(CaudalAspersor(q360, rNom, 360, 4.6), "0.00") & "  (tabla 5.49)" & vbCrLf & vbCrLf & _
        "Precipitacion (360, R=4.3, Q=4.81):" & vbCrLf & _
        "  cuadro:    " & Format(PrecipCuadro(4.81, 4.3), "0.0") & " mm/h  (tabla 16)" & vbCrLf & _
        "  triangulo: " & Format(PrecipTriangulo(4.81, 4.3), "0.0") & " mm/h  (tabla 18)" & vbCrLf & vbCrLf & _
        "EXPONENTE n (Q~R^n) por familia -- NO es universal:" & vbCrLf & _
        "  R-VAN14 : n = " & Format(ExponenteRadio(4#, 4.16, 4.6, 5.49), "0.00") & "  (~2 matched)" & vbCrLf & _
        "  HE-VAN8 : n = " & Format(ExponenteRadio(1.5, 3.14, 2.4, 4.43), "0.00") & vbCrLf & _
        "  VAN10   : n = " & Format(ExponenteRadio(2.1, 7.3, 3.1, 9.8), "0.00") & vbCrLf & _
        "  -> use ExponenteRadio() / CaudalInterp() para HE-VAN y VAN."
    MsgBox s, vbInformation, "Modelo de caudal y precipitacion"
End Sub
