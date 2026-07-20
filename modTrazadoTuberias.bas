Attribute VB_Name = "modTrazadoTuberias"
Option Explicit
'==============================================================================
' TRAZADO DE TUBERIAS DE RIEGO MEDIANTE POLILINEAS  (v3)
'------------------------------------------------------------------------------
' Complemento de "modDisenoAspersion". Conecta los aspersores colocados
' (bloques ASPERSOR_RIEGO_* en capa RIEGO_ASPERSOR) con un punto de FUENTE
' (cabezal / valvula) trazando la red con POLILINEAS y dimensionandola segun
' las mejores practicas de riego.
'
' DISPOSICIONES:
'   1 = ANILLO (looped main)  - perimetro en bucle, caudal repartido en dos
'       ramas -> presion mas uniforme. Recomendado para aspersores en perimetro.
'   2 = PRINCIPAL + LATERALES - troncal en el eje dominante con laterales
'       perpendiculares (pocos emisores en serie).
'   3 = ARBOL (MST, Prim)     - minima longitud total de tuberia.
'
' CRITERIO HIDRAULICO (regla del 20%):
'   La variacion de presion entre emisores del sector <= ~20% de la presion
'   nominal. Se dimensiona el diametro de cada tramo (Hazen-Williams) partiendo
'   del minimo por velocidad y agrandando el tramo mas critico hasta cumplir.
'   La variacion incluye FRICCION + DESNIVEL (cota Z de los bloques y/o una
'   pendiente de terreno indicada por el usuario).
'
' COMPARATIVO AUTOMATICO:
'   Evalua las 3 disposiciones (longitud, variacion de presion, indice de costo,
'   cumple/falla), muestra la tabla y dibuja la que el usuario elija (por
'   defecto la recomendada: menor costo entre las que cumplen).
'
' CAPAS: RIEGO_TUB_<diametro>, RIEGO_TUB_TXT (rotulos), RIEGO_FUENTE (marcador).
'==============================================================================

Private Const PI As Double = 3.14159265358979

'--- Aspersores detectados (indice 0 = FUENTE tras InsertarFuente) ------------
Private pX() As Double
Private pY() As Double
Private pZ() As Double          ' cota Z (desnivel) del bloque / fuente
Private pQ() As Double          ' demanda propia del nodo (l/min); fuente = 0
Private pNum() As String        ' etiqueta NUM del aspersor
Private pZona() As String       ' atributo ZONA del aspersor
Private pElevRel() As Double    ' elevacion relativa a la fuente (Z + pendiente)
Private nP As Long

'--- Red construida como ARBOL enraizado en la fuente (nodo 0) ----------------
Private gNX() As Double
Private gNY() As Double
Private gDem() As Double
Private gPar() As Long          ' padre (gPar(0) = -1 = raiz)
Private gOrd() As Long          ' orden topologico (padre antes que hijo)
Private gAcc() As Double        ' caudal aguas abajo por nodo (l/min)
Private gSeg() As Long          ' indice de diametro del tramo que llega al nodo
Private gNN As Long
'--- Aristas EXTRA que cierran bucles (anillo) -------------------------------
Private gExA() As Long
Private gExB() As Long
Private gNEx As Long

'--- Catalogo de diametros comerciales (mm) ----------------------------------
Private gDiam() As Double
Private gDiamColor() As Long
Private nDiam As Long

'==============================================================================
' PUNTO DE ENTRADA
'==============================================================================
Public Sub TrazadoTuberias()
    On Error GoTo errH

    '======================================================================
    ' 1) RECOLECTAR ASPERSORES
    '======================================================================
    RecolectarAspersores
    If nP = 0 Then
        Dim nBlk As Long, nCir As Long, nPnt As Long
        ContarEntidades nBlk, nCir, nPnt
        Dim op As String
        op = Trim$(InputBox( _
            "No se detectaron aspersores automaticamente." & vbCrLf & _
            "En el ESPACIO MODELO hay:" & vbCrLf & _
            "   - Bloques (INSERT):  " & nBlk & vbCrLf & _
            "   - Circulos:          " & nCir & vbCrLf & _
            "   - Puntos:            " & nPnt & vbCrLf & vbCrLf & _
            "De donde tomo los aspersores?" & vbCrLf & _
            "   1 = TODOS los bloques" & vbCrLf & _
            "   2 = TODOS los circulos" & vbCrLf & _
            "   3 = SELECCIONARLOS a mano", _
            "Trazado de tuberias - diagnostico", "1"))
        Select Case op
            Case "1": ColectarBloques
            Case "2": ColectarCirculos
            Case Else: SeleccionManual
        End Select
    End If
    If nP < 1 Then
        MsgBox "No hay aspersores para conectar.", vbExclamation, "Trazado de tuberias"
        Exit Sub
    End If

    FiltrarPorZona
    If nP < 1 Then Exit Sub

    If MsgBox("Se detectaron " & nP & " aspersores para conectar." & vbCrLf & vbCrLf & _
              "Se pedira el punto de FUENTE (cabezal/valvula). Continuar?", _
              vbOKCancel + vbInformation, "Trazado de tuberias") = vbCancel Then Exit Sub

    '======================================================================
    ' 2) PUNTO DE FUENTE
    '======================================================================
    Dim src As Variant
    On Error Resume Next
    src = ThisDrawing.Utility.GetPoint(, vbCrLf & _
          "Indique el punto de FUENTE (cabezal / valvula): ")
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo errH
    InsertarFuente CDbl(src(0)), CDbl(src(1)), CDbl(src(2))

    '======================================================================
    ' 3) CRITERIOS DE DISENO
    '======================================================================
    Dim vMax As Double
    vMax = Val(InputBox("Velocidad maxima admisible (m/s):" & vbCrLf & _
                        "  1.5 = recomendado PVC/PE", "Criterio hidraulico", "1.5"))
    If vMax <= 0.1 Then vMax = 1.5

    Dim pNom As Double, pctVar As Double, hwC As Double
    pNom = Val(InputBox("Presion NOMINAL del aspersor (m.c.a.):" & vbCrLf & _
                        "  20 m.c.a. = 2.0 bar", "Criterio de presion", "20"))
    If pNom <= 0# Then pNom = 20#
    pctVar = Val(InputBox("Variacion MAXIMA de presion admisible (%):" & vbCrLf & _
                          "  20 % = criterio estandar", "Criterio de presion", "20"))
    If pctVar <= 0# Then pctVar = 20#
    hwC = Val(InputBox("Coeficiente de Hazen-Williams (C):" & vbCrLf & _
                       "  150 = PVC/PE liso  |  140 = PVC usado", "Criterio de presion", "150"))
    If hwC <= 0# Then hwC = 150#

    ' --- DESNIVEL ---
    Dim pend As Double, az As Double
    pend = Val(InputBox("DESNIVEL - pendiente media del terreno (%):" & vbCrLf & _
                        "  0 = plano (o si el dibujo ya trae la cota Z real de" & vbCrLf & _
                        "  cada bloque, dejela en 0 y se usara esa Z).", _
                        "Desnivel", "0"))
    az = Val(InputBox("DESNIVEL - azimut de la SUBIDA (grados):" & vbCrLf & _
                      "  0 = +X (este)   90 = +Y (norte)", "Desnivel", "0"))
    CalcularElevaciones pend, az

    Dim rotular As Boolean
    rotular = (UCase$(Trim$(InputBox("Rotular cada tramo (diametro y caudal)?  (S/N)", _
              "Rotulos", "S"))) = "S")

    CargarCatalogoDiametros

    '======================================================================
    ' 4) COMPARATIVO AUTOMATICO DE LAS 3 DISPOSICIONES
    '======================================================================
    Dim mLong(1 To 3) As Double, mVar(1 To 3) As Double, mCost(1 To 3) As Double
    Dim mCump(1 To 3) As Boolean, mName(1 To 3) As String
    mName(1) = "Anillo (looped main)"
    mName(2) = "Principal + laterales"
    mName(3) = "Arbol (MST)"

    Dim tp As Long, vM As Double, cmpl As Boolean, lt As Double, ct As Double
    For tp = 1 To 3
        Construir tp
        If gNN >= 2 Then
            DimensionarRed vMax, pNom, pctVar, hwC, vM, cmpl, lt, ct
            mVar(tp) = vM: mCump(tp) = cmpl: mLong(tp) = lt: mCost(tp) = ct
        Else
            mVar(tp) = 1E+30
        End If
    Next

    Dim rec As Long: rec = Recomendar(mCump, mCost, mVar)

    ' desnivel solo (referencia): variacion de cota entre emisores
    Dim dzMax As Double, dzMin As Double, i As Long
    dzMax = -1E+30: dzMin = 1E+30
    For i = 1 To nP - 1
        If pElevRel(i) > dzMax Then dzMax = pElevRel(i)
        If pElevRel(i) < dzMin Then dzMin = pElevRel(i)
    Next
    Dim dzVar As Double: dzVar = dzMax - dzMin

    Dim cmp As String
    cmp = "COMPARATIVO DE DISPOSICIONES  (" & (nP - 1) & " aspersores)" & vbCrLf & _
          "Presion nominal " & Format(pNom, "0.0") & " m.c.a.  |  admisible " & _
          Format(pctVar, "0") & "% = " & Format(pctVar / 100# * pNom, "0.00") & " m" & vbCrLf & _
          "Desnivel entre emisores: " & Format(dzVar, "0.00") & " m.c.a." & vbCrLf & _
          String(56, "-") & vbCrLf
    For tp = 1 To 3
        cmp = cmp & tp & ") " & mName(tp) & IIf(tp = rec, "   <== RECOMENDADA", "") & vbCrLf & _
              "     Longitud " & Format(mLong(tp), "0.0") & " m" & _
              "  |  Var.presion " & Format(mVar(tp), "0.00") & " m (" & _
              Format(100# * mVar(tp) / pNom, "0.0") & "%)  " & _
              IIf(mCump(tp), "CUMPLE", "FALLA") & vbCrLf & _
              "     Indice de costo (long x diam): " & Format(mCost(tp), "0") & vbCrLf
    Next
    cmp = cmp & String(56, "-") & vbCrLf & _
          "Recomendada = menor costo entre las que cumplen." & vbCrLf & _
          IIf(dzVar > pctVar / 100# * pNom, _
              "AVISO: el desnivel por si solo supera el " & Format(pctVar, "0") & _
              "%. Ninguna tuberia lo corrige: sectorice o regule presion." & vbCrLf, "")
    MsgBox cmp, vbInformation, "Comparativo de disposiciones"

    Dim elec As String
    elec = Trim$(InputBox(cmp & vbCrLf & "Cual DIBUJAR?  1 / 2 / 3   (Enter = recomendada " & rec & ")", _
                          "Elegir disposicion", CStr(rec)))
    If elec = "" Then elec = CStr(rec)
    Dim chosen As Long: chosen = Val(elec)
    If chosen < 1 Or chosen > 3 Then chosen = rec

    '======================================================================
    ' 5) CONSTRUIR + DIMENSIONAR + DIBUJAR LA ELEGIDA
    '======================================================================
    Construir chosen
    If gNN < 2 Then Exit Sub
    DimensionarRed vMax, pNom, pctVar, hwC, vM, cmpl, lt, ct

    Dim htxt As Double: htxt = AlturaTexto()
    Dim usados(0 To 63) As Boolean
    Dim longD(0 To 63) As Double
    Dim longTot As Double

    CrearCapa "RIEGO_FUENTE", 1
    Dim cf(0 To 2) As Double
    cf(0) = gNX(0): cf(1) = gNY(0): cf(2) = 0#
    Dim mkr As AcadCircle
    Set mkr = ThisDrawing.ModelSpace.AddCircle(cf, htxt * 1.5)
    mkr.Layer = "RIEGO_FUENTE": mkr.color = 1

    For i = 1 To gNN - 1
        DibujarTramo gPar(i), i, gSeg(i), gAcc(i), rotular, htxt, longTot, longD, usados
    Next
    For i = 0 To gNEx - 1
        DibujarTramo gExA(i), gExB(i), 0, 0#, False, htxt, longTot, longD, usados
    Next
    ThisDrawing.Regen acActiveViewport

    '======================================================================
    ' 6) REPORTE
    '======================================================================
    Dim admis As Double: admis = pctVar / 100# * pNom
    Dim pctReal As Double: pctReal = 100# * vM / pNom
    Dim veredicto As String
    If cmpl Then
        veredicto = "CUMPLE  (<= " & Format(pctVar, "0") & "%)"
    Else
        veredicto = "NO CUMPLE - sectorice / acorte / use el ANILLO" & vbCrLf & _
                    "                   o regule presion si domina el desnivel."
    End If

    Dim rep As String
    rep = "TRAZADO DE TUBERIAS COMPLETADO" & vbCrLf & String(46, "-") & vbCrLf & _
          "Disposicion:            " & mName(chosen) & vbCrLf & _
          "Aspersores conectados:  " & (nP - 1) & vbCrLf & _
          "Tramos de tuberia:      " & (gNN - 1 + gNEx) & vbCrLf & _
          "Longitud TOTAL de red:  " & Format(longTot, "0.00") & " m" & vbCrLf & _
          "Velocidad maxima:       " & Format(vMax, "0.0") & " m/s" & vbCrLf & vbCrLf & _
          "LONGITUD POR DIAMETRO:" & vbCrLf
    For i = 0 To nDiam - 1
        If usados(i) Then
            rep = rep & "  D" & Format(gDiam(i), "0") & " mm :  " & _
                  Format(longD(i), "0.00") & " m" & vbCrLf
        End If
    Next
    rep = rep & vbCrLf & _
          "Caudal total en la fuente: " & Format(gAcc(0), "0.0") & " l/min" & vbCrLf & vbCrLf & _
          "CRITERIO DE PRESION (friccion + desnivel):" & vbCrLf & _
          "  Presion nominal:      " & Format(pNom, "0.0") & " m.c.a." & vbCrLf & _
          "  Variacion admisible:  " & Format(pctVar, "0") & " % = " & Format(admis, "0.00") & " m" & vbCrLf & _
          "  Variacion obtenida:   " & Format(vM, "0.00") & " m.c.a. (" & Format(pctReal, "0.0") & " %)" & vbCrLf & _
          "  Aporte del desnivel:  " & Format(dzVar, "0.00") & " m.c.a." & vbCrLf & _
          "  Estado:               " & veredicto & vbCrLf & vbCrLf & _
          "(Friccion Hazen-Williams C=" & Format(hwC, "0") & _
          ". Metros por diametro con DATAEXTRACTION por capa.)"
    MsgBox rep, vbInformation, "Trazado de tuberias"
    Exit Sub

errH:
    If Err.Number = -2147352567 Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Trazado de tuberias"
End Sub

'==============================================================================
' ELEVACION RELATIVA A LA FUENTE (cota Z del bloque + pendiente de terreno)
'==============================================================================
Private Sub CalcularElevaciones(pend As Double, az As Double)
    ReDim pElevRel(nP - 1)
    Dim azr As Double: azr = az * PI / 180#
    Dim i As Long
    For i = 0 To nP - 1
        pElevRel(i) = (pZ(i) - pZ(0)) + _
                      (pend / 100#) * ((pX(i) - pX(0)) * Cos(azr) + (pY(i) - pY(0)) * Sin(azr))
    Next
    pElevRel(0) = 0#
End Sub

'==============================================================================
' DIMENSIONAR LA RED YA CONSTRUIDA (gPar/gOrd/...) POR EL CRITERIO DE PRESION.
' Devuelve variacion de presion (m), cumple, longitud total y un indice de
' costo relativo (suma de longitud x diametro).  Llena gAcc y gSeg.
'==============================================================================
Private Sub DimensionarRed(vMax As Double, pNom As Double, pctVar As Double, hwC As Double, _
                           ByRef varM As Double, ByRef cumple As Boolean, _
                           ByRef longTot As Double, ByRef costo As Double)
    Dim i As Long, k As Long, nd As Long
    ReDim gAcc(gNN - 1)
    For i = 0 To gNN - 1: gAcc(i) = gDem(i): Next
    For k = gNN - 1 To 1 Step -1
        nd = gOrd(k)
        gAcc(gPar(nd)) = gAcc(gPar(nd)) + gAcc(nd)
    Next

    ReDim gSeg(gNN - 1)
    For i = 1 To gNN - 1: gSeg(i) = ElegirDiametro(gAcc(i), vMax): Next

    Dim admis As Double: admis = pctVar / 100# * pNom
    Dim cumHf() As Double: ReDim cumHf(gNN - 1)
    Dim iterD As Long, nMax As Long: nMax = gNN * nDiam + 20
    Dim maxL As Double, minL As Double, mn As Long, carga As Double

    For iterD = 1 To nMax
        ComputarCumHf cumHf, gSeg, gAcc, hwC
        maxL = -1E+30: minL = 1E+30: mn = -1
        For i = 1 To nP - 1
            carga = cumHf(i) + pElevRel(i)
            If carga > maxL Then maxL = carga: mn = i
            If carga < minL Then minL = carga
        Next
        varM = maxL - minL
        If varM <= admis Or mn = -1 Then Exit For
        ' agrandar el tramo de mayor friccion del camino al emisor de mayor carga
        Dim best As Long, bestHf As Double, nn As Long, hh As Double
        best = -1: bestHf = -1#: nn = mn
        Do While nn <> 0
            If gSeg(nn) < nDiam - 1 Then
                hh = HfTramo(nn, gSeg(nn), gAcc, hwC)
                If hh > bestHf Then bestHf = hh: best = nn
            End If
            nn = gPar(nn)
        Loop
        If best = -1 Then Exit For
        gSeg(best) = gSeg(best) + 1
    Next

    ComputarCumHf cumHf, gSeg, gAcc, hwC
    maxL = -1E+30: minL = 1E+30
    For i = 1 To nP - 1
        carga = cumHf(i) + pElevRel(i)
        If carga > maxL Then maxL = carga
        If carga < minL Then minL = carga
    Next
    varM = maxL - minL
    cumple = (varM <= admis + 0.0000001)

    longTot = 0#: costo = 0#
    Dim L As Double
    For i = 1 To gNN - 1
        L = Sqr((gNX(i) - gNX(gPar(i))) ^ 2 + (gNY(i) - gNY(gPar(i))) ^ 2)
        longTot = longTot + L: costo = costo + L * gDiam(gSeg(i))
    Next
    For i = 0 To gNEx - 1
        L = Sqr((gNX(gExB(i)) - gNX(gExA(i))) ^ 2 + (gNY(gExB(i)) - gNY(gExA(i))) ^ 2)
        longTot = longTot + L: costo = costo + L * gDiam(0)
    Next
End Sub

'--- Recomendada: menor costo entre las que cumplen; si ninguna, menor variacion
Private Function Recomendar(cumple() As Boolean, costo() As Double, varM() As Double) As Long
    Dim tp As Long, best As Long: best = -1
    For tp = 1 To 3
        If cumple(tp) Then
            If best = -1 Then
                best = tp
            ElseIf costo(tp) < costo(best) Then
                best = tp
            End If
        End If
    Next
    If best = -1 Then
        best = 1
        For tp = 2 To 3
            If varM(tp) < varM(best) Then best = tp
        Next
    End If
    Recomendar = best
End Function

'==============================================================================
' CONSTRUCTORES DE LA RED (llenan gNX/gNY/gDem/gPar/gOrd/gNN y gEx*)
'==============================================================================
Private Sub Construir(topo As Long)
    Select Case topo
        Case 2: ConstruirPrincipalLaterales
        Case 3: ConstruirArbolMST
        Case Else: ConstruirAnillo
    End Select
End Sub

'--- 3) ARBOL DE EXPANSION MINIMA (Prim) -------------------------------------
Private Sub ConstruirArbolMST()
    gNN = nP
    ReDim gNX(gNN - 1): ReDim gNY(gNN - 1): ReDim gDem(gNN - 1)
    ReDim gPar(gNN - 1): ReDim gOrd(gNN - 1)
    Dim i As Long
    For i = 0 To nP - 1
        gNX(i) = pX(i): gNY(i) = pY(i): gDem(i) = pQ(i)
    Next
    Prim gPar, gOrd
    gPar(0) = -1
    gNEx = 0
End Sub

'--- 1) ANILLO (looped main) -------------------------------------------------
Private Sub ConstruirAnillo()
    Dim ns As Long: ns = nP - 1
    If ns < 3 Then ConstruirArbolMST: Exit Sub

    Dim chain() As Long: ReDim chain(ns - 1)
    Dim used() As Boolean: ReDim used(nP - 1)
    Dim i As Long, j As Long, start As Long, dmin As Double, d As Double

    start = 1: dmin = 1E+30
    For j = 1 To nP - 1
        d = (pX(0) - pX(j)) ^ 2 + (pY(0) - pY(j)) ^ 2
        If d < dmin Then dmin = d: start = j
    Next
    chain(0) = start: used(start) = True
    Dim cur As Long: cur = start
    Dim nc As Long: nc = 1
    Do While nc < ns
        Dim nxt As Long: nxt = -1: dmin = 1E+30
        For j = 1 To nP - 1
            If Not used(j) Then
                d = (pX(cur) - pX(j)) ^ 2 + (pY(cur) - pY(j)) ^ 2
                If d < dmin Then dmin = d: nxt = j
            End If
        Next
        chain(nc) = nxt: used(nxt) = True: cur = nxt: nc = nc + 1
    Loop

    Dim D As Double: D = 0#
    For i = 0 To ns - 1: D = D + pQ(chain(i)): Next
    Dim m As Long
    If D <= 0# Then
        m = ns \ 2
    Else
        Dim cum As Double: cum = 0#: m = 0
        For i = 1 To ns - 1
            cum = cum + pQ(chain(i))
            If cum >= D / 2# Then m = i: Exit For
        Next
        If m = 0 Then m = ns \ 2
    End If
    If m < 1 Then m = 1
    If m > ns - 2 Then m = ns - 2

    gNN = nP
    ReDim gNX(gNN - 1): ReDim gNY(gNN - 1): ReDim gDem(gNN - 1)
    ReDim gPar(gNN - 1): ReDim gOrd(gNN - 1)
    For i = 0 To nP - 1
        gNX(i) = pX(i): gNY(i) = pY(i): gDem(i) = pQ(i)
    Next
    gPar(0) = -1
    gPar(chain(0)) = 0
    For i = 1 To m
        gPar(chain(i)) = chain(i - 1)
    Next
    gPar(chain(ns - 1)) = chain(0)
    For i = ns - 2 To m + 1 Step -1
        gPar(chain(i)) = chain(i + 1)
    Next

    gNEx = 1
    ReDim gExA(0): ReDim gExB(0)
    gExA(0) = chain(m): gExB(0) = chain(m + 1)
    OrdenarArbol
End Sub

'--- 2) PRINCIPAL + LATERALES ------------------------------------------------
Private Sub ConstruirPrincipalLaterales()
    Dim ns As Long: ns = nP - 1
    If ns < 2 Then ConstruirArbolMST: Exit Sub

    Dim i As Long, j As Long, a As Long, b As Long, dmax As Double, d As Double
    a = 1: b = 1: dmax = -1#
    For i = 1 To nP - 1
        For j = i + 1 To nP - 1
            d = (pX(i) - pX(j)) ^ 2 + (pY(i) - pY(j)) ^ 2
            If d > dmax Then dmax = d: a = i: b = j
        Next
    Next
    Dim ux As Double, uy As Double, ln As Double
    ux = pX(b) - pX(a): uy = pY(b) - pY(a)
    ln = Sqr(ux * ux + uy * uy)
    If ln < 0.000001 Then ux = 1#: uy = 0# Else ux = ux / ln: uy = uy / ln

    gNN = nP + ns
    ReDim gNX(gNN - 1): ReDim gNY(gNN - 1): ReDim gDem(gNN - 1)
    ReDim gPar(gNN - 1): ReDim gOrd(gNN - 1)
    For i = 0 To nP - 1
        gNX(i) = pX(i): gNY(i) = pY(i): gDem(i) = pQ(i)
    Next

    Dim t() As Double: ReDim t(nP - 1)
    t(0) = 0#
    For i = 1 To nP - 1
        t(i) = (pX(i) - pX(0)) * ux + (pY(i) - pY(0)) * uy
        Dim jn As Long: jn = nP + (i - 1)
        gNX(jn) = pX(0) + t(i) * ux
        gNY(jn) = pY(0) + t(i) * uy
        gDem(jn) = 0#
    Next

    Dim ord() As Long: ReDim ord(ns - 1)
    For i = 0 To ns - 1: ord(i) = i + 1: Next
    Dim p As Long, q As Long, tmp As Long
    For p = 1 To ns - 1
        tmp = ord(p): q = p - 1
        Do While q >= 0
            If t(ord(q)) > t(tmp) Then ord(q + 1) = ord(q): q = q - 1 Else Exit Do
        Loop
        ord(q + 1) = tmp
    Next

    Dim prevR As Long: prevR = 0
    For p = 0 To ns - 1
        If t(ord(p)) >= 0# Then
            Dim s As Long: s = ord(p)
            Dim jr As Long: jr = nP + (s - 1)
            gPar(jr) = prevR: gPar(s) = jr: prevR = jr
        End If
    Next
    Dim prevL As Long: prevL = 0
    For p = ns - 1 To 0 Step -1
        If t(ord(p)) < 0# Then
            Dim s2 As Long: s2 = ord(p)
            Dim jl As Long: jl = nP + (s2 - 1)
            gPar(jl) = prevL: gPar(s2) = jl: prevL = jl
        End If
    Next

    gPar(0) = -1
    gNEx = 0
    OrdenarArbol
End Sub

'--- Orden topologico (BFS desde la raiz 0) ----------------------------------
Private Sub OrdenarArbol()
    Dim cnt As Long, h As Long, cur As Long, j As Long
    gOrd(0) = 0: cnt = 1: h = 0
    Do While h < cnt
        cur = gOrd(h)
        For j = 1 To gNN - 1
            If gPar(j) = cur Then gOrd(cnt) = j: cnt = cnt + 1
        Next
        h = h + 1
    Loop
End Sub

'==============================================================================
' PRIM (arbol de expansion minima, metrica euclidiana)
'==============================================================================
Private Sub Prim(ByRef parent() As Long, ByRef orden() As Long)
    Dim enArbol() As Boolean, best() As Double
    ReDim enArbol(nP - 1): ReDim best(nP - 1)
    Dim i As Long, j As Long
    For i = 0 To nP - 1: best(i) = 1E+30: parent(i) = 0: Next
    best(0) = 0#
    Dim nAdd As Long: nAdd = 0
    For i = 0 To nP - 1
        Dim u As Long, mejor As Double
        u = -1: mejor = 1E+30
        For j = 0 To nP - 1
            If Not enArbol(j) Then
                If best(j) < mejor Then mejor = best(j): u = j
            End If
        Next
        If u = -1 Then Exit For
        enArbol(u) = True: orden(nAdd) = u: nAdd = nAdd + 1
        For j = 0 To nP - 1
            If Not enArbol(j) Then
                Dim d As Double
                d = (pX(j) - pX(u)) ^ 2 + (pY(j) - pY(u)) ^ 2
                If d < best(j) Then best(j) = d: parent(j) = u
            End If
        Next
    Next
End Sub

'==============================================================================
' HIDRAULICA (Hazen-Williams)
'==============================================================================
Private Function HfTramo(i As Long, di As Long, acc() As Double, hwC As Double) As Double
    Dim Q As Double: Q = acc(i)
    If Q <= 0# Then Exit Function
    Dim L As Double, Dm As Double, Qm As Double
    L = Sqr((gNX(i) - gNX(gPar(i))) ^ 2 + (gNY(i) - gNY(gPar(i))) ^ 2)
    Dm = gDiam(di) / 1000#
    Qm = Q / 60000#
    HfTramo = 10.67 * L * (Qm ^ 1.852) / ((hwC ^ 1.852) * (Dm ^ 4.871))
End Function

Private Sub ComputarCumHf(ByRef cumHf() As Double, segDi() As Long, _
                          acc() As Double, hwC As Double)
    Dim k As Long, nd As Long
    cumHf(0) = 0#
    For k = 1 To gNN - 1
        nd = gOrd(k)
        cumHf(nd) = cumHf(gPar(nd)) + HfTramo(nd, segDi(nd), acc, hwC)
    Next
End Sub

'==============================================================================
' DIAMETRO minimo por velocidad
'==============================================================================
Private Function ElegirDiametro(caudalLmin As Double, vMax As Double) As Long
    Dim qm3s As Double, dMin As Double, k As Long
    If caudalLmin <= 0# Then ElegirDiametro = 0: Exit Function
    qm3s = caudalLmin / 60000#
    dMin = Sqr(4# * qm3s / (PI * vMax)) * 1000#
    For k = 0 To nDiam - 1
        If gDiam(k) >= dMin - 0.000001 Then ElegirDiametro = k: Exit Function
    Next
    ElegirDiametro = nDiam - 1
End Function

Private Sub CargarCatalogoDiametros()
    Dim dd As Variant, cc As Variant, i As Long
    dd = Array(16#, 20#, 25#, 32#, 40#, 50#, 63#, 75#, 90#, 110#)
    cc = Array(5, 3, 4, 1, 6, 2, 30, 8, 40, 200)
    nDiam = UBound(dd) + 1
    ReDim gDiam(nDiam - 1): ReDim gDiamColor(nDiam - 1)
    For i = 0 To nDiam - 1
        gDiam(i) = CDbl(dd(i)): gDiamColor(i) = CLng(cc(i))
    Next
End Sub

'==============================================================================
' DETECCION DE ASPERSORES
'==============================================================================
Private Sub RecolectarAspersores()
    ReDim pX(255): ReDim pY(255): ReDim pZ(255)
    ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, br As AcadBlockReference
    Dim nombre As String, capa As String, esAspersor As Boolean
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            Set br = ent
            nombre = ""
            On Error Resume Next
            nombre = br.EffectiveName
            If nombre = "" Then nombre = br.Name
            capa = br.Layer
            On Error GoTo 0
            esAspersor = (InStr(1, UCase$(nombre), "ASPERSOR", vbTextCompare) > 0)
            If Not esAspersor Then _
                esAspersor = (InStr(1, UCase$(capa), "RIEGO_ASPERSOR", vbTextCompare) > 0)
            If Not esAspersor Then _
                esAspersor = TieneAtributo(br, "NUM") Or TieneAtributo(br, "CAUDAL")
            If esAspersor Then
                Dim ip As Variant: ip = br.InsertionPoint
                AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), _
                             Val(AtributoBloque(br, "CAUDAL")), _
                             AtributoBloque(br, "NUM"), _
                             UCase$(Trim$(AtributoBloque(br, "ZONA")))
            End If
        End If
    Next
End Sub

Private Function TieneAtributo(br As AcadBlockReference, tag As String) As Boolean
    On Error Resume Next
    Dim atts As Variant, i As Long
    If br.HasAttributes Then
        atts = br.GetAttributes
        For i = LBound(atts) To UBound(atts)
            If UCase$(atts(i).TagString) = UCase$(tag) Then TieneAtributo = True: Exit Function
        Next
    End If
End Function

Private Function AtributoBloque(br As AcadBlockReference, tag As String) As String
    On Error Resume Next
    Dim atts As Variant, i As Long
    If br.HasAttributes Then
        atts = br.GetAttributes
        For i = LBound(atts) To UBound(atts)
            If UCase$(atts(i).TagString) = UCase$(tag) Then _
                AtributoBloque = atts(i).TextString: Exit Function
        Next
    End If
End Function

Private Sub ContarEntidades(ByRef nBlk As Long, ByRef nCir As Long, ByRef nPnt As Long)
    Dim ent As AcadEntity
    nBlk = 0: nCir = 0: nPnt = 0
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            nBlk = nBlk + 1
        ElseIf TypeOf ent Is AcadCircle Then
            nCir = nCir + 1
        ElseIf TypeOf ent Is AcadPoint Then
            nPnt = nPnt + 1
        End If
    Next
End Sub

Private Sub ColectarBloques()
    ReDim pX(255): ReDim pY(255): ReDim pZ(255)
    ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, br As AcadBlockReference, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            Set br = ent
            ip = br.InsertionPoint
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), _
                         Val(AtributoBloque(br, "CAUDAL")), _
                         AtributoBloque(br, "NUM"), _
                         UCase$(Trim$(AtributoBloque(br, "ZONA")))
        End If
    Next
End Sub

Private Sub ColectarCirculos()
    ReDim pX(255): ReDim pY(255): ReDim pZ(255)
    ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadCircle Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), 0#, ""
        End If
    Next
End Sub

Private Function SeleccionManual() As Boolean
    On Error Resume Next
    Dim ss As AcadSelectionSet
    Set ss = ThisDrawing.SelectionSets.Item("TUB_SEL")
    If Not ss Is Nothing Then ss.Delete
    On Error GoTo 0
    Set ss = ThisDrawing.SelectionSets.Add("TUB_SEL")
    ThisDrawing.Utility.Prompt vbCrLf & _
        "Seleccione los aspersores (bloques / circulos / puntos): "
    ss.SelectOnScreen
    ReDim pX(255): ReDim pY(255): ReDim pZ(255)
    ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, ip As Variant
    For Each ent In ss
        If TypeOf ent Is AcadBlockReference Then
            ip = ent.InsertionPoint
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), _
                         Val(AtributoBloque(ent, "CAUDAL")), AtributoBloque(ent, "NUM")
        ElseIf TypeOf ent Is AcadCircle Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), 0#, ""
        ElseIf TypeOf ent Is AcadArc Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), 0#, ""
        ElseIf TypeOf ent Is AcadEllipse Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), 0#, ""
        ElseIf TypeOf ent Is AcadPoint Then
            ip = ent.Coordinates
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), CDbl(ip(2)), 0#, ""
        Else
            Dim x As Double, y As Double, z As Double
            If CentroEntidad(ent, x, y, z) Then AgregarPunto x, y, z, 0#, ""
        End If
    Next
    ss.Delete
    SeleccionManual = (nP > 0)
End Function

Private Function CentroEntidad(ent As AcadEntity, ByRef x As Double, _
                               ByRef y As Double, ByRef z As Double) As Boolean
    On Error GoTo fin
    Dim lo As Variant, hi As Variant
    ent.GetBoundingBox lo, hi
    x = (CDbl(lo(0)) + CDbl(hi(0))) / 2#
    y = (CDbl(lo(1)) + CDbl(hi(1))) / 2#
    z = (CDbl(lo(2)) + CDbl(hi(2))) / 2#
    CentroEntidad = True
fin:
End Function

'==============================================================================
' FILTRO POR ZONA
'==============================================================================
Private Sub FiltrarPorZona()
    Dim zonas(63) As String, nz As Long
    Dim i As Long, j As Long, existe As Boolean
    nz = 0
    For i = 0 To nP - 1
        If pZona(i) <> "" Then
            existe = False
            For j = 0 To nz - 1
                If zonas(j) = pZona(i) Then existe = True
            Next
            If Not existe And nz < 64 Then zonas(nz) = pZona(i): nz = nz + 1
        End If
    Next
    If nz <= 1 Then Exit Sub

    Dim lista As String
    For i = 0 To nz - 1: lista = lista & "   " & zonas(i) & vbCrLf: Next
    Dim sel As String
    sel = UCase$(Trim$(InputBox( _
        "Hay aspersores de varias ZONAS / VALVULAS:" & vbCrLf & lista & vbCrLf & _
        "Escriba la ZONA a trazar, o deje VACIO para TODAS.", "Filtro por zona", "")))
    If sel = "" Then Exit Sub

    Dim k As Long: k = 0
    For i = 0 To nP - 1
        If pZona(i) = sel Then
            pX(k) = pX(i): pY(k) = pY(i): pZ(k) = pZ(i): pQ(k) = pQ(i)
            pNum(k) = pNum(i): pZona(k) = pZona(i)
            k = k + 1
        End If
    Next
    nP = k
    If nP = 0 Then MsgBox "Ningun aspersor tiene la zona '" & sel & "'.", vbExclamation
End Sub

'==============================================================================
' AUXILIARES DE PUNTOS
'==============================================================================
Private Sub AgregarPunto(x As Double, y As Double, z As Double, q As Double, _
                         num As String, Optional zona As String = "")
    Dim i As Long
    For i = 0 To nP - 1
        If Abs(pX(i) - x) < 0.000001 And Abs(pY(i) - y) < 0.000001 Then Exit Sub
    Next
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pZ(UBound(pZ) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
        ReDim Preserve pZona(UBound(pZona) + 256)
    End If
    pX(nP) = x: pY(nP) = y: pZ(nP) = z: pQ(nP) = q
    pNum(nP) = num: pZona(nP) = zona
    nP = nP + 1
End Sub

Private Sub InsertarFuente(x As Double, y As Double, z As Double)
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pZ(UBound(pZ) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
        ReDim Preserve pZona(UBound(pZona) + 256)
    End If
    Dim i As Long
    For i = nP To 1 Step -1
        pX(i) = pX(i - 1): pY(i) = pY(i - 1): pZ(i) = pZ(i - 1)
        pQ(i) = pQ(i - 1): pNum(i) = pNum(i - 1): pZona(i) = pZona(i - 1)
    Next
    pX(0) = x: pY(0) = y: pZ(0) = z: pQ(0) = 0#
    pNum(0) = "FUENTE": pZona(0) = ""
    nP = nP + 1
End Sub

'==============================================================================
' DIBUJO Y UTILIDADES
'==============================================================================
Private Sub DibujarTramo(a As Long, b As Long, di As Long, flujo As Double, _
                         rotular As Boolean, htxt As Double, _
                         ByRef longTot As Double, ByRef longD() As Double, _
                         ByRef usados() As Boolean)
    Dim L As Double
    L = Sqr((gNX(b) - gNX(a)) ^ 2 + (gNY(b) - gNY(a)) ^ 2)
    If L < 0.000000001 Then Exit Sub
    Dim dmm As Double: dmm = gDiam(di)
    longTot = longTot + L
    longD(di) = longD(di) + L
    usados(di) = True
    Dim capa As String: capa = "RIEGO_TUB_" & Format(dmm, "0")
    CrearCapa capa, gDiamColor(di)
    Dim pts(0 To 3) As Double
    pts(0) = gNX(a): pts(1) = gNY(a): pts(2) = gNX(b): pts(3) = gNY(b)
    Dim tub As AcadLWPolyline
    Set tub = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
    tub.Layer = capa: tub.color = gDiamColor(di)
    If rotular Then RotularTramo gNX(a), gNY(a), gNX(b), gNY(b), dmm, flujo, htxt
End Sub

Private Sub RotularTramo(x1 As Double, y1 As Double, x2 As Double, y2 As Double, _
                         dmm As Double, caudal As Double, h As Double)
    CrearCapa "RIEGO_TUB_TXT", 8
    Dim mx As Double, my As Double, ang As Double
    mx = (x1 + x2) / 2#: my = (y1 + y2) / 2#
    Dim ins(0 To 2) As Double
    ins(0) = mx: ins(1) = my + h * 0.4: ins(2) = 0#
    Dim txt As AcadText
    Set txt = ThisDrawing.ModelSpace.AddText( _
        "D" & Format(dmm, "0") & " (" & Format(caudal, "0.0") & " l/min)", ins, h)
    txt.Layer = "RIEGO_TUB_TXT"
    ang = Atan2(y2 - y1, x2 - x1)
    If ang > PI / 2# Then ang = ang - PI
    If ang < -PI / 2# Then ang = ang + PI
    txt.Rotation = ang
    txt.Alignment = acAlignmentLeft
End Sub

Private Function AlturaTexto() As Double
    Dim xmn As Double, xmx As Double, ymn As Double, ymx As Double, i As Long
    xmn = gNX(0): xmx = gNX(0): ymn = gNY(0): ymx = gNY(0)
    For i = 1 To gNN - 1
        If gNX(i) < xmn Then xmn = gNX(i)
        If gNX(i) > xmx Then xmx = gNX(i)
        If gNY(i) < ymn Then ymn = gNY(i)
        If gNY(i) > ymx Then ymx = gNY(i)
    Next
    Dim diag As Double: diag = Sqr((xmx - xmn) ^ 2 + (ymx - ymn) ^ 2)
    If diag <= 0# Then diag = 10#
    AlturaTexto = diag * 0.012
End Function

Private Sub CrearCapa(nombre As String, color As Long)
    Dim ly As AcadLayer
    On Error Resume Next
    Set ly = ThisDrawing.Layers(nombre)
    If ly Is Nothing Then
        Set ly = ThisDrawing.Layers.Add(nombre)
        ly.color = color
    End If
    On Error GoTo 0
End Sub

Private Function Atan2(y As Double, x As Double) As Double
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atan2 = Atn(y / x) + PI
    ElseIf x < 0# Then
        Atan2 = Atn(y / x) - PI
    ElseIf y > 0# Then
        Atan2 = PI / 2#
    ElseIf y < 0# Then
        Atan2 = -PI / 2#
    Else
        Atan2 = 0#
    End If
End Function
