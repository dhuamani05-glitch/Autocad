Attribute VB_Name = "modTrazadoTuberias"
Option Explicit
'==============================================================================
' TRAZADO DE TUBERIAS DE RIEGO MEDIANTE POLILINEAS  (v2)
'------------------------------------------------------------------------------
' Complemento de "modDisenoAspersion". Conecta los aspersores colocados
' (bloques ASPERSOR_RIEGO_* en capa RIEGO_ASPERSOR) con un punto de FUENTE
' (cabezal / valvula) trazando la red con POLILINEAS.
'
' DISPOSICIONES (el usuario elige al ejecutar):
'   1 = ANILLO (looped main): cierra el perimetro en bucle alimentado desde
'       la fuente; el caudal se reparte por las dos ramas -> menor friccion
'       y presion mas uniforme. Recomendado para aspersores en perimetro.
'   2 = PRINCIPAL + LATERALES: una troncal a lo largo del eje dominante con
'       laterales perpendiculares a cada aspersor (pocos emisores en serie).
'   3 = ARBOL (MST, Prim): minima longitud total de tuberia (menor material).
'
' CRITERIO HIDRAULICO (mejor practica de riego):
'   La variacion de presion dentro del sector no debe superar ~20% de la
'   presion nominal del emisor. Se dimensiona el diametro de cada tramo
'   (Hazen-Williams) partiendo del minimo por velocidad y agrandando el
'   tramo mas critico hasta cumplir el criterio. El reporte indica la
'   variacion real y si CUMPLE / NO CUMPLE.
'
' CAPAS: RIEGO_TUB_<diametro> (una por diametro), RIEGO_TUB_TXT (rotulos),
'        RIEGO_FUENTE (marcador de la fuente).
'==============================================================================

Private Const PI As Double = 3.14159265358979

'--- Aspersores detectados (indice 0 = FUENTE tras InsertarFuente) ------------
Private pX() As Double
Private pY() As Double
Private pQ() As Double         ' demanda propia del nodo (l/min); fuente = 0
Private pNum() As String       ' etiqueta NUM del aspersor
Private pZona() As String      ' atributo ZONA del aspersor
Private nP As Long

'--- Red construida como ARBOL enraizado en la fuente (nodo 0) ----------------
'    Los nodos incluyen la fuente, los aspersores y (en principal+laterales)
'    nodos de union (tees) sobre la troncal.
Private gNX() As Double        ' coordenada X del nodo
Private gNY() As Double        ' coordenada Y del nodo
Private gDem() As Double       ' demanda del nodo (l/min)
Private gPar() As Long         ' nodo padre (gPar(0) = -1 = raiz)
Private gOrd() As Long         ' orden topologico (padre antes que hijo)
Private gNN As Long            ' numero de nodos
'--- Aristas EXTRA que cierran bucles (no son del arbol): anillo -------------
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
    ' 1) RECOLECTAR ASPERSORES (por firma del bloque de modDisenoAspersion)
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
            "   1 = TODOS los bloques (recomendado si hay bloques)" & vbCrLf & _
            "   2 = TODOS los circulos" & vbCrLf & _
            "   3 = SELECCIONARLOS a mano en pantalla", _
            "Trazado de tuberias - diagnostico", "1"))
        Select Case op
            Case "1": ColectarBloques
            Case "2": ColectarCirculos
            Case Else: SeleccionManual
        End Select
    End If

    If nP < 1 Then
        MsgBox "No hay aspersores para conectar." & vbCrLf & _
               "Coloque primero los aspersores (modDisenoAspersion)" & vbCrLf & _
               "o seleccionelos a mano.", vbExclamation, "Trazado de tuberias"
        Exit Sub
    End If

    FiltrarPorZona
    If nP < 1 Then Exit Sub

    If MsgBox("Se detectaron " & nP & " aspersores para conectar." & vbCrLf & vbCrLf & _
              "A continuacion se pedira el punto de FUENTE (cabezal/valvula)." & vbCrLf & _
              "Desea continuar?", vbOKCancel + vbInformation, _
              "Trazado de tuberias") = vbCancel Then Exit Sub

    '======================================================================
    ' 2) PUNTO DE FUENTE (cabezal / valvula) = nodo 0
    '======================================================================
    Dim src As Variant
    On Error Resume Next
    src = ThisDrawing.Utility.GetPoint(, vbCrLf & _
          "Indique el punto de FUENTE (cabezal / valvula): ")
    If Err.Number <> 0 Then Exit Sub          ' ESC
    On Error GoTo errH
    InsertarFuente CDbl(src(0)), CDbl(src(1))

    '======================================================================
    ' 3) CRITERIOS DE DISENO
    '======================================================================
    Dim vMax As Double
    vMax = Val(InputBox("Velocidad maxima admisible en la tuberia (m/s):" & vbCrLf & _
                        "  1.5 = recomendado para PVC/PE", _
                        "Criterio hidraulico", "1.5"))
    If vMax <= 0.1 Then vMax = 1.5

    Dim pNom As Double, pctVar As Double, hwC As Double
    pNom = Val(InputBox("Presion NOMINAL de operacion del aspersor (m.c.a.):" & vbCrLf & _
                        "  20 m.c.a. = 2.0 bar (tipico en aspersion)", _
                        "Criterio de presion", "20"))
    If pNom <= 0# Then pNom = 20#
    pctVar = Val(InputBox("Variacion MAXIMA de presion admisible en el sector (%):" & vbCrLf & _
                          "  20 % = criterio estandar de diseno", _
                          "Criterio de presion", "20"))
    If pctVar <= 0# Then pctVar = 20#
    hwC = Val(InputBox("Coeficiente de Hazen-Williams (C):" & vbCrLf & _
                       "  150 = PVC / PE liso   |   140 = PVC usado", _
                       "Criterio de presion", "150"))
    If hwC <= 0# Then hwC = 150#

    Dim rotular As Boolean
    rotular = (UCase$(Trim$(InputBox( _
        "Rotular cada tramo con su diametro y caudal?  (S/N)", _
        "Rotulos", "S"))) = "S")

    CargarCatalogoDiametros

    '======================================================================
    ' 4) DISPOSICION DE LA RED (el usuario elige) -> construye el ARBOL
    '======================================================================
    Dim topo As String, topoName As String
    topo = Trim$(InputBox( _
        "DISPOSICION de la red de tuberias:" & vbCrLf & vbCrLf & _
        "  1 = ANILLO (looped main)" & vbCrLf & _
        "        presion mas uniforme; recomendado para aspersores" & vbCrLf & _
        "        distribuidos en el perimetro" & vbCrLf & _
        "  2 = PRINCIPAL + LATERALES" & vbCrLf & _
        "        troncal en el eje con laterales perpendiculares" & vbCrLf & _
        "  3 = ARBOL (MST)" & vbCrLf & _
        "        minima longitud total de tuberia", _
        "Disposicion de tuberias", "1"))

    Select Case topo
        Case "2": topoName = "Principal + laterales": ConstruirPrincipalLaterales
        Case "3": topoName = "Arbol de expansion minima (MST)": ConstruirArbolMST
        Case Else: topoName = "Anillo (looped main)": ConstruirAnillo
    End Select
    If gNN < 2 Then Exit Sub

    '======================================================================
    ' 5) CAUDAL AGUAS ABAJO EN CADA TRAMO (subarbol enraizado en la fuente)
    '======================================================================
    Dim acc() As Double
    ReDim acc(gNN - 1)
    Dim i As Long, k As Long, nd As Long
    For i = 0 To gNN - 1
        acc(i) = gDem(i)
    Next
    For k = gNN - 1 To 1 Step -1
        nd = gOrd(k)
        acc(gPar(nd)) = acc(gPar(nd)) + acc(nd)
    Next

    '======================================================================
    ' 6) DIMENSIONAMIENTO POR PRESION (criterio del 20%)
    '======================================================================
    Dim segDi() As Long
    ReDim segDi(gNN - 1)
    For i = 1 To gNN - 1
        segDi(i) = ElegirDiametro(acc(i), vMax)      ' piso por velocidad
    Next

    Dim admis As Double: admis = pctVar / 100# * pNom      ' variacion admisible (m.c.a.)
    Dim cumHf() As Double: ReDim cumHf(gNN - 1)
    Dim worst As Double, worstNode As Long
    Dim iterD As Long, nMax As Long: nMax = gNN * nDiam + 20

    For iterD = 1 To nMax
        ComputarCumHf cumHf, segDi, acc, hwC
        worst = 0#: worstNode = -1
        For i = 1 To gNN - 1
            If cumHf(i) > worst Then worst = cumHf(i): worstNode = i
        Next
        If worst <= admis Or worstNode = -1 Then Exit For
        ' agrandar el tramo de mayor perdida en el camino al nodo peor
        Dim best As Long, bestHf As Double, nn As Long, hh As Double
        best = -1: bestHf = -1#
        nn = worstNode
        Do While nn <> 0
            If segDi(nn) < nDiam - 1 Then
                hh = HfTramo(nn, segDi(nn), acc, hwC)
                If hh > bestHf Then bestHf = hh: best = nn
            End If
            nn = gPar(nn)
        Loop
        If best = -1 Then Exit For              ' catalogo agotado
        segDi(best) = segDi(best) + 1
    Next

    ComputarCumHf cumHf, segDi, acc, hwC
    worst = 0#
    For i = 1 To gNN - 1
        If cumHf(i) > worst Then worst = cumHf(i)
    Next

    '======================================================================
    ' 7) DIBUJO
    '======================================================================
    Dim htxt As Double: htxt = AlturaTexto()
    Dim usados(0 To 63) As Boolean
    Dim longD(0 To 63) As Double
    Dim longTot As Double

    CrearCapa "RIEGO_FUENTE", 1
    Dim cf(0 To 2) As Double
    cf(0) = gNX(0): cf(1) = gNY(0): cf(2) = 0#
    Dim mkr As AcadCircle
    Set mkr = ThisDrawing.ModelSpace.AddCircle(cf, htxt * 1.5)
    mkr.Layer = "RIEGO_FUENTE"
    mkr.color = 1

    ' tramos del arbol
    For i = 1 To gNN - 1
        DibujarTramo gPar(i), i, segDi(i), acc(i), rotular, htxt, _
                     longTot, longD, usados
    Next
    ' aristas extra (cierre del anillo): diametro minimo, caudal ~0
    For i = 0 To gNEx - 1
        DibujarTramo gExA(i), gExB(i), 0, 0#, False, htxt, _
                     longTot, longD, usados
    Next

    ThisDrawing.Regen acActiveViewport

    '======================================================================
    ' 8) REPORTE
    '======================================================================
    Dim pctReal As Double: pctReal = 100# * worst / pNom
    Dim veredicto As String
    If worst <= admis + 0.0000001 Then
        veredicto = "CUMPLE  (<= " & Format(pctVar, "0") & "%)"
    Else
        veredicto = "NO CUMPLE - fraccione el sector / acorte el" & vbCrLf & _
                    "                   recorrido / use el ANILLO."
    End If

    Dim rep As String
    rep = "TRAZADO DE TUBERIAS COMPLETADO" & vbCrLf & String(46, "-") & vbCrLf & _
          "Disposicion:            " & topoName & vbCrLf & _
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
          "Caudal total en la fuente: " & Format(acc(0), "0.0") & " l/min" & vbCrLf & vbCrLf & _
          "CRITERIO DE PRESION (mejor practica de riego):" & vbCrLf & _
          "  Presion nominal:      " & Format(pNom, "0.0") & " m.c.a." & vbCrLf & _
          "  Variacion admisible:  " & Format(pctVar, "0") & " % = " & _
                Format(admis, "0.00") & " m.c.a." & vbCrLf & _
          "  Perdida en el emisor mas desfavorable:" & vbCrLf & _
          "                        " & Format(worst, "0.00") & " m.c.a. (" & _
                Format(pctReal, "0.0") & " % de la nominal)" & vbCrLf & _
          "  Estado:               " & veredicto & vbCrLf & vbCrLf & _
          "(Friccion Hazen-Williams C=" & Format(hwC, "0") & "; sin desnivel." & vbCrLf & _
          "Metros por diametro con DATAEXTRACTION por capa.)"
    MsgBox rep, vbInformation, "Trazado de tuberias"
    Exit Sub

errH:
    If Err.Number = -2147352567 Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Trazado de tuberias"
End Sub

'==============================================================================
' CONSTRUCTORES DE LA RED (cada uno llena gNX/gNY/gDem/gPar/gOrd/gNN y gEx*)
'==============================================================================

'--- 3) ARBOL DE EXPANSION MINIMA (Prim) -------------------------------------
Private Sub ConstruirArbolMST()
    gNN = nP
    ReDim gNX(gNN - 1): ReDim gNY(gNN - 1): ReDim gDem(gNN - 1)
    ReDim gPar(gNN - 1): ReDim gOrd(gNN - 1)
    Dim i As Long
    For i = 0 To nP - 1
        gNX(i) = pX(i): gNY(i) = pY(i): gDem(i) = pQ(i)
    Next
    Prim gPar, gOrd          ' usa pX/pY/nP; llena gPar/gOrd
    gPar(0) = -1
    gNEx = 0
End Sub

'--- 1) ANILLO (looped main) -------------------------------------------------
' Cadena perimetral por vecino mas cercano; la fuente alimenta el nodo mas
' proximo; el caudal se reparte en dos ramas que se encuentran en el "punto
' neutro" (donde la demanda acumulada de cada lado se equilibra); el tramo de
' cierre entre ambas ramas se dibuja aparte y transporta ~0.
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

    ' punto neutro por equilibrio de demanda (o por conteo si no hay caudal)
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
    gPar(chain(0)) = 0                       ' alimentador fuente -> nodo de entrada
    For i = 1 To m                           ' rama adelante
        gPar(chain(i)) = chain(i - 1)
    Next
    gPar(chain(ns - 1)) = chain(0)           ' rama atras
    For i = ns - 2 To m + 1 Step -1
        gPar(chain(i)) = chain(i + 1)
    Next

    gNEx = 1                                 ' cierre del anillo (~0 de caudal)
    ReDim gExA(0): ReDim gExB(0)
    gExA(0) = chain(m): gExB(0) = chain(m + 1)

    OrdenarArbol
End Sub

'--- 2) PRINCIPAL + LATERALES ------------------------------------------------
' Eje dominante = recta entre los dos aspersores mas alejados. La troncal
' pasa por la fuente en esa direccion; cada aspersor se conecta con un lateral
' perpendicular a un nodo de union (tee) sobre la troncal.
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
    If ln < 0.000001 Then
        ux = 1#: uy = 0#
    Else
        ux = ux / ln: uy = uy / ln
    End If

    ' nodos: 0..nP-1 (fuente+aspersores) ; nP..nP+ns-1 (uniones, una por aspersor)
    gNN = nP + ns
    ReDim gNX(gNN - 1): ReDim gNY(gNN - 1): ReDim gDem(gNN - 1)
    ReDim gPar(gNN - 1): ReDim gOrd(gNN - 1)
    For i = 0 To nP - 1
        gNX(i) = pX(i): gNY(i) = pY(i): gDem(i) = pQ(i)
    Next

    Dim t() As Double: ReDim t(nP - 1)
    t(0) = 0#
    For i = 1 To nP - 1
        t(i) = (pX(i) - pX(0)) * ux + (pY(i) - pY(0)) * uy      ' proyeccion sobre el eje
        Dim jn As Long: jn = nP + (i - 1)
        gNX(jn) = pX(0) + t(i) * ux
        gNY(jn) = pY(0) + t(i) * uy
        gDem(jn) = 0#
    Next

    ' ordenar aspersores por t (insercion)
    Dim ord() As Long: ReDim ord(ns - 1)
    For i = 0 To ns - 1: ord(i) = i + 1: Next
    Dim p As Long, q As Long, tmp As Long
    For p = 1 To ns - 1
        tmp = ord(p): q = p - 1
        Do While q >= 0
            If t(ord(q)) > t(tmp) Then
                ord(q + 1) = ord(q): q = q - 1
            Else
                Exit Do
            End If
        Loop
        ord(q + 1) = tmp
    Next

    ' troncal desde la fuente hacia t>=0 (ascendente) y hacia t<0 (descendente)
    Dim prevR As Long: prevR = 0
    For p = 0 To ns - 1
        If t(ord(p)) >= 0# Then
            Dim s As Long: s = ord(p)
            Dim jr As Long: jr = nP + (s - 1)
            gPar(jr) = prevR
            gPar(s) = jr
            prevR = jr
        End If
    Next
    Dim prevL As Long: prevL = 0
    For p = ns - 1 To 0 Step -1
        If t(ord(p)) < 0# Then
            Dim s2 As Long: s2 = ord(p)
            Dim jl As Long: jl = nP + (s2 - 1)
            gPar(jl) = prevL
            gPar(s2) = jl
            prevL = jl
        End If
    Next

    gPar(0) = -1
    gNEx = 0
    OrdenarArbol
End Sub

'--- Orden topologico (BFS desde la raiz 0): padre antes que hijo -------------
Private Sub OrdenarArbol()
    Dim cnt As Long, h As Long, cur As Long, j As Long
    gOrd(0) = 0
    cnt = 1: h = 0
    Do While h < cnt
        cur = gOrd(h)
        For j = 1 To gNN - 1
            If gPar(j) = cur Then
                gOrd(cnt) = j: cnt = cnt + 1
            End If
        Next
        h = h + 1
    Loop
End Sub

'==============================================================================
' ALGORITMO DE PRIM  (arbol de expansion minima, metrica euclidiana)
'==============================================================================
Private Sub Prim(ByRef parent() As Long, ByRef orden() As Long)
    Dim enArbol() As Boolean
    Dim best() As Double
    ReDim enArbol(nP - 1)
    ReDim best(nP - 1)

    Dim i As Long, j As Long
    For i = 0 To nP - 1
        best(i) = 1E+30
        parent(i) = 0
    Next
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
        enArbol(u) = True
        orden(nAdd) = u: nAdd = nAdd + 1
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
' HIDRAULICA
'==============================================================================
' Perdida de carga de un tramo (Hazen-Williams), m.c.a.
'   hf = 10.67 * L * Q^1.852 / ( C^1.852 * D^4.871 )   [Q m3/s, D m, L m]
'   El tramo que alimenta al nodo i lleva el caudal acc(i).
Private Function HfTramo(i As Long, di As Long, acc() As Double, hwC As Double) As Double
    Dim Q As Double: Q = acc(i)
    If Q <= 0# Then Exit Function
    Dim L As Double, Dm As Double, Qm As Double
    L = Sqr((gNX(i) - gNX(gPar(i))) ^ 2 + (gNY(i) - gNY(gPar(i))) ^ 2)
    Dm = gDiam(di) / 1000#
    Qm = Q / 60000#
    HfTramo = 10.67 * L * (Qm ^ 1.852) / ((hwC ^ 1.852) * (Dm ^ 4.871))
End Function

' Perdida acumulada desde la fuente a cada nodo (recorriendo el orden topol.)
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
' DIMENSIONAMIENTO: menor diametro comercial que cumple la velocidad maxima.
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
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
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
                AgregarPunto CDbl(ip(0)), CDbl(ip(1)), _
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
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, br As AcadBlockReference, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            Set br = ent
            ip = br.InsertionPoint
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), _
                         Val(AtributoBloque(br, "CAUDAL")), _
                         AtributoBloque(br, "NUM"), _
                         UCase$(Trim$(AtributoBloque(br, "ZONA")))
        End If
    Next
End Sub

Private Sub ColectarCirculos()
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadCircle Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
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
        "Seleccione los aspersores a conectar (bloques / circulos / puntos): "
    ss.SelectOnScreen

    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255): ReDim pZona(255)
    nP = 0
    Dim ent As AcadEntity, ip As Variant
    For Each ent In ss
        If TypeOf ent Is AcadBlockReference Then
            ip = ent.InsertionPoint
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), _
                         Val(AtributoBloque(ent, "CAUDAL")), _
                         AtributoBloque(ent, "NUM")
        ElseIf TypeOf ent Is AcadCircle Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
        ElseIf TypeOf ent Is AcadArc Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
        ElseIf TypeOf ent Is AcadEllipse Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
        ElseIf TypeOf ent Is AcadPoint Then
            ip = ent.Coordinates
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
        Else
            Dim x As Double, y As Double
            If CentroEntidad(ent, x, y) Then AgregarPunto x, y, 0#, ""
        End If
    Next
    ss.Delete
    SeleccionManual = (nP > 0)
End Function

Private Function CentroEntidad(ent As AcadEntity, ByRef x As Double, ByRef y As Double) As Boolean
    On Error GoTo fin
    Dim lo As Variant, hi As Variant
    ent.GetBoundingBox lo, hi
    x = (CDbl(lo(0)) + CDbl(hi(0))) / 2#
    y = (CDbl(lo(1)) + CDbl(hi(1))) / 2#
    CentroEntidad = True
fin:
End Function

'==============================================================================
' FILTRO POR ZONA (solo si hay varias zonas distintas)
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
        "Escriba la ZONA a trazar, o deje VACIO para trazar TODAS.", _
        "Filtro por zona", "")))
    If sel = "" Then Exit Sub

    Dim k As Long: k = 0
    For i = 0 To nP - 1
        If pZona(i) = sel Then
            pX(k) = pX(i): pY(k) = pY(i): pQ(k) = pQ(i)
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
Private Sub AgregarPunto(x As Double, y As Double, q As Double, num As String, _
                         Optional zona As String = "")
    Dim i As Long
    For i = 0 To nP - 1
        If Abs(pX(i) - x) < 0.000001 And Abs(pY(i) - y) < 0.000001 Then Exit Sub
    Next
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
        ReDim Preserve pZona(UBound(pZona) + 256)
    End If
    pX(nP) = x: pY(nP) = y: pQ(nP) = q: pNum(nP) = num: pZona(nP) = zona
    nP = nP + 1
End Sub

Private Sub InsertarFuente(x As Double, y As Double)
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
        ReDim Preserve pZona(UBound(pZona) + 256)
    End If
    Dim i As Long
    For i = nP To 1 Step -1
        pX(i) = pX(i - 1): pY(i) = pY(i - 1)
        pQ(i) = pQ(i - 1): pNum(i) = pNum(i - 1): pZona(i) = pZona(i - 1)
    Next
    pX(0) = x: pY(0) = y: pQ(0) = 0#: pNum(0) = "FUENTE": pZona(0) = ""
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
    If L < 0.000000001 Then Exit Sub          ' tramo nulo (union sobre la fuente)
    Dim dmm As Double: dmm = gDiam(di)
    longTot = longTot + L
    longD(di) = longD(di) + L
    usados(di) = True

    Dim capa As String: capa = "RIEGO_TUB_" & Format(dmm, "0")
    CrearCapa capa, gDiamColor(di)

    Dim pts(0 To 3) As Double
    pts(0) = gNX(a): pts(1) = gNY(a)
    pts(2) = gNX(b): pts(3) = gNY(b)
    Dim tub As AcadLWPolyline
    Set tub = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
    tub.Layer = capa
    tub.color = gDiamColor(di)

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
