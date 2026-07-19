Attribute VB_Name = "modTrazadoTuberias"
Option Explicit
'==============================================================================
' TRAZADO DE TUBERIAS DE RIEGO MEDIANTE POLILINEAS  (v1)
'------------------------------------------------------------------------------
' Complemento de "modDisenoAspersion" (distribucion de aspersores).
'
' OBJETIVO: dado un conjunto de aspersores ya colocados (bloques
' ASPERSOR_RIEGO_* en la capa RIEGO_ASPERSOR) y un punto de FUENTE
' (cabezal / valvula), trazar la red de tuberias que los conecta a TODOS
' con la MENOR longitud total posible, dibujada con POLILINEAS.
'
' POR QUE ES "LO MAS EFICIENTE POSIBLE"?
'   Conectar N puntos con tuberia, sin bucles (un arbol), gastando la
'   menor cantidad de tuberia, es EXACTAMENTE el problema del ARBOL DE
'   EXPANSION MINIMA (Minimum Spanning Tree). Este modulo lo resuelve con
'   el algoritmo de PRIM: garantiza el arbol de menor longitud total que
'   une la fuente con cada aspersor (optimo, no heuristico). No existe
'   ningun otro arbol de aristas aspersor-aspersor mas corto.
'
' ADEMAS (dimensionamiento hidraulico):
'   - Enraiza el arbol en la FUENTE y calcula, para cada tramo, el CAUDAL
'     que transporta = suma de los caudales de los aspersores "aguas
'     abajo" de ese tramo.
'   - Elige para cada tramo el DIAMETRO comercial mas pequeno que respeta
'     una velocidad maxima (por defecto 1.5 m/s): la red mas barata que
'     cumple el criterio hidraulico.
'   - Coloca cada tramo en una capa por diametro (RIEGO_TUB_xx) para
'     poder totalizar metros por diametro con DATAEXTRACTION, y opcional-
'     mente rotula cada tramo con su diametro y caudal.
'
' ENTRADAS:
'   - Aspersores: se detectan solos (bloques con "ASPERSOR" en el nombre).
'     Si no hay, se pueden seleccionar a mano (bloques / circulos / puntos).
'   - Punto de FUENTE (cabezal/valvula): se pide con el mouse.
'   - Filtro opcional por ZONA (atributo ZONA del bloque).
'
' CAPAS QUE CREA: RIEGO_TUB_16, _20, _25, ... (una por diametro usado) y
'                 RIEGO_TUB_TXT (rotulos).
'==============================================================================

Private Const PI As Double = 3.14159265358979

'--- Puntos a conectar (indice 0 = FUENTE) -----------------------------------
Private pX() As Double
Private pY() As Double
Private pQ() As Double        ' demanda propia del nodo (l/min); fuente = 0
Private pNum() As String      ' etiqueta NUM del aspersor (para reportes)
Private nP As Long

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
    Dim zonaFiltro As String
    zonaFiltro = UCase$(Trim$(InputBox( _
        "Filtrar por ZONA / VALVULA (atributo ZONA del bloque)." & vbCrLf & _
        "Deje VACIO para trazar TODOS los aspersores.", _
        "Trazado de tuberias", "")))

    ' Deteccion automatica de los bloques de aspersor (por nombre/atributo).
    RecolectarAspersores zonaFiltro

    ' Si la deteccion automatica no encontro nada, mostrar un DIAGNOSTICO de
    ' lo que hay en el dibujo y dejar elegir de donde tomar los aspersores.
    If nP = 0 Then
        Dim nBlk As Long, nCir As Long, nPnt As Long
        ContarEntidades nBlk, nCir, nPnt

        Dim op As String
        op = Trim$(InputBox( _
            "No se detectaron aspersores por nombre/atributo." & vbCrLf & _
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
            Case "1": ColectarBloques zonaFiltro
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

    ' Diagnostico: informar cuantos se van a conectar.
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

    ' Inserta la fuente al inicio del arreglo (desplaza el resto).
    InsertarFuente CDbl(src(0)), CDbl(src(1))

    '======================================================================
    ' 3) CRITERIO HIDRAULICO (velocidad maxima) Y CATALOGO DE DIAMETROS
    '======================================================================
    Dim vTxt As String, vMax As Double
    vTxt = InputBox("Velocidad maxima admisible en la tuberia (m/s):" & vbCrLf & _
                    "  1.5  = recomendado para PVC/PE" & vbCrLf & _
                    "  (menor velocidad -> diametros mas grandes)", _
                    "Criterio hidraulico", "1.5")
    vMax = Val(vTxt)
    If vMax <= 0.1 Then vMax = 1.5      ' vacio / cancelar -> valor por defecto

    CargarCatalogoDiametros

    Dim rotular As Boolean
    rotular = (UCase$(Trim$(InputBox( _
        "Rotular cada tramo con su diametro y caudal?  (S/N)", _
        "Rotulos", "S"))) = "S")

    '======================================================================
    ' 4) ARBOL DE EXPANSION MINIMA (PRIM)  -> longitud total minima
    '======================================================================
    Dim parent() As Long
    Dim orden() As Long
    ReDim parent(nP - 1)
    ReDim orden(nP - 1)
    Prim parent, orden

    '======================================================================
    ' 5) CAUDAL AGUAS ABAJO EN CADA TRAMO (enraizado en la fuente)
    '    acc(i) = suma de demandas del subarbol que cuelga de i.
    '    El tramo (parent(i) -> i) transporta acc(i).
    '======================================================================
    Dim acc() As Double
    ReDim acc(nP - 1)
    Dim i As Long
    For i = 0 To nP - 1
        acc(i) = pQ(i)
    Next
    ' orden(0) = fuente; cada nodo se agrego despues de su padre, asi que
    ' recorriendo en orden inverso el hijo siempre se procesa antes del padre.
    Dim k As Long, nd As Long
    For k = nP - 1 To 1 Step -1
        nd = orden(k)
        acc(parent(nd)) = acc(parent(nd)) + acc(nd)
    Next

    '======================================================================
    ' 6) DIBUJO DE LA RED  (una polilinea recta por tramo)
    '======================================================================
    Dim htxt As Double: htxt = AlturaTexto()
    Dim usados(0 To 63) As Boolean       ' que diametros se usaron
    Dim longD(0 To 63) As Double         ' metros por diametro
    Dim longTot As Double
    Dim di As Long, dmm As Double, flujo As Double

    ' --- marcador de la FUENTE (cabezal / valvula) ---
    CrearCapa "RIEGO_FUENTE", 1
    Dim cf(0 To 2) As Double
    cf(0) = pX(0): cf(1) = pY(0): cf(2) = 0#
    Dim mkr As AcadCircle
    Set mkr = ThisDrawing.ModelSpace.AddCircle(cf, htxt * 1.5)
    mkr.Layer = "RIEGO_FUENTE"
    mkr.color = 1

    For i = 1 To nP - 1
        flujo = acc(i)
        di = ElegirDiametro(flujo, vMax)
        dmm = gDiam(di)

        Dim L As Double
        L = Sqr((pX(i) - pX(parent(i))) ^ 2 + (pY(i) - pY(parent(i))) ^ 2)
        longTot = longTot + L
        longD(di) = longD(di) + L
        usados(di) = True

        Dim capa As String
        capa = "RIEGO_TUB_" & Format(dmm, "0")
        CrearCapa capa, gDiamColor(di)

        ' --- polilinea del tramo (2 vertices) ---
        Dim pts(0 To 3) As Double
        pts(0) = pX(parent(i)): pts(1) = pY(parent(i))
        pts(2) = pX(i):         pts(3) = pY(i)
        Dim tub As AcadLWPolyline
        Set tub = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
        tub.Layer = capa
        tub.color = gDiamColor(di)

        ' --- rotulo opcional en el punto medio ---
        If rotular Then
            RotularTramo pX(parent(i)), pY(parent(i)), pX(i), pY(i), _
                         dmm, flujo, htxt
        End If
    Next

    ThisDrawing.Regen acActiveViewport

    '======================================================================
    ' 7) REPORTE
    '======================================================================
    Dim rep As String
    rep = "TRAZADO DE TUBERIAS COMPLETADO" & vbCrLf & String(46, "-") & vbCrLf & _
          "Metodo: Arbol de Expansion Minima (Prim) = red de" & vbCrLf & _
          "menor longitud total que conecta la fuente con" & vbCrLf & _
          "todos los aspersores, sin bucles." & vbCrLf & vbCrLf & _
          "Aspersores conectados:  " & (nP - 1) & vbCrLf & _
          "Tramos de tuberia:      " & (nP - 1) & vbCrLf & _
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
          "Caudal total en la fuente: " & Format(acc(0), "0.0") & " l/min" & vbCrLf & _
          "Cada tramo esta en la capa RIEGO_TUB_<diametro>; use" & vbCrLf & _
          "DATAEXTRACTION filtrando por capa para totalizar metros" & vbCrLf & _
          "por diametro (lista de materiales)."
    MsgBox rep, vbInformation, "Trazado de tuberias"
    Exit Sub

errH:
    If Err.Number = -2147352567 Then Exit Sub
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Trazado de tuberias"
End Sub

'==============================================================================
' ALGORITMO DE PRIM  (arbol de expansion minima, metrica euclidiana)
'   O(n^2): adecuado para cientos/miles de aspersores.
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

    ' Nodo inicial = fuente (0)
    best(0) = 0#
    Dim nAdd As Long: nAdd = 0

    For i = 0 To nP - 1
        ' elegir el nodo no incluido con menor distancia al arbol
        Dim u As Long, mejor As Double
        u = -1: mejor = 1E+30
        For j = 0 To nP - 1
            If Not enArbol(j) Then
                If best(j) < mejor Then
                    mejor = best(j)
                    u = j
                End If
            End If
        Next
        If u = -1 Then Exit For

        enArbol(u) = True
        orden(nAdd) = u
        nAdd = nAdd + 1

        ' actualizar distancias de los nodos restantes hacia u
        For j = 0 To nP - 1
            If Not enArbol(j) Then
                Dim d As Double
                d = (pX(j) - pX(u)) ^ 2 + (pY(j) - pY(u)) ^ 2   ' cuadrado: basta para comparar
                If d < best(j) Then
                    best(j) = d
                    parent(j) = u
                End If
            End If
        Next
    Next
End Sub

'==============================================================================
' DIMENSIONAMIENTO: menor diametro comercial que cumple la velocidad maxima.
'   Q(l/min) -> m3/s: Q/60000 ;  v = Qm3s / (PI*d^2/4)  con d en metros.
'   d_min = raiz( 4*Qm3s / (PI*vMax) ).
'==============================================================================
Private Function ElegirDiametro(caudalLmin As Double, vMax As Double) As Long
    Dim qm3s As Double, dMin As Double, dm As Double, k As Long
    If caudalLmin <= 0# Then
        ElegirDiametro = 0                     ' sin caudal: el menor
        Exit Function
    End If
    qm3s = caudalLmin / 60000#
    dMin = Sqr(4# * qm3s / (PI * vMax)) * 1000#   ' en mm
    For k = 0 To nDiam - 1
        If gDiam(k) >= dMin - 0.000001 Then
            ElegirDiametro = k
            Exit Function
        End If
    Next
    ElegirDiametro = nDiam - 1                  ' supera el catalogo: el mayor
End Function

'==============================================================================
' CATALOGO DE DIAMETROS COMERCIALES (mm) Y COLOR ASOCIADO
'==============================================================================
Private Sub CargarCatalogoDiametros()
    Dim dd As Variant, cc As Variant, i As Long
    dd = Array(16#, 20#, 25#, 32#, 40#, 50#, 63#, 75#, 90#, 110#)
    '            azul verde cian rojo mag amar 30  5   40  1
    cc = Array(5, 3, 4, 1, 6, 2, 30, 8, 40, 200)
    nDiam = UBound(dd) + 1
    ReDim gDiam(nDiam - 1)
    ReDim gDiamColor(nDiam - 1)
    For i = 0 To nDiam - 1
        gDiam(i) = CDbl(dd(i))
        gDiamColor(i) = CLng(cc(i))
    Next
End Sub

'==============================================================================
' RECOLECCION AUTOMATICA de los aspersores colocados.
'   Se considera "aspersor" cualquier bloque cuyo nombre contenga "ASPERSOR"
'   O que tenga el atributo NUM o CAUDAL (firma del bloque de riego), asi
'   funciona aunque el bloque tenga otro nombre.
'   Lee INSERTIONPOINT y los atributos CAUDAL / NUM / ZONA.
'==============================================================================
Private Function RecolectarAspersores(zonaFiltro As String) As Boolean
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255)
    nP = 0

    Dim ent As AcadEntity, br As AcadBlockReference
    Dim nombre As String, esAspersor As Boolean
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            Set br = ent
            nombre = ""
            On Error Resume Next
            nombre = br.EffectiveName
            If nombre = "" Then nombre = br.Name
            On Error GoTo 0

            esAspersor = (InStr(1, UCase$(nombre), "ASPERSOR", vbTextCompare) > 0)
            If Not esAspersor Then
                esAspersor = TieneAtributo(br, "NUM") Or TieneAtributo(br, "CAUDAL")
            End If

            If esAspersor Then
                Dim zna As String
                zna = UCase$(Trim$(AtributoBloque(br, "ZONA")))
                If zonaFiltro = "" Or zna = zonaFiltro Then
                    Dim ip As Variant: ip = br.InsertionPoint
                    Dim qs As String: qs = AtributoBloque(br, "CAUDAL")
                    AgregarPunto CDbl(ip(0)), CDbl(ip(1)), Val(qs), _
                                 AtributoBloque(br, "NUM")
                End If
            End If
        End If
    Next

    RecolectarAspersores = (nP > 0)
End Function

'------------------------------------------------------------------------------
' Verdadero si el bloque tiene un atributo con ese TAG.
'------------------------------------------------------------------------------
Private Function TieneAtributo(br As AcadBlockReference, tag As String) As Boolean
    On Error Resume Next
    Dim atts As Variant, i As Long
    If br.HasAttributes Then
        atts = br.GetAttributes
        For i = LBound(atts) To UBound(atts)
            If UCase$(atts(i).TagString) = UCase$(tag) Then
                TieneAtributo = True
                Exit Function
            End If
        Next
    End If
End Function

'------------------------------------------------------------------------------
' Devuelve el texto del atributo cuyo TAG coincide (o "" si no existe).
'------------------------------------------------------------------------------
Private Function AtributoBloque(br As AcadBlockReference, tag As String) As String
    On Error Resume Next
    Dim atts As Variant, i As Long
    If br.HasAttributes Then
        atts = br.GetAttributes
        For i = LBound(atts) To UBound(atts)
            If UCase$(atts(i).TagString) = UCase$(tag) Then
                AtributoBloque = atts(i).TextString
                Exit Function
            End If
        Next
    End If
End Function

'==============================================================================
' SELECCION MANUAL (respaldo): bloques, circulos o puntos.
'==============================================================================
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

    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255)
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
            ' cualquier otra entidad: centro de su caja delimitadora
            Dim x As Double, y As Double
            If CentroEntidad(ent, x, y) Then AgregarPunto x, y, 0#, ""
        End If
    Next

    ss.Delete
    SeleccionManual = (nP > 0)
End Function

'------------------------------------------------------------------------------
' Centro de la caja delimitadora de una entidad (respaldo universal).
'------------------------------------------------------------------------------
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
' DIAGNOSTICO: cuenta entidades por tipo en el espacio modelo.
'==============================================================================
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

'==============================================================================
' COLECTA TODOS LOS BLOQUES (cualquier nombre) como aspersores.
'==============================================================================
Private Sub ColectarBloques(zonaFiltro As String)
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255)
    nP = 0
    Dim ent As AcadEntity, br As AcadBlockReference, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadBlockReference Then
            Set br = ent
            Dim zna As String
            zna = UCase$(Trim$(AtributoBloque(br, "ZONA")))
            If zonaFiltro = "" Or zna = zonaFiltro Then
                ip = br.InsertionPoint
                AgregarPunto CDbl(ip(0)), CDbl(ip(1)), _
                             Val(AtributoBloque(br, "CAUDAL")), _
                             AtributoBloque(br, "NUM")
            End If
        End If
    Next
End Sub

'==============================================================================
' COLECTA TODOS LOS CIRCULOS como aspersores (por su centro).
'==============================================================================
Private Sub ColectarCirculos()
    ReDim pX(255): ReDim pY(255): ReDim pQ(255): ReDim pNum(255)
    nP = 0
    Dim ent As AcadEntity, ip As Variant
    For Each ent In ThisDrawing.ModelSpace
        If TypeOf ent Is AcadCircle Then
            ip = ent.Center
            AgregarPunto CDbl(ip(0)), CDbl(ip(1)), 0#, ""
        End If
    Next
End Sub

'==============================================================================
' AUXILIARES DE PUNTOS
'==============================================================================
Private Sub AgregarPunto(x As Double, y As Double, q As Double, num As String)
    ' evita duplicados exactos (mismo aspersor contado dos veces)
    Dim i As Long
    For i = 0 To nP - 1
        If Abs(pX(i) - x) < 0.000001 And Abs(pY(i) - y) < 0.000001 Then Exit Sub
    Next
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
    End If
    pX(nP) = x: pY(nP) = y: pQ(nP) = q: pNum(nP) = num
    nP = nP + 1
End Sub

' Inserta la fuente como nodo 0, desplazando el resto una posicion.
Private Sub InsertarFuente(x As Double, y As Double)
    If nP > UBound(pX) Then
        ReDim Preserve pX(UBound(pX) + 256)
        ReDim Preserve pY(UBound(pY) + 256)
        ReDim Preserve pQ(UBound(pQ) + 256)
        ReDim Preserve pNum(UBound(pNum) + 256)
    End If
    Dim i As Long
    For i = nP To 1 Step -1
        pX(i) = pX(i - 1): pY(i) = pY(i - 1)
        pQ(i) = pQ(i - 1): pNum(i) = pNum(i - 1)
    Next
    pX(0) = x: pY(0) = y: pQ(0) = 0#: pNum(0) = "FUENTE"
    nP = nP + 1
End Sub

'==============================================================================
' ROTULO DE UN TRAMO (diametro y caudal) en su punto medio
'==============================================================================
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

    ' orienta el texto a lo largo del tramo (legible)
    ang = Atan2(y2 - y1, x2 - x1)
    If ang > PI / 2# Then ang = ang - PI
    If ang < -PI / 2# Then ang = ang + PI
    txt.Rotation = ang
    txt.Alignment = acAlignmentLeft
End Sub

'==============================================================================
' UTILIDADES DE DIBUJO
'==============================================================================
Private Function AlturaTexto() As Double
    Dim xmn As Double, xmx As Double, ymn As Double, ymx As Double, i As Long
    xmn = pX(0): xmx = pX(0): ymn = pY(0): ymx = pY(0)
    For i = 1 To nP - 1
        If pX(i) < xmn Then xmn = pX(i)
        If pX(i) > xmx Then xmx = pX(i)
        If pY(i) < ymn Then ymn = pY(i)
        If pY(i) > ymx Then ymx = pY(i)
    Next
    Dim diag As Double
    diag = Sqr((xmx - xmn) ^ 2 + (ymx - ymn) ^ 2)
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
