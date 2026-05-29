Attribute VB_Name = "Modulo_qPCR"
' qPCR - Pegar export StepOne COMPLETO en RAW (celda A1). Macro: ProcesarPlaca
' Version 3.8 - Grupos configurables; botones decorados en RAW

Option Explicit

Private Const MACRO_VER As String = "3.8"
Private Const MAX_GRUPOS_GLOBAL As Long = 24
Private Const COL_BLOQUE_SYP As Long = 17
Private Const CT_MIN As Double = 5#
Private Const CT_MAX As Double = 50#
Private Const SH_DATOS As String = "Datos"
Private Const SH_CALC As String = "Calculos"
Private Const HK_PPIA As String = "PPIA"
Private Const HK_SYP As String = "SYP"
Private Const SD_UMBRAL As Double = 0.3
Private Const FC_EXTREMO As Double = 1000#
Private Const MAX_REG As Long = 500

Private G_Sample(1 To MAX_REG) As String
Private G_Target(1 To MAX_REG) As String
Private G_Ct1(1 To MAX_REG) As Double
Private G_Ct2(1 To MAX_REG) As Double
Private G_CtMean(1 To MAX_REG) As Double
Private G_CtSd(1 To MAX_REG) As Double
Private G_nDup(1 To MAX_REG) As Long
Private G_nValidCt(1 To MAX_REG) As Long
Private G_nUndet(1 To MAX_REG) As Long
Private G_Ct1Txt(1 To MAX_REG) As String
Private G_Ct2Txt(1 To MAX_REG) As String
Private G_EsIndet(1 To MAX_REG) As Boolean
Private G_UnReplica(1 To MAX_REG) As Boolean
Private G_HasMean(1 To MAX_REG) As Boolean
Private G_N As Long
Private G_ConfigCtrl As String
Private G_ConfigMap As String
Private mGlobN As Long
Private mGlobPref(1 To MAX_GRUPOS_GLOBAL) As String
Private mGlobS(1 To MAX_GRUPOS_GLOBAL) As Collection
Private mGlobP(1 To MAX_GRUPOS_GLOBAL) As Collection
Private mGlobY(1 To MAX_GRUPOS_GLOBAL) As Collection

Public Sub ProcesarPlaca()
    Dim wsRaw As Worksheet
    Dim wsDat As Worksheet
    Dim wsCalc As Worksheet
    Dim wsRes As Worksheet
    Dim wsGlob As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim cS As Long, cT As Long, cCt As Long, cM As Long, cSD As Long
    Dim bloque As Variant
    Dim orden As Collection
    Dim goi As String

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsRaw = ThisWorkbook.Worksheets("RAW")
    AsegurarHoja SH_DATOS
    AsegurarHoja SH_CALC
    AsegurarHoja "Resultados"
    AsegurarHoja "GLOBAL"
    Set wsDat = ThisWorkbook.Worksheets(SH_DATOS)
    Set wsCalc = ThisWorkbook.Worksheets(SH_CALC)
    Set wsRes = ThisWorkbook.Worksheets("Resultados")
    Set wsGlob = ThisWorkbook.Worksheets("GLOBAL")
    wsDat.Cells.Clear
    wsCalc.Cells.Clear
    wsRes.Cells.Clear
    wsGlob.Cells.Clear

    Set orden = New Collection
    G_N = 0
    Call CargarConfigGrupos
    Call LeerTodoRAW(wsRaw, orden)

    If G_N = 0 Then Err.Raise 5, , "No hay muestras validas."
    goi = DetectarGOI()
    If goi = "" Then Err.Raise 5, , MsgGenesEncontrados()
    If Not ExisteGen(HK_PPIA) Or Not ExisteGen(HK_SYP) Then
        Err.Raise 5, , "Faltan PPIA o SYP. Pegue el export COMPLETO del termociclador (los 3 bloques de genes)."
    End If

    Call EscribirDatos(wsDat)
    Call EscribirCalculos(wsCalc, orden, goi)
    Call EscribirTodo(wsRes, wsGlob, orden, goi)
    Call InstalarBotones

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Analisis completado: " & goi & " (macro " & MACRO_VER & ")", vbInformation, "qPCR"
    Exit Sub

ErrH:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Macro " & MACRO_VER & vbCrLf & vbCrLf & Err.Description, vbCritical, "qPCR"
End Sub

'--- Se ejecuta al abrir el libro (crea botones en RAW) ---
Public Sub Auto_Open()
    Call InstalarBotones
End Sub

'--- Botones decorados a la derecha (no tapar zona de pegado A3+) ---
Public Sub InstalarBotones()
    Dim ws As Worksheet
    Dim L As Single, T As Single, esp As Single
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RAW")
    If ws Is Nothing Then Exit Sub
    ws.Shapes("btn_qPCR_Procesar").Delete
    ws.Shapes("btn_qPCR_Limpiar").Delete
    ws.Shapes("btn_qPCR_Anadir").Delete
    ws.Shapes("qPCR_panel_titulo").Delete
    On Error GoTo 0
    ws.Rows(1).RowHeight = 18
    ws.Rows(2).RowHeight = 34
    ws.Range("A1").Value = "qPCR — Pegar export StepOne desde fila 3 (celda A3)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 11
    ws.Range("A1").Font.Color = RGB(0, 70, 130)
    L = ws.Range("N2").Left
    T = ws.Range("N2").Top
    esp = 6
    Call CrearBotonDecorado(ws, "btn_qPCR_Procesar", "ProcesarPlaca", "Procesar placa", _
        L, T, 118, 30, RGB(0, 120, 212), RGB(255, 255, 255))
    Call CrearBotonDecorado(ws, "btn_qPCR_Limpiar", "LimpiarDatos", "Limpiar", _
        L + 118 + esp, T, 88, 30, RGB(245, 245, 245), RGB(180, 0, 0))
    Call CrearBotonDecorado(ws, "btn_qPCR_Anadir", "AnadirPlacaRAW", "+ Placa", _
        L + 118 + esp + 88 + esp, T, 88, 30, RGB(255, 192, 0), RGB(50, 50, 50))
    Call CrearEtiquetaPanel(ws, "qPCR_panel_titulo", "Acciones", L, T - 14, 300, 12)
End Sub

Private Sub CrearBotonDecorado(ws As Worksheet, nombre As String, macro As String, _
    texto As String, L As Single, T As Single, W As Single, H As Single, _
    rgbFill As Long, rgbText As Long)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, W, H)
    sh.Name = nombre
    sh.OnAction = macro
    sh.Fill.ForeColor.RGB = rgbFill
    sh.Line.ForeColor.RGB = RGB(180, 180, 180)
    sh.Line.Weight = 0.75
    With sh.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 4
        .MarginRight = 4
        .TextRange.Text = texto
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = rgbText
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub CrearEtiquetaPanel(ws As Worksheet, nombre As String, texto As String, _
    L As Single, T As Single, W As Single, H As Single)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, L, T, W, H)
    sh.Name = nombre
    sh.Fill.Visible = msoFalse
    sh.Line.Visible = msoFalse
    With sh.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Size = 8
        .TextRange.Font.Fill.ForeColor.RGB = RGB(100, 100, 100)
    End With
End Sub

'--- Borra RAW, Resultados y GLOBAL ---
Public Sub LimpiarDatos()
    Dim r As VbMsgBoxResult
    r = MsgBox("Borrar todos los datos de RAW, Datos, Calculos, Resultados y GLOBAL?", vbYesNo + vbQuestion, "qPCR")
    If r <> vbYes Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Worksheets("RAW").Cells.Clear
    ThisWorkbook.Worksheets(SH_DATOS).Cells.Clear
    ThisWorkbook.Worksheets(SH_CALC).Cells.Clear
    ThisWorkbook.Worksheets("Resultados").Cells.Clear
    ThisWorkbook.Worksheets("GLOBAL").Cells.Clear
    On Error GoTo 0
    With ThisWorkbook.Worksheets("RAW")
        .Range("A3").Select
    End With
    Call InstalarBotones
    MsgBox "Datos borrados.", vbInformation, "qPCR"
End Sub

'--- Deja una fila en blanco al final de RAW para pegar otra placa ---
Public Sub AnadirPlacaRAW()
    Dim ws As Worksheet
    Dim lr As Long
    Set ws = ThisWorkbook.Worksheets("RAW")
    lr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    ws.Cells(lr + 2, 1).Value = "--- Pegar aqui la siguiente placa (export completo) ---"
    ws.Cells(lr + 3, 1).Select
    MsgBox "Pegue el export de la nueva placa debajo de la linea indicada." & vbCrLf & _
        "Luego pulse Procesar placa (procesa TODAS las placas pegadas en RAW).", vbInformation, "qPCR"
End Sub

Private Sub AsegurarHoja(ByVal nombreHoja As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombreHoja
    End If
End Sub

Private Function NormTxt(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace$(t, ChrW$(1090), "t")
    NormTxt = t
End Function

Private Function BuscarCabecera(ws As Worksheet, ByRef cS As Long, ByRef cT As Long, _
    ByRef cCt As Long, ByRef cM As Long, ByRef cSD As Long) As Long
    Dim r As Long, c As Long, t As String
    Dim okS As Boolean, okT As Boolean, okC As Boolean
    BuscarCabecera = 0
    For r = 1 To 20
        cS = 0: cT = 0: cCt = 0: cM = 0: cSD = 0
        okS = False: okT = False: okC = False
        For c = 1 To 30
            t = NormTxt(CStr(ws.Cells(r, c).Value2))
            If t = "" Then GoTo Nx
            If InStr(t, "sample name") > 0 Then cS = c: okS = True
            If InStr(t, "target name") > 0 Then cT = c: okT = True
            If EsColumnaCtSuelto(t) Then cCt = c: okC = True
            If InStr(t, "ct mean") > 0 And InStr(t, "delta") = 0 And InStr(t, ChrW$(916)) = 0 Then cM = c
            If InStr(t, "ct sd") > 0 Then cSD = c
Nx:
        Next c
        If okS And okT And okC Then BuscarCabecera = r: Exit Function
    Next r
End Function

Private Function EsColumnaCtSuelto(ByVal t As String) As Boolean
    ' Solo la columna "Ct", no "dCt" / "ddCt"
    EsColumnaCtSuelto = False
    If t = "ct" Then EsColumnaCtSuelto = True: Exit Function
    If InStr(t, " ") > 0 Then Exit Function
    If InStr(t, "delta") > 0 Then Exit Function
    If InStr(t, ChrW$(916)) > 0 Then Exit Function
    If Len(t) = 2 And Left$(t, 1) = "c" And Right$(t, 1) = "t" Then EsColumnaCtSuelto = True
End Function

Private Function UltimaFila(ws As Worksheet, headerRow As Long, cS As Long, cT As Long) As Long
    Dim lrS As Long, lrT As Long, lr As Long
    lrS = ws.Cells(ws.Rows.Count, cS).End(xlUp).Row
    lrT = ws.Cells(ws.Rows.Count, cT).End(xlUp).Row
    If lrS > lrT Then lr = lrS Else lr = lrT
    If lr > headerRow + 400 Then lr = headerRow + 400
    If lr < headerRow + 1 Then lr = headerRow + 1
    UltimaFila = lr
End Function

Private Function LeerRango2D(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long) As Variant
    Dim v As Variant
    If r2 < r1 Then r2 = r1
    v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    LeerRango2D = AsegurarMatriz2D(v, r2 - r1 + 1, c2 - c1 + 1)
End Function

Private Function AsegurarMatriz2D(v As Variant, nFilas As Long, nCols As Long) As Variant
    Dim m() As Variant
    Dim i As Long, j As Long
    If nFilas < 1 Then nFilas = 1
    If nCols < 1 Then nCols = 1
    ReDim m(1 To nFilas, 1 To nCols)
    If Not IsArray(v) Then
        m(1, 1) = v
        AsegurarMatriz2D = m
        Exit Function
    End If
    On Error Resume Next
    j = UBound(v, 2)
    If Err.Number <> 0 Then
        Err.Clear
        For i = 1 To nFilas
            m(i, 1) = v(i)
        Next i
        AsegurarMatriz2D = m
        Exit Function
    End If
    On Error GoTo 0
    For i = 1 To nFilas
        For j = 1 To nCols
            m(i, j) = v(i, j)
        Next j
    Next i
    AsegurarMatriz2D = m
End Function

Private Function NormalizarGen(ByVal t As String) As String
    Dim u As String
    u = UCase$(Trim$(t))
    If u = HK_PPIA Or InStr(u, "PPIA") > 0 Then NormalizarGen = HK_PPIA: Exit Function
    If u = HK_SYP Or InStr(u, "SYP") > 0 Then NormalizarGen = HK_SYP: Exit Function
    NormalizarGen = u
End Function

'--- Configuracion de grupos (hoja Instrucciones B21-B22) ---
Private Sub CargarConfigGrupos()
    Dim ws As Worksheet
    Dim ctrl As String, mapa As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Instrucciones")
    On Error GoTo 0
    ctrl = "C"
    mapa = ""
    If Not ws Is Nothing Then
        If Trim$(CStr(ws.Range("B21").Value2)) <> "" Then ctrl = Trim$(CStr(ws.Range("B21").Value2))
        If Trim$(CStr(ws.Range("B22").Value2)) <> "" Then mapa = Trim$(CStr(ws.Range("B22").Value2))
    End If
    G_ConfigCtrl = PrefijosAPatron(ctrl)
    G_ConfigMap = MapaEtiquetasAPatron(mapa)
    If G_ConfigCtrl = "||" Then G_ConfigCtrl = "|C|"
End Sub

Private Function PrefijosAPatron(ByVal lista As String) As String
    Dim partes() As String, i As Long, p As String, r As String
    r = "|"
    partes = Split(Replace$(lista, " ", ""), ",")
    For i = LBound(partes) To UBound(partes)
        p = UCase$(Trim$(partes(i)))
        If Len(p) > 0 Then r = r & p & "|"
    Next i
    If r = "|" Then r = "|C|"
    PrefijosAPatron = r
End Function

Private Function MapaEtiquetasAPatron(ByVal mapa As String) As String
    Dim partes() As String, i As Long, kv() As String, r As String
    r = ""
    If Len(Trim$(mapa)) = 0 Then MapaEtiquetasAPatron = "": Exit Function
    mapa = Replace$(mapa, vbCrLf, ";")
    mapa = Replace$(mapa, ",", ";")
    partes = Split(mapa, ";")
    For i = LBound(partes) To UBound(partes)
        If InStr(partes(i), "=") > 0 Then
            kv = Split(partes(i), "=")
            If UBound(kv) >= 1 Then
                r = r & "|" & UCase$(Trim$(kv(0))) & "=" & Trim$(kv(1)) & "|"
            End If
        End If
    Next i
    MapaEtiquetasAPatron = r
End Function

Private Function PrefijoMuestra(ByVal s As String) As String
    Dim u As String, i As Long, c As String
    u = UCase$(Trim$(s))
    PrefijoMuestra = ""
    For i = 1 To Len(u)
        c = Mid$(u, i, 1)
        If c >= "0" And c <= "9" Then Exit For
        PrefijoMuestra = PrefijoMuestra & c
    Next i
End Function

Private Function NumeroMuestra(ByVal s As String) As Long
    Dim u As String, i As Long
    u = UCase$(Trim$(s))
    For i = 1 To Len(u)
        If Mid$(u, i, 1) >= "0" And Mid$(u, i, 1) <= "9" Then
            NumeroMuestra = CLng(Val(Mid$(u, i)))
            Exit Function
        End If
    Next i
End Function

Private Function EsMuestraControl(ByVal s As String) As Boolean
    Dim pref As String
    pref = PrefijoMuestra(s)
    If pref = "" Then Exit Function
    EsMuestraControl = (InStr(1, G_ConfigCtrl, "|" & pref & "|", vbTextCompare) > 0)
End Function

Private Function EtiquetaGrupo(ByVal pref As String) As String
    Dim clave As String, p As Long, q As Long, u As String
    u = UCase$(Trim$(pref))
    clave = "|" & u & "="
    p = InStr(1, G_ConfigMap, clave, vbTextCompare)
    If p > 0 Then
        q = InStr(p + Len(clave), G_ConfigMap, "|")
        If q > p Then
            EtiquetaGrupo = Mid$(G_ConfigMap, p + Len(clave), q - p - Len(clave))
            Exit Function
        End If
    End If
    EtiquetaGrupo = EtiquetaGrupoAuto(u)
End Function

Private Function EtiquetaGrupoAuto(ByVal pref As String) As String
    Select Case pref
        Case "C": EtiquetaGrupoAuto = "CONTROLES"
        Case "S": EtiquetaGrupoAuto = "SUJETOS"
        Case "A": EtiquetaGrupoAuto = "ALCOHOLICOS"
        Case "T": EtiquetaGrupoAuto = "TEST"
        Case Else: EtiquetaGrupoAuto = "GRUPO " & pref
    End Select
End Function

Private Function TituloPromedioControles() As String
  Dim p As Long
    p = InStr(2, G_ConfigCtrl, "|")
    If p > 2 Then
        TituloPromedioControles = "PROMEDIO " & EtiquetaGrupo(Mid$(G_ConfigCtrl, 2, p - 2))
    Else
        TituloPromedioControles = "PROMEDIO controles"
    End If
End Function

Private Function MuestraOK(ByVal s As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(s))
    If Len(u) < 2 Then Exit Function
    If Len(PrefijoMuestra(u)) < 1 Then Exit Function
    If NumeroMuestra(u) < 1 Then Exit Function
    MuestraOK = True
End Function

Private Sub ResetGruposGlobales()
    Dim i As Long
    mGlobN = 0
    For i = 1 To MAX_GRUPOS_GLOBAL
        Set mGlobS(i) = Nothing
        Set mGlobP(i) = Nothing
        Set mGlobY(i) = Nothing
        mGlobPref(i) = ""
    Next i
End Sub

Private Function IdxGrupoGlobal(ByVal pref As String) As Long
    Dim i As Long
    For i = 1 To mGlobN
        If mGlobPref(i) = pref Then IdxGrupoGlobal = i: Exit Function
    Next i
    If mGlobN >= MAX_GRUPOS_GLOBAL Then IdxGrupoGlobal = 0: Exit Function
    mGlobN = mGlobN + 1
    IdxGrupoGlobal = mGlobN
    mGlobPref(mGlobN) = pref
    Set mGlobS(mGlobN) = New Collection
    Set mGlobP(mGlobN) = New Collection
    Set mGlobY(mGlobN) = New Collection
End Function

Private Sub AnadirGlobalNoControl(ByVal sample As String, ByVal pref As String, ByVal vP As Variant, ByVal vY As Variant)
    Dim ix As Long
    ix = IdxGrupoGlobal(pref)
    If ix < 1 Then Exit Sub
    mGlobS(ix).Add sample
    mGlobP(ix).Add vP
    mGlobY(ix).Add vY
End Sub

Private Function ColorGrupoGlobal(ByVal ix As Long) As Long
    Select Case ((ix - 1) Mod 4)
        Case 0: ColorGrupoGlobal = RGB(146, 208, 80)
        Case 1: ColorGrupoGlobal = RGB(180, 198, 231)
        Case 2: ColorGrupoGlobal = RGB(255, 192, 0)
        Case Else: ColorGrupoGlobal = RGB(200, 230, 201)
    End Select
End Function

Private Function ValTxt(bloque As Variant, f As Long, col As Long) As String
    Dim v As Variant
    v = CeldaMatriz(bloque, f, col)
    If IsEmpty(v) Then ValTxt = "" Else ValTxt = Trim$(CStr(v))
End Function

Private Function ValNum(bloque As Variant, f As Long, col As Long) As Variant
    ValNum = CeldaMatriz(bloque, f, col)
End Function

Private Function CeldaMatriz(bloque As Variant, f As Long, col As Long) As Variant
    On Error Resume Next
    CeldaMatriz = Empty
    If Not IsArray(bloque) Then Exit Function
    If f < 1 Or col < 1 Then Exit Function
    If f > UBound(bloque, 1) Or col > UBound(bloque, 2) Then Exit Function
    CeldaMatriz = bloque(f, col)
End Function

Private Function EnCola(c As Collection, s As String) As Boolean
    Dim i As Long
    EnCola = False
    For i = 1 To c.Count
        If c(i) = s Then EnCola = True: Exit Function
    Next i
End Function

Private Function BuscarIdx(sample As String, tgt As String) As Long
    Dim i As Long
    BuscarIdx = 0
    For i = 1 To G_N
        If G_Sample(i) = sample And G_Target(i) = tgt Then BuscarIdx = i: Exit Function
    Next i
End Function

Private Function NuevoIdx(sample As String, tgt As String) As Long
    G_N = G_N + 1
    If G_N > MAX_REG Then G_N = MAX_REG
    G_Sample(G_N) = sample
    G_Target(G_N) = tgt
    G_Ct1(G_N) = 0#
    G_Ct2(G_N) = 0#
    G_CtMean(G_N) = 0#
    G_CtSd(G_N) = 0#
    G_nDup(G_N) = 0
    G_nValidCt(G_N) = 0
    G_nUndet(G_N) = 0
    G_Ct1Txt(G_N) = ""
    G_Ct2Txt(G_N) = ""
    G_EsIndet(G_N) = False
    G_UnReplica(G_N) = False
    G_HasMean(G_N) = False
    NuevoIdx = G_N
End Function

Private Function EsCtIndeterminado(v As Variant) As Boolean
    Dim t As String
    If IsEmpty(v) Then EsCtIndeterminado = True: Exit Function
    If IsNumeric(v) Then
        If CtEnRango(CDbl(v)) Then EsCtIndeterminado = False Else EsCtIndeterminado = True
        Exit Function
    End If
    t = UCase$(Trim$(CStr(v)))
    If t = "" Then EsCtIndeterminado = True: Exit Function
    If InStr(t, "UNDETERMIN") > 0 Then EsCtIndeterminado = True: Exit Function
    If InStr(t, "INDETERMIN") > 0 Then EsCtIndeterminado = True: Exit Function
    If t = "N/A" Or t = "#N/A" Or t = "NA" Then EsCtIndeterminado = True
End Function

Private Function CtEnRango(d As Double) As Boolean
    CtEnRango = (d >= CT_MIN And d <= CT_MAX)
End Function

' Convierte Ct del export (coma/punto, Excel ES) a numero; Empty si no valido
Private Function ParseCtNum(v As Variant) As Variant
    Dim d As Double
  ParseCtNum = Empty
    If EsCtIndeterminado(v) Then Exit Function
    If IsNumeric(v) Then
        d = CDbl(v)
        If CtEnRango(d) Then ParseCtNum = d: Exit Function
    End If
    d = ParseCtDesdeTexto(Trim$(CStr(v)))
    If CtEnRango(d) Then ParseCtNum = d
End Function

Private Function ParseCtDesdeTexto(s As String) As Double
    Dim t As String, p() As String, d As Double, nDot As Long, nCom As Long
    ParseCtDesdeTexto = 0#
    t = Replace$(Replace$(s, " ", ""), Chr$(160), "")
    If t = "" Then Exit Function
    nDot = Len(t) - Len(Replace$(t, ".", ""))
    nCom = Len(t) - Len(Replace$(t, ",", ""))
    If nCom = 1 And nDot = 0 Then
        t = Replace$(t, ",", ".")
    ElseIf nCom = 1 And nDot >= 1 Then
        t = Replace$(t, ".", "")
        t = Replace$(t, ",", ".")
    ElseIf nDot = 1 And nCom = 0 Then
        p = Split(t, ".")
        If UBound(p) = 1 And Len(p(0)) <= 2 Then
            On Error Resume Next
            d = CDbl(p(0) & Application.International(xlDecimalSeparator) & p(1))
            On Error GoTo 0
            If CtEnRango(d) Then ParseCtDesdeTexto = d: Exit Function
        End If
    ElseIf nDot > 1 And nCom = 0 Then
        t = Replace$(t, ".", "")
    End If
    On Error Resume Next
    d = CDbl(Val(t))
    On Error GoTo 0
    If CtEnRango(d) Then ParseCtDesdeTexto = d
End Function

Private Function ParseSdNum(v As Variant) As Variant
    Dim d As Double
    ParseSdNum = Empty
    If IsEmpty(v) Or v = "" Then Exit Function
    If IsNumeric(v) Then
        d = CDbl(v)
        If d >= 0# And d <= 5# Then ParseSdNum = d: Exit Function
    End If
    d = ParseCtDesdeTexto(Trim$(CStr(v)))
    If d >= 0# And d <= 5# Then ParseSdNum = d
End Function

Private Sub AnadirCtValor(idx As Long, vCt As Variant, m As Variant, s As Variant)
    Dim ctv As Variant, mv As Variant, sv As Variant
    If idx < 1 Or idx > G_N Then Exit Sub
    G_nDup(idx) = G_nDup(idx) + 1
    ctv = ParseCtNum(vCt)
    If IsEmpty(ctv) Then
        G_nUndet(idx) = G_nUndet(idx) + 1
        If G_nDup(idx) = 1 Then G_Ct1Txt(idx) = "Indeterminado"
        If G_nDup(idx) = 2 Then G_Ct2Txt(idx) = "Indeterminado"
    Else
        G_nValidCt(idx) = G_nValidCt(idx) + 1
        If G_nDup(idx) = 1 Then
            G_Ct1(idx) = CDbl(ctv)
            G_Ct1Txt(idx) = ""
        ElseIf G_nDup(idx) = 2 Then
            G_Ct2(idx) = CDbl(ctv)
            G_Ct2Txt(idx) = ""
        End If
    End If
    mv = ParseCtNum(m)
    If Not IsEmpty(mv) Then
        G_CtMean(idx) = CDbl(mv)
        G_HasMean(idx) = True
    End If
    sv = ParseSdNum(s)
    If Not IsEmpty(sv) Then G_CtSd(idx) = CDbl(sv)
End Sub

Private Sub FinalizarEstadoCt(idx As Long)
    If idx < 1 Or idx > G_N Then Exit Sub
    G_EsIndet(idx) = False
    G_UnReplica(idx) = False
    If G_nValidCt(idx) = 0 Then
        G_EsIndet(idx) = True
        Exit Sub
    End If
    If G_nValidCt(idx) = 1 Then G_UnReplica(idx) = True
    If G_nUndet(idx) >= 2 And G_nValidCt(idx) = 0 Then G_EsIndet(idx) = True
End Sub

Private Function MediaIdx(idx As Long) As Double
    Call FinalizarEstadoCt(idx)
    If G_EsIndet(idx) Then MediaIdx = 0#: Exit Function
    If G_HasMean(idx) Then
        MediaIdx = G_CtMean(idx)
        Exit Function
    End If
    If G_nValidCt(idx) >= 2 Then
        MediaIdx = (G_Ct1(idx) + G_Ct2(idx)) / 2#
    ElseIf G_nValidCt(idx) = 1 Then
        If G_Ct1Txt(idx) = "" Then
            MediaIdx = G_Ct1(idx)
        Else
            MediaIdx = G_Ct2(idx)
        End If
    Else
        MediaIdx = 0#
    End If
End Function

Private Function EsIndetIdx(idx As Long) As Boolean
    Call FinalizarEstadoCt(idx)
    EsIndetIdx = G_EsIndet(idx)
End Function

Private Function UnReplicaIdx(idx As Long) As Boolean
    Call FinalizarEstadoCt(idx)
    UnReplicaIdx = G_UnReplica(idx) And Not G_EsIndet(idx)
End Function

' Lee una o varias placas apiladas en RAW (cada una con su fila de cabecera)
Private Sub LeerTodoRAW(ws As Worksheet, orden As Collection)
    Dim r As Long
    Dim headerRow As Long
    Dim lastRow As Long
    Dim cS As Long, cT As Long, cCt As Long, cM As Long, cSD As Long
    Dim bloque As Variant
    Dim maxR As Long

    maxR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If maxR < 20 Then maxR = 400
    r = 1
    Do While r <= maxR
        headerRow = BuscarCabeceraDesde(ws, r, cS, cT, cCt, cM, cSD)
        If headerRow = 0 Then Exit Do
        lastRow = UltimaFila(ws, headerRow, cS, cT)
        bloque = LeerRango2D(ws, headerRow + 1, 1, lastRow, 30)
        Call LeerBloque(bloque, cS, cT, cCt, cM, cSD, orden)
        r = lastRow + 1
    Loop
End Sub

Private Function BuscarCabeceraDesde(ws As Worksheet, desdeFila As Long, _
    ByRef cS As Long, ByRef cT As Long, ByRef cCt As Long, _
    ByRef cM As Long, ByRef cSD As Long) As Long
    Dim r As Long, c As Long, t As String
    Dim okS As Boolean, okT As Boolean, okC As Boolean
  BuscarCabeceraDesde = 0
    For r = desdeFila To desdeFila + 25
        cS = 0: cT = 0: cCt = 0: cM = 0: cSD = 0
        okS = False: okT = False: okC = False
        For c = 1 To 30
            t = NormTxt(CStr(ws.Cells(r, c).Value2))
            If t = "" Then GoTo Nx2
            If InStr(t, "sample name") > 0 Then cS = c: okS = True
            If InStr(t, "target name") > 0 Then cT = c: okT = True
            If EsColumnaCtSuelto(t) Then cCt = c: okC = True
            If InStr(t, "ct mean") > 0 And InStr(t, "delta") = 0 Then cM = c
            If InStr(t, "ct sd") > 0 Then cSD = c
Nx2:
        Next c
        If okS And okT And okC Then BuscarCabeceraDesde = r: Exit Function
    Next r
End Function

Private Sub LeerBloque(bloque As Variant, cS As Long, cT As Long, cCt As Long, _
    cM As Long, cSD As Long, orden As Collection)

    Dim f As Long, nf As Long
    Dim sample As String, tgt As String
    Dim idx As Long
    Dim v As Variant
    Dim vacias As Long

    If Not IsArray(bloque) Then Exit Sub
    nf = UBound(bloque, 1)

    For f = 1 To nf
        sample = UCase$(ValTxt(bloque, f, cS))
        tgt = NormalizarGen(ValTxt(bloque, f, cT))
        If Not MuestraOK(sample) Or tgt = "" Then
            vacias = vacias + 1
            If vacias > 20 Then Exit For
            GoTo Sig
        End If
        vacias = 0
        If Not EnCola(orden, sample) Then orden.Add sample

        idx = BuscarIdx(sample, tgt)
        If idx = 0 Then idx = NuevoIdx(sample, tgt)
        Call AnadirCtValor(idx, ValNum(bloque, f, cCt), ValNum(bloque, f, cM), ValNum(bloque, f, cSD))
Sig:
    Next f
End Sub

Private Function EsTargetIgnorar(ByVal t As String) As Boolean
  ' Excluye controles, basura del export y muestras leidas por error como Target
    Select Case UCase$(Trim$(t))
        Case "", "UNKNOWN", "UNDETERMINED", "N/A", "-", "NTC", "NEGATIVE", "POSITIVE", "TASK"
            EsTargetIgnorar = True
        Case Else
            If MuestraOK(t) Then EsTargetIgnorar = True Else EsTargetIgnorar = False
    End Select
End Function

Private Function DetectarGOI() As String
    Dim i As Long
    Dim t As String
    Dim unico As String
    Dim n As Long

    unico = ""
    n = 0
    For i = 1 To G_N
        t = UCase$(Trim$(G_Target(i)))
        If t = HK_PPIA Or t = HK_SYP Then GoTo Sig
        If EsTargetIgnorar(t) Then GoTo Sig
        If n = 0 Then
            unico = t
            n = 1
        ElseIf unico <> t Then
            DetectarGOI = ""
            Exit Function
        End If
Sig:
    Next i
    If n = 1 Then DetectarGOI = unico
End Function

Private Function MsgGenesEncontrados() As String
    Dim i As Long
    Dim t As String
    Dim lista As String
    Dim visto As String

    lista = ""
    For i = 1 To G_N
        t = UCase$(Trim$(G_Target(i)))
        If t = HK_PPIA Or t = HK_SYP Or EsTargetIgnorar(t) Then GoTo Sig
        If InStr(1, "|" & visto & "|", "|" & t & "|", vbTextCompare) = 0 Then
            visto = visto & "|" & t
            If lista = "" Then lista = t Else lista = lista & ", " & t
        End If
Sig:
    Next i
    If lista = "" Then
        MsgGenesEncontrados = "No se detecto el gen de interes." & vbCrLf & vbCrLf & _
            "Causas habituales:" & vbCrLf & _
            "1) Solo pego el bloque RGS (falta PPIA y SYP debajo)." & vbCrLf & _
            "2) Macro antigua: importe Modulo_qPCR.bas version " & MACRO_VER & "."
    Else
        MsgGenesEncontrados = "Genes detectados: " & lista & vbCrLf & vbCrLf & _
            "Debe haber UN solo gen de interes + PPIA + SYP." & vbCrLf & _
            "Pegue el export COMPLETO del StepOne (los 3 bloques)."
    End If
End Function

Private Function ExisteGen(tgt As String) As Boolean
    Dim i As Long
    ExisteGen = False
    For i = 1 To G_N
        If G_Target(i) = tgt Then ExisteGen = True: Exit Function
    Next i
End Function

Private Function PromedioC(dCt() As Double, n As Long, orden As Collection, goi As String, usarPPI As Boolean) As Double
    Dim i As Long, suma As Double, cnt As Long
    Dim ixG As Long
    suma = 0#: cnt = 0
    For i = 1 To n
        If Not EsMuestraControl(CStr(orden(i))) Then GoTo SigP
        ixG = BuscarIdx(orden(i), goi)
        If ixG > 0 Then
            If EsIndetIdx(ixG) Then GoTo SigP
            If usarPPI Then
                If EsIndetIdx(BuscarIdx(orden(i), HK_PPIA)) Then GoTo SigP
            Else
                If EsIndetIdx(BuscarIdx(orden(i), HK_SYP)) Then GoTo SigP
            End If
        End If
        suma = suma + dCt(i)
        cnt = cnt + 1
SigP:
    Next i
    If cnt > 0 Then PromedioC = suma / CDbl(cnt)
End Function

Private Function OrdenClaveMuestra(s As String) As Long
    Dim n As Long, ordP As Long
    Dim pref As String
    n = NumeroMuestra(s)
    pref = PrefijoMuestra(s)
    If Len(pref) < 1 Then pref = "Z"
    If EsMuestraControl(s) Then
        ordP = 100 + Asc(Left$(pref, 1))
    Else
        ordP = 1000 + Asc(Left$(pref, 1))
    End If
    OrdenClaveMuestra = ordP * 10000 + n
End Function

Private Sub OrdenarMuestras(orden As Collection)
    Dim arr() As String
    Dim i As Long, j As Long
    Dim n As Long, tmp As String
    n = orden.Count
    If n < 2 Then Exit Sub
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = orden(i)
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If OrdenClaveMuestra(arr(j)) < OrdenClaveMuestra(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    Do While orden.Count > 0
        orden.Remove 1
    Loop
    For i = 1 To n
        orden.Add arr(i)
    Next i
End Sub

Private Sub EscribirCeldaCt(ws As Worksheet, r As Long, c As Long, txt As String, val As Double, tieneVal As Boolean)
    If txt = "Indeterminado" Then
        ws.Cells(r, c).Value = "Indeterminado"
    ElseIf tieneVal Then
        ws.Cells(r, c).Value = val
        ws.Cells(r, c).NumberFormat = "0.000000"
    End If
End Sub

'--- Tabla plana leida de RAW: numeros reales (no texto con miles mal parseados) ---
Private Sub EscribirDatos(ws As Worksheet)
    Dim i As Long, r As Long
    ws.Cells(1, 1).Value = "Muestra"
    ws.Cells(1, 2).Value = "Gen"
    ws.Cells(1, 3).Value = "Ct 1"
    ws.Cells(1, 4).Value = "Ct 2"
    ws.Cells(1, 5).Value = "Ct Mean"
    ws.Cells(1, 6).Value = "Ct SD"
    ws.Cells(1, 7).Value = "Indet"
    ws.Rows(1).Font.Bold = True
    r = 2
    For i = 1 To G_N
        ws.Cells(r, 1).Value = G_Sample(i)
        ws.Cells(r, 2).Value = G_Target(i)
        Call EscribirCeldaCt(ws, r, 3, G_Ct1Txt(i), G_Ct1(i), (G_Ct1Txt(i) = "" And G_nValidCt(i) >= 1))
        If G_nDup(i) >= 2 Then
            Call EscribirCeldaCt(ws, r, 4, G_Ct2Txt(i), G_Ct2(i), _
                (G_Ct2Txt(i) = "" And (G_nValidCt(i) >= 2 Or (G_nValidCt(i) = 1 And G_Ct1Txt(i) = "Indeterminado"))))
        End If
        If EsIndetIdx(i) Then
            ws.Cells(r, 5).Value = "Indeterminado"
            ws.Cells(r, 7).Value = 1
        Else
            ws.Cells(r, 5).Value = MediaIdx(i)
            ws.Cells(r, 5).NumberFormat = "0.000000"
            ws.Cells(r, 7).Value = 0
        End If
        If G_CtSd(i) > 0# Or Not EsIndetIdx(i) Then
            ws.Cells(r, 6).Value = G_CtSd(i)
            ws.Cells(r, 6).NumberFormat = "0.000000"
        End If
        r = r + 1
    Next i
    ws.Columns("A:G").ColumnWidth = 14
    ws.Columns("C:F").ColumnWidth = 12
End Sub

'--- Una fila por muestra: Ct y dCt (solo valores, sin formulas Excel) ---
Private Sub EscribirCalculos(ws As Worksheet, orden As Collection, goi As String)
    Dim n As Long, i As Long, r As Long
    Dim sample As String
    Dim ixG As Long, ixP As Long, ixS As Long
    Dim ctG As Variant, ctP As Variant, ctS As Variant
    Dim dP() As Double, dS() As Double
    Dim avgP As Double, avgS As Double

    Call OrdenarMuestras(orden)
    n = orden.Count
    If n = 0 Then Exit Sub
    ReDim dP(1 To n)
    ReDim dS(1 To n)

    ws.Cells(1, 1).Value = "Muestra"
    ws.Cells(1, 2).Value = "Ct GOI"
    ws.Cells(1, 3).Value = "Ct PPIA"
    ws.Cells(1, 4).Value = "Ct SYP"
    ws.Cells(1, 5).Value = "dCt PPIA"
    ws.Cells(1, 6).Value = "dCt SYP"
    ws.Cells(1, 8).Value = "Prom dCt (C) PPIA"
    ws.Cells(1, 9).Value = "Prom dCt (C) SYP"
    ws.Rows(1).Font.Bold = True
    ws.Range("G1").Value = goi
    ws.Range("G1").Font.Italic = True

    For i = 1 To n
        sample = orden(i)
        ixG = BuscarIdx(sample, goi)
        ixP = BuscarIdx(sample, HK_PPIA)
        ixS = BuscarIdx(sample, HK_SYP)
        r = i + 2
        ws.Cells(r, 1).Value = sample
        If ixG > 0 And Not EsIndetIdx(ixG) Then
            ctG = MediaIdx(ixG)
            ws.Cells(r, 2).Value = ctG
            ws.Cells(r, 2).NumberFormat = "0.000000"
        Else
            ws.Cells(r, 2).Value = "Indeterminado"
        End If
        If ixP > 0 And Not EsIndetIdx(ixP) Then
            ctP = MediaIdx(ixP)
            ws.Cells(r, 3).Value = ctP
            ws.Cells(r, 3).NumberFormat = "0.000000"
        Else
            ws.Cells(r, 3).Value = "Indeterminado"
        End If
        If ixS > 0 And Not EsIndetIdx(ixS) Then
            ctS = MediaIdx(ixS)
            ws.Cells(r, 4).Value = ctS
            ws.Cells(r, 4).NumberFormat = "0.000000"
        Else
            ws.Cells(r, 4).Value = "Indeterminado"
        End If
        If ixG > 0 And ixP > 0 And ixS > 0 Then
            If Not EsIndetIdx(ixG) And Not EsIndetIdx(ixP) And Not EsIndetIdx(ixS) Then
                dP(i) = CDbl(ctG) - CDbl(ctP)
                dS(i) = CDbl(ctG) - CDbl(ctS)
                ws.Cells(r, 5).Value = dP(i)
                ws.Cells(r, 6).Value = dS(i)
                ws.Cells(r, 5).NumberFormat = "0.000000"
                ws.Cells(r, 6).NumberFormat = "0.000000"
            Else
                ws.Cells(r, 5).Value = "Indeterminado"
                ws.Cells(r, 6).Value = "Indeterminado"
            End If
        Else
            ws.Cells(r, 5).Value = "Indeterminado"
            ws.Cells(r, 6).Value = "Indeterminado"
        End If
    Next i

    avgP = PromedioC(dP, n, orden, goi, True)
    avgS = PromedioC(dS, n, orden, goi, False)
    ws.Cells(2, 8).Value = avgP
    ws.Cells(2, 9).Value = avgS
    ws.Cells(2, 8).NumberFormat = "0.000000"
    ws.Cells(2, 9).NumberFormat = "0.000000"
    ws.Columns("A:I").ColumnWidth = 14
End Sub

Private Sub EscribirTodo(ws As Worksheet, wsG As Worksheet, orden As Collection, goi As String)
    Dim n As Long, i As Long, rowOut As Long
    Dim sample As String
    Dim ixG As Long, ixP As Long, ixS As Long
    Dim dP() As Double, dS() As Double
    Dim avgP As Double, avgS As Double
    Dim tit As String
    Dim cS As Collection, cP As Collection, cY As Collection

    Call OrdenarMuestras(orden)

    n = orden.Count
    If n = 0 Then Exit Sub
    ReDim dP(1 To n)
    ReDim dS(1 To n)
    Dim okCalc() As Boolean
    ReDim okCalc(1 To n)
    Set cS = New Collection: Set cP = New Collection: Set cY = New Collection
    Call ResetGruposGlobales

    Call Cabeceras(ws, goi)

    For i = 1 To n
        okCalc(i) = False
        sample = orden(i)
        ixG = BuscarIdx(sample, goi)
        ixP = BuscarIdx(sample, HK_PPIA)
        ixS = BuscarIdx(sample, HK_SYP)
        If ixG > 0 And ixP > 0 And ixS > 0 Then
            If Not EsIndetIdx(ixG) And Not EsIndetIdx(ixP) And Not EsIndetIdx(ixS) Then
                dP(i) = MediaIdx(ixG) - MediaIdx(ixP)
                dS(i) = MediaIdx(ixG) - MediaIdx(ixS)
                okCalc(i) = True
            End If
        End If
    Next i

    avgP = PromedioC(dP, n, orden, goi, True)
    avgS = PromedioC(dS, n, orden, goi, False)
    Call EscribirFilaPromedios(ws, avgP, avgS)
    tit = goi
    If Len(tit) > 0 Then If Right$(tit, 1) = "r" Or Right$(tit, 1) = "R" Then tit = Left$(tit, Len(tit) - 1)

    rowOut = 3
    For i = 1 To n
        sample = orden(i)
        ixG = BuscarIdx(sample, goi)
        If ixG = 0 Then GoTo SigO
        If okCalc(i) Then
            Call FilaCalc(ws, rowOut, 1, sample, goi, ixG, dP(i), avgP, _
                (2 ^ (-(dP(i) - avgP)) > FC_EXTREMO Or 2 ^ (-(dP(i) - avgP)) < 1# / FC_EXTREMO))
            Call FilaCalc(ws, rowOut, COL_BLOQUE_SYP, sample, goi, ixG, dS(i), avgS, _
                (2 ^ (-(dS(i) - avgS)) > FC_EXTREMO Or 2 ^ (-(dS(i) - avgS)) < 1# / FC_EXTREMO))
            If EsMuestraControl(sample) Then
                cS.Add sample: cP.Add 2 ^ (-(dP(i) - avgP)): cY.Add 2 ^ (-(dS(i) - avgS))
            Else
                Call AnadirGlobalNoControl(sample, PrefijoMuestra(sample), _
                    2 ^ (-(dP(i) - avgP)), 2 ^ (-(dS(i) - avgS)))
            End If
        Else
            Call FilaIndeterminado(ws, rowOut, 1, sample, goi, ixG)
            Call FilaIndeterminado(ws, rowOut, COL_BLOQUE_SYP, sample, goi, ixG)
            If EsMuestraControl(sample) Then
                cS.Add sample: cP.Add "Indeterminado": cY.Add "Indeterminado"
            Else
                Call AnadirGlobalNoControl(sample, PrefijoMuestra(sample), "Indeterminado", "Indeterminado")
            End If
        End If
        rowOut = rowOut + 2
SigO:
    Next i

    For i = 1 To n
        sample = orden(i)
        ixP = BuscarIdx(sample, HK_PPIA)
        ixS = BuscarIdx(sample, HK_SYP)
        If ixP > 0 Then Call Fila(ws, rowOut, 1, sample, HK_PPIA, ixP, 0, 0, 1, False, False)
        If ixS > 0 Then Call Fila(ws, rowOut, COL_BLOQUE_SYP, sample, HK_SYP, ixS, 0, 0, 1, False, False)
        rowOut = rowOut + 2
    Next i

    Call EscribirGlobalDinamico(wsG, tit, cS, cP, cY)
End Sub

Private Sub EscribirGlobalDinamico(wsG As Worksheet, tit As String, _
    cS As Collection, cP As Collection, cY As Collection)
    Dim i As Long, col As Long
    Dim titCtrl As String
    titCtrl = EtiquetaGrupo(Mid$(G_ConfigCtrl, 2, InStr(2, G_ConfigCtrl, "|") - 2))
    If titCtrl = "" Then titCtrl = "CONTROLES"
    Call TablaOrdenada(wsG, 1, UCase$(titCtrl) & " " & tit & " PFC", RGB(255, 255, 0), cS, cP, cY)
    col = 6
    For i = 1 To mGlobN
        Call TablaOrdenada(wsG, col, UCase$(EtiquetaGrupo(mGlobPref(i))) & " " & tit & " PFC", _
            ColorGrupoGlobal(i), mGlobS(i), mGlobP(i), mGlobY(i))
        col = col + 5
        If col > 26 Then Exit For
    Next i
End Sub

Private Sub Cabeceras(ws As Worksheet, goi As String)
    Dim k As Long
    Dim h As Variant
    h = Array("Sample Name", "Target Name", "Ct", "Ct Mean", "Ct SD", "dCt", "Prom. dCt (C)", "ddCt", "2^(-ddCt)")
    For k = 0 To 8
        ws.Cells(1, k + 1).Value = h(k)
        ws.Cells(1, k + 1).Font.Bold = True
        ws.Cells(1, k + COL_BLOQUE_SYP).Value = h(k)
        ws.Cells(1, k + COL_BLOQUE_SYP).Font.Bold = True
    Next k
    ws.Cells(1, 7).Value = "Prom. dCt (" & Mid$(G_ConfigCtrl, 2, InStr(2, G_ConfigCtrl, "|") - 2) & ")"
    ws.Cells(1, COL_BLOQUE_SYP + 6).Value = ws.Cells(1, 7).Value
    ws.Rows(2).Font.Bold = True
    ws.Rows(2).Interior.Color = RGB(242, 242, 242)
    On Error Resume Next
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Sub EscribirFilaPromedios(ws As Worksheet, avgP As Double, avgS As Double)
    ' Promedio C en columna "Prom. dCt (C)" (G y W), como en PLACA referencia
    Dim cPromP As Long, cPromS As Long
    cPromP = 7
    cPromS = COL_BLOQUE_SYP + 6
    ws.Cells(2, 1).Value = TituloPromedioControles()
    ws.Cells(2, cPromP).Value = avgP
    ws.Cells(2, cPromS).Value = avgS
    ws.Cells(2, cPromP).NumberFormat = "0.000000"
    ws.Cells(2, cPromS).NumberFormat = "0.000000"
    ws.Cells(2, 5).Value = "(fijo)"
    ws.Cells(2, COL_BLOQUE_SYP + 4).Value = "(fijo)"
End Sub

Private Sub MarcarCelda(c As Range, rojo As Boolean, naranja As Boolean)
    If rojo Then
        c.Font.Color = RGB(255, 0, 0)
        c.Font.Bold = True
    ElseIf naranja Then
        c.Font.Color = RGB(255, 128, 0)
        c.Font.Bold = True
    End If
End Sub

Private Sub EscribirCtCeldas(ws As Worksheet, r As Long, sc As Long, ix As Long)
    Call EscribirCeldaCt(ws, r, sc + 2, G_Ct1Txt(ix), G_Ct1(ix), (G_Ct1Txt(ix) = "" And G_nValidCt(ix) >= 1))
    If G_nDup(ix) >= 2 Then
        Call EscribirCeldaCt(ws, r + 1, sc + 2, G_Ct2Txt(ix), G_Ct2(ix), (G_Ct2Txt(ix) = "" And G_nValidCt(ix) >= 1))
        If G_Ct2Txt(ix) = "" And G_nValidCt(ix) = 1 And G_Ct1Txt(ix) = "Indeterminado" Then
            Call EscribirCeldaCt(ws, r + 1, sc + 2, "", G_Ct2(ix), True)
        End If
    End If
    ws.Columns(sc + 2).ColumnWidth = 12
    If EsIndetIdx(ix) Then
        ws.Cells(r, sc + 3).Value = "Indeterminado"
    Else
        ws.Cells(r, sc + 3).Value = MediaIdx(ix)
        ws.Cells(r, sc + 3).NumberFormat = "0.000000"
    End If
    ws.Cells(r, sc + 4).Value = G_CtSd(ix)
    ws.Cells(r, sc + 4).NumberFormat = "0.000000"
End Sub

Private Sub FilaCalc(ws As Worksheet, r As Long, sc As Long, sample As String, tgt As String, ix As Long, _
    dCt As Double, avgC As Double, flagFcExtremo As Boolean)
    Dim ddCt As Double, fc As Double
    Dim rojo As Boolean, naranja As Boolean

    ws.Cells(r, sc).Value = sample
    ws.Cells(r, sc + 1).Value = tgt
    Call EscribirCtCeldas(ws, r, sc, ix)
    rojo = (G_CtSd(ix) > SD_UMBRAL)
    naranja = UnReplicaIdx(ix)
    Call MarcarCelda(ws.Cells(r, sc), rojo Or EsIndetIdx(ix), naranja)
    ddCt = dCt - avgC
    fc = 2 ^ (-ddCt)
    ws.Cells(r, sc + 5).Value = dCt
    ws.Cells(r, sc + 5).NumberFormat = "0.000000"
    ws.Cells(r, sc + 7).Value = ddCt
    ws.Cells(r, sc + 7).NumberFormat = "0.000000"
    ws.Cells(r, sc + 8).Value = fc
    ws.Cells(r, sc + 8).NumberFormat = "0.000000"
    If flagFcExtremo Then rojo = True
    Call MarcarCelda(ws.Cells(r, sc + 8), rojo, naranja)
    Call MarcarCelda(ws.Cells(r, sc + 5), rojo, naranja)
End Sub

Private Sub FilaIndeterminado(ws As Worksheet, r As Long, sc As Long, sample As String, tgt As String, ix As Long)
    ws.Cells(r, sc).Value = sample
    ws.Cells(r, sc + 1).Value = tgt
    Call EscribirCtCeldas(ws, r, sc, ix)
    ws.Cells(r, sc + 5).Value = "Indeterminado"
    ws.Cells(r, sc + 7).Value = "Indeterminado"
    ws.Cells(r, sc + 8).Value = "Indeterminado"
    Call MarcarCelda(ws.Cells(r, sc), True, UnReplicaIdx(ix))
    Call MarcarCelda(ws.Cells(r, sc + 8), True, False)
End Sub

Private Sub Fila(ws As Worksheet, r As Long, sc As Long, sample As String, tgt As String, ix As Long, _
    dCt As Double, ddCt As Double, fc As Double, red As Boolean, calc As Boolean)
    ws.Cells(r, sc).Value = sample
    ws.Cells(r, sc + 1).Value = tgt
    Call EscribirCtCeldas(ws, r, sc, ix)
    If calc Then
        ws.Cells(r, sc + 5).Value = dCt
        ws.Cells(r, sc + 7).Value = ddCt
        ws.Cells(r, sc + 8).Value = fc
    End If
End Sub

Private Sub TablaOrdenada(ws As Worksheet, sc As Long, tit As String, clr As Long, _
    suj As Collection, vP As Collection, vY As Collection)
    Dim n As Long, i As Long, j As Long, r As Long
    Dim ord() As Long, tmp As Long
    Dim p As Variant, y As Variant, media As Variant
    n = suj.Count
    If n = 0 Then Exit Sub
    ReDim ord(1 To n)
    For i = 1 To n
        ord(i) = i
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If OrdenClaveMuestra(CStr(suj(ord(j)))) < OrdenClaveMuestra(CStr(suj(ord(i)))) Then
                tmp = ord(i): ord(i) = ord(j): ord(j) = tmp
            End If
        Next j
    Next i
    ws.Range(ws.Cells(1, sc), ws.Cells(1, sc + 3)).Merge
    ws.Cells(1, sc).Value = tit
    ws.Cells(1, sc).Interior.Color = clr
    ws.Cells(1, sc).Font.Bold = True
    ws.Cells(2, sc).Value = "SUJETO"
    ws.Cells(2, sc + 1).Value = "PPIA"
    ws.Cells(2, sc + 2).Value = "SYP"
    ws.Cells(2, sc + 3).Value = "MEDIA"
    For i = 1 To n
        r = i + 2
        p = vP(ord(i)): y = vY(ord(i))
        ws.Cells(r, sc).Value = suj(ord(i))
        ws.Cells(r, sc + 1).Value = p
        ws.Cells(r, sc + 2).Value = y
        If IsNumeric(p) And IsNumeric(y) Then
            media = (CDbl(p) + CDbl(y)) / 2#
            ws.Cells(r, sc + 3).Value = media
            If CDbl(p) > FC_EXTREMO Or CDbl(y) > FC_EXTREMO Then
                Call MarcarCelda(ws.Cells(r, sc + 3), True, False)
            End If
        Else
            ws.Cells(r, sc + 3).Value = "Indeterminado"
            Call MarcarCelda(ws.Cells(r, sc + 3), True, False)
        End If
        If VarType(p) = vbString Or VarType(y) = vbString Then
            Call MarcarCelda(ws.Cells(r, sc), True, False)
        End If
    Next i
End Sub
