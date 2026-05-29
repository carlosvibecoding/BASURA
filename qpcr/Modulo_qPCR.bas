Attribute VB_Name = "Modulo_qPCR"
' qPCR - Pegar export StepOne COMPLETO en RAW (celda A1). Macro: ProcesarPlaca
' Version 3.3 - Indeterminado, 1 replica marcada, orden GLOBAL

Option Explicit

Private Const MACRO_VER As String = "3.3"
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

Public Sub ProcesarPlaca()
    Dim wsRaw As Worksheet
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
    AsegurarHoja "Resultados"
    AsegurarHoja "GLOBAL"
    Set wsRes = ThisWorkbook.Worksheets("Resultados")
    Set wsGlob = ThisWorkbook.Worksheets("GLOBAL")
    wsRes.Cells.Clear
    wsGlob.Cells.Clear

    Set orden = New Collection
    G_N = 0
    Call LeerTodoRAW(wsRaw, orden)

    If G_N = 0 Then Err.Raise 5, , "No hay muestras validas."
    goi = DetectarGOI()
    If goi = "" Then Err.Raise 5, , MsgGenesEncontrados()
    If Not ExisteGen(HK_PPIA) Or Not ExisteGen(HK_SYP) Then
        Err.Raise 5, , "Faltan PPIA o SYP. Pegue el export COMPLETO del termociclador (los 3 bloques de genes)."
    End If

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

'--- Botones en hoja RAW ---
Public Sub InstalarBotones()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RAW")
    If ws Is Nothing Then Exit Sub
    ws.Shapes("btn_qPCR_Procesar").Delete
    ws.Shapes("btn_qPCR_Limpiar").Delete
    ws.Shapes("btn_qPCR_Anadir").Delete
    On Error GoTo 0
    Call CrearBotonFormulario(ws, "btn_qPCR_Procesar", "ProcesarPlaca", "Procesar placa", 8, 4, 128, 26)
    Call CrearBotonFormulario(ws, "btn_qPCR_Limpiar", "LimpiarDatos", "Limpiar datos", 142, 4, 110, 26)
    Call CrearBotonFormulario(ws, "btn_qPCR_Anadir", "AnadirPlacaRAW", "Anadir placa abajo", 258, 4, 128, 26)
End Sub

Private Sub CrearBotonFormulario(ws As Worksheet, nombre As String, macro As String, _
    texto As String, L As Single, T As Single, W As Single, H As Single)
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(Type:=xlButtonControl, Left:=L, Top:=T, Width:=W, Height:=H)
    btn.Name = nombre
    btn.OnAction = macro
    btn.TextFrame.Characters.Text = texto
End Sub

'--- Borra RAW, Resultados y GLOBAL ---
Public Sub LimpiarDatos()
    Dim r As VbMsgBoxResult
    r = MsgBox("Borrar todos los datos de RAW, Resultados y GLOBAL?", vbYesNo + vbQuestion, "qPCR")
    If r <> vbYes Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Worksheets("RAW").Cells.Clear
    ThisWorkbook.Worksheets("Resultados").Cells.Clear
    ThisWorkbook.Worksheets("GLOBAL").Cells.Clear
    On Error GoTo 0
    With ThisWorkbook.Worksheets("RAW")
        .Range("A1").Value = "Pegue aqui el export COMPLETO del StepOne (A1). Luego: Procesar placa."
        .Range("A1").Font.Italic = True
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

Private Function MuestraOK(ByVal s As String) As Boolean
    Dim p As Long
    Dim u As String
    u = UCase$(Trim$(s))
    If Len(u) < 2 Then Exit Function
    If Left$(u, 1) <> "C" And Left$(u, 1) <> "S" Then Exit Function
    For p = 2 To Len(u)
        If Mid$(u, p, 1) < "0" Or Mid$(u, p, 1) > "9" Then Exit Function
    Next p
    MuestraOK = True
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
    If IsNumeric(v) Then EsCtIndeterminado = False: Exit Function
    t = UCase$(Trim$(CStr(v)))
    If t = "" Then EsCtIndeterminado = True: Exit Function
    If InStr(t, "UNDETERMIN") > 0 Then EsCtIndeterminado = True: Exit Function
    If InStr(t, "INDETERMIN") > 0 Then EsCtIndeterminado = True: Exit Function
    If t = "N/A" Or t = "#N/A" Or t = "NA" Then EsCtIndeterminado = True
End Function

Private Sub AnadirCtValor(idx As Long, vCt As Variant, m As Variant, s As Variant)
    If idx < 1 Or idx > G_N Then Exit Sub
    G_nDup(idx) = G_nDup(idx) + 1
    If EsCtIndeterminado(vCt) Then
        G_nUndet(idx) = G_nUndet(idx) + 1
        If G_nDup(idx) = 1 Then G_Ct1Txt(idx) = "Indeterminado"
        If G_nDup(idx) = 2 Then G_Ct2Txt(idx) = "Indeterminado"
    ElseIf IsNumeric(vCt) Then
        G_nValidCt(idx) = G_nValidCt(idx) + 1
        If G_nDup(idx) = 1 Then
            G_Ct1(idx) = CDbl(vCt)
            G_Ct1Txt(idx) = Trim$(CStr(vCt))
        ElseIf G_nDup(idx) = 2 Then
            G_Ct2(idx) = CDbl(vCt)
            G_Ct2Txt(idx) = Trim$(CStr(vCt))
        End If
    End If
    If IsNumeric(m) And Not EsCtIndeterminado(m) Then
        G_CtMean(idx) = CDbl(m)
        G_HasMean(idx) = True
    End If
    If IsNumeric(s) Then G_CtSd(idx) = CDbl(s)
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
        MediaIdx = G_Ct1(idx)
        If G_Ct2Txt(idx) <> "" And G_nValidCt(idx) = 1 And G_Ct1Txt(idx) = "" Then MediaIdx = G_Ct2(idx)
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
        tgt = UCase$(ValTxt(bloque, f, cT))
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
        If Left$(orden(i), 1) <> "C" Then GoTo SigP
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
    Dim n As Long
    n = 0
    On Error Resume Next
    n = CLng(Val(Mid$(s, 2)))
    On Error GoTo 0
    If Left$(s, 1) = "C" Then
        OrdenClaveMuestra = n
    Else
        OrdenClaveMuestra = 1000000 + n
    End If
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

Private Sub EscribirTodo(ws As Worksheet, wsG As Worksheet, orden As Collection, goi As String)
    Dim n As Long, i As Long, rowOut As Long
    Dim sample As String
    Dim ixG As Long, ixP As Long, ixS As Long
    Dim dP() As Double, dS() As Double
    Dim avgP As Double, avgS As Double
    Dim tit As String
    Dim cS As Collection, cP As Collection, cY As Collection
    Dim sS As Collection, sP As Collection, sY As Collection

    Call OrdenarMuestras orden

    n = orden.Count
    If n = 0 Then Exit Sub
    ReDim dP(1 To n)
    ReDim dS(1 To n)
    Dim okCalc() As Boolean
    ReDim okCalc(1 To n)
    Set cS = New Collection: Set cP = New Collection: Set cY = New Collection
    Set sS = New Collection: Set sP = New Collection: Set sY = New Collection

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
            Call FilaCalc(ws, rowOut, 1, sample, goi, ixG, dP(i), avgP, True)
            Call FilaCalc(ws, rowOut, 17, sample, goi, ixG, dS(i), avgS, True)
            If Left$(sample, 1) = "C" Then
                cS.Add sample: cP.Add 2 ^ (-(dP(i) - avgP)): cY.Add 2 ^ (-(dS(i) - avgS))
            ElseIf Left$(sample, 1) = "S" Then
                sS.Add sample: sP.Add 2 ^ (-(dP(i) - avgP)): sY.Add 2 ^ (-(dS(i) - avgS))
            End If
        Else
            Call FilaIndeterminado(ws, rowOut, 1, sample, goi, ixG)
            Call FilaIndeterminado(ws, rowOut, 17, sample, goi, ixG)
            If Left$(sample, 1) = "C" Then
                cS.Add sample: cP.Add "Indeterminado": cY.Add "Indeterminado"
            ElseIf Left$(sample, 1) = "S" Then
                sS.Add sample: sP.Add "Indeterminado": sY.Add "Indeterminado"
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
        If ixS > 0 Then Call Fila(ws, rowOut, 17, sample, HK_SYP, ixS, 0, 0, 1, False, False)
        rowOut = rowOut + 2
    Next i

    Call TablaOrdenada(wsG, 1, "CONTROLES " & tit & " PFC", RGB(255, 255, 0), cS, cP, cY)
    Call TablaOrdenada(wsG, 6, "SUICIDAS " & tit & " PFC", RGB(146, 208, 80), sS, sP, sY)
End Sub

Private Sub Cabeceras(ws As Worksheet, goi As String)
    Dim k As Long
    Dim h As Variant
    h = Array("Sample Name", "Target Name", "Ct", "Ct Mean", "Ct SD", "dCt", "Prom. dCt (C)", "ddCt", "2^(-ddCt)")
    For k = 0 To 8
        ws.Cells(1, k + 1).Value = h(k)
        ws.Cells(1, k + 1).Font.Bold = True
        ws.Cells(1, k + 18).Value = h(k)
        ws.Cells(1, k + 18).Font.Bold = True
    Next k
    ws.Rows(2).Font.Bold = True
    ws.Rows(2).Interior.Color = RGB(242, 242, 242)
    On Error Resume Next
    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

Private Sub EscribirFilaPromedios(ws As Worksheet, avgP As Double, avgS As Double)
    ws.Cells(2, 1).Value = "PROMEDIO controles (C)"
    ws.Cells(2, 6).Value = avgP
    ws.Cells(2, 22).Value = avgS
    ws.Cells(2, 5).Value = "(fijo)"
    ws.Cells(2, 21).Value = "(fijo)"
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
    If G_Ct1Txt(ix) <> "" Then
        ws.Cells(r, sc + 2).Value = G_Ct1Txt(ix)
    ElseIf G_nValidCt(ix) >= 1 Then
        ws.Cells(r, sc + 2).Value = G_Ct1(ix)
    End If
    If G_Ct2Txt(ix) <> "" Then
        ws.Cells(r + 1, sc + 2).Value = G_Ct2Txt(ix)
    ElseIf G_nDup(ix) >= 2 And G_nValidCt(ix) >= 2 Then
        ws.Cells(r + 1, sc + 2).Value = G_Ct2(ix)
    End If
    If EsIndetIdx(ix) Then
        ws.Cells(r, sc + 3).Value = "Indeterminado"
    Else
        ws.Cells(r, sc + 3).Value = MediaIdx(ix)
    End If
    ws.Cells(r, sc + 4).Value = G_CtSd(ix)
End Sub

Private Sub FilaCalc(ws As Worksheet, r As Long, sc As Long, sample As String, tgt As String, ix As Long, _
    dCt As Double, avgC As Double, conCalcs As Boolean)
    Dim ddCt As Double, fc As Double
    Dim rojo As Boolean, naranja As Boolean
    ws.Cells(r, sc).Value = sample
    ws.Cells(r, sc + 1).Value = tgt
    Call EscribirCtCeldas(ws, r, sc, ix)
    rojo = (G_CtSd(ix) > SD_UMBRAL)
    naranja = UnReplicaIdx(ix)
    Call MarcarCelda(ws.Cells(r, sc), rojo Or EsIndetIdx(ix), naranja)
    If conCalcs Then
        ddCt = dCt - avgC
        fc = 2 ^ (-ddCt)
        ws.Cells(r, sc + 5).Value = dCt
        ws.Cells(r, sc + 7).Value = ddCt
        ws.Cells(r, sc + 8).Value = fc
        If fc > FC_EXTREMO Or fc < 1# / FC_EXTREMO Then rojo = True
        Call MarcarCelda(ws.Cells(r, sc + 8), rojo, naranja)
        Call MarcarCelda(ws.Cells(r, sc + 5), rojo, naranja)
    End If
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
