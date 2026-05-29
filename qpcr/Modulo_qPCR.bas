Attribute VB_Name = "Modulo_qPCR"
' qPCR - Pegar export StepOne completo en RAW (celda A1). Macro: ProcesarPlaca
' Sin Dictionary ni ReDim Preserve en bucle (evita error de pila)

Option Explicit

Private Const HK_PPIA As String = "PPIA"
Private Const HK_SYP As String = "SYP"
Private Const SD_UMBRAL As Double = 0.3
Private Const MAX_REG As Long = 500

Private G_Sample(1 To MAX_REG) As String
Private G_Target(1 To MAX_REG) As String
Private G_Ct1(1 To MAX_REG) As Double
Private G_Ct2(1 To MAX_REG) As Double
Private G_CtMean(1 To MAX_REG) As Double
Private G_CtSd(1 To MAX_REG) As Double
Private G_nDup(1 To MAX_REG) As Long
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

    headerRow = BuscarCabecera(wsRaw, cS, cT, cCt, cM, cSD)
    If headerRow = 0 Then Err.Raise 5, , "No se encontro cabecera Sample Name / Target Name / Ct."

    lastRow = UltimaFila(wsRaw, headerRow, cS)
    bloque = wsRaw.Range(wsRaw.Cells(headerRow + 1, 1), wsRaw.Cells(lastRow, 30)).Value2

    Set orden = New Collection
    G_N = 0
    goi = ""
    Call LeerBloque(bloque, cS, cT, cCt, cM, cSD, orden, goi)

    If G_N = 0 Then Err.Raise 5, , "No hay muestras validas."
    If goi = "" Then Err.Raise 5, , "No hay un unico gen de interes."
    If Not ExisteGen(HK_PPIA) Or Not ExisteGen(HK_SYP) Then Err.Raise 5, , "Faltan PPIA o SYP."

    Call EscribirTodo(wsRes, wsGlob, orden, goi)

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Analisis completado: " & goi, vbInformation, "qPCR"
    Exit Sub

ErrH:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "qPCR"
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
            If t = "ct" Then cCt = c: okC = True
            If InStr(t, "ct mean") > 0 Then cM = c
            If InStr(t, "ct sd") > 0 Then cSD = c
Nx:
        Next c
        If okS And okT And okC Then BuscarCabecera = r: Exit Function
    Next r
End Function

Private Function UltimaFila(ws As Worksheet, headerRow As Long, cS As Long) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, cS).End(xlUp).Row
    If lr > headerRow + 250 Then lr = headerRow + 250
    UltimaFila = lr
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
    If Not IsArray(bloque) Then ValTxt = "": Exit Function
    If col < 1 Or col > UBound(bloque, 2) Then ValTxt = "": Exit Function
    If f < 1 Or f > UBound(bloque, 1) Then ValTxt = "": Exit Function
    ValTxt = Trim$(CStr(bloque(f, col)))
End Function

Private Function ValNum(bloque As Variant, f As Long, col As Long) As Variant
    If Not IsArray(bloque) Then ValNum = Empty: Exit Function
    If col < 1 Or col > UBound(bloque, 2) Then ValNum = Empty: Exit Function
    If f < 1 Or f > UBound(bloque, 1) Then ValNum = Empty: Exit Function
    ValNum = bloque(f, col)
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
    NuevoIdx = G_N
End Function

Private Sub AnadirCt(idx As Long, ct As Double, m As Variant, s As Variant)
    If idx < 1 Or idx > G_N Then Exit Sub
    G_nDup(idx) = G_nDup(idx) + 1
    If G_nDup(idx) = 1 Then
        G_Ct1(idx) = ct
    ElseIf G_nDup(idx) = 2 Then
        G_Ct2(idx) = ct
    End If
    If IsNumeric(m) Then G_CtMean(idx) = CDbl(m)
    If IsNumeric(s) Then G_CtSd(idx) = CDbl(s)
End Sub

Private Function MediaIdx(idx As Long) As Double
    If G_CtMean(idx) <> 0# Then
        MediaIdx = G_CtMean(idx)
    ElseIf G_nDup(idx) >= 2 Then
        MediaIdx = (G_Ct1(idx) + G_Ct2(idx)) / 2#
    Else
        MediaIdx = G_Ct1(idx)
    End If
End Function

Private Sub LeerBloque(bloque As Variant, cS As Long, cT As Long, cCt As Long, _
    cM As Long, cSD As Long, orden As Collection, ByRef goi As String)

    Dim f As Long, nf As Long
    Dim sample As String, tgt As String
    Dim idx As Long
    Dim v As Variant
    Dim g1 As String, gN As Long
    Dim vacias As Long

    g1 = "": gN = 0: goi = ""
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

        If tgt <> HK_PPIA And tgt <> HK_SYP Then
            If g1 = "" Then g1 = tgt: gN = 1
            ElseIf g1 <> tgt Then gN = 2
            If Not EnCola(orden, sample) Then orden.Add sample
        End If

        idx = BuscarIdx(sample, tgt)
        If idx = 0 Then idx = NuevoIdx(sample, tgt)
        v = ValNum(bloque, f, cCt)
        If IsNumeric(v) Then Call AnadirCt(idx, CDbl(v), ValNum(bloque, f, cM), ValNum(bloque, f, cSD))
Sig:
    Next f

    If gN = 1 Then goi = g1
End Sub

Private Function ExisteGen(tgt As String) As Boolean
    Dim i As Long
    ExisteGen = False
    For i = 1 To G_N
        If G_Target(i) = tgt Then ExisteGen = True: Exit Function
    Next i
End Function

Private Function PromedioC(dCt() As Double, n As Long, orden As Collection) As Double
    Dim i As Long, suma As Double, cnt As Long
    suma = 0#: cnt = 0
    For i = 1 To n
        If Left$(orden(i), 1) = "C" Then
            suma = suma + dCt(i)
            cnt = cnt + 1
        End If
    Next i
    If cnt > 0 Then PromedioC = suma / CDbl(cnt)
End Function

Private Sub EscribirTodo(ws As Worksheet, wsG As Worksheet, orden As Collection, goi As String)
    Dim n As Long, i As Long, rowOut As Long
    Dim sample As String
    Dim ixG As Long, ixP As Long, ixS As Long
    Dim dP() As Double, dS() As Double
    Dim avgP As Double, avgS As Double
    Dim tit As String
    Dim cS As Collection, cP As Collection, cY As Collection
    Dim sS As Collection, sP As Collection, sY As Collection

    n = orden.Count
    If n = 0 Then Exit Sub
    ReDim dP(1 To n)
    ReDim dS(1 To n)
    Set cS = New Collection: Set cP = New Collection: Set cY = New Collection
    Set sS = New Collection: Set sP = New Collection: Set sY = New Collection

    Call Cabeceras(ws)

    For i = 1 To n
        sample = orden(i)
        ixG = BuscarIdx(sample, goi)
        ixP = BuscarIdx(sample, HK_PPIA)
        ixS = BuscarIdx(sample, HK_SYP)
        If ixG > 0 And ixP > 0 And ixS > 0 Then
            dP(i) = MediaIdx(ixG) - MediaIdx(ixP)
            dS(i) = MediaIdx(ixG) - MediaIdx(ixS)
        End If
    Next i

    avgP = PromedioC(dP, n, orden)
    avgS = PromedioC(dS, n, orden)
    tit = goi
    If Len(tit) > 0 Then If Right$(tit, 1) = "r" Or Right$(tit, 1) = "R" Then tit = Left$(tit, Len(tit) - 1)

    rowOut = 2
    For i = 1 To n
        sample = orden(i)
        ixG = BuscarIdx(sample, goi)
        If ixG = 0 Then GoTo SigO
        Call Fila(ws, rowOut, 1, sample, goi, ixG, dP(i), avgP, dP(i) - avgP, 2 ^ (-(dP(i) - avgP)), G_CtSd(ixG) > SD_UMBRAL, True)
        Call Fila(ws, rowOut, 17, sample, goi, ixG, dS(i), avgS, dS(i) - avgS, 2 ^ (-(dS(i) - avgS)), G_CtSd(ixG) > SD_UMBRAL, True)
        If Left$(sample, 1) = "C" Then
            cS.Add sample: cP.Add 2 ^ (-(dP(i) - avgP)): cY.Add 2 ^ (-(dS(i) - avgS))
        ElseIf Left$(sample, 1) = "S" Then
            sS.Add sample: sP.Add 2 ^ (-(dP(i) - avgP)): sY.Add 2 ^ (-(dS(i) - avgS))
        End If
        rowOut = rowOut + 2
SigO:
    Next i

    For i = 1 To n
        sample = orden(i)
        ixP = BuscarIdx(sample, HK_PPIA)
        ixS = BuscarIdx(sample, HK_SYP)
        If ixP > 0 Then Call Fila(ws, rowOut, 1, sample, HK_PPIA, ixP, 0, 0, 0, 1, False, False)
        If ixS > 0 Then Call Fila(ws, rowOut, 17, sample, HK_SYP, ixS, 0, 0, 0, 1, False, False)
        rowOut = rowOut + 2
    Next i

    Call Tabla(wsG, 1, "CONTROLES " & tit & " PFC", RGB(255, 255, 0), cS, cP, cY)
    Call Tabla(wsG, 6, "SUICIDAS " & tit & " PFC", RGB(146, 208, 80), sS, sP, sY)
End Sub

Private Sub Cabeceras(ws As Worksheet)
    Dim k As Long
    Dim h As Variant
    h = Array("Sample Name", "Target Name", "Ct", "Ct Mean", "Ct SD", "dCt", "Prom C", "ddCt", "2^-ddCt")
    For k = 0 To 8
        ws.Cells(1, k + 1).Value = h(k)
        ws.Cells(1, k + 1).Font.Bold = True
        ws.Cells(1, k + 18).Value = h(k)
        ws.Cells(1, k + 18).Font.Bold = True
    Next k
End Sub

Private Sub Fila(ws As Worksheet, r As Long, sc As Long, sample As String, tgt As String, ix As Long, _
    dCt As Double, avgC As Double, ddCt As Double, fc As Double, red As Boolean, calc As Boolean)
    ws.Cells(r, sc).Value = sample
    ws.Cells(r, sc + 1).Value = tgt
    If G_nDup(ix) >= 1 Then ws.Cells(r, sc + 2).Value = G_Ct1(ix)
    ws.Cells(r, sc + 3).Value = MediaIdx(ix)
    ws.Cells(r, sc + 4).Value = G_CtSd(ix)
    If G_nDup(ix) >= 2 Then ws.Cells(r + 1, sc + 2).Value = G_Ct2(ix)
    If red Then ws.Cells(r, sc).Font.Color = RGB(255, 0, 0)
    If calc Then
        ws.Cells(r, sc + 5).Value = dCt
        ws.Cells(r, sc + 6).Value = avgC
        ws.Cells(r, sc + 7).Value = ddCt
        ws.Cells(r, sc + 8).Value = fc
    End If
End Sub

Private Sub Tabla(ws As Worksheet, sc As Long, tit As String, clr As Long, _
    suj As Collection, vP As Collection, vY As Collection)
    Dim i As Long, r As Long, p As Double, y As Double
    ws.Range(ws.Cells(1, sc), ws.Cells(1, sc + 3)).Merge
    ws.Cells(1, sc).Value = tit
    ws.Cells(1, sc).Interior.Color = clr
    ws.Cells(1, sc).Font.Bold = True
    ws.Cells(2, sc).Value = "SUJETO"
    ws.Cells(2, sc + 1).Value = "PPIA"
    ws.Cells(2, sc + 2).Value = "SYP"
    ws.Cells(2, sc + 3).Value = "MEDIA"
    For i = 1 To suj.Count
        r = i + 2
        p = vP(i): y = vY(i)
        ws.Cells(r, sc).Value = suj(i)
        ws.Cells(r, sc + 1).Value = p
        ws.Cells(r, sc + 2).Value = y
        ws.Cells(r, sc + 3).Value = (p + y) / 2#
    Next i
End Sub
