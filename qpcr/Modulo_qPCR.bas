Attribute VB_Name = "Modulo_qPCR"
' Análisis qPCR - ΔΔCt con PPIA y SYP
' Pegar datos en hoja RAW y ejecutar ProcesarPlaca (o botón asignado)

Option Explicit

Private Const HK_PPIA As String = "PPIA"
Private Const HK_SYP As String = "SYP"
Private Const SD_UMBRAL As Double = 0.3

Public Sub ProcesarPlaca()
    Dim wsRaw As Worksheet
    Dim wsRes As Worksheet
    Dim wsGlob As Worksheet
    Dim lastRow As Long
    Dim headerRow As Long
    Dim colSample As Long, colTarget As Long, colCt As Long
    Dim colCtMean As Long, colCtSd As Long
    Dim goi As String
    Dim sampleOrder As Collection
    Dim data As Object
    Dim i As Long
    Dim sample As String, tgt As String
    Dim key As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrHandler
    
    Set wsRaw = ThisWorkbook.Worksheets("RAW")
    EnsureSheet "Resultados"
    EnsureSheet "GLOBAL"
    Set wsRes = ThisWorkbook.Worksheets("Resultados")
    Set wsGlob = ThisWorkbook.Worksheets("GLOBAL")
    
    wsRes.Cells.Clear
    wsGlob.Cells.Clear
    
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1, , "No hay datos en la hoja RAW."
    
    headerRow = FindHeaderRow(wsRaw, lastRow, colSample, colTarget, colCt, colCtMean, colCtSd)
    If headerRow = 0 Then Err.Raise vbObjectError + 2, , "No se encontró la cabecera (Sample Name, Target Name, Ct)."
    
    Set data = CreateObject("Scripting.Dictionary")
    Set sampleOrder = New Collection
    
    For i = headerRow + 1 To lastRow
        sample = Trim$(UCase$(CStr(wsRaw.Cells(i, colSample).Value)))
        tgt = Trim$(UCase$(CStr(wsRaw.Cells(i, colTarget).Value)))
        If Not EsMuestraValida(sample) Then GoTo NextRow
        If tgt = "" Then GoTo NextRow
        
        key = sample & "|" & tgt
        If Not data.Exists(key) Then
            data.Add key, NuevoLectura()
            If tgt <> HK_PPIA And tgt <> HK_SYP Then
                If Not MuestraEnOrden(sampleOrder, sample) Then sampleOrder.Add sample
            End If
        End If
        Call AgregarLectura(data(key), wsRaw, i, colCt, colCtMean, colCtSd)
NextRow:
    Next i
    
    goi = DetectarGenInteres(data)
    If goi = "" Then Err.Raise vbObjectError + 3, , "No se detectó un único gen de interés."
    If Not TieneControl(data, HK_PPIA) Or Not TieneControl(data, HK_SYP) Then
        Err.Raise vbObjectError + 4, , "Faltan PPIA o SYP en los datos."
    End If
    
    Call EscribirResultados(wsRes, wsGlob, data, sampleOrder, goi)
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Análisis completado: " & goi, vbInformation, "qPCR"
    Exit Sub
    
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "qPCR"
End Sub

Private Sub EnsureSheet(ByVal name As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = name
    End If
End Sub

Private Function FindHeaderRow(ws As Worksheet, lastRow As Long, _
    ByRef colSample As Long, ByRef colTarget As Long, ByRef colCt As Long, _
    ByRef colCtMean As Long, ByRef colCtSd As Long) As Long
    Dim r As Long, c As Long
    Dim lbl As String
    Dim foundSample As Boolean, foundTarget As Boolean, foundCt As Boolean
    
    FindHeaderRow = 0
    For r = 1 To lastRow
        colSample = 0: colTarget = 0: colCt = 0: colCtMean = 0: colCtSd = 0
        foundSample = False: foundTarget = False: foundCt = False
        For c = 1 To 30
            lbl = NormalizarCabecera(CStr(ws.Cells(r, c).Value))
            If lbl = "" Then GoTo NextCol
            If InStr(lbl, "sample name") > 0 Or lbl = "sample" Then colSample = c: foundSample = True
            If InStr(lbl, "target name") > 0 Or lbl = "target" Then colTarget = c: foundTarget = True
            If lbl = "ct" Then colCt = c: foundCt = True
            If InStr(lbl, "ct mean") > 0 Then colCtMean = c
            If InStr(lbl, "ct sd") > 0 Then colCtSd = c
NextCol:
        Next c
        If foundSample And foundTarget And foundCt Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function NormalizarCabecera(s As String) As String
    Dim lbl As String
    lbl = LCase$(Trim$(s))
    lbl = Replace(lbl, ChrW$(1090), "t")  ' т cirílica → t
    NormalizarCabecera = lbl
End Function

Private Function EsMuestraValida(s As String) As Boolean
    Dim p As Long
    If Len(s) < 2 Then Exit Function
    If UCase$(Left$(s, 1)) <> "C" And UCase$(Left$(s, 1)) <> "S" Then Exit Function
    For p = 2 To Len(s)
        If Mid$(s, p, 1) < "0" Or Mid$(s, p, 1) > "9" Then Exit Function
    Next p
    EsMuestraValida = True
End Function

Private Function MuestraEnOrden(col As Collection, sample As String) As Boolean
    Dim i As Long
    MuestraEnOrden = False
    For i = 1 To col.Count
        If col(i) = sample Then MuestraEnOrden = True: Exit Function
    Next i
End Function

Private Function NuevoLectura() As Object
    Set NuevoLectura = CreateObject("Scripting.Dictionary")
    NuevoLectura("ct1") = Empty
    NuevoLectura("ct2") = Empty
    NuevoLectura("ctMean") = Empty
    NuevoLectura("ctSd") = Empty
    NuevoLectura("nCt") = 0
End Function

Private Sub AgregarLectura(lec As Object, ws As Worksheet, rowNum As Long, _
    colCt As Long, colCtMean As Long, colCtSd As Long)
    Dim v As Variant
    v = ws.Cells(rowNum, colCt).Value
    If IsNumeric(v) Then
        lec("nCt") = lec("nCt") + 1
        If lec("nCt") = 1 Then lec("ct1") = CDbl(v)
        If lec("nCt") = 2 Then lec("ct2") = CDbl(v)
    End If
    v = ws.Cells(rowNum, colCtMean).Value
    If IsNumeric(v) And IsEmpty(lec("ctMean")) Then lec("ctMean") = CDbl(v)
    v = ws.Cells(rowNum, colCtSd).Value
    If IsNumeric(v) And IsEmpty(lec("ctSd")) Then lec("ctSd") = CDbl(v)
End Sub

Private Function LeeMean(lec As Object) As Double
    If Not IsEmpty(lec("ctMean")) Then
        LeeMean = lec("ctMean")
    ElseIf lec("nCt") > 0 Then
        LeeMean = (NzDbl(lec("ct1")) + NzDbl(lec("ct2"))) / lec("nCt")
    Else
        LeeMean = 0
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNumeric(v) Then NzDbl = CDbl(v) Else NzDbl = 0
End Function

Private Function GetLectura(data As Object, sample As String, tgt As String) As Object
    Dim k As String
    k = sample & "|" & tgt
    If data.Exists(k) Then Set GetLectura = data(k) Else Set GetLectura = Nothing
End Function

Private Function DetectarGenInteres(data As Object) As String
    Dim k As Variant, parts() As String, t As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    For Each k In data.Keys
        parts = Split(k, "|")
        t = parts(1)
        If t <> HK_PPIA And t <> HK_SYP Then dict(t) = 1
    Next k
    If dict.Count = 1 Then
        DetectarGenInteres = dict.Keys()(0)
    Else
        DetectarGenInteres = ""
    End If
End Function

Private Function TieneControl(data As Object, hk As String) As Boolean
    Dim k As Variant
    For Each k In data.Keys
        If Right$(k, Len(hk) + 1) = "|" & hk Then TieneControl = True: Exit Function
    Next k
    TieneControl = False
End Function

Private Sub EscribirResultados(ws As Worksheet, wsG As Worksheet, data As Object, _
    sampleOrder As Collection, goi As String)
    Dim headers
    Dim rowOut As Long, i As Long, sample As String
    Dim lecGOI As Object, lecPPI As Object, lecSYP As Object
    Dim dctPPI As Object, dctSYP As Object
    Dim avgPPI As Double, avgSYP As Double
    Dim dCt As Double, ddCt As Double, fc As Double
    Dim goiMean As Double, goiSd As Double
    Dim flagRed As Boolean
    Dim controls As Collection, suicides As Collection
    Dim cSamples As Collection
    Dim goiShort As String
    
    Set dctPPI = CreateObject("Scripting.Dictionary")
    Set dctSYP = CreateObject("Scripting.Dictionary")
    Set cSamples = New Collection
    Set controls = New Collection
    Set suicides = New Collection
    
    headers = Array("Sample Name", "Target Name", "Ct", "Ct Mean", "Ct SD", _
        "ΔCt", "Prom. ΔCt (C)", "ΔΔCt", "2^(-ΔΔCt)")
    
    Call EscribirCabecerasBloque(ws, 1, "PPIA", headers, goi)
    Call EscribirCabecerasBloque(ws, 17, "SYP", headers, goi)
    
    ' Paso 1: ΔCt
    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        Set lecGOI = GetLectura(data, sample, goi)
        Set lecPPI = GetLectura(data, sample, HK_PPIA)
        Set lecSYP = GetLectura(data, sample, HK_SYP)
        If lecGOI Is Nothing Or lecPPI Is Nothing Or lecSYP Is Nothing Then GoTo NextS
        goiMean = LeeMean(lecGOI)
        dctPPI(sample) = goiMean - LeeMean(lecPPI)
        dctSYP(sample) = goiMean - LeeMean(lecSYP)
        If Left$(sample, 1) = "C" Then cSamples.Add sample
NextS:
    Next i
    
    If cSamples.Count = 0 Then Err.Raise vbObjectError + 5, , "No hay muestras C para el promedio."
    
    avgPPI = PromedioDictMuestras(dctPPI, cSamples)
    avgSYP = PromedioDictMuestras(dctSYP, cSamples)
    
    goiShort = NombreCortoGOI(goi)
    rowOut = 2
    
    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        If Not dctPPI.Exists(sample) Then GoTo NextOut
        Set lecGOI = GetLectura(data, sample, goi)
        goiMean = LeeMean(lecGOI)
        goiSd = NzDbl(lecGOI("ctSd"))
        flagRed = (goiSd > SD_UMBRAL)
        
        dCt = dctPPI(sample)
        ddCt = dCt - avgPPI
        fc = 2 ^ (-ddCt)
        Call EscribirFilaMuestra(ws, rowOut, 1, sample, goi, lecGOI, dCt, avgPPI, ddCt, fc, flagRed, True)
        
        dCt = dctSYP(sample)
        ddCt = dCt - avgSYP
        fc = 2 ^ (-ddCt)
        Call EscribirFilaMuestra(ws, rowOut, 17, sample, goi, lecGOI, dCt, avgSYP, ddCt, fc, flagRed, True)
        
        If Left$(sample, 1) = "C" Then controls.Add Array(sample, 2 ^ (-(dctPPI(sample) - avgPPI)), 2 ^ (-(dctSYP(sample) - avgSYP))
        If Left$(sample, 1) = "S" Then suicides.Add Array(sample, 2 ^ (-(dctPPI(sample) - avgPPI)), 2 ^ (-(dctSYP(sample) - avgSYP))
        
        rowOut = rowOut + 2
NextOut:
    Next i
    
    ' Genes control
    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        Set lecPPI = GetLectura(data, sample, HK_PPIA)
        Set lecSYP = GetLectura(data, sample, HK_SYP)
        If Not lecPPI Is Nothing Then Call EscribirFilaMuestra(ws, rowOut, 1, sample, HK_PPIA, lecPPI, 0, 0, 0, 0, False, False)
        If Not lecSYP Is Nothing Then Call EscribirFilaMuestra(ws, rowOut, 17, sample, HK_SYP, lecSYP, 0, 0, 0, 0, False, False)
        rowOut = rowOut + 2
    Next i
    
    Call EscribirGlobal(wsG, controls, suicides, goiShort)
End Sub

Private Function PromedioDictMuestras(d As Object, samples As Collection) As Double
    Dim i As Long, s As String, sum As Double
    sum = 0
    For i = 1 To samples.Count
        s = samples(i)
        sum = sum + d(s)
    Next i
    PromedioDictMuestras = sum / samples.Count
End Function

Private Function NombreCortoGOI(goi As String) As String
    Dim s As String
    s = goi
    If Right$(s, 1) = "r" Or Right$(s, 1) = "R" Then s = Left$(s, Len(s) - 1)
    NombreCortoGOI = s
End Function

Private Sub EscribirCabecerasBloque(ws As Worksheet, startCol As Long, refName As String, headers, goi As String)
    Dim j As Long
    For j = 0 To UBound(headers)
        ws.Cells(1, startCol + j).Value = headers(j)
        ws.Cells(1, startCol + j).Font.Bold = True
    Next j
End Sub

Private Sub EscribirFilaMuestra(ws As Worksheet, rowOut As Long, startCol As Long, _
    sample As String, tgt As String, lec As Object, _
    dCt As Double, avgC As Double, ddCt As Double, fc As Double, _
    flagRed As Boolean, conCalcs As Boolean)
    Dim fnt As Font
    Set fnt = ws.Cells(rowOut, startCol).Font
    If flagRed Then ws.Cells(rowOut, startCol).Font.Color = RGB(255, 0, 0)
    
    ws.Cells(rowOut, startCol).Value = sample
    ws.Cells(rowOut, startCol + 1).Value = tgt
    If Not IsEmpty(lec("ct1")) Then ws.Cells(rowOut, startCol + 2).Value = lec("ct1")
    ws.Cells(rowOut, startCol + 3).Value = LeeMean(lec)
    ws.Cells(rowOut, startCol + 4).Value = NzDbl(lec("ctSd"))
    If Not IsEmpty(lec("ct2")) Then ws.Cells(rowOut + 1, startCol + 2).Value = lec("ct2")
    
    If conCalcs Then
        ws.Cells(rowOut, startCol + 5).Value = dCt
        ws.Cells(rowOut, startCol + 6).Value = avgC
        ws.Cells(rowOut, startCol + 7).Value = ddCt
        ws.Cells(rowOut, startCol + 8).Value = fc
        If flagRed Then
            ws.Cells(rowOut, startCol + 5).Font.Color = RGB(255, 0, 0)
            ws.Cells(rowOut, startCol + 8).Font.Color = RGB(255, 0, 0)
        End If
    End If
End Sub

Private Sub EscribirGlobal(ws As Worksheet, controls As Collection, suicides As Collection, goiShort As String)
    Call EscribirTablaGlobal(ws, 1, "CONTROLES " & goiShort & " PFC", RGB(255, 255, 0), controls)
    Call EscribirTablaGlobal(ws, 6, "SUICIDAS " & goiShort & " PFC", RGB(146, 208, 80), suicides)
End Sub

Private Sub EscribirTablaGlobal(ws As Worksheet, startCol As Long, title As String, fillColor As Long, items As Collection)
    Dim i As Long, r As Long
    Dim ppi As Double, syp As Double, media As Double
    ws.Range(ws.Cells(1, startCol), ws.Cells(1, startCol + 3)).Merge
    ws.Cells(1, startCol).Value = title
    ws.Cells(1, startCol).Font.Bold = True
    ws.Cells(1, startCol).Interior.Color = fillColor
    ws.Cells(2, startCol).Value = "SUJETO"
    ws.Cells(2, startCol + 1).Value = "PPIA"
    ws.Cells(2, startCol + 2).Value = "SYP"
    ws.Cells(2, startCol + 3).Value = "MEDIA"
    For i = 1 To items.Count
        r = i + 2
        ppi = items(i)(1)
        syp = items(i)(2)
        media = (ppi + syp) / 2
        ws.Cells(r, startCol).Value = items(i)(0)
        ws.Cells(r, startCol + 1).Value = ppi
        ws.Cells(r, startCol + 2).Value = syp
        ws.Cells(r, startCol + 3).Value = media
    Next i
End Sub
