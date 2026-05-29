Attribute VB_Name = "Modulo_qPCR"
' qPCR - Acepta el export del StepOne pegado tal cual en RAW (cabecera fila ~8)
' Ejecutar: ProcesarPlaca

Option Explicit

Private Const HK_PPIA As String = "PPIA"
Private Const HK_SYP As String = "SYP"
Private Const SD_UMBRAL As Double = 0.3

Public Sub ProcesarPlaca()
    Dim wsRaw As Worksheet
    Dim wsRes As Worksheet
    Dim wsGlob As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim colSample As Long, colTarget As Long, colCt As Long
    Dim colCtMean As Long, colCtSd As Long
    Dim goi As String
    Dim sampleOrder As Collection
    Dim data As Object
    Dim i As Long
    Dim sample As String, tgt As String, clave As String
    Dim vacias As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsRaw = ThisWorkbook.Worksheets("RAW")
    Call AsegurarHoja("Resultados")
    Call AsegurarHoja("GLOBAL")
    Set wsRes = ThisWorkbook.Worksheets("Resultados")
    Set wsGlob = ThisWorkbook.Worksheets("GLOBAL")
    wsRes.Cells.Clear
    wsGlob.Cells.Clear

    headerRow = BuscarFilaCabecera(wsRaw, colSample, colTarget, colCt, colCtMean, colCtSd)
    If headerRow = 0 Then Err.Raise vbObjectError + 2, , "No se encontró la cabecera (Sample Name, Target Name, Ct)."

    lastRow = UltimaFilaDatos(wsRaw, headerRow, colSample)
    If lastRow <= headerRow Then Err.Raise vbObjectError + 1, , "No hay filas de datos debajo de la cabecera."

    Set data = CreateObject("Scripting.Dictionary")
    Set sampleOrder = New Collection
    vacias = 0

    For i = headerRow + 1 To lastRow
        sample = Trim$(UCase$(CStr(wsRaw.Cells(i, colSample).Value)))
        tgt = Trim$(UCase$(CStr(wsRaw.Cells(i, colTarget).Value)))

        If Not EsMuestraValida(sample) Or tgt = "" Then
            vacias = vacias + 1
            If vacias > 15 Then Exit For
            GoTo SiguienteFila
        End If
        vacias = 0

        clave = sample & "|" & tgt
        If Not data.Exists(clave) Then
            data.Add clave, NuevoLectura()
            If tgt <> HK_PPIA And tgt <> HK_SYP Then
                If Not MuestraEnCola(sampleOrder, sample) Then sampleOrder.Add sample
            End If
        End If
        Call AgregarLectura(data(clave), wsRaw, i, colCt, colCtMean, colCtSd)
SiguienteFila:
    Next i

    goi = DetectarGenInteres(data)
    If goi = "" Then Err.Raise vbObjectError + 3, , "No se detectó un único gen de interés."
    If Not ExisteTarget(data, HK_PPIA) Or Not ExisteTarget(data, HK_SYP) Then
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

Private Function BuscarFilaCabecera(ws As Worksheet, _
    ByRef colSample As Long, ByRef colTarget As Long, ByRef colCt As Long, _
    ByRef colCtMean As Long, ByRef colCtSd As Long) As Long

    Dim r As Long, c As Long
    Dim lbl As String
    Dim okSample As Boolean, okTarget As Boolean, okCt As Boolean
    Dim maxScan As Long

    BuscarFilaCabecera = 0
    maxScan = 25
    If ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1 < maxScan Then
        maxScan = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    End If

    For r = 1 To maxScan
        colSample = 0: colTarget = 0: colCt = 0: colCtMean = 0: colCtSd = 0
        okSample = False: okTarget = False: okCt = False
        For c = 1 To 35
            lbl = NormalizarTexto(CStr(ws.Cells(r, c).Value))
            If lbl = "" Then GoTo SigCol
            If InStr(1, lbl, "sample name", vbTextCompare) > 0 Then colSample = c: okSample = True
            If InStr(1, lbl, "target name", vbTextCompare) > 0 Then colTarget = c: okTarget = True
            If lbl = "ct" Then colCt = c: okCt = True
            If InStr(1, lbl, "ct mean", vbTextCompare) > 0 Then colCtMean = c
            If InStr(1, lbl, "ct sd", vbTextCompare) > 0 Then colCtSd = c
SigCol:
        Next c
        If okSample And okTarget And okCt Then
            BuscarFilaCabecera = r
            Exit Function
        End If
    Next r
End Function

Private Function UltimaFilaDatos(ws As Worksheet, headerRow As Long, colSample As Long) As Long
    Dim lr As Long
    Dim ur As Long
    lr = ws.Cells(ws.Rows.Count, colSample).End(xlUp).Row
    ur = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If ur > lr Then lr = ur
    If lr < headerRow + 1 Then lr = headerRow + 1
    UltimaFilaDatos = lr
End Function

Private Function NormalizarTexto(s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW$(1090), "t")
    NormalizarTexto = t
End Function

Private Function EsMuestraValida(s As String) As Boolean
    Dim p As Long
    Dim ch As String
    If Len(s) < 2 Then Exit Function
    ch = UCase$(Left$(s, 1))
    If ch <> "C" And ch <> "S" Then Exit Function
    For p = 2 To Len(s)
        If Mid$(s, p, 1) < "0" Or Mid$(s, p, 1) > "9" Then Exit Function
    Next p
    EsMuestraValida = True
End Function

Private Function MuestraEnCola(col As Collection, sample As String) As Boolean
    Dim i As Long
    MuestraEnCola = False
    For i = 1 To col.Count
        If col(i) = sample Then MuestraEnCola = True: Exit Function
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
    v = ws.Cells(rowNum, colCt).Value2
    If IsNumeric(v) Then
        lec("nCt") = lec("nCt") + 1
        If lec("nCt") = 1 Then lec("ct1") = CDbl(v)
        If lec("nCt") = 2 Then lec("ct2") = CDbl(v)
    End If
    If colCtMean > 0 Then
        v = ws.Cells(rowNum, colCtMean).Value2
        If IsNumeric(v) And IsEmpty(lec("ctMean")) Then lec("ctMean") = CDbl(v)
    End If
    If colCtSd > 0 Then
        v = ws.Cells(rowNum, colCtSd).Value2
        If IsNumeric(v) And IsEmpty(lec("ctSd")) Then lec("ctSd") = CDbl(v)
    End If
End Sub

Private Function LeeMean(lec As Object) As Double
    If Not IsEmpty(lec("ctMean")) Then
        LeeMean = CDbl(lec("ctMean"))
    ElseIf lec("nCt") > 0 Then
        LeeMean = (NzDbl(lec("ct1")) + NzDbl(lec("ct2"))) / CDbl(lec("nCt"))
    Else
        LeeMean = 0#
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsNumeric(v) Then NzDbl = CDbl(v) Else NzDbl = 0#
End Function

Private Function GetLectura(data As Object, sample As String, tgt As String) As Object
    Dim k As String
    k = sample & "|" & tgt
    If data.Exists(k) Then Set GetLectura = data(k) Else Set GetLectura = Nothing
End Function

Private Function DetectarGenInteres(data As Object) As String
    Dim k As Variant
    Dim partes() As String
    Dim t As String
    Dim unico As String
    Dim n As Long

    unico = ""
    n = 0
    For Each k In data.Keys
        partes = Split(CStr(k), "|")
        If UBound(partes) >= 1 Then
            t = partes(1)
            If t <> HK_PPIA And t <> HK_SYP Then
                unico = t
                n = n + 1
                If n > 1 Then
                    DetectarGenInteres = ""
                    Exit Function
                End If
            End If
        End If
    Next k
    DetectarGenInteres = unico
End Function

Private Function ExisteTarget(data As Object, hk As String) As Boolean
    Dim k As Variant
    ExisteTarget = False
    For Each k In data.Keys
        If Right$(CStr(k), Len(hk) + 1) = "|" & hk Then ExisteTarget = True: Exit Function
    Next k
End Function

Private Sub EscribirResultados(ws As Worksheet, wsG As Worksheet, data As Object, _
    sampleOrder As Collection, goi As String)

    Dim i As Long, rowOut As Long
    Dim sample As String
    Dim lecGOI As Object, lecPPI As Object, lecSYP As Object
    Dim dctPPI As Object, dctSYP As Object
    Dim avgPPI As Double, avgSYP As Double
    Dim dCt As Double, ddCt As Double, fc As Double
    Dim goiSd As Double
    Dim flagRed As Boolean
    Dim cSamples As Collection
    Dim ctrlSuj As Collection, ctrlPPI As Collection, ctrlSYP As Collection
    Dim suiSuj As Collection, suiPPI As Collection, suiSYP As Collection
    Dim goiShort As String

    Set dctPPI = CreateObject("Scripting.Dictionary")
    Set dctSYP = CreateObject("Scripting.Dictionary")
    Set cSamples = New Collection
    Set ctrlSuj = New Collection
    Set ctrlPPI = New Collection
    Set ctrlSYP = New Collection
    Set suiSuj = New Collection
    Set suiPPI = New Collection
    Set suiSYP = New Collection

    Call PonerCabeceras(ws)

    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        Set lecGOI = GetLectura(data, sample, goi)
        Set lecPPI = GetLectura(data, sample, HK_PPIA)
        Set lecSYP = GetLectura(data, sample, HK_SYP)
        If lecGOI Is Nothing Or lecPPI Is Nothing Or lecSYP Is Nothing Then GoTo SigS
        dctPPI.Add sample, LeeMean(lecGOI) - LeeMean(lecPPI)
        dctSYP.Add sample, LeeMean(lecGOI) - LeeMean(lecSYP)
        If Left$(sample, 1) = "C" Then cSamples.Add sample
SigS:
    Next i

    If cSamples.Count = 0 Then Err.Raise vbObjectError + 5, , "No hay muestras C para el promedio."

    avgPPI = PromedioMuestras(dctPPI, cSamples)
    avgSYP = PromedioMuestras(dctSYP, cSamples)
    goiShort = NombreCortoGOI(goi)
    rowOut = 2

    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        If Not dctPPI.Exists(sample) Then GoTo SigO
        Set lecGOI = GetLectura(data, sample, goi)
        goiSd = NzDbl(lecGOI("ctSd"))
        flagRed = (goiSd > SD_UMBRAL)

        dCt = dctPPI(sample)
        ddCt = dCt - avgPPI
        fc = 2 ^ (-ddCt)
        Call EscribirFila(ws, rowOut, 1, sample, goi, lecGOI, dCt, avgPPI, ddCt, fc, flagRed, True)

        dCt = dctSYP(sample)
        ddCt = dCt - avgSYP
        fc = 2 ^ (-ddCt)
        Call EscribirFila(ws, rowOut, 17, sample, goi, lecGOI, dCt, avgSYP, ddCt, fc, flagRed, True)

        If Left$(sample, 1) = "C" Then
            ctrlSuj.Add sample
            ctrlPPI.Add 2 ^ (-(dctPPI(sample) - avgPPI))
            ctrlSYP.Add 2 ^ (-(dctSYP(sample) - avgSYP))
        ElseIf Left$(sample, 1) = "S" Then
            suiSuj.Add sample
            suiPPI.Add 2 ^ (-(dctPPI(sample) - avgPPI))
            suiSYP.Add 2 ^ (-(dctSYP(sample) - avgSYP))
        End If
        rowOut = rowOut + 2
SigO:
    Next i

    For i = 1 To sampleOrder.Count
        sample = sampleOrder(i)
        Set lecPPI = GetLectura(data, sample, HK_PPIA)
        Set lecSYP = GetLectura(data, sample, HK_SYP)
        If Not lecPPI Is Nothing Then Call EscribirFila(ws, rowOut, 1, sample, HK_PPIA, lecPPI, 0, 0, 0, 0, False, False)
        If Not lecSYP Is Nothing Then Call EscribirFila(ws, rowOut, 17, sample, HK_SYP, lecSYP, 0, 0, 0, 0, False, False)
        rowOut = rowOut + 2
    Next i

    Call EscribirGlobal(wsG, goiShort, ctrlSuj, ctrlPPI, ctrlSYP, suiSuj, suiPPI, suiSYP)
End Sub

Private Sub PonerCabeceras(ws As Worksheet)
    Dim h
    Dim cols As Variant
    cols = Array("Sample Name", "Target Name", "Ct", "Ct Mean", "Ct SD", _
        "ΔCt", "Prom. ΔCt (C)", "ΔΔCt", "2^(-ΔΔCt)")
    For h = 0 To 8
        ws.Cells(1, 1 + h).Value = cols(h)
        ws.Cells(1, 1 + h).Font.Bold = True
        ws.Cells(1, 17 + h).Value = cols(h)
        ws.Cells(1, 17 + h).Font.Bold = True
    Next h
End Sub

Private Function PromedioMuestras(d As Object, samples As Collection) As Double
    Dim i As Long
    Dim s As String
    Dim suma As Double
    suma = 0#
    For i = 1 To samples.Count
        s = samples(i)
        suma = suma + d(s)
    Next i
    PromedioMuestras = suma / CDbl(samples.Count)
End Function

Private Function NombreCortoGOI(goi As String) As String
    Dim s As String
    s = goi
    If Len(s) > 0 Then
        If Right$(s, 1) = "r" Or Right$(s, 1) = "R" Then s = Left$(s, Len(s) - 1)
    End If
    NombreCortoGOI = s
End Function

Private Sub EscribirFila(ws As Worksheet, rowOut As Long, startCol As Long, _
    sample As String, tgt As String, lec As Object, _
    dCt As Double, avgC As Double, ddCt As Double, fc As Double, _
    flagRed As Boolean, conCalcs As Boolean)

    ws.Cells(rowOut, startCol).Value = sample
    ws.Cells(rowOut, startCol + 1).Value = tgt
    If Not IsEmpty(lec("ct1")) Then ws.Cells(rowOut, startCol + 2).Value = lec("ct1")
    ws.Cells(rowOut, startCol + 3).Value = LeeMean(lec)
    ws.Cells(rowOut, startCol + 4).Value = NzDbl(lec("ctSd"))
    If Not IsEmpty(lec("ct2")) Then ws.Cells(rowOut + 1, startCol + 2).Value = lec("ct2")

    If flagRed Then ws.Cells(rowOut, startCol).Font.Color = RGB(255, 0, 0)

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

Private Sub EscribirGlobal(ws As Worksheet, goiShort As String, _
    ctrlSuj As Collection, ctrlPPI As Collection, ctrlSYP As Collection, _
    suiSuj As Collection, suiPPI As Collection, suiSYP As Collection)

    Call TablaGlobal(ws, 1, "CONTROLES " & goiShort & " PFC", RGB(255, 255, 0), ctrlSuj, ctrlPPI, ctrlSYP)
    Call TablaGlobal(ws, 6, "SUICIDAS " & goiShort & " PFC", RGB(146, 208, 80), suiSuj, suiPPI, suiSYP)
End Sub

Private Sub TablaGlobal(ws As Worksheet, startCol As Long, titulo As String, colorFondo As Long, _
    sujetos As Collection, valsPPI As Collection, valsSYP As Collection)

    Dim i As Long, r As Long
    Dim ppi As Double, syp As Double

    ws.Range(ws.Cells(1, startCol), ws.Cells(1, startCol + 3)).Merge
    ws.Cells(1, startCol).Value = titulo
    ws.Cells(1, startCol).Font.Bold = True
    ws.Cells(1, startCol).Interior.Color = colorFondo
    ws.Cells(2, startCol).Value = "SUJETO"
    ws.Cells(2, startCol + 1).Value = "PPIA"
    ws.Cells(2, startCol + 2).Value = "SYP"
    ws.Cells(2, startCol + 3).Value = "MEDIA"

    For i = 1 To sujetos.Count
        r = i + 2
        ppi = valsPPI(i)
        syp = valsSYP(i)
        ws.Cells(r, startCol).Value = sujetos(i)
        ws.Cells(r, startCol + 1).Value = ppi
        ws.Cells(r, startCol + 2).Value = syp
        ws.Cells(r, startCol + 3).Value = (ppi + syp) / 2#
    Next i
End Sub
