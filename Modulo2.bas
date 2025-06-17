Attribute VB_Name = "Módulo2"
Sub GenerarInformeDesdeBase()
    Dim wsBase As Worksheet
    Dim wsDatos As Worksheet
    Dim wbNuevo As Workbook
    Dim fila As Long
    Dim rutaImagen As String
    Dim hojaH As Worksheet
    Dim hojaNueva As Worksheet
    Dim hojaPlantilla As Worksheet
    Dim hojaResumen As Worksheet
    Dim hojaVR As Worksheet
    Dim img As Shape
    Dim nombreHoja As String
    Dim nombreBase As String
    Dim topArea As Double, leftArea As Double, anchoArea As Double, altoArea As Double
    Dim celdaSuperiorIzq As Range, celdaInferiorDer As Range
    Dim contador As Long
    Dim dict As Object

    Set wsBase = ThisWorkbook.Sheets("ImagenesCargadas")
    Set wsDatos = ThisWorkbook.Sheets("Interfaz")
    Set dict = CreateObject("Scripting.Dictionary")

    If wsBase.Cells(wsBase.rows.Count, "A").End(xlUp).Row < 2 Then
        MsgBox "No hay imágenes en 'ImagenesCargadas'."
        Exit Sub
    End If

    Set wbNuevo = Workbooks.Add

    ThisWorkbook.Sheets("Valorización de riesgos").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
    Set hojaVR = wbNuevo.Sheets(wbNuevo.Sheets.Count)

    Dim celda As Range
    For Each celda In hojaVR.UsedRange
        If celda.HasFormula Then
            celda.Formula = Replace(celda.Formula, "'[" & ThisWorkbook.Name & "]", "'")
        End If
    Next celda

    Set hojaPlantilla = ThisWorkbook.Sheets("Referencia")
    
    ThisWorkbook.Sheets("Introducción").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
    
    Set hojaIntro = wbNuevo.Sheets(wbNuevo.Sheets.Count)


    Dim celdaIntro As Range
    For Each celdaIntro In hojaIntro.UsedRange
        If Not IsError(celdaIntro.Value) And Not IsEmpty(celdaIntro.Value) Then
            If InStr(celdaIntro.Value, "{licenciado}") > 0 Then
                celdaIntro.Value = Replace(celdaIntro.Value, "{licenciado}", wsDatos.Range("I16").Value)
            End If
            If InStr(celdaIntro.Value, "{cliente}") > 0 Then
                celdaIntro.Value = Replace(celdaIntro.Value, "{cliente}", wsDatos.Range("J16").Value)
            End If
            If InStr(celdaIntro.Value, "{localizacion, fecha}") > 0 Then
                celdaIntro.Value = Replace(celdaIntro.Value, "{localizacion, fecha}", wsDatos.Range("H16").Value)
            End If
        End If
    Next celdaIntro

    contador = 1
    fila = 16

    Do While wsBase.Cells(contador + 1, 2).Value <> ""
        rutaImagen = wsBase.Cells(contador + 1, 2).Value
        nombreBase = ObtenerNombreBase(rutaImagen)

        If Not dict.exists(nombreBase) Then
            hojaPlantilla.Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
            Set hojaH = wbNuevo.Sheets(wbNuevo.Sheets.Count)

            nombreHoja = "H (" & dict.Count + 1 & ")"
            On Error Resume Next
            hojaH.Name = nombreHoja
            On Error GoTo 0

            dict.Add nombreBase, hojaH

            hojaH.Range("B2").Value = dict.Count
            hojaH.Range("B5").Value = wsDatos.Cells(fila, 1).Value
            hojaH.Range("B6").Value = wsDatos.Cells(fila, 2).Value
            hojaH.Range("B7").Value = wsDatos.Cells(fila, 3).Value
            hojaH.Range("B10").Value = wsDatos.Cells(fila, 4).Value
            hojaH.Range("D10").Value = wsDatos.Cells(fila, 5).Value
            hojaH.Range("A16").Value = wsDatos.Cells(fila, 6).Value

            fila = fila + 1

            Dim celdaRef As Range
            For Each celdaRef In hojaH.UsedRange
                If celdaRef.HasFormula Then
                    celdaRef.Formula = Replace(celdaRef.Formula, "'[" & ThisWorkbook.Name & "]", "'")
                End If
            Next celdaRef
        Else
            Set hojaH = dict(nombreBase)
        End If

        Set img = hojaH.Shapes.AddPicture( _
            Filename:=rutaImagen, _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=0, Top:=0, Width:=-1, Height:=-1)

        Dim totalImg As Long: totalImg = 0
        Dim shpIter As Shape
        For Each shpIter In hojaH.Shapes
            If shpIter.Type = msoPicture Then totalImg = totalImg + 1
        Next shpIter

        Set celdaSuperiorIzq = hojaH.Range("A13")
        Set celdaInferiorDer = hojaH.Range("F14")
        topArea = celdaSuperiorIzq.Top
        leftArea = celdaSuperiorIzq.Left
        anchoArea = celdaInferiorDer.Left + celdaInferiorDer.Width - leftArea
        altoArea = celdaInferiorDer.Top + celdaInferiorDer.Height - topArea

        Dim cols As Long, rows As Long
        cols = Application.WorksheetFunction.RoundUp(Sqr(totalImg), 0)
        rows = Application.WorksheetFunction.RoundUp(totalImg / cols, 0)

        Dim anchoCelda As Double, altoCelda As Double
        anchoCelda = anchoArea / cols
        altoCelda = altoArea / rows

        Dim szW As Double, szH As Double
        If totalImg = 1 Then
            szW = anchoArea - 20
            szH = altoArea - 20
        Else
            szW = anchoCelda - 10
            szH = altoCelda - 10
        End If

        Dim idx As Long: idx = 0
        For Each shpIter In hojaH.Shapes
            If shpIter.Type = msoPicture Then
                idx = idx + 1
                With shpIter
                    .LockAspectRatio = msoTrue
                    If .Width >= .Height Then
                        If .Width > szW Then .Width = szW
                        If .Height > szH Then .Height = szH
                    Else
                        If .Height > szH Then .Height = szH
                        If .Width > szW Then .Width = szW
                    End If

                    Dim c As Long, r As Long
                    c = (idx - 1) Mod cols
                    r = (idx - 1) \ cols

                    .Left = Round(leftArea + c * anchoCelda + (anchoCelda - .Width) / 2, 0)
                    .Top = Round(topArea + r * altoCelda + (altoCelda - .Height) / 2, 0)
                End With
            End If
        Next shpIter

        contador = contador + 1
    Loop
    
    wbNuevo.Sheets("Valorización de riesgos").Move After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)

    ThisWorkbook.Sheets("Resumen").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)

    Application.DisplayAlerts = False
    For Each hojaNueva In wbNuevo.Sheets
        If hojaNueva.UsedRange.Address = "$A$1" And hojaNueva.Cells(1, 1).Value = "" Then
            hojaNueva.Delete
        End If
    Next hojaNueva
    Application.DisplayAlerts = True

    MsgBox "Informe generado con " & (contador - 1) & " observaciones e imágenes."

    Set hojaResumen = Nothing
    For Each hojaNueva In wbNuevo.Sheets
        If hojaNueva.Name = "Resumen" Then
            Set hojaResumen = hojaNueva
            Exit For
        End If
    Next hojaNueva

    If hojaResumen Is Nothing Then
        MsgBox "No se encontró la hoja 'Resumen'."
        Exit Sub
    End If

    Dim clave As Variant
    Dim filaResumen As Long
    filaResumen = 20

    For Each clave In dict.Keys
        Set hojaH = dict(clave)
        Dim hojaNombre As String
        hojaNombre = hojaH.Name

        With hojaResumen
            .Cells(filaResumen, 1).Formula = "='" & hojaNombre & "'!B2"
            .Cells(filaResumen, 2).Formula = "='" & hojaNombre & "'!D2"
            .Cells(filaResumen, 3).Formula = "='" & hojaNombre & "'!F2"
            .Cells(filaResumen, 4).Formula = "='" & hojaNombre & "'!B5"
            .Cells(filaResumen, 5).Formula = "='" & hojaNombre & "'!B6"
            .Cells(filaResumen, 6).Formula = "='" & hojaNombre & "'!B7"
            .Cells(filaResumen, 7).Formula = "='" & hojaNombre & "'!F10"
            .Cells(filaResumen, 8).Formula = "='" & hojaNombre & "'!A16"
        End With

        filaResumen = filaResumen + 1
    Next clave
End Sub

Function ObtenerNombreBase(rutaImagen As String) As String
    Dim fso As Object, nombreArchivo As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    nombreArchivo = fso.GetFileName(rutaImagen)
    nombreArchivo = Left(nombreArchivo, InStrRev(nombreArchivo, ".") - 1)

    If Right(nombreArchivo, 1) = ")" Then
        Dim pos As Long
        pos = InStrRev(nombreArchivo, " (")
        If pos > 0 Then
            nombreArchivo = Left(nombreArchivo, pos - 1)
        End If
    End If

    ObtenerNombreBase = nombreArchivo
End Function

