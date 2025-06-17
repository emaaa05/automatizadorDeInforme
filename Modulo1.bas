Attribute VB_Name = "Módulo1"
Sub AgregarImagenABase()
    Dim fd As FileDialog
    Dim filePath As Variant
    Dim wsBase As Worksheet
    Dim i As Long

    On Error Resume Next
    Set wsBase = ThisWorkbook.Sheets("ImagenesCargadas")
    On Error GoTo 0
    If wsBase Is Nothing Then
        Set wsBase = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsBase.Name = "ImagenesCargadas"
        wsBase.Range("A1:B1").Value = Array("N°", "RutaImagen")
    End If

    i = wsBase.Cells(wsBase.rows.Count, "A").End(xlUp).Row + 1

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccioná una o más imágenes"
        .Filters.Clear
        .Filters.Add "JPG", "*.jpg"
        .Filters.Add "JPEG", "*.jpeg"
        .Filters.Add "PNG", "*.png"
        .Filters.Add "BMP", "*.bmp"
        .Filters.Add "GIF", "*.gif"
        .Filters.Add "Todos los archivos", "*.*"
        .AllowMultiSelect = True
        If .Show <> -1 Then
            MsgBox "No se seleccionó ninguna imagen."
            Exit Sub
        End If
    End With

    Dim j As Integer
    For j = 1 To fd.SelectedItems.Count
        wsBase.Cells(i, 1).Value = i - 1
        wsBase.Cells(i, 2).Value = fd.SelectedItems(j)
        i = i + 1
    Next j

    MsgBox "Se agregaron " & fd.SelectedItems.Count & " imágenes a la base."
End Sub
