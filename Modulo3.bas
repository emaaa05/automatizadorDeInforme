Attribute VB_Name = "Módulo3"
Sub BorrarImagenes()
    Dim wsBase As Worksheet
    On Error Resume Next
    Set wsBase = ThisWorkbook.Sheets("ImagenesCargadas")
    On Error GoTo 0

    If wsBase Is Nothing Then
        MsgBox "No existe la hoja 'ImagenesCargadas'."
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsBase.Cells(wsBase.rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No hay datos para borrar."
        Exit Sub
    End If
    
    wsBase.Range("A2:B" & lastRow).ClearContents

    MsgBox "Se borraron " & (lastRow - 1) & " rutas de imágenes de la hoja 'ImagenesCargadas'."
End Sub

