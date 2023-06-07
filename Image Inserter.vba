'Image Insert VBA Code

Sub InsertPicture()
    Dim i As Integer

    On Error Resume Next ' Skip errors

    For i = 4 To 2398
        If Not IsEmpty(Cells(i, 3).Value) Then
            Cells(i, 4).Select
            ActiveSheet.Shapes.AddPicture _
                Filename:=Cells(i, 3).Value, _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=ActiveCell.Left, _
                Top:=ActiveCell.Top, _
                Width:=100, Height:=100
        End If
    Next i

    On Error GoTo 0 ' Disable error handling
End Sub


