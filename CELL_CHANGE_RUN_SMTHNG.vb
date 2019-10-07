Private Sub Worksheet_Change(ByVal Target As Range)

If Not Intersect(Target, Range("A1:B100")) Is Nothing Then
Call Mymacro
End If

End Sub
