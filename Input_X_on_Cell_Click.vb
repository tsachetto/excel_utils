Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'Adicionar função em uma sheet, não em um módulo! Talkei?

    Dim rInt As Range
    Dim rCell As Range

    Set rInt = Intersect(Target, Range("F3:G312"))
    
    If Not rInt Is Nothing Then
        
        If Application.Selection.Cells.Count > 1 Then
        Exit Sub
        Else: End If
    
        For Each rCell In rInt
        
            If rCell.Value = "X" Then
            rCell.Value = ""
            ElseIf rCell.Value <> "" Then
            Exit Sub
            Else
            rCell.Value = "X"
            End If
            
        Next
    End If
    
    Set rInt = Nothing
    Set rCell = Nothing

End Sub
