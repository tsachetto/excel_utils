Dim fldr As FileDialog
Dim diretorio as String

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

    With fldr
        .Title = "Escolha um diretório:"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:

    If sItem = "" Then
        Exit Sub
    Else: End If
    
    diretorio = sItem & "\"
    Set fldr = Nothing
