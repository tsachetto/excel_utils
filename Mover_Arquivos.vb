Sub Mover()

    Dim FSO As Object
    Dim Origem As String, Destino As String

    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    Origem = "C:\Users\Jun.xlsx"
    Destino = "C:\Users\Desktop\Jun.xlsx"

    FSO.MoveFile Source:=Origem, Destination:=Destino

End Sub
