    Dim myFile as String

    myFile = "C:\Users\Arquivo.txt"

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & myFile, Destination:=Range("$C$1"))
        .Name = "TextoImportado "
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = True
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    For i = ActiveWorkbook.Connections.Count To 1 Step -1
    ActiveWorkbook.Connections(i).Delete
    Next
