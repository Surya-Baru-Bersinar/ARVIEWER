Sub UpdateDataAnalisisC()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim searchCrit As String, formulaStr As String
    Dim lastRow As Long
    
    On Error GoTo ErrHndl
    
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsTarget = ActiveSheet
    
    searchCrit = wsTarget.Range("B24").Value

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    If lastRow < 25 Then lastRow = 25
    
    wsTarget.Range("B26:N" & lastRow + 100).ClearContents

    If searchCrit = "" Then GoTo ExitSub

    formulaStr = "IFERROR(SORT(FILTER(Source!A:M, ISNUMBER(SEARCH(""" & searchCrit & """, Source!A:A)), ""Data Tidak Ditemukan""), {10,2,9}, 1), ""Data Tidak Ditemukan"")"

    With wsTarget.Range("B26")
        .Formula2 = "=" & formulaStr
        
        DoEvents
        
        On Error Resume Next
        
        Dim resultRange As Range
        Set resultRange = .SpillArea
        
        If resultRange Is Nothing Then Set resultRange = .CurrentRegion
        On Error GoTo ErrHndl
        
        With resultRange
            .Value = .Value
            .EntireColumn.AutoFit
        End With
    End With

ExitSub:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHndl:
    MsgBox "Terjadi kesalahan: " & Err.Description, vbCritical
    Resume ExitSub
End Sub
