Sub UpdateDataAnalisisC()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim searchCrit As String, formulaStr As String
    Dim lastRow As Long
    Dim flagRow As Variant
    Dim isFraudActive As Boolean
    
    On Error GoTo ErrHndl
    
    Set wsSource = ThisWorkbook.Sheets("Source")
    Set wsTarget = ActiveSheet
    
    searchCrit = wsTarget.Range("B24").Value
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    isFraudActive = False
    flagRow = Application.Match("ar_data_fraud", wsTarget.Columns("K"), 0)
    
    If Not IsError(flagRow) Then
        If UCase(Trim(wsTarget.Cells(flagRow, "L").Value)) = "YA" Then
            isFraudActive = True
        End If
    End If
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    If lastRow < 25 Then lastRow = 25
    
    wsTarget.Range("B26:N" & lastRow + 100).ClearContents
    
    If searchCrit = "" Then GoTo ExitSub
    
    If isFraudActive Then
        formulaStr = "IFERROR(SORT(FILTER(Source!A:M, ISNUMBER(SEARCH(""" & searchCrit & """, Source!A:A)), ""Data Tidak Ditemukan""), {10,2,9}, 1), ""Data Tidak Ditemukan"")"
    Else
        formulaStr = "IFERROR(SORT(FILTER(Source!A:M, (ISNUMBER(SEARCH(""" & searchCrit & """, Source!A:A))) * (ISERROR(SEARCH(""FRAUD"", Source!K:K))), ""Data Tidak Ditemukan""), {10,2,9}, 1), ""Data Tidak Ditemukan"")"
    End If
    
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
