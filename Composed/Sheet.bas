Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B24")) Is Nothing Then
        On Error GoTo SafeExit
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        Call UpdateDataAnalisisC
        
        Application.Calculate
        
        Call ProsesKolomHijau
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
    Exit Sub

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub ProsesKolomHijau()
    Dim wsComposed As Worksheet
    Dim wsAvg As Worksheet
    Set wsComposed = Me
    
    If Trim(wsComposed.Range("B26").Value) = "" Then
        wsComposed.Range("D3:D11").Value = "-"
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsAvg = ThisWorkbook.Sheets("AVG")
    On Error GoTo 0
    
    If wsAvg Is Nothing Then
        MsgBox "Sheet dengan nama 'AVG' tidak ditemukan! Periksa kembali nama sheet Anda.", vbCritical, "Error Sheet"
        Exit Sub
    End If
    
    Dim custID As Variant
    custID = wsComposed.Range("B24").Value
    
    Dim colNoPelanggan As Long
    Dim rowCust As Variant
    colNoPelanggan = FindColumnHeader(wsAvg, "NO. PELANGGAN")
    
    If colNoPelanggan > 0 Then
        rowCust = Application.Match(custID, wsAvg.Columns(colNoPelanggan), 0)
    Else
        rowCust = CVErr(xlErrNA)
    End If
    
    Dim r As Long
    Dim keyName As String
    Dim targetHeader As String
    Dim flagRow As Variant
    Dim flagVal As String
    Dim finalResult As Variant
    Dim colIndexAvg As Long
    Dim lastRowData As Long
    
    For r = 3 To 11
        targetHeader = ""
        
        Select Case r
            Case 3: keyName = "ar_data_calc":     targetHeader = "SUM_SISA_PIUTANG"
            Case 4: keyName = "ar_data_inv_val":  targetHeader = "SUM_NILAI_FAKTUR"
            Case 5: keyName = "ar_data_avg_plaf": targetHeader = "PLAFON"
            Case 6: keyName = "ar_data_avg_pay":  targetHeader = "AVG BAYAR"
            Case 7: keyName = "ar_data_avg_his":  targetHeader = "AVG HISTORY BAYAR (HARI)"
            Case 8: keyName = "ar_data_avg_age":  targetHeader = "AVG UMUR PIUTANG"
            Case 9: keyName = "ar_data_avg_val":  targetHeader = "AVG NILAI FAKTUR"
            Case 10: keyName = "ar_data_avg_inv": targetHeader = "JUMLAH INVOICE"
            Case 11: keyName = "ar_data_avg_tier": targetHeader = "TIERING"
        End Select
        
        flagRow = Application.Match(keyName, wsComposed.Columns("K"), 0)
        flagVal = ""
        
        If Not IsError(flagRow) Then
            flagVal = UCase(Trim(wsComposed.Cells(flagRow, "L").Value))
        End If
        
        If flagVal = "YA" Then
            finalResult = ""
            
            If targetHeader = "SUM_SISA_PIUTANG" Then
                lastRowData = wsComposed.Cells(wsComposed.Rows.Count, "I").End(xlUp).Row
                If lastRowData >= 26 Then
                    finalResult = Application.WorksheetFunction.Sum(wsComposed.Range("I26:I" & lastRowData))
                Else
                    finalResult = 0
                End If
                
            ElseIf targetHeader = "SUM_NILAI_FAKTUR" Then
                lastRowData = wsComposed.Cells(wsComposed.Rows.Count, "G").End(xlUp).Row
                If lastRowData >= 26 Then
                    finalResult = Application.WorksheetFunction.Sum(wsComposed.Range("G26:G" & lastRowData))
                Else
                    finalResult = 0
                End If
                
            Else
                If Not IsError(rowCust) Then
                    colIndexAvg = FindColumnHeader(wsAvg, targetHeader)
                    If colIndexAvg > 0 Then
                        finalResult = wsAvg.Cells(rowCust, colIndexAvg).Value
                    Else
                        finalResult = "Kolom tidak ditemukan"
                    End If
                Else
                    finalResult = "Pelanggan tidak ditemukan di AVG"
                End If
            End If
            
            wsComposed.Cells(r, "D").Value = finalResult
        Else
            wsComposed.Cells(r, "D").Value = "-"
        End If
    Next r
End Sub

Private Function FindColumnHeader(ws As Worksheet, headerName As String) As Long
    Dim cell As Range
    Set cell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not cell Is Nothing Then
        FindColumnHeader = cell.Column
    Else
        FindColumnHeader = 0
    End If
End Function
