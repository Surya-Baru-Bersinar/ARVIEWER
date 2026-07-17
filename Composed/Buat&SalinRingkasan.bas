Sub BuatDanSalinRingkasan()
    Dim ws As Worksheet
    Dim txt As String
    Dim r As Long, lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Composed")
    
    Dim flag_fraud As Boolean:      flag_fraud = GetFlagStatus(ws, "ar_data_fraud")
    Dim flag_codecus As Boolean:    flag_codecus = GetFlagStatus(ws, "ar_data_codecus")
    Dim flag_namecus As Boolean:    flag_namecus = GetFlagStatus(ws, "ar_data_namecus")
    Dim flag_calc As Boolean:       flag_calc = GetFlagStatus(ws, "ar_data_calc")
    Dim flag_avg_age As Boolean:    flag_avg_age = GetFlagStatus(ws, "ar_data_avg_age")
    Dim flag_avg_val As Boolean:    flag_avg_val = GetFlagStatus(ws, "ar_data_avg_val")
    Dim flag_avg_inv As Boolean:    flag_avg_inv = GetFlagStatus(ws, "ar_data_avg_inv")
    Dim flag_avg_plaf As Boolean:   flag_avg_plaf = GetFlagStatus(ws, "ar_data_avg_plaf")
    Dim flag_avg_pay As Boolean:    flag_avg_pay = GetFlagStatus(ws, "ar_data_avg_pay")
    Dim flag_avg_his As Boolean:    flag_avg_his = GetFlagStatus(ws, "ar_data_avg_his")
    Dim flag_avg_tier As Boolean:   flag_avg_tier = GetFlagStatus(ws, "ar_data_avg_tier")
    
    Dim flag_inv_numb As Boolean:   flag_inv_numb = GetFlagStatus(ws, "ar_data_inv_numb")
    Dim flag_inv_dt As Boolean:     flag_inv_dt = GetFlagStatus(ws, "ar_data_inv_dt")
    Dim flag_inv_due As Boolean:    flag_inv_due = GetFlagStatus(ws, "ar_data_inv_due")
    Dim flag_inv_val As Boolean:    flag_inv_val = GetFlagStatus(ws, "ar_data_inv_val")
    Dim flag_inv_ar As Boolean:     flag_inv_ar = GetFlagStatus(ws, "ar_data_inv_ar")
    Dim flag_inv_orig As Boolean:   flag_inv_orig = GetFlagStatus(ws, "ar_data_inv_orig")
    Dim flag_inv_pay As Boolean:    flag_inv_pay = GetFlagStatus(ws, "ar_data_inv_pay")
    Dim flag_owing As Boolean:      flag_owing = GetFlagStatus(ws, "ar_data_owing")
    Dim flag_giro As Boolean:       flag_giro = GetFlagStatus(ws, "ar_data_giro")
    Dim flag_age As Boolean:        flag_age = GetFlagStatus(ws, "ar_data_age")
    
    Dim headerText As String
    headerText = ""
    If flag_codecus Then headerText = ws.Range("B24").Value
    If flag_namecus Then
        If headerText <> "" Then headerText = headerText & vbTab
        headerText = headerText & ws.Range("K26").Value
    End If
    
    txt = headerText & vbCrLf
    txt = txt & "Tanggal Order: " & Format(Date, "dd/mm/yyyy") & vbCrLf & vbCrLf
    
    txt = txt & "========================================" & vbCrLf
    txt = txt & "RINGKASAN PERFORMA PIUTANG" & vbCrLf
    txt = txt & "========================================" & vbCrLf
    
    If flag_calc Then txt = txt & "Piutang" & vbTab & vbTab & vbTab & vbTab & vbTab & " :  " & FormatIDR(ws.Range("D3").Value) & " " & vbCrLf
    If flag_inv_val Then txt = txt & "Total Faktur Aktif (Inv) :  " & GetSafeText(ws.Range("D4")) & " " & vbCrLf
    
    If flag_avg_tier Then txt = txt & "Tiering" & vbTab & vbTab & vbTab & vbTab & vbTab & " : " & GetSafeText(ws.Range("D11")) & vbCrLf
    If flag_avg_age Then txt = txt & "Avg Umur Piutang" & vbTab & vbTab & " : " & GetSafeText(ws.Range("D8")) & vbCrLf
    If flag_avg_val Then txt = txt & "Avg Nilai Faktur" & vbTab & vbTab & " : " & FormatIDR(ws.Range("D9").Value) & vbCrLf
    If flag_avg_inv Then txt = txt & "Avg Jumlah Faktur" & vbTab & vbTab & " : " & GetSafeText(ws.Range("D10")) & vbCrLf
    
    txt = txt & vbCrLf
    If flag_avg_plaf Then txt = txt & "Plafon Kredibilitas" & vbTab & vbTab & " :  " & FormatIDR(ws.Range("D5").Value) & " " & vbCrLf
    If flag_avg_pay Then txt = txt & "Rata-Rata Bayar" & vbTab & vbTab & vbTab & " :  " & FormatIDR(ws.Range("D6").Value) & " " & vbCrLf
    If flag_avg_his Then
        Dim hisVal As String
        hisVal = GetSafeText(ws.Range("D7"))
        If hisVal <> "-" And hisVal <> "" Then hisVal = hisVal & " Hari"
        txt = txt & "Rata-Rata History Bayar" & vbTab & " : " & hisVal & vbCrLf
    End If
    
    txt = txt & vbCrLf & "========================================" & vbCrLf
    txt = txt & "DAFTAR RINCIAN FAKTUR AKTIF" & vbCrLf
    txt = txt & "========================================" & vbCrLf
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For r = 26 To lastRow
        If Not IsError(ws.Cells(r, "C").Value) Then
            If Trim(ws.Cells(r, "C").Text) <> "" Then
                Dim lineParts As String
                lineParts = ""
                
                If flag_inv_numb Then lineParts = AppendTextPart(lineParts, ws.Cells(r, "C").Text)
                
                If flag_inv_dt Then
                    If Not IsError(ws.Cells(r, "D").Value) And IsDate(ws.Cells(r, "D").Value) Then
                        lineParts = AppendTextPart(lineParts, Format(ws.Cells(r, "D").Value, "dd mmm yyyy"))
                    Else
                        lineParts = AppendTextPart(lineParts, "-")
                    End If
                End If
                
                If flag_inv_due Then
                    If Not IsError(ws.Cells(r, "F").Value) And IsDate(ws.Cells(r, "F").Value) Then
                        lineParts = AppendTextPart(lineParts, Format(ws.Cells(r, "F").Value, "dd mmm yyyy"))
                    Else
                        lineParts = AppendTextPart(lineParts, "-")
                    End If
                End If
                
                If flag_inv_orig Then lineParts = AppendTextPart(lineParts, FormatIDR(ws.Cells(r, "H").Value))
                If flag_inv_ar Then lineParts = AppendTextPart(lineParts, FormatIDR(ws.Cells(r, "I").Value))
                
                If flag_inv_pay Then
                    If Not IsError(ws.Cells(r, "H").Value) And Not IsError(ws.Cells(r, "I").Value) Then
                        Dim origVal As Double, arVal As Double
                        origVal = Val(ws.Cells(r, "H").Value)
                        arVal = Val(ws.Cells(r, "I").Value)
                        If (origVal - arVal) > 0 Then
                            lineParts = AppendTextPart(lineParts, "Ttp Byr: " & FormatIDR(origVal - arVal))
                        End If
                    End If
                End If
                
                If flag_age Then
                    If Not IsError(ws.Cells(r, "D").Value) And IsDate(ws.Cells(r, "D").Value) Then
                        Dim ageDays As Long
                        ageDays = DateDiff("d", ws.Cells(r, "D").Value, Date)
                        lineParts = AppendTextPart(lineParts, ageDays & vbTab & "HR")
                    Else
                        lineParts = AppendTextPart(lineParts, "-" & vbTab & "HR")
                    End If
                End If
                
                If flag_owing Then
                    Dim wsOwing As Worksheet
                    Dim checkOwing As Variant
                    On Error Resume Next
                    Set wsOwing = ThisWorkbook.Sheets("OWING")
                    On Error GoTo 0
                    
                    If Not wsOwing Is Nothing And Not IsError(ws.Cells(r, "C").Value) Then
                        checkOwing = Application.Match(ws.Cells(r, "C").Value, wsOwing.Columns("D"), 0)
                        If Not IsError(checkOwing) Then lineParts = lineParts & " (OWING)"
                    End If
                End If
                
                If flag_giro Then
                    If Not IsError(ws.Cells(r, "N").Value) Then
                        If Trim(ws.Cells(r, "N").Text) <> "" And Val(ws.Cells(r, "N").Value) <> 0 Then
                            lineParts = lineParts & " (" & Format(ws.Cells(r, "N").Value, "dd mmm yyyy") & ")"
                        End If
                    End If
                End If
                
                If flag_fraud Then
                    If Not IsError(ws.Cells(r, "L").Value) Then
                        If InStr(1, UCase(ws.Cells(r, "L").Text), "FRAUD") > 0 Then
                            lineParts = lineParts & " (FRAUD)"
                        End If
                    End If
                End If
                
                txt = txt & lineParts & vbCrLf
            End If
        End If
    Next r
    
    Dim DataObj As Object
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText txt
    DataObj.PutInClipboard
    
    MsgBox "Ringkasan Performa berhasil dibuat & disalin ke Clipboard!", vbInformation, "Sukses"
End Sub

Private Function GetFlagStatus(ws As Worksheet, key As String) As Boolean
    Dim rngMatch As Variant
    rngMatch = Application.Match(key, ws.Columns("K"), 0)
    If Not IsError(rngMatch) Then
        GetFlagStatus = (UCase(Trim(ws.Cells(rngMatch, "L").Value)) = "YA")
    Else
        GetFlagStatus = False
    End If
End Function

Private Function AppendTextPart(mainStr As String, newPart As String) As String
    If mainStr = "" Then
        AppendTextPart = newPart
    Else
        AppendTextPart = mainStr & vbTab & newPart
    End If
End Function

Private Function GetSafeText(cell As Range) As String
    If IsError(cell.Value) Then
        GetSafeText = "-"
    Else
        Dim valStr As String
        valStr = Trim(cell.Text)
        If valStr = "" Then GetSafeText = "-" Else GetSafeText = valStr
    End If
End Function

Private Function FormatIDR(val As Variant) As String
    If IsError(val) Then
        FormatIDR = "-"
        Exit Function
    End If
    
    If IsNumeric(val) And Not IsEmpty(val) And val <> 0 Then
        Dim temp As String
        temp = Format(val, "#,##0")
        FormatIDR = Replace(temp, ",", ".")
    ElseIf val = 0 Then
        FormatIDR = "0"
    Else
        FormatIDR = "-"
    End If
End Function
