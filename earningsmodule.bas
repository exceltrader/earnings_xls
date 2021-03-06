Attribute VB_Name = "earningsmodule"
Option Explicit
Function yahoo()
    Dim yahoo_ As site_class
    Set yahoo_ = New site_class
    With yahoo_
        'a unique html string, near the date we want
        .pre_tar_arr(0) = "Earnings Date:"
        'another unique html string between .pre_tar_arr(0) and the date we want
        .pre_tar_arr(1) = "yfnc_tabledata1"">"
        'add more as needed until the value is the string before date
        '.pre_tar_arr(2) = "any string before the date"
        .aft_target = "<" 'The first character after the date we want
    End With
    Set yahoo = yahoo_
End Function
Function zacks()
    Dim zacks_ As site_class
    Set zacks_ = New site_class
    With zacks_
        .pre_tar_arr(0) = "Earnings Date"
        .pre_tar_arr(1) = "</sup>"
        .aft_target = "</td"
    End With
    Set zacks = zacks_
End Function
Function earnwhisp()
    Dim ew_ As site_class
    Set ew_ = New site_class
    With ew_
        .pre_tar_arr(0) = "Earnings Release Date"",""startDate"" : """
        '.pre_tar_arr(1) = "</sup>"
        .aft_target = "T2"
    End With
    Set earnwhisp = ew_
End Function
Function yahoo_url(ByVal symbol As String)
    Dim url As String
    url = "http://finance.yahoo.com/q?s=" & LCase(symbol)
    yahoo_url = url
End Function
Function zacks_url(ByVal symbol As String)
    Dim url As String
    url = "http://www.zacks.com/stock/quote/" & UCase(symbol)
    zacks_url = url
End Function
Function ew_url(ByVal symbol As String)
    Dim url As String
    url = "https://www.earningswhispers.com/stocks/" & LCase(symbol)
    ew_url = url
End Function
Sub symbols()
    Dim yahoo_site As site_class
    Dim zacks_site As site_class
    Dim sym As Range
    Dim lr As Long
    Dim ws As Worksheet
    Const rowstart = 2 'row of first symbol
    Dim objWinHttp As Object
    'Dim url As String
    Dim eclass As parse_class
    Set eclass = New parse_class
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set ws = ThisWorkbook.Worksheets("E")
    With ws
        lr = .Cells(Rows.Count, "A").End(xlUp).Row
        If lr < rowstart Then Exit Sub
        .Range("B2:D" & lr).ClearContents
        .Range("B1").Value = "Yahoo"
        .Range("C1").Value = "Zacks"
        .Range("D1").Value = "Earnings Whisper"
        For Each sym In ws.Range("A2:A" & lr)
            '''''yahoo
            eclass.url = yahoo_url(sym)
            eclass.site = yahoo
            ws.Range("B" & sym.Row).Value = eclass.getDate 'column B
            '''''zacks
            eclass.url = zacks_url(sym)
            eclass.site = zacks
            Err.Clear
            On Error Resume Next
            ws.Range("C" & sym.Row).Value = eclass.getDate 'column C
            If Err.Number > 0 Then
                ws.Range("C" & sym.Row).Value = "error"
            End If
            '''''
            '''''ew
            eclass.url = ew_url(sym)
            eclass.site = earnwhisp
            Err.Clear
            On Error Resume Next
            ws.Range("D" & sym.Row).Value = eclass.getDate
                If Err.Number > 0 Then
                ws.Range("D" & sym.Row).Value = "error"
            End If
            ''''''
            disable
            Application.Wait (Now + TimeValue("00:00:01"))
            enable
        Next sym
    End With
End Sub
Sub disable()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub
Sub enable()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

