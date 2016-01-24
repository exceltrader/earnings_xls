Attribute VB_Name = "earningsmodule"
Option Explicit
Function yahoo()
    '''''
    Dim search(0 To 2) As String
    search(0) = "Earnings Date:"
    search(1) = """>"
    search(2) = "test2"
    Dim after_target As String
    after_target = "x"
    ''''''
    Dim yahoo_ As site_class
    Set yahoo_ = New site_class
    With yahoo_
        .pre_tar_arr = search
        .aft_target = "x"
    End With
    Set yahoo = yahoo_
End Function
Function yahoo_url(ByVal symbol As String)
    Dim url As String
    url = "http://finance.yahoo.com/q?s=" & LCase(symbol)
    yahoo_url = url
End Function
Function zacks()
    Dim zacks_ As site_class
    Set zacks_ = New site_class
    With zacks_
        .pre_target_arr(0) = "XXXXXXEarnings Date:"
        .pre_target_arr(1) = """>"
        .aft_target = "<span"
    End With
    zacks = zacks_
End Function
Function zacks_url(ByVal symbol As String)
    Dim url As String
    url = "http://www.zacks.com/stock/quote/" & UCase(symbol)
    zacks_url = url
End Function
Sub symbols()
    Dim yahoo_site As site_class
    Dim zacks_site As site_class
    Dim sym As Range
    Dim lr As Long
    Dim ws As Worksheet
    'Dim url1() As String 'yahoo url
    'Dim url2() As String 'zacks url
    Const rowstart = 2 'row of first symbol
    Dim objWinHttp As Object
    Dim url As String
    'Dim res As String
    Dim eclass As parse_class
    Set eclass = New parse_class
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set ws = ThisWorkbook.Worksheets("E")
    With ws
        lr = .Cells(Rows.Count, "A").End(xlUp).Row
        If lr < rowstart Then Exit Sub
        ReDim url1(0 To lr - rowstart)
        ReDim url2(0 To lr - rowstart)
        For Each sym In ws.Range("A2:A" & lr)
        
        



            eclass.url = yahoo_url(sym)
   
            'Set yahoo_site = yahoo
            eclass.site = yahoo
                
                
            
            'eclass.sitetype = yahoo
            ws.Range("B" & sym.Row).Value = eclass.getDate
            

            
            
            
            
        Next sym
    End With
End Sub







