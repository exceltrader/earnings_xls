VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "parse_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private P_start As Variant
Private P_end() As String
Private P_site As site_class
Public rawHtml As String
Public url As String
Private objWinHttp As Object
Public Property Let start(s As Variant)
    P_start = s
End Property
Public Property Let site(s As Variant)
    Set P_site = New site_class
    Set P_site = s
End Property
Public Function getDate()
   Dim tmp As Long
   Dim cnt As Long
   Dim x() As String
   Dim datedebug As String
   tmp = Len(rawHtml)
   Dim Srchtrm As String
   Dim a As Long, b As Long
   Dim s_array() As String
   Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    objWinHttp.Open "GET", url
    objWinHttp.send
    objWinHttp.WaitForResponse
    rawHtml = objWinHttp.ResponseText
    a = 1 'start first char
    For cnt = 0 To P_site.Pre_tar_arr_COUNT
         a = InStr(a, rawHtml, P_site.pre_tar_arr(cnt))
         a = a + Len(P_site.pre_tar_arr(cnt))
    Next cnt
    'a = position of last character before target date
    'b will get first character position that is after the target date.
    'error is b is not higher than a
    b = InStr(a, rawHtml, P_site.aft_target)
    datedebug = Mid(rawHtml, a, b - a)
    If IsDate(datedebug) Then
       datedebug = Format(datedebug, "MM/DD/YY")
    ElseIf InStr(1, datedebug, "<") > 0 Then 'fail since string still contains html
        datedebug = "error"
    End If
   getDate = datedebug
End Function

