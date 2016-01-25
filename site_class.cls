VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "site_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private P_pre_target_arr() As String
Private P_aft_target As String
Private Sub Class_Initialize()
    ReDim P_pre_target_arr(0)
End Sub
Public Property Let pre_tar_arr(index As Long, s As String)
    If index > UBound(P_pre_target_arr) Then ReDim Preserve P_pre_target_arr(index)
    P_pre_target_arr(index) = s
End Property
Public Property Get pre_tar_arr(index As Long) As String
    pre_tar_arr = P_pre_target_arr(index)
End Property
Public Property Get Pre_tar_arr_COUNT() As Long
    Pre_tar_arr_COUNT = UBound(P_pre_target_arr)
End Property
Public Property Let aft_target(s As String)
    P_aft_target = s
End Property
Public Property Get aft_target() As String
    aft_target = P_aft_target
End Property


