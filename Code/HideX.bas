Attribute VB_Name = "HideX"
'Include this code at the top of the module
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000

#If VBA7 Then

    Private Declare PtrSafe Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
        
#Else

    Private Declare Function GetWindowLong _
        Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong _
        Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function DrawMenuBar _
        Lib "user32" (ByVal hWnd As Long) As Long
#End If

'Include this code in the same module as the API calls above
Public Sub SystemButtonSettings(frm As Object, show As Boolean)
Dim windowStyle As Long
Dim windowHandle As Long

windowHandle = FindWindowA(vbNullString, frm.Caption)
windowStyle = GetWindowLong(windowHandle, GWL_STYLE)

If show = False Then

    SetWindowLong windowHandle, GWL_STYLE, (windowStyle And Not WS_SYSMENU)

Else

    SetWindowLong windowHandle, GWL_STYLE, (windowStyle + WS_SYSMENU)

End If

DrawMenuBar (windowHandle)

End Sub



