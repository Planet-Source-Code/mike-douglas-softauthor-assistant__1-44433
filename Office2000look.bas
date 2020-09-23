Attribute VB_Name = "Office2000Look"
Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, _
        ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Public Sub Do2000Look(frm As Form)
    For Each Control In frm.Controls
        If (TypeOf Control Is CommandButton) Then
            AddOfficeBorder Control.hwnd
        End If
    Next Control
End Sub

Public Sub AddOfficeBorder(ByVal hwnd As Long)
    'AddOfficeBorder(Text1.hWnd)
    Dim lngRetVal As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(hwnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
    'Apply the changes
    SetWindowLong hwnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
