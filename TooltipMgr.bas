Attribute VB_Name = "TooltipMgr"
Public Type Tooltip
    Form As String
    Control As String
    Text As String
End Type

Public Tooltips() As Tooltip
Private TipsLoaded As Boolean

Public Sub LoadTooltips()
    Dim etmp() As String
    Dim frm As String
    Dim X As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReDim Tooltips(0)
    TipsLoaded = True
    frm = "frmMain"
    If Dir(DataPath & "tooltips.dat") = "" Then Exit Sub
    Open DataPath & "tooltips.dat" For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, tmp
            etmp = Split(tmp, "=")
            If UBound(etmp) > 0 Then
                Select Case UCase(etmp(0))
                    Case "FORM"
                        frm = etmp(1)
                    Case Else
                        AddToolTip frm, etmp(0), etmp(1)
                End Select
            End If
        Loop
    Close #FileNum
End Sub

Public Sub SetTooltips(frm As Form)
    Dim X As Integer
    Dim y As Integer
    Dim frmName As String
    Dim ctrlName As String
    
    If TipsLoaded = False Then LoadTooltips
    frmName = UCase(frm.Name)
    For Each Control In frm.Controls
        ctrlName = UCase(Control.Name)
        For X = 0 To UBound(Tooltips)
            If frmName = Tooltips(X).Form Then
                If ctrlName = Tooltips(X).Control Then Control.ToolTipText = Tooltips(X).Text
            End If
        Next X
    Next
End Sub

Private Sub AddToolTip(ByVal tForm As String, ByVal tControl As String, ByVal tText As String)
    Dim X As Integer
    
    X = UBound(Tooltips) + 1
    ReDim Preserve Tooltips(X)
    Tooltips(X).Form = UCase(tForm)
    Tooltips(X).Control = UCase(tControl)
    Tooltips(X).Text = tText
End Sub

