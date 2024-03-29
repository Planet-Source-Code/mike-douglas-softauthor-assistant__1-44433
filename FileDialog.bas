Attribute VB_Name = "FileDialog"
'Description: Calls the "Open File Dialog" without need for an OCX
'Be careful when dealing with this and the "Save File Dialog", the
'Type and examples are the same. It can be confusing...

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function SelectFile() As String
    Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hwnd
    ofn.hInstance = App.hInstance
'    ofn.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Rich Text Files (*.rtf)" + Chr$(0) + "*.rtf" + Chr$(0)
    ofn.lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = CurDir
        ofn.lpstrTitle = "Select Distribution File"
        ofn.flags = 0
        Dim a
        a = GetOpenFileName(ofn)

        If (a) Then
                SelectFile = Trim$(ofn.lpstrFile)
        Else
                SelectFile = "" ' Cancel was pressed
        End If
End Function
