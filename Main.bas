Attribute VB_Name = "Main"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public ProgramModified As Boolean
Public CompanyModified As Boolean
Public AuthorModified As Boolean
Public PermissionsModified As Boolean
Public BlurbModified As Boolean
Public SubmitModified As Boolean

Public PostitLeft As Long
Public PostitTop As Long
Public PostitHeight As Long
Public PostitWidth As Long

Public WebSubmiting As Boolean
Public LastUpdateNo As Integer

Public AdminMode As Boolean


Public Sub Status(ByVal Message As String)
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If frmConsole.Visible = True Then
        frmConsole.Message Message
        Open App.Path & "\EMAILSUBMITLOG.TXT" For Append As #FileNum
            Print #FileNum, Message
        Close #FileNum
    Else
        frmMain.lblStatus = Message
        frmMain.Refresh
    End If
End Sub

Public Sub CheckModifiedData()
    If ProgramModified = True And Trim(frmMain.txtProductName.Text) > "" Then
        Status "Saving current program data..."
        SaveProgram Trim(frmMain.txtProductName.Text) & " v" & frmMain.txtProgramVersion.Text
        Status "Data saved."
    End If
    
    If CompanyModified = True And Trim(frmMain.txtCompanyName.Text) > "" Then
        Status "Saving current company data..."
        SaveCompany frmMain.txtCompanyName.Text
        Status "Data saved."
    End If
    
    If AuthorModified = True And Trim(frmMain.txtAuthorFName.Text & frmMain.txtAuthorLName.Text) > "" Then
        Status "Saving current author data..."
        SaveAuthor Trim(frmMain.txtAuthorFName.Text & " " & frmMain.txtAuthorLName.Text)
        Status "Data saved."
    End If
    
    If PermissionsModified = True And Trim(frmMain.txtPermissionName.Text) > "" Then
        Status "Saving current permission data..."
        SavePermission frmMain.txtPermissionName.Text
        Status "Data saved."
    End If

    If BlurbModified = True And Trim(frmMain.txtBlurbName.Text) > "" Then
        Status "Saving current blurb data..."
        SaveBlurb frmMain.txtBlurbName.Text
        Status "Data saved."
    End If
    
End Sub
