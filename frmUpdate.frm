VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkUpdateAll 
      Caption         =   "Update All"
      Height          =   225
      Left            =   1793
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2070
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   345
      Left            =   3593
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2010
      Width           =   915
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3990
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   173
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2010
      Width           =   915
   End
   Begin VB.Label lblDataNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   1118
      TabIndex        =   5
      Top             =   1710
      Width           =   2445
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Press [Update] to update data files."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   188
      TabIndex        =   3
      Top             =   1290
      Width           =   4305
   End
   Begin VB.Label lblBanner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Data Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   398
      TabIndex        =   2
      Top             =   660
      Width           =   3765
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const AutoUpdateBaseURL = "http://www.aesgard.com/product/softauthor/update/"
'Private Const AutoUpdateBaseURL = "http://localhost/aesgard/product/softauthor/update/"
Private UpdateCfg As String
Private UpdateQueue() As String
Private AbortUpdate As Boolean

Private Sub cmdCancel_Click()
'    Me.SetFocus
    AbortUpdate = True
    If Inet1.StillExecuting Then Inet1.CANCEL
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
'    Me.SetFocus
    
    ReDim UpdateQueue(0)
    AbortUpdate = False
    lblBanner = "Update in progress..."
    lblStatus = "Contacting server"
    chkUpdateAll.Enabled = False
    UpdateCfg = DLFile("update.cfg")
    If UpdateCfg = "" Then
        lblStatus = "Updates unavailable."
    Else
        If DemoVersion = True Then
            lblStatus = "Updates unavailable in unregistered version."
        Else
            lblStatus = "Processing updates."
            ParseCfg
            ProcessQueue
        End If
    End If
    chkUpdateAll.Enabled = True

    cmdCancel.Caption = "OK"
End Sub

Private Sub Form_Load()
    Do2000Look Me
    SetTooltips Me
    Me.Refresh
    lblDataNo = "Current data revision = v" & LastUpdateNo + 1
End Sub

Private Sub ProcessQueue()
    Dim XX As Integer
    Dim FileContents As String
    
    If AbortUpdate = True Then Exit Sub
    For X = 0 To UBound(UpdateQueue)
        If UpdateQueue(X) > "" Then
            lblStatus = "Downloading " & UpdateQueue(X)
            FileContents = DLFile(UpdateQueue(X))
            If FileContents > "" Then
                SaveFile UpdateQueue(X), FileContents
                XX = XX + 1
            End If
        End If
    Next X
    
    lblDataNo = "Current data revision=v" & LastUpdateNo + 1

    lblStatus = "Done. (" & XX & " files updated)"
End Sub

Private Sub SaveFile(ByVal FileName As String, ByVal Contents As String)
    Dim Extension As String
    Dim FilePath As String
    Dim PermissionFirst As Boolean
    Dim FullFileName As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If FileName = "" Or Contents = "" Then Exit Sub
    'on error resume next
    X = InStr(1, FileName, ".")
    If X = 0 Then Exit Sub
    Extension = Mid$(FileName, X)
    Select Case UCase(Extension)
        Case ".DAT"
            FilePath = DataPath
        Case ".HTM", ".HTML", ".GIF", ".JPG", ".JPEG", ".PNG", ".MPG", ".MPEG", ".XML"
            FilePath = HelpPath
        Case ".TEMPLATE"
            FilePath = TemplatePath
            PermissionFirst = True
        Case Else '".BLURB", ".PERMISSION"
            FilePath = DataPath
            PermissionFirst = True
    End Select
    
    FullFileName = FilePath & FileName
    If Dir(FullFileName) > "" And PermissionFirst = True Then
        XX = MsgBox("Overwrite " & FileName & " with new version?", vbYesNo, "Overwrite?")
        If XX = vbNo Then
            MsgBox "New version saved as NEW." & FileName & ". Review or delete as necessary."
            FullFileName = FilePath & "NEW." & FileName
        End If
    End If
    
    Open FullFileName For Output As #FileNum
        Print #FileNum, Contents;
    Close #FileNum
End Sub

Private Sub ParseCfg()
    Dim eCfg() As String
    Dim eLine() As String
    Dim key As String
    Dim Value As String
    Dim UpdateVer As String
    Dim AppVer As String
    Dim MaxVer As String
    Dim UpdateNo As Integer
    Dim MaxUpdateNo As Integer
    
    AppVer = App.Major & "." & App.Minor & "." & App.Revision
    eCfg = Split(UpdateCfg, vbCrLf)
    For X = 0 To UBound(eCfg)
        eLine = Split(eCfg(X), ",")
        If UBound(eLine) = 2 Then
            UpdateVer = Trim(eLine(0))      'the app.version needed to use this file
            UpdateNo = CInt(eLine(1))        'compared to LastUpdate# to see if done before
            Value = Trim(eLine(2))          'the file name
            If UpdateVer > MaxVer Then MaxVer = UpdateVer
            If UpdateNo > 0 And Value > "" And UpdateVer <= AppVer Then
                If (chkUpdateAll.Value = 1) Or (UpdateNo > LastUpdateNo) Then
                    If UpdateNo > MaxUpdateNo Then MaxUpdateNo = UpdateNo
                    y = UBound(UpdateQueue) + 1
                    ReDim Preserve UpdateQueue(y)
                    UpdateQueue(y) = Value
                End If
            End If
        End If
    Next X
    If MaxUpdateNo > 0 Then LastUpdateNo = MaxUpdateNo
    
    If MaxVer > AppVer Then MsgBox "Update to version " & MaxVer & " from www.aesgard.com for best results."
End Sub

Private Function DLFile(ByVal FileName As String) As String
    Dim URL As String
    Dim FileContents As String
    
    URL = AutoUpdateBaseURL & FileName
    FileContents = Inet1.OpenURL(URL, icString)
    
    If CheckHeader = False Then Exit Function
    
    DLFile = FileContents
End Function

Private Function CheckHeader() As Boolean
    Dim eHeader() As String
    Dim Header As String
    
    tmp = True
    
    Header = Inet1.GetHeader
    eHeader = Split(Header, vbCrLf)
    'find and compare response header
    For X = 0 To UBound(eHeader)
        If InStr(1, UCase(eHeader(X)), "HTTP") > 0 And InStr(1, eHeader(X), "404") > 0 Then
            FileContents = ""
            tmp = False
        End If
    Next X
    
    CheckHeader = tmp
End Function

