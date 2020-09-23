VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "E-Mail Submission Console"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "EXIT"
      Height          =   315
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3840
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4350
      Top             =   390
   End
   Begin VB.TextBox txtConsole 
      Enabled         =   0   'False
      Height          =   2865
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   780
      Width           =   945
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub cmdExit_Click()
    SMTPClient.AbortSend
    Unload SMTPClient
    txtConsole.Text = ""
    Unload Me
End Sub

Private Sub Form_Load()
    frmPreviewEmail.Hide
'    Unload frmPreviewEmail
    
    Do2000Look Me
    SetTooltips Me
    Me.Top = frmMain.Top
    Me.Left = frmMain.Left
    Me.Width = frmMain.Width
    Me.Height = frmMain.Height
    
    OnTop Me
    txtConsole.Height = Me.Height - cmdExit.Height - 400
    txtConsole.Width = Me.Width - (txtConsole.Left * 2) - 150
    txtConsole.Text = "Beginning mail send..." & vbCrLf & vbCrLf
    
    cmdExit.Top = txtConsole.Height + txtConsole.Top + 50
    cmdExit.Left = txtConsole.Left + txtConsole.Width - cmdExit.Width
End Sub

Public Sub Message(ByVal Message As String)
    txtConsole = txtConsole & Message & vbCrLf
    txtConsole = Right(txtConsole.Text, 5000)
    txtConsole.SelStart = Len(txtConsole.Text)
End Sub

Public Sub Go()
    Timer1.Enabled = True
    OnTop Me
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
    Notontop Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If Dir(App.Path & "\EMAILSUBMIT.TXT") > "" Then Kill App.Path & "\EMAILSUBMIT.TXT"
    SMTPClient.ToAddress = frmMain.txtEmailAddress
    SMTPClient.FromName = frmMain.txtFromName
    SMTPClient.FromAddress = frmMain.txtFromAddress
    SMTPClient.MailServer = frmMain.txtMailserver
    SMTPClient.SMTPPort = Val(frmMain.txtSMTPPort)
    SMTPClient.TimeOut = Val(frmMain.txtTimeOut)
    SMTPClient.Message = frmPreviewEmail.txtMailMessage
    SMTPClient.RelayUser = frmMain.txtSMTPUser
    SMTPClient.RelayPassword = frmMain.txtSMTPPassword
    SMTPClient.Retries = frmMain.txtEmailRetries
    
    SMTPClient.Send
    Status ""
    Status "***** COMPLETE *****"
End Sub
