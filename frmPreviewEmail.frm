VERSION 5.00
Begin VB.Form frmPreviewEmail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "E-Mail"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMailMessage 
      Height          =   2505
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2790
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2790
      Width           =   915
   End
End
Attribute VB_Name = "frmPreviewEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    frmConsole.Go
    frmConsole.Show 'vbModal, Me
End Sub

Private Sub Form_Load()
    Dim tmpEmailFile As String
    
    Do2000Look Me
    SetTooltips Me
    Me.Top = frmMain.Top
    Me.Left = frmMain.Left
    Me.Width = frmMain.Width
    Me.Height = frmMain.Height
End Sub

Private Sub Form_Resize()
    txtMailMessage.Width = Me.Width - 150
    txtMailMessage.Height = Me.Height - cmdOK.Height - 500
    cmdOK.Top = txtMailMessage.Height + 100
    cmdOK.Left = Me.Width - cmdOK.Width - 150
    cmdCancel.Top = cmdOK.Top
End Sub
