VERSION 5.00
Begin VB.Form frmPostIt 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "x"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5175
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPostIt 
      BackColor       =   &H00C0FFFF&
      Height          =   3165
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Use this space to save information regarding this site."
      Top             =   10
      Width           =   4665
   End
End
Attribute VB_Name = "frmPostIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Do2000Look Me
    OnTop Me
    SetTooltips Me
    
    If PostitLeft > 0 Then
        Me.Left = PostitLeft
    Else
        Me.Left = frmMain.Left
    End If
    
    If PostitTop > 0 Then
        Me.Top = PostitTop
    Else
        Me.Top = frmMain.Top
    End If
    
    If PostitWidth > 0 Then Me.Width = PostitWidth
    If PostitHeight > 0 Then Me.Height = PostitHeight
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
    Notontop Me
    PostitLeft = Me.Left
    PostitTop = Me.Top
    PostitHeight = Me.Height
    PostitWidth = Me.Width
End Sub

Private Sub Form_Resize()
    txtPostIt.Top = 10
    txtPostIt.Left = 10
    txtPostIt.Width = Me.Width - 20
    txtPostIt.Height = Me.Height - 310
End Sub

Private Sub txtPostIt_Change()
    SubmitList(SubmitListIndex).Notes = txtPostIt.Text
End Sub
