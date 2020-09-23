VERSION 5.00
Begin VB.Form frmExpireInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expiration Information"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtExpirationOtherInfo 
      Height          =   990
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1650
      Width           =   2535
   End
   Begin VB.TextBox txtExpirationCount 
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   90
      Width           =   2535
   End
   Begin VB.TextBox txtExpirationDate 
      Height          =   300
      Left            =   1500
      TabIndex        =   3
      Top             =   1230
      Width           =   2535
   End
   Begin VB.CommandButton cmdSelectExpireBasedOn 
      Height          =   300
      Left            =   3690
      Picture         =   "frmExpireInfo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   345
   End
   Begin VB.TextBox txtExpireBasedOn 
      Height          =   300
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   2205
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Other Info:"
      Height          =   255
      Left            =   -150
      TabIndex        =   8
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Count:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   135
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date:"
      Height          =   255
      Left            =   -150
      TabIndex        =   4
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Base:"
      Height          =   255
      Left            =   -150
      TabIndex        =   2
      Top             =   630
      Width           =   1515
   End
End
Attribute VB_Name = "frmExpireInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Sub cmdSelectExpireBasedOn_Click()
    Notontop Me
    txtExpireBasedOn = frmSelect.SelectVals(frmMain.Predef.INIValue("Expire_Based_On"), txtExpireBasedOn, False, "Expire Based On")
    OnTop Me
End Sub

Private Sub Form_Load()
    txtExpirationCount = curExpirationCount
    txtExpireBasedOn = curExpireBasedOn
    txtExpirationDate = curExpirationDate
    txtExpirationOtherInfo = curExpirationOtherInfo

    Do2000Look Me
    OnTop Me
    SetTooltips Me
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
    curExpirationCount = txtExpirationCount
    curExpireBasedOn = txtExpireBasedOn
    curExpirationDate = txtExpirationDate
    curExpirationOtherInfo = txtExpirationOtherInfo

'    If Trim(curExpirationCount & curExpireBasedOn & curExpirationDate & curExpirationOtherInfo) > "" Then
'        frmMain.chkProgramExpires.Value = 1
'    Else
'        frmMain.chkProgramExpires.Value = 0
'    End If
    
    Unload frmSelect
    Notontop Me
End Sub

Public Sub OnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Sub Notontop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Private Sub txtExpirationCount_Change()
    ProgramModified = True
End Sub

Private Sub txtExpirationCount_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyNumbers(KeyAscii)
End Sub

Private Sub txtExpirationDate_Change()
    ProgramModified = True
End Sub

Private Sub txtExpirationOtherInfo_Change()
    ProgramModified = True
End Sub

Private Sub txtExpireBasedOn_Change()
    ProgramModified = True
End Sub
