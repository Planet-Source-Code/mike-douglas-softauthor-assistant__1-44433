VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   4590
      Begin VB.Timer tmrSplash 
         Interval        =   1000
         Left            =   510
         Top             =   180
      End
      Begin VB.Label lblLoading 
         Alignment       =   2  'Center
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image imgLogo 
         Height          =   1005
         Left            =   270
         Stretch         =   -1  'True
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   1590
         Width           =   4425
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   1800
         Width           =   4425
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   60
         TabIndex        =   2
         Top             =   2160
         Width           =   4485
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3765
         TabIndex        =   5
         Top             =   990
         Width           =   720
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3525
         TabIndex        =   6
         Top             =   690
         Width           =   960
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   7
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Platform = "Win32"

'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'Const HWND_NOTOPMOST = -2
'Const HWND_TOPMOST = -1
'Const SWP_NOMOVE = &H2
'Const SWP_NOSIZE = &H1
'Const flags = SWP_NOMOVE Or SWP_NOSIZE

Dim TimerCycles As Integer

Private Sub Form_Load()
    Dim temp As String
    
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.ProductName
    lblCopyright = App.LegalCopyright
    lblCompany = App.CompanyName
    lblPlatform = Platform
    lblWarning = "Warning: This computer program is protected by international copyright law and treaties. " & _
                 "Unauthorized reproduction or distribution of this program or portions thereof, may result " & _
                 "in severe civil and criminal penalties and will be prosecuted to the fullest extent of the law."
    
    OnTop Me

'    Me.Refresh
End Sub

Private Sub Form_Paint()
    imgLogo.Picture = frmMain.Icon
    lblLoading.Visible = False
End Sub

'Public Sub OnTop(FormName As Form)
'    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
'End Sub
'
'Public Sub Notontop(FormName As Form)
'    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags)
'End Sub

Private Sub tmrSplash_Timer()
    TimerCycles = TimerCycles + 1

    Select Case TimerCycles
        Case 1
            frmMain.Visible = False
            imgLogo.Picture = frmMain.Icon
            lblLoading.Visible = False
            Me.Refresh
        Case 5
            Unload Me
    End Select
End Sub

Private Sub Form_QueryUnload(CANCEL As Integer, UnloadMode As Integer)
    Notontop Me
    frmMain.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    DoUnload
End Sub

Private Sub imgLogo_Click()
    DoUnload
End Sub

Private Sub lblCompany_Click()
    DoUnload
End Sub

Private Sub lblCopyright_Click()
    DoUnload
End Sub

Private Sub lblLicenseTo_Click()
    DoUnload
End Sub

Private Sub lblPlatform_Click()
    DoUnload
End Sub

Private Sub lblProductName_Click()
    DoUnload
End Sub

Private Sub lblVersion_Click()
    DoUnload
End Sub

Private Sub lblWarning_Click()
    DoUnload
End Sub

Private Sub Frame1_Click()
    DoUnload
End Sub

Private Sub DoUnload()
'    imgLogo.Picture = frmMain.Icon
    Unload Me
End Sub
