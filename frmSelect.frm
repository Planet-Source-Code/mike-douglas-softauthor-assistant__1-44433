VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "\/"
      Height          =   285
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3660
      Width           =   765
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "/\"
      Height          =   285
      Left            =   930
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3660
      Width           =   765
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Index           =   1
      Left            =   120
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   3255
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Height          =   285
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3210
      Width           =   765
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "None"
      Height          =   285
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3210
      Width           =   765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3210
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3210
      Width           =   765
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   3255
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'Const HWND_NOTOPMOST = -2
'Const HWND_TOPMOST = -1
'Const SWP_NOMOVE = &H2
'Const SWP_NOSIZE = &H1
'Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Dim DoneHere As Boolean
Dim mList As Integer

Public Function SelectVals(ByVal Choices As String, ByVal PrevChoices As String, ByVal MultiSelect As Boolean, Optional Title As String, Optional ByVal DelimiterIN As String, Optional ByVal DelimiterOUT As String, Optional Ordered As Boolean) As String
    Dim choice() As String
    Dim tmp As String
    
    List1(0).Clear
    List1(1).Clear
    
    cmdNone.Visible = MultiSelect
    cmdAll.Visible = MultiSelect
    If MultiSelect = False And Ordered = True Then
        cmdMoveUp.Visible = True
        cmdMoveDown.Visible = True
        cmdMoveUp.Left = cmdAll.Left
        cmdMoveUp.Top = cmdAll.Top
        cmdMoveDown.Left = cmdNone.Left
        cmdMoveDown.Top = cmdNone.Top
    Else
        cmdMoveUp.Visible = False
        cmdMoveDown.Visible = False
    End If
    
    If MultiSelect = True Then
        mList = 1
        List1(1).Visible = True
        List1(0).Visible = False
    Else
        mList = 0
        List1(0).Visible = True
        List1(1).Visible = False
    End If
    
    If DelimiterIN = "" Then DelimiterIN = ","
    If DelimiterOUT = "" Then DelimiterOUT = ","
    If Title = "" Then Title = "Select"
    Me.Caption = Title
    
    choice = Split(Choices, DelimiterIN)
    
    PrevChoices = DelimiterIN & PrevChoices & DelimiterIN
    
    For X = 0 To UBound(choice)
        If Trim(choice(X)) > "" Then
            List1(mList).AddItem choice(X)
        End If
    Next X
    
    For X = 0 To List1(mList).ListCount - 1
        If InStr(1, PrevChoices, DelimiterIN & List1(mList).List(X) & DelimiterIN) > 0 Then List1(mList).Selected(X) = True
    Next X
    
    If List1(mList).ListCount > 0 And List1(mList).SelCount = 0 Then List1(mList).Selected(0) = True

    Me.Show
    
    Do Until DoneHere = True      'wait for cmd button input
        DoEvents
    Loop
    DoneHere = False
    
    If Ordered = False Then
        'build by selection
        For X = 0 To List1(mList).ListCount - 1
            If List1(mList).Selected(X) = True Then tmp = tmp & List1(mList).List(X) & DelimiterOUT
        Next X
    Else
        'build ordered list
        tmp = ""
        For X = 0 To List1(mList).ListCount - 1
            tmp = tmp & List1(mList).List(X) & DelimiterOUT
        Next X
    End If
    
    If tmp > "" Then tmp = Left(tmp, Len(tmp) - Len(DelimiterOUT))
    
    If tmp = "" Then
        Status ""
    Else
        If List1(mList).SelCount > 1 Then
            Status List1(mList).SelCount & " items selected."
        Else
            Status tmp & " selected."
        End If
    End If
    
    SelectVals = tmp
    Notontop Me
    Unload Me
End Function

Private Sub cmdAll_Click()
    Me.SetFocus
    For X = 0 To List1(mList).ListCount - 1
        List1(mList).Selected(X) = True
    Next X
End Sub

Private Sub cmdCancel_Click()
    Me.SetFocus
    List1(mList).Clear
    DoneHere = True
End Sub

Private Sub cmdMoveDown_Click()
    Me.SetFocus
    If List1(0).ListCount = 0 Then Exit Sub
    X = List1(0).ListIndex
    If X < List1(0).ListCount - 1 Then
        temptxt1 = List1(0).List(X + 1)
        tempsel1 = List1(0).Selected(X + 1)
        
        temptxt2 = List1(0).List(X)
        tempsel2 = List1(0).Selected(X)
        
        List1(0).RemoveItem X
        List1(0).RemoveItem X
        
        List1(0).AddItem temptxt2, X
        List1(0).AddItem temptxt1, X
        
        List1(0).Selected(X + 1) = tempsel2
        List1(0).Selected(X) = tempsel1
        
        List1(0).ListIndex = X + 1
    End If
End Sub

Private Sub cmdMoveUp_Click()
    Me.SetFocus
    If List1(0).ListCount = 0 Then Exit Sub
    X = List1(0).ListIndex
    If X > 0 Then
        temptxt1 = List1(0).List(X - 1)
        tempsel1 = List1(0).Selected(X - 1)
        
        temptxt2 = List1(0).List(X)
        tempsel2 = List1(0).Selected(X)
        
        List1(0).RemoveItem X - 1
        List1(0).RemoveItem X - 1
        
        List1(0).AddItem temptxt1, X - 1
        List1(0).AddItem temptxt2, X - 1
        
        List1(0).Selected(X) = tempsel1
        List1(0).Selected(X - 1) = tempsel2
        
        List1(0).ListIndex = X - 1
    End If
End Sub

Private Sub cmdNone_Click()
    Me.SetFocus
    For X = 0 To List1(mList).ListCount - 1
        List1(mList).Selected(X) = False
    Next X
End Sub

Private Sub cmdOK_Click()
    Me.SetFocus
    DoneHere = True
End Sub

Private Sub Form_Load()
    Do2000Look Me
    SetTooltips Me
    OnTop Me
End Sub

'Public Sub OnTop(FormName As Form)
'    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
'End Sub
'
'Public Sub Notontop(FormName As Form)
'    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
'End Sub


