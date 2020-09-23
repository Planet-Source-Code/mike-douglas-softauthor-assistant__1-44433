VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public LogActivity As Boolean       ' determines whether logging is enabled
Public FileNameBase As String       ' used as beginning of file name that date is added to
Public Silent As Boolean            ' determines if message boxes are shown (always shown on
                                    ' view if no file exists.

Public Function LogContents(FileDate As Date) As String
    '--------
    ' returns contents of logfile from given date
    '--------
    Dim LogFileName As String
    Dim FileNum As Integer
    Dim Temp As String
    Dim Message As String
        
    LogFileName = DateFile(FileDate)
    If Dir(LogFileName) = "" Then
        Message = "Log file is empty."
    Else
        FileNum = FreeFile
        Open LogFileName For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, Temp
                Message = Message & Temp & vbCrLf
                DoEvents
            Loop
        Close FileNum
    End If
    
    LogContents = Message
End Function

Public Sub ViewLog()
    Dim LogFileName As String
    Dim XX As Long
    
    LogFileName = DateFile(Date)
    If Dir(LogFileName) = "" Then
        If Not Silent Then MsgBox "Today's log file is empty"
    Else
        XX = ShellExecute(hwnd, vbNullString, LogFileName, vbNullString, App.Path, 1)
    End If
End Sub

Public Sub WriteLog(EventMessage As String)
    '-----------
    ' writes one line log entry
    '-----------
    Dim LogFileName As String
    Dim FileNum As Integer
    
    If LogActivity = False Then Exit Sub
    
    LogFileName = DateFile(Date)

    FileNum = FreeFile
    Open LogFileName For Append As #FileNum
        Print #FileNum, Now & " - " & EventMessage
    Close FileNum
End Sub

Public Sub Delete()
    '-----------
    ' deletes current logfile
    '-----------
    On Error Resume Next
    Kill DateFile(Date)
    If Not Silent Then MsgBox "Log file deleted."
End Sub

Public Sub Purge()
    '-----------
    ' deletes all logfiles of this type
    '-----------
    On Error Resume Next
    Kill (App.Path & "\" & FileNameBase & "." & App.EXEName & "*.txt")
    If Not Silent Then MsgBox "All log files deleted."
End Sub

Private Function DateFile(FileDate As Date) As String
    '------------
    ' creates name for file based on date
    '------------
    If FileNameBase = "" Then FileNameBase = "LOG"
    DateFile = App.Path & "\" & FileNameBase & "." & App.EXEName & Format(FileDate, "yyyy-mm-dd") & ".txt"
End Function

