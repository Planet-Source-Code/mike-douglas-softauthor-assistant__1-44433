VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' MICROINI.CLS
'
' Simple procedures to build tuple array representing
' contents of INI file.
'
'---------------------------------------------------------
' Two simple functions:
'       ReadINI "project.ini" - reads INI file into tuple array ignoring
'                               lines after a "#" (REMARK), blank lines
'                               and lines not containing an "=". Example
'                               uses "project.ini" as ini file read in.
'
'       INIValue("ZERO")      - returns string containing value of INI key
'                               searched for.  Example uses "ZERO" as key.
'
'---------------------------------------------------------

Private Type Tuple
    key As String
    Value As String
End Type

Dim INISet() As Tuple
Public INITupleCount As Integer

Public Function INIValue(key As String) As String
    '------------
    ' Takes INI Key and returns the value
    '------------
    Dim X As Integer
    
    For X = 1 To INITupleCount
        If UCase(INISet(X).key) = UCase(key) Then
            INIValue = (INISet(X).Value)
        End If
    Next X
    
End Function

Public Sub ReadINI(FileName As String)
    '------------
    ' Takes filename and inputs contents into
    ' INISet() tuple array
    '------------
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If Dir(FileName) > "" Then
        Open FileName For Input As #FileNum
            Do While Not EOF(FileNum)
                Line Input #FileNum, Temp
                SetTuple (Temp)
                DoEvents
            Loop
        Close #FileNum
    End If

End Sub

Private Sub SetTuple(INILine As String)
    '------------
    ' Takes line from INI file and formats it to INISet()
    ' tuple array
    '------------
    Dim REMFound As Integer
    Dim EQUFound As Integer
    Dim key As String
    Dim Value As String
    
    REMFound = InStr(1, INILine, "#")           ' look for REM statements
    If REMFound Then
        INILine = Left(INILine, REMFound - 1)
    End If
    
    INILine = Trim(INILine)
    EQUFound = InStr(1, INILine, "=")           ' look for =, to split key/value
    If EQUFound > 1 Then
        INITupleCount = INITupleCount + 1
        ReDim Preserve INISet(INITupleCount)
        
        key = Trim(Left(INILine, EQUFound - 1))
        Value = Trim(Right(INILine, Len(INILine) - EQUFound))
        
        INISet(INITupleCount).key = key
        INISet(INITupleCount).Value = Value
    End If

End Sub


