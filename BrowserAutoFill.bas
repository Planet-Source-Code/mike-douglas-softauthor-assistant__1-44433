Attribute VB_Name = "BrowserAutoFill"
Public Type Tuple
    key As String
    Value As String
End Type

Private FormTuple() As Tuple
Private FormTuplesInit As Boolean
Private UnknownTuple() As String

Public Sub FillBrowserForm()
    Dim allCol
    Dim TagName As String
    Dim ItemName As String
    Dim PTList As String        'list of page tuples
    Dim FTKey As Integer
    Dim InpType As String
    
    If FormTuplesInit = False Then InitFormTuples
    SetPageTuples
    PTList = PageTupleList
    
'    For X = 0 To frmMain.WebBrowser1.Document.frames.Length - 1
        Set allCol = frmMain.WebBrowser1.Document.All
        allcount = allCol.Length
        For i = 0 To allcount - 1
            TagName = allCol.Item(i).TagName
            If (TagName = "INPUT" Or TagName = "TEXTAREA") Then
                InpType = UCase(allCol.Item(i).Type) & "**"
                If InStr(1, "SUBMIT**RESET**HIDDEN**RADIO**CHECKBOX**BUTTON**FILE**IMAGE**", InpType) = 0 Then
                    ItemName = allCol.Item(i).Name
                    FTKey = FindFormTuple(ItemName)
                    If FTKey > 0 Then
                        allCol.Item(i).Value = FindPageTuple(FormTuple(FTKey).Value)
                    Else
                        AddUnknownTuple (ItemName)
                    End If
                End If
            End If
        Next i
'    Next X
    
    If AdminMode = True Then
        For X = 0 To UBound(UnknownTuple)
            If UnknownTuple(X) > "" Then
                AddFormTuple UnknownTuple(X), frmSelect.SelectVals(PTList, "", False, "Unknown field=" & UnknownTuple(X)) 'select a pagetuple
            End If
        Next X
        SaveFormTupleDat
    End If
    
    ReDim UnknownTuple(0)
End Sub

Private Function InitFormTuples()
    FormTuplesInit = True
    ReDim FormTuple(0)
    ReDim UnknownTuple(0)
    
    SetPageTuples
    For X = 0 To UBound(PageTuple)
        AddFormTuple PageTuple(X).key, PageTuple(X).key
    Next X
        
    LoadFormTupleDat
End Function

Private Function FindUnknownTuple(ByVal key As String) As Integer
    '------------
    ' return form tuple index for given key, 0 on fail
    '------------
    Dim X As Integer
    Dim Found As Integer
    
    key = Trim(UCase(key))
    For X = 0 To UBound(UnknownTuple)
        If UCase(UnknownTuple(X)) = key Then
            Found = X
            Exit For
        End If
    Next X
    
    FindUnknownTuple = Found
End Function

Private Sub AddUnknownTuple(ByVal key As String)
    Dim X As Integer
    
    key = Trim(key)
    If key = "" Then Exit Sub
    X = UBound(UnknownTuple) + 1
    ReDim Preserve UnknownTuple(X)
    UnknownTuple(X) = key
End Sub

Private Function FindFormTuple(ByVal key As String) As Integer
    '------------
    ' return form tuple index for given key, 0 on fail
    '------------
    Dim X As Integer
    Dim Found As Integer
    
    key = Trim(UCase(key))
    For X = 0 To UBound(FormTuple)
        If FormTuple(X).key = key Then
            Found = X
            Exit For
        End If
    Next X
    
    FindFormTuple = Found
End Function

Public Sub LoadFormTupleDat()
    Dim FileNum As Integer
    Dim tmp As String
    Dim etmp() As String
    
    If FormTuplesInit = False Then InitFormTuples
    FileNum = FreeFile
'    On Error Resume Next
    If Dir(DataPath & "formtuples.dat") = "" Then Exit Sub
    Open DataPath & "formtuples.dat" For Input As #1
        Do Until EOF(FileNum)
            Line Input #1, tmp
            etmp = Split(tmp, ",")
            If UBound(etmp) = 1 Then
                If etmp(0) > "" And etmp(1) > "" Then
                    AddFormTuple etmp(0), etmp(1)
                End If
            End If
        Loop
    Close #1
End Sub

Public Sub SaveFormTupleDat()
    Dim FileNum As Integer
    
    If FormTuplesInit = False Then InitFormTuples
    FileNum = FreeFile
'    On Error Resume Next
    Open DataPath & "formtuples.dat" For Output As #1
        For X = 0 To UBound(FormTuple)
            If FormTuple(X).key > "" And FormTuple(X).Value > "" Then
                Print #1, FormTuple(X).key & "," & FormTuple(X).Value
            End If
        Next X
    Close #1
End Sub

Private Sub AddFormTuple(ByVal key As String, ByVal Value As String)
    Dim X As Integer
    
    key = Trim(UCase(key))
    Value = Trim(UCase(Value))
    If key = "" Or Value = "" Then Exit Sub
    If FindFormTuple(key) > 0 Then Exit Sub
    X = UBound(FormTuple) + 1
    ReDim Preserve FormTuple(X)
    FormTuple(X).key = key
    FormTuple(X).Value = Value
End Sub
