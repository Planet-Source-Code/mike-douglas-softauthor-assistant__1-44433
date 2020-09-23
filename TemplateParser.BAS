Attribute VB_Name = "TemplateParser"
Type Tuple                              ' A Tuple is a structure that holds name\value pairs
    key As String
    Value As String
End Type

Public TemplatePath As String
Public OutputPath As String

Public PageTuple() As Tuple             ' an array of page tuples
Public TuplesInit As Boolean

Public Sub InitPageTuples()
    ReDim PageTuple(0)
    TuplesInit = True
End Sub

Public Function strParseTemplateText(ByVal Text As String) As String
    Dim FileContents As String
    Dim Search As String
    Dim Found As Long
    Dim X As Integer
    Dim temp$
    
    If InStr(1, Text, "<!--") = 0 Then
        strParseTemplateText = Text
        Exit Function
    End If
    
    If TuplesInit = False Then InitPageTuples
    
    SetStandardPageValues
    
    FileContents = Text
    
    For X = 1 To UBound(PageTuple)
        Search = "<!--" & PageTuple(X).key & "-->"
        Do While InStr(1, FileContents, Search, vbTextCompare)
            Found = InStr(1, FileContents, Search, vbTextCompare)
            If Found Then
                FileContents = Left(FileContents, Found - 1) & _
                    PageTuple(X).Value & Mid(FileContents, _
                    Found + Len(Search))
            End If
        Loop
    Next X
    
    If FileContents = "" Then Exit Function
    
    strParseTemplateText = FileContents
End Function

Public Function strParseTemplate(ByVal TemplateFile As String) As String
    Dim FileContents As String
    Dim Search As String
    Dim Found As Long
    Dim X As Integer
    Dim temp$
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If TuplesInit = False Then InitPageTuples
    
    SetStandardPageValues
    
    FileContents = ""
    If InStr(TemplateFile, "\") = 0 Then TemplateFile = TemplatePath & TemplateFile

    If Dir(TemplateFile) = "" Then Exit Function
    
    Open (TemplateFile) For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, temp$
            FileContents = FileContents & temp$ & vbCrLf
        Loop
    Close #FileNum
    
    If InStr(1, FileContents, "<!--") = 0 Then
        strParseTemplate = FileContents
        Exit Function
    End If
        
    For X = 1 To UBound(PageTuple)
        Search = "<!--" & PageTuple(X).key & "-->"
        Do While InStr(1, FileContents, Search, vbTextCompare)
            Found = InStr(1, FileContents, Search, vbTextCompare)
            If Found Then
                FileContents = Left(FileContents, Found - 1) & _
                    PageTuple(X).Value & Mid(FileContents, _
                    Found + Len(Search))
            End If
        Loop
    Next X
    
    If FileContents = "" Then Exit Function
    
    strParseTemplate = FileContents
End Function

Public Sub ParseTemplate(ByVal TemplateFile As String, ByVal OutputFile As String)
    Dim tmpContents As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    tmpContents = strParseTemplate(TemplateFile)
    If tmpContents = "" Then Exit Sub
    If InStr(OutputFile, "\") = 0 Then OutputFile = OutputPath & OutputFile

    Open OutputFile For Output As #FileNum
        Print #FileNum, tmpContents
    Close #FileNum
End Sub

Public Function FindPageTuple(ByVal key As String) As String
    '------------------------
    ' give it a key, it'll give you the value ("" if not found)
    '------------------------
    Dim X As Integer
    Dim Found As Integer
    
    If TuplesInit = False Then InitPageTuples
    
    key = Trim(UCase(key))
    For X = 0 To UBound(PageTuple)
        If UCase(PageTuple(X).key) = key Then
            Found = X
            Exit For
        End If
    Next X
    
    If Found > 0 Then
        FindPageTuple = PageTuple(Found).Value
    Else
        FindPageTuple = ""
    End If
End Function

Public Function PageTupleList() As String
    '-----------------------
    ' returns comma-delimited list of existing pagetuple keys
    '-----------------------
    Dim X As Integer
    Dim PTList As String
    
    If TuplesInit = False Then InitPageTuples
    
    PTList = ""
    For X = 0 To UBound(PageTuple)
        If PageTuple(X).key > "" Then PTList = PTList & PageTuple(X).key & ","
    Next X
    
    If PTList > "" Then PTList = Left(PTList, Len(PTList) - 1)
    
    PageTupleList = PTList
End Function

Public Sub AddPageTuple(key As String, Value As Variant)
    '-------------------------
    ' This sub called when adding a new pagetuple to the PageTuple array
    '-------------------------
    Dim PageTupleCount As Integer
    
    If TuplesInit = False Then InitPageTuples
    
    PageTupleCount = UBound(PageTuple) + 1
    ReDim Preserve PageTuple(PageTupleCount)
    PageTuple(PageTupleCount).key = UCase(key)
    PageTuple(PageTupleCount).Value = Value
End Sub

Private Sub SetStandardPageValues()
    ''''''''''''''''''''''''''''''''''''''
    ' Comment: Sets common page variables from CGI environment
    ''''''''''''''''''''''''''''''''''''''
    Dim Value, key As String         ' dummy variable
    Dim wDay$, NMonth$
    
    '----------- BEGIN CALENDAR VARIABLES --------------------------
    key = "NOW": AddPageTuple key, Now
    key = "DATE": AddPageTuple key, Date
    key = "DAY_OF_THE_MONTH": AddPageTuple key, Day(Now)
    key = "WEEKDAY"
        Select Case Weekday(Now)
            Case 1
                wDay$ = "Monday"
            Case 2
                wDay$ = "Tuesday"
            Case 3
                wDay$ = "Wednesday"
            Case 4
                wDay$ = "Thursday"
            Case 5
                wDay$ = "Friday"
            Case 6
                wDay$ = "Saturday"
            Case 7
                wDay$ = "Sunday"
        End Select
        AddPageTuple key, wDay$
        
    key = "WEEKDAY_SHORT": AddPageTuple key, Left(UCase(wDay$), 3)

    key = "MONTH"
        Select Case Month(Now)
            Case 1
                NMonth$ = "January"
            Case 2
                NMonth$ = "February"
            Case 3
                NMonth$ = "March"
            Case 4
                NMonth$ = "April"
            Case 5
                NMonth$ = "May"
            Case 6
                NMonth$ = "June"
            Case 7
                NMonth$ = "July"
            Case 8
                NMonth$ = "August"
            Case 9
                NMonth$ = "September"
            Case 10
                NMonth$ = "October"
            Case 11
                NMonth$ = "November"
            Case 12
                NMonth$ = "December"
        End Select
        AddPageTuple key, NMonth$
        
    key = "MONTH_SHORT": AddPageTuple key, Left(UCase(NMonth$), 3)
    key = "MONTH_NUMBER": AddPageTuple key, Month(Now)
    
    key = "TIME": AddPageTuple key, Time
    key = "YEAR": AddPageTuple key, Year(Now)
    '----------- END CALENDAR VARIABLES --------------------------
End Sub

