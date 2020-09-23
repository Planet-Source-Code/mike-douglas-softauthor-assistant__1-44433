Attribute VB_Name = "HTMLCode"
Public Type HTMLTranslation
    ASCIIValue As String
    Number As Integer
    EnglishCode As String
    HTMLCode As String
    Description As String
End Type

Public HTMLTable() As HTMLTranslation
Private HTMLCodesLoaded As Boolean


Public Sub LoadHTMLCodes()
    Dim Number As Integer
    Dim EnglishCode As String
    Dim HTMLCode As String
    Dim Description As String
    Dim X As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReDim HTMLTable(1)
    HTMLTable(1).ASCIIValue = vbCrLf
    HTMLTable(1).HTMLCode = "<BR>" & vbCrLf
    
    On Error Resume Next
    Open DataPath & "HTMLCodes.dat" For Input As #FileNum
        Do While Not EOF(FileNum)
            Input #FileNum, ASCIIValue, Number, EnglishCode, HTMLCode, Description
            If Len(ASCIIValue) > 0 And HTMLCode > "" Then
                X = UBound(HTMLTable) + 1
                ReDim Preserve HTMLTable(X)
                HTMLTable(X).ASCIIValue = ASCIIValue
                HTMLTable(X).HTMLCode = HTMLCode
            End If
        Loop
    Close #FileNum
    HTMLCodesLoaded = True
End Sub

Public Function Txt2HTMLCodes(ByVal Text As String) As String
    Dim Newline As String
    Dim Y As Integer
    Dim FoundIt As Long
    
    If HTMLCodesLoaded = False Then LoadHTMLCodes
    
    Newline = Text
    For Y = 1 To UBound(HTMLTable)
        FoundIt = 0
        Do
            FoundIt = InStr(FoundIt + 1, Text, HTMLTable(Y).ASCIIValue)
            If FoundIt Then
                Newline = Left$(Text, FoundIt - 1) & HTMLTable(Y).HTMLCode & _
                  Right$(Text, (Len(Text) - FoundIt - Len(HTMLTable(Y).ASCIIValue) + 1))
                  FoundIt = FoundIt + Len(HTMLTable(Y).HTMLCode)
            End If
            DoEvents
        Loop Until Val(FoundIt) < 1
    Next Y
        
    Txt2HTMLCodes = Newline
End Function


