Attribute VB_Name = "XMLCode"
Public Type XMLTranslation
    ASCIIValue As String
    XMLCode As String
End Type

Public XMLTable() As XMLTranslation
Private XMLCodesLoaded As Boolean

Public Sub LoadXMLCodes()
    Dim ASCIIValue As String
    Dim XMLCode As String
    Dim X As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReDim XMLTable(0)
'    On Error Resume Next
    Open DataPath & "XMLCodes.dat" For Input As #FileNum
        Do While Not EOF(FileNum)
            Input #FileNum, ASCIIValue, XMLCode
            If Len(ASCIIValue) > 0 And XMLCode > "" Then
                X = UBound(XMLTable) + 1
                ReDim Preserve XMLTable(X)
                XMLTable(X).ASCIIValue = ASCIIValue
                XMLTable(X).XMLCode = XMLCode
            End If
        Loop
    Close #FileNum
    XMLCodesLoaded = True
End Sub

Public Function Txt2XMLCodes(ByVal Text As String) As String
    Dim Newline As String
    Dim y As Integer
    Dim Foundit As Long
    
    If XMLCodesLoaded = False Then LoadXMLCodes
    
    Newline = Text
    For y = 1 To UBound(XMLTable)
        Do
            Foundit = InStr(Foundit + 1, Newline, XMLTable(y).ASCIIValue)
            If Foundit Then
                Newline = Left$(Newline, Foundit - 1) & XMLTable(y).XMLCode & _
                  Right$(Newline, (Len(Newline) - Foundit - Len(XMLTable(y).ASCIIValue) + 1))
                  Foundit = Foundit + Len(XMLTable(y).XMLCode)
            End If
        Loop Until Val(Foundit) < 1
        
        Foundit = 0
    Next y
        
    Txt2XMLCodes = Newline
End Function




