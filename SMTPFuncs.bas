Attribute VB_Name = "SMTPFuncs"
' SMTPFUNCS.BAS
'
'Common SMTP email functions.
'--------------------------
'IsSMTPAddress(S$) returns true if S$ is a validly formed SMTP address
'UUEncode(D$) returns encoded string based on D$
'UUDecode(D$) returns decoded string based on D$
'Base64Encode(F$) returns Base64 encoded data based on filename (F$)


Public Function IsSMTPAddress(Incoming As String) As Boolean
    '-------------------
    ' Tests "Incoming" to see if it is a validly formed SMTP address
    ' and returns TRUE if so.
    '-------------------
    Dim ATFound As Integer                          ' location of "@" in incoming string
    Dim DotFound As Integer                         ' location of first "." in incoming string
    Dim NameSection As String                       ' contents before "@" in address
    Dim SubDomain As String                         ' contents between "@" and "."
    Dim HighDomain As String                        ' contents after "."
    
    ATFound = InStr(1, Incoming, "@")
    
    If ATFound > 1 Then
        DotFound = InStr(ATFound, Incoming, ".")    ' only look for first dot after "@"
        Do While InStr(DotFound + 1, Incoming, ".")
            DotFound = InStr(DotFound + 1, Incoming, ".")
        Loop
        
        If DotFound > 1 And DotFound < Len(Incoming) Then
            NameSection = Left(Incoming, ATFound - 1)
            SubDomain = Mid(Incoming, ATFound + 1, (DotFound - ATFound - 1))
            HighDomain = Right(Incoming, (Len(Incoming) + 1) - DotFound)
            
            If (NameSection > "") And (SubDomain > "") And (HighDomain > "") Then
                IsSMTPAddress = True
            End If
        Else
            IsSMTPAddress = False
        End If
    Else
        IsSMTPAddress = False
    End If

End Function

'-----------------------------------------------------
'        BEGIN UUENCODE FUNCTIONS
'-----------------------------------------------------

'Public Function UUDecode(Data As String) As String
'    '-----------
'    ' UU Decode string (data) from transmission.
'    '-----------
'    On Error GoTo ErrorHandler
'
'    Dim szOut   As String
'    Dim nChar   As Integer
'    Dim i       As Integer
'
'    For i = 1 To Len(Data) Step 4
'        szOut = szOut & Chr((Asc(Mid(Data, i, 1)) - 32) * 4 + (Asc(Mid(Data, i + 1, 1)) - 32) \ 16)
'        szOut = szOut & Chr((Asc(Mid(Data, i + 1, 1)) Mod 16) * 16 + (Asc(Mid(Data, i + 2, 1)) - 32) \ 4)
'        szOut = szOut & Chr((Asc(Mid(Data, i + 2, 1)) Mod 4) * 64 + Asc(Mid(Data, i + 3, 1)) - 32)
'    Next i
'
'    UUDecode = szOut
'
'    Exit Function
'ErrorHandler:
'    UUDecode = ""
'End Function
'
'Public Function UUEncode(Data As String) As String
'    '-----------
'    ' UU Encode string (data) for transmission.
'    '-----------
'    Dim szOut   As String
'    Dim nChar   As Integer
'    Dim i       As Integer
'
'    '   pad to 3 byte multiple
'    If Len(Data) Mod 3 <> 0 Then Data = Data & Space(3 - Len(Data) Mod 3)
'
'    For i = 1 To Len(szData) Step 3
'        szOut = szOut & Chr(Asc(Mid(Data, i, 1)) \ 4 + 32)
'        szOut = szOut & Chr((Asc(Mid(Data, i, 1)) Mod 4) * 16 + Asc(Mid(Data, i + 1, 1)) \ 16 + 32)
'        szOut = szOut & Chr((Asc(Mid(Data, i + 1, 1)) Mod 16) * 4 + Asc(Mid(Data, i + 2, 1)) \ 64 + 32)
'        szOut = szOut & Chr(Asc(Mid(Data, i + 2, 1)) Mod 64 + 32)
'    Next i
'
'    Encode = szOut
'End Function

'-------------------------------------------------
'      END UUENCODE FUNCTIONS
'-------------------------------------------------


'-------------------------------------------------
'      BEGIN BASE64 FUNCTIONS
'-------------------------------------------------
'Public Function Base64Encode(ByVal vsFullPathname As String) As String
'    Dim b           As Integer
'    Dim Base64Tab   As Variant
'    Dim bin(3)      As Byte
'    Dim s           As String
'    Dim l           As Long
'    Dim i           As Long
'    Dim FileIn      As Long
'    Dim sResult     As String
'    Dim n           As Long
'
'    Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
'
'    Erase bin
'    l = 0: i = 0: FileIn = 0: b = 0:
'    s = ""
'
'    FileIn = FreeFile
'    Open vsFullPathname For Binary As FileIn
'
'    sResult = s & vbCrLf
'    s = ""
'
'    l = LOF(FileIn) - (LOF(FileIn) Mod 3)
'
'    For i = 1 To l Step 3
'
'        'Read three bytes
'        Get FileIn, , bin(0)
'        Get FileIn, , bin(1)
'        Get FileIn, , bin(2)
'
'        'Always wait until there're more then 64 characters
'        If Len(s) > 64 Then
'
'            s = s & vbCrLf
'            sResult = sResult & s
'            s = ""
'
'        End If
'
'        'Calc Base64-encoded char
'        b = (bin(n) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
'        s = s & Base64Tab(b) 'the character s holds the encoded chars
'
'        b = ((bin(n) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
'        s = s & Base64Tab(b)
'
'        b = ((bin(n + 1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
'        s = s & Base64Tab(b)
'
'        b = bin(n + 2) And &H3F
'        s = s & Base64Tab(b)
'
'    Next i
'
'    'Now, you need to check if there is something left
'    If Not (LOF(FileIn) Mod 3 = 0) Then
'
'        'Reads the number of bytes left
'        For i = 1 To (LOF(FileIn) Mod 3)
'            Get FileIn, , bin(i - 1)
'        Next i
'
'        'If there are only 2 chars left
'        If (LOF(FileIn) Mod 3) = 2 Then
'            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
'            s = s & Base64Tab(b)
'
'            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
'            s = s & Base64Tab(b)
'
'            b = ((bin(1) And &HF) * 4) Or ((bin(2) \ 64) And &H3)
'            s = s & Base64Tab(b)
'
'            s = s & "="
'
'        Else 'If there is only one char left
'            b = (bin(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
'            s = s & Base64Tab(b)
'
'            b = ((bin(0) And &H3) * 16) Or ((bin(1) \ 16) And &HF)
'            s = s & Base64Tab(b)
'
'            s = s & "=="
'        End If
'    End If
'
'    'Send the characters left
'    If s <> "" Then
'        s = s & vbCrLf
'        sResult = sResult & s
'    End If
'
'    'Send the last part of the MIME Body
'    s = ""
'
'    Close FileIn
'    EncodeBase64 = sResult
'
'End Function
'
''-------------------------------------------------
''      END BASE64 FUNCTIONS
''-------------------------------------------------
