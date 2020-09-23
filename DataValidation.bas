Attribute VB_Name = "DataValidation"

Public Function KeyNumbers(ByVal KeyAscii As Integer, Optional ByVal Positive As Boolean) As Integer
    Dim strValid As String
    
    strValid = "0123456789." & vbBack & vbKeyDelete
    If Positive = False Then strValid = strValid & "-"
    
    KeyNumbers = ValidateASCII(KeyAscii, strValid)
End Function

Public Function KeyAlphaNum(ByVal KeyAscii As Integer) As Integer
    Dim strValid As String
    Const strAlpha = " abcdefghijklmnopqrstuvwxyz"
    
    strValid = "0123456789.-" & vbBack & vbKeyDelete & strAlpha & UCase(strAlpha)
    
    KeyAlphaNum = ValidateASCII(KeyAscii, strValid)
End Function

Public Function ValidateASCII(ByVal KeyAscii As Integer, ByVal ValidKeys As String) As Integer
    If InStr(1, ValidKeys, Chr(KeyAscii)) > 0 Then
        ValidateASCII = KeyAscii
    Else
        ValidateASCII = 0
    End If
End Function

Public Function CheckForUnsavedData() As Boolean
    Dim tmp As Boolean
    
    tmp = CheckURLs
        If tmp = False Then Exit Function
    tmp = CheckEmail
        If tmp = False Then Exit Function
    
    CheckForUnsavedData = True
End Function


Public Function CheckURLs() As Boolean
    'All URL fields that have text in them must start with: http:// or https://
    'In addition ONLY download URL fields should additionally allow ftp:// as a prefix.
    Dim tmp As Boolean
    
    tmp = isURL(frmMain.txtXMLfileURL, "http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtIconFileURL, "http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtScreenshotURL, "http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtOrderURL, "http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtInfoURL, "http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtDownloadURL(0), "ftp://,http://,https://", 1, False)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtDownloadURL(1), "ftp://,http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtDownloadURL(2), "ftp://,http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtDownloadURL(3), "ftp://,http://,https://", 1, True)
        If tmp = False Then Exit Function
    tmp = isURL(frmMain.txtCompanyWebsiteURL, "http://,https://", 2, True)
        If tmp = False Then Exit Function
    
    CheckURLs = True
End Function

Public Function isURL(ByRef URL As TextBox, ByVal Protocols As String, ByVal SSTabNumber As Integer, ByVal EmptyAllowed As Boolean)
    Dim eProtocols() As String
    Dim tmpURL As String
    Dim ProtocolFound As Boolean
    
    If URL.Text = "" And EmptyAllowed = False Then
        URL.Text = "http://"
        frmMain.SSTab1.Tab = SSTabNumber
        URL.SetFocus
        URL.SelStart = 0
        URL.SelLength = Len(URL.Text)
        MsgBox "URL must have a value."
        Exit Function
    End If
    
    eProtocols = Split(UCase(Protocols), ",")
    tmpURL = UCase(Trim(URL.Text))
    For X = 0 To UBound(eProtocols)
        If Len(tmpURL) >= Len(eProtocols(X)) Then
            If Left(tmpURL, Len(eProtocols(X))) = eProtocols(X) Then ProtocolFound = True
        End If
    Next X
    
    If URL.Text = "" And EmptyAllowed = True Then ProtocolFound = True
    
    If ProtocolFound = False Then
        frmMain.SSTab1.Tab = SSTabNumber
        URL.SetFocus
        URL.SelStart = 0
        URL.SelLength = Len(URL.Text)
        MsgBox "'" & URL.Text & "' is not a valid URL."
    End If
    
    isURL = ProtocolFound
End Function

Public Function CheckEmail() As Boolean
    'All email address fields should be verified to be sure they contain at least
    'an '@' and '.' characters.
    Dim tmp As Boolean
    
    tmp = isEmailAddress(frmMain.txtGeneralEmail, 2, True)
        If tmp = False Then Exit Function
    tmp = isEmailAddress(frmMain.txtGeneralEmail, 2, True)
        If tmp = False Then Exit Function
    tmp = isEmailAddress(frmMain.txtSalesEmail, 2, True)
        If tmp = False Then Exit Function
    tmp = isEmailAddress(frmMain.txtSupportEmail, 2, True)
        If tmp = False Then Exit Function
    tmp = isEmailAddress(frmMain.txtAuthorEmail, 3, True)
        If tmp = False Then Exit Function
    tmp = isEmailAddress(frmMain.txtContactEMail, 3, True)
        If tmp = False Then Exit Function

    CheckEmail = True
End Function

Public Function isEmailAddress(ByRef EmailAddress As TextBox, ByVal SSTabNumber As Integer, ByVal EmptyAllowed As Boolean)
    Dim eProtocols() As String
    Dim tmpEmailAddress As String
    Dim tmp1 As Integer
    Dim tmp2 As Integer
    
    tmpEmailAddress = EmailAddress.Text
    tmp1 = InStr(1, tmpEmailAddress, "@")
    tmp2 = InStr(tmp1 + 1, tmpEmailAddress, ".")
        
    If tmpEmailAddress = "" Then
        If EmptyAllowed = False Then
            EmailAddress.Text = "???"
            frmMain.SSTab1.Tab = SSTabNumber
            EmailAddress.SetFocus
            EmailAddress.SelStart = 0
            EmailAddress.SelLength = Len(EmailAddress.Text)
            MsgBox "Email address must have a value."
            Exit Function
        Else
            tmp1 = 1
            tmp2 = 1
        End If
    End If
    
    If tmp1 = 0 Or tmp2 = 0 Then
        frmMain.SSTab1.Tab = SSTabNumber
        EmailAddress.SetFocus
        EmailAddress.SelStart = 0
        EmailAddress.SelLength = Len(EmailAddress.Text)
        MsgBox "'" & EmailAddress.Text & "' is not a valid email address."
        isEmailAddress = False
    Else
        isEmailAddress = True
    End If
End Function

'Public Function XMLClean(ByVal inText As String) As String
'    Dim tmp As String
'    Dim XX As String * 1
'
'    For X = 1 To Len(inText)
'        XX = Mid$(inText, X, 1)
'        Select Case XX
'            Case "&"
'                tmp = tmp & "&#38;"
'            Case "<"
'                tmp = tmp & "&#60;"
'            Case Else
'                tmp = tmp & XX$
'        End Select
'    Next X
'End Function
