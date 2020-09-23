Attribute VB_Name = "curValues"
Public curAddress1  As String
Public curAddress2  As String
Public curAuthor  As String
Public curAuthorEmail  As String
Public curAuthorFName  As String
Public curAuthorLName  As String
Public curBlurbName  As String
Public curBlurbs  As String
Public curBlurbText  As String
Public curBrowserURL  As String
Public curCategories  As String
Public curChangeInfo  As String
Public curCityTown  As String
Public curCompany  As String
Public curCompanyAbout  As String
Public curCompanyName  As String
Public curCompanyWebsiteURL  As String
Public curContactEMail  As String
Public curContactFName  As String
Public curContactLName  As String
Public curContactPhone  As String
Public curCountry  As String
Public curDescription(4)  As String
Public curDistributionIncludes  As String
Public curDistributionPermisions  As String
Public curDownloadURL(3)  As String
Public curEmailAddress  As String
Public curEULA  As String
Public curFaxPhone  As String
Public curFilenameGeneric  As String
Public curFilenameLong  As String
Public curFilenamePrevious  As String
Public curFilenameVersioned  As String
Public curFileSize  As String
Public curGeneralEmail  As String
Public curIconFileURL  As String
Public curInfoURL  As String
Public curInstallSupport  As String
Public curKeywords  As String
Public curLanguage  As String
Public curOrderURL  As String
Public curOSSupport  As String
Public curPermissionName  As String
Public curPermissions  As String
Public curProductName  As String
Public curProgramType  As String
Public curProgramVersion  As String
Public curRegistrationCostOther  As String
Public curRegistrationCostUSD  As String
Public curReleaseDate  As String
Public curReleaseStatus  As String
Public curSalesEmail  As String
Public curSalesPhone  As String
Public curScreenshotURL  As String
Public curSiteList  As String
Public curSpecificCategory  As String
Public curStateProvince  As String
Public curSupportEmail  As String
Public curSupportPhone  As String
Public curSystemRequirements  As String
Public curXMLfileURL  As String
Public curZIPPostal  As String
Public curProgramExpires As Integer
Public curExpirationCount  As String
Public curExpireBasedOn  As String
Public curExpirationDate  As String
Public curExpirationOtherInfo  As String
Public curOrganizations As String
Public curCompanyLogoURL As String

Type FileTuple
    key As String
    Value As String
    File As String
End Type

Public DataPath As String
Public HelpPath As String
Public BlurbOrder As String
Public TOC As String
Public Articles As String


Public Sub SaveProgram(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameS(FileName, ".program")
    If fFilename = "" Then Exit Sub
    
    curProductName = frmMain.txtProductName
    curProgramVersion = frmMain.txtProgramVersion
    curReleaseStatus = frmMain.txtReleaseStatus
    curCompany = frmMain.txtCompany
    curAuthor = frmMain.txtAuthor
    curProgramType = frmMain.txtProgramType
    curPermissions = frmMain.txtPermissions
    curBlurbs = frmMain.txtBlurbs
    curOSSupport = frmMain.txtOSSupport
    curChangeInfo = StripCRLF(frmMain.txtChangeInfo)
    curReleaseDate = frmMain.txtReleaseDate
    curRegistrationCostUSD = frmMain.txtRegistrationCostUSD
    curRegistrationCostOther = frmMain.txtRegistrationCostOther
    curLanguage = frmMain.txtLanguage
    curDistributionIncludes = frmMain.txtDistributionIncludes
    curInstallSupport = frmMain.txtInstallSupport
    curSystemRequirements = StripCRLF(frmMain.txtSystemRequirements)
    curFilenameVersioned = frmMain.txtFilenameVersioned
    curFilenamePrevious = frmMain.txtFilenamePrevious
    curFilenameGeneric = frmMain.txtFilenameGeneric
    curFilenameLong = frmMain.txtFilenameLong
    curFileSize = frmMain.txtFileSize
    curXMLfileURL = frmMain.txtXMLfileURL
    curIconFileURL = frmMain.txtIconFileURL
    curScreenshotURL = frmMain.txtScreenshotURL
    curOrderURL = frmMain.txtOrderURL
    curInfoURL = frmMain.txtInfoURL
    curCategories = frmMain.txtCategories
    curDownloadURL(0) = frmMain.txtDownloadURL(0)
    curDownloadURL(1) = frmMain.txtDownloadURL(1)
    curDownloadURL(2) = frmMain.txtDownloadURL(2)
    curDownloadURL(3) = frmMain.txtDownloadURL(3)
    curKeywords = frmMain.txtKeywords
    curSpecificCategory = frmMain.txtSpecificCategory
    curDescription(0) = StripCRLF(frmMain.txtDescription(0))
    curDescription(1) = StripCRLF(frmMain.txtDescription(1))
    curDescription(2) = StripCRLF(frmMain.txtDescription(2))
    curDescription(3) = StripCRLF(frmMain.txtDescription(3))
    curDescription(4) = StripCRLF(frmMain.txtDescription(4))
    curProgramExpires = frmMain.chkProgramExpires.Value
    
    Open fFilename For Output As #FileNum
        Print #FileNum, curProductName
        Print #FileNum, curProgramVersion
        Print #FileNum, curReleaseStatus
        Print #FileNum, curCompany
        Print #FileNum, curAuthor
        Print #FileNum, curProgramType
        Print #FileNum, curPermissions
        Print #FileNum, curBlurbs
        Print #FileNum, curOSSupport
        Print #FileNum, curChangeInfo
        Print #FileNum, curReleaseDate
        Print #FileNum, curRegistrationCostUSD
        Print #FileNum, curRegistrationCostOther
        Print #FileNum, curLanguage
        Print #FileNum, curDistributionIncludes
        Print #FileNum, curInstallSupport
        Print #FileNum, curSystemRequirements
        Print #FileNum, curFilenameVersioned
        Print #FileNum, curFilenamePrevious
        Print #FileNum, curFilenameGeneric
        Print #FileNum, curFilenameLong
        Print #FileNum, curFileSize
        Print #FileNum, curXMLfileURL
        Print #FileNum, curIconFileURL
        Print #FileNum, curScreenshotURL
        Print #FileNum, curOrderURL
        Print #FileNum, curInfoURL
        Print #FileNum, curCategories
        Print #FileNum, curDownloadURL(0)
        Print #FileNum, curDownloadURL(1)
        Print #FileNum, curDownloadURL(2)
        Print #FileNum, curDownloadURL(3)
        Print #FileNum, curKeywords
        Print #FileNum, curSpecificCategory
        Print #FileNum, curDescription(0)
        Print #FileNum, curDescription(1)
        Print #FileNum, curDescription(2)
        Print #FileNum, curDescription(3)
        Print #FileNum, curDescription(4)
        Print #FileNum, curProgramExpires
        Print #FileNum, curExpirationCount
        Print #FileNum, curExpireBasedOn
        Print #FileNum, curExpirationDate
        Print #FileNum, curExpirationOtherInfo
    Close #FileNum
    Status "Program '" & FileName & "' saved."
    ProgramModified = False
End Sub

Public Sub LoadProgram(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameL(FileName, ".program")
    
    If fFilename = "" Then Exit Sub
    Open fFilename For Input As #FileNum
        Line Input #FileNum, curProductName
        Line Input #FileNum, curProgramVersion
        Line Input #FileNum, curReleaseStatus
        Line Input #FileNum, curCompany
        Line Input #FileNum, curAuthor
        Line Input #FileNum, curProgramType
        Line Input #FileNum, curPermissions
        Line Input #FileNum, curBlurbs
        Line Input #FileNum, curOSSupport
        Line Input #FileNum, tmp
            curChangeInfo = StripCRLF(tmp)
        Line Input #FileNum, curReleaseDate
        Line Input #FileNum, curRegistrationCostUSD
        Line Input #FileNum, curRegistrationCostOther
        Line Input #FileNum, curLanguage
        Line Input #FileNum, curDistributionIncludes
        Line Input #FileNum, curInstallSupport
        Line Input #FileNum, tmp
            curSystemRequirements = StripCRLF(tmp)
        Line Input #FileNum, curFilenameVersioned
        Line Input #FileNum, curFilenamePrevious
        Line Input #FileNum, curFilenameGeneric
        Line Input #FileNum, curFilenameLong
        Line Input #FileNum, curFileSize
        Line Input #FileNum, curXMLfileURL
        Line Input #FileNum, curIconFileURL
        Line Input #FileNum, curScreenshotURL
        Line Input #FileNum, curOrderURL
        Line Input #FileNum, curInfoURL
        Line Input #FileNum, curCategories
        Line Input #FileNum, curDownloadURL(0)
        Line Input #FileNum, curDownloadURL(1)
        Line Input #FileNum, curDownloadURL(2)
        Line Input #FileNum, curDownloadURL(3)
        Line Input #FileNum, curKeywords
        Line Input #FileNum, curSpecificCategory
        Line Input #FileNum, tmp
            curDescription(0) = StripCRLF(tmp)
        Line Input #FileNum, tmp
            curDescription(1) = StripCRLF(tmp)
        Line Input #FileNum, tmp
            curDescription(2) = StripCRLF(tmp)
        Line Input #FileNum, tmp
            curDescription(3) = StripCRLF(tmp)
        Line Input #FileNum, tmp
            curDescription(4) = StripCRLF(tmp)
        Line Input #FileNum, tmp
            curProgramExpires = CInt(tmp)
        Line Input #FileNum, curExpirationCount
        Line Input #FileNum, curExpireBasedOn
        Line Input #FileNum, curExpirationDate
        Line Input #FileNum, curExpirationOtherInfo
    Close #FileNum
    
    frmMain.txtProductName = curProductName
    frmMain.txtProgramVersion = curProgramVersion
    frmMain.txtReleaseStatus = curReleaseStatus
    frmMain.txtCompany = curCompany
    frmMain.txtAuthor = curAuthor
    frmMain.txtProgramType = curProgramType
    frmMain.txtPermissions = curPermissions
    frmMain.txtBlurbs = curBlurbs
    frmMain.txtOSSupport = curOSSupport
    frmMain.txtChangeInfo = PadCRLF(curChangeInfo)
    frmMain.txtReleaseDate = curReleaseDate
    frmMain.txtRegistrationCostUSD = curRegistrationCostUSD
    frmMain.txtRegistrationCostOther = curRegistrationCostOther
    frmMain.txtLanguage = curLanguage
    frmMain.txtDistributionIncludes = curDistributionIncludes
    frmMain.txtInstallSupport = curInstallSupport
    frmMain.txtSystemRequirements = PadCRLF(curSystemRequirements)
    frmMain.txtFilenameVersioned = curFilenameVersioned
    frmMain.txtFilenamePrevious = curFilenamePrevious
    frmMain.txtFilenameGeneric = curFilenameGeneric
    frmMain.txtFilenameLong = curFilenameLong
    frmMain.txtFileSize = curFileSize
    frmMain.txtXMLfileURL = curXMLfileURL
    frmMain.txtIconFileURL = curIconFileURL
    frmMain.txtScreenshotURL = curScreenshotURL
    frmMain.txtOrderURL = curOrderURL
    frmMain.txtInfoURL = curInfoURL
    frmMain.txtCategories = curCategories
    frmMain.txtDownloadURL(0) = curDownloadURL(0)
    frmMain.txtDownloadURL(1) = curDownloadURL(1)
    frmMain.txtDownloadURL(2) = curDownloadURL(2)
    frmMain.txtDownloadURL(3) = curDownloadURL(3)
    frmMain.txtKeywords = curKeywords
    frmMain.txtSpecificCategory = curSpecificCategory
    frmMain.txtDescription(0) = PadCRLF(curDescription(0))
    frmMain.txtDescription(1) = PadCRLF(curDescription(1))
    frmMain.txtDescription(2) = PadCRLF(curDescription(2))
    frmMain.txtDescription(3) = PadCRLF(curDescription(3))
    frmMain.txtDescription(4) = PadCRLF(curDescription(4))
    frmExpireInfo.Visible = False
    frmMain.chkProgramExpires.Value = curProgramExpires
    Unload frmExpireInfo
    frmMain.chkProgramExpires.Value = curProgramExpires
    
    LoadCompany curCompany
    LoadAuthor curAuthor
    LoadPermission curPermissions
    
    Status "Program '" & FileName & "' loaded."
    ProgramModified = False
End Sub

Public Sub SaveCompany(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameS(FileName, ".company")
    If fFilename = "" Then Exit Sub
    
    curCompanyName = frmMain.txtCompanyName
    curAddress1 = frmMain.txtAddress1
    curAddress2 = frmMain.txtAddress2
    curCompanyAbout = StripCRLF(frmMain.txtCompanyAbout)
    curFaxPhone = frmMain.txtFaxPhone
    curSalesEmail = frmMain.txtSalesEmail
    curSalesPhone = frmMain.txtSalesPhone
    curGeneralEmail = frmMain.txtGeneralEmail
    curContactPhone = frmMain.txtContactPhone
    curSupportEmail = frmMain.txtSupportEmail
    curSupportPhone = frmMain.txtSupportPhone
    curCityTown = frmMain.txtCityTown
    curStateProvince = frmMain.txtStateProvince
    curZIPPostal = frmMain.txtZIPPostal
    curCountry = frmMain.txtCountry
    curCompanyWebsiteURL = frmMain.txtCompanyWebsiteURL
    curOrganizations = frmMain.txtOrganizations
    curCompanyLogoURL = frmMain.txtCompanyLogoURL
    
    Open fFilename For Output As #FileNum
        Print #FileNum, curCompanyName
        Print #FileNum, curAddress1
        Print #FileNum, curAddress2
        Print #FileNum, StripCRLF(curCompanyAbout)
        Print #FileNum, curFaxPhone
        Print #FileNum, curSalesEmail
        Print #FileNum, curSalesPhone
        Print #FileNum, curGeneralEmail
        Print #FileNum, curContactPhone
        Print #FileNum, curSupportEmail
        Print #FileNum, curSupportPhone
        Print #FileNum, curCityTown
        Print #FileNum, curStateProvince
        Print #FileNum, curZIPPostal
        Print #FileNum, curCountry
        Print #FileNum, curCompanyWebsiteURL
        Print #FileNum, curOrganizations
        Print #FileNum, curCompanyLogoURL
    Close #FileNum
    
    Status "Company '" & FileName & "' saved."
    CompanyModified = False
End Sub

Public Sub LoadCompany(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameL(FileName, ".company")
    If fFilename = "" Then Exit Sub
    Open fFilename For Input As #FileNum
        Line Input #FileNum, curCompanyName
        Line Input #FileNum, curAddress1
        Line Input #FileNum, curAddress2
        Line Input #FileNum, tmp
            curCompanyAbout = StripCRLF(tmp)
        Line Input #FileNum, curFaxPhone
        Line Input #FileNum, curSalesEmail
        Line Input #FileNum, curSalesPhone
        Line Input #FileNum, curGeneralEmail
        Line Input #FileNum, curContactPhone
        Line Input #FileNum, curSupportEmail
        Line Input #FileNum, curSupportPhone
        Line Input #FileNum, curCityTown
        Line Input #FileNum, curStateProvince
        Line Input #FileNum, curZIPPostal
        Line Input #FileNum, curCountry
        Line Input #FileNum, curCompanyWebsiteURL
        Line Input #FileNum, curOrganizations
        Line Input #FileNum, curCompanyLogoURL
    Close #FileNum
    
    frmMain.txtCompanyName = curCompanyName
    frmMain.txtAddress1 = curAddress1
    frmMain.txtAddress2 = curAddress2
    frmMain.txtCompanyAbout = PadCRLF(curCompanyAbout)
    frmMain.txtFaxPhone = curFaxPhone
    frmMain.txtSalesEmail = curSalesEmail
    frmMain.txtSalesPhone = curSalesPhone
    frmMain.txtGeneralEmail = curGeneralEmail
    frmMain.txtContactPhone = curContactPhone
    frmMain.txtSupportEmail = curSupportEmail
    frmMain.txtSupportPhone = curSupportPhone
    frmMain.txtCityTown = curCityTown
    frmMain.txtStateProvince = curStateProvince
    frmMain.txtZIPPostal = curZIPPostal
    frmMain.txtCountry = curCountry
    frmMain.txtCompanyWebsiteURL = curCompanyWebsiteURL
    frmMain.txtOrganizations = curOrganizations
    frmMain.txtCompanyLogoURL = curCompanyLogoURL
    
    Status "Company '" & FileName & "' loaded."
    CompanyModified = False
End Sub

Public Sub SaveAuthor(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameS(FileName, ".author")
    If fFilename = "" Then Exit Sub
    
    curAuthorEmail = frmMain.txtAuthorEmail
    curAuthorFName = frmMain.txtAuthorFName
    curAuthorLName = frmMain.txtAuthorLName
    curContactEMail = frmMain.txtContactEMail
    curContactFName = frmMain.txtContactFName
    curContactLName = frmMain.txtContactLName
    
    Open fFilename For Output As #FileNum
        Print #FileNum, curAuthorEmail
        Print #FileNum, curAuthorFName
        Print #FileNum, curAuthorLName
        Print #FileNum, curContactEMail
        Print #FileNum, curContactFName
        Print #FileNum, curContactLName
    Close #FileNum
    Status "Author '" & FileName & "' saved."
    AuthorModified = False
End Sub

Public Sub LoadAuthor(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameL(FileName, ".author")
    If fFilename = "" Then Exit Sub
    Open fFilename For Input As #FileNum
        Line Input #FileNum, curAuthorEmail
        Line Input #FileNum, curAuthorFName
        Line Input #FileNum, curAuthorLName
        Line Input #FileNum, curContactEMail
        Line Input #FileNum, curContactFName
        Line Input #FileNum, curContactLName
    Close #FileNum
    
    frmMain.txtAuthorEmail = curAuthorEmail
    frmMain.txtAuthorFName = curAuthorFName
    frmMain.txtAuthorLName = curAuthorLName
    frmMain.txtContactEMail = curContactEMail
    frmMain.txtContactFName = curContactFName
    frmMain.txtContactLName = curContactLName
    
    Status "Author '" & FileName & "' loaded."
    AuthorModified = False
End Sub

Public Sub SaveBlurb(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameS(FileName, ".blurb")
    If fFilename = "" Then Exit Sub
    
    curBlurbName = frmMain.txtBlurbName
    curBlurbText = frmMain.txtBlurbText
    
    Open fFilename For Output As #FileNum
        Print #FileNum, curBlurbName
        Print #FileNum, StripCRLF(curBlurbText)
    Close #FileNum
    Status "Blurb '" & FileName & "' saved."
    BlurbModified = False
End Sub

Public Sub LoadBlurb(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameL(FileName, ".blurb")
    If fFilename = "" Then Exit Sub
    Open fFilename For Input As #FileNum
        Line Input #FileNum, curBlurbName
        Line Input #FileNum, tmp
            curBlurbText = PadCRLF(tmp)
    Close #FileNum
    
    frmMain.txtBlurbName = curBlurbName
    frmMain.txtBlurbText = PadCRLF(curBlurbText)
    
    Status "Blurb '" & FileName & "' loaded."
    BlurbModified = False
End Sub

Public Sub SavePermission(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameS(FileName, ".permission")
    If fFilename = "" Then Exit Sub
    
    curPermissionName = frmMain.txtPermissionName
    curDistributionPermisions = frmMain.txtDistributionPermisions
    curEULA = frmMain.txtEULA
    
    Open fFilename For Output As #FileNum
        Print #FileNum, curPermissionName
        Print #FileNum, curDistributionPermisions
        Print #FileNum, StripCRLF(curEULA)
    Close #FileNum
    Status "Permission '" & FileName & "' saved."
    PermissionsModified = False
End Sub

Public Sub LoadPermission(ByVal FileName As String)
    Dim fFilename As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    fFilename = GetFilenameL(FileName, ".permission")
    If fFilename = "" Then Exit Sub
    Open fFilename For Input As #FileNum
        Line Input #FileNum, curPermissionName
        Line Input #FileNum, curDistributionPermisions
        Line Input #FileNum, tmp
            curEULA = PadCRLF(tmp)
    Close #FileNum
    
    frmMain.txtPermissionName = curPermissionName
    frmMain.txtDistributionPermisions = curDistributionPermisions
    frmMain.txtEULA = PadCRLF(curEULA)
    
    Status "Permission '" & FileName & "' loaded."
    PermissionsModified = False
End Sub

Public Function StripCRLF(ByVal sVal As String) As String
    '------
    'strip crlf from sval to save as one liner
    '------
    del = vbCrLf
    Rep = "}c{"
    dlen = Len(del)
    
    X = InStr(1, sVal, del)
    Do Until X = 0
        sVal = Left(sVal, X - 1) & Rep & Right(sVal, Len(sVal) - X - dlen + 1)
        X = InStr(1, sVal, del)
    Loop
    
    StripCRLF = sVal
End Function


Public Function PadCRLF(ByVal sVal As String) As String
    '------
    'put crlf into sval to load as one liner
    '------
    del = "}c{"
    Rep = vbCrLf
    dlen = Len(del)
    
    X = InStr(1, sVal, del)
    Do Until X = 0
        sVal = Left(sVal, X - 1) & Rep & Right(sVal, Len(sVal) - X - dlen + 1)
        X = InStr(1, sVal, del)
    Loop
    
    PadCRLF = sVal
End Function

Public Function GetFilenameL(ByVal FileName As String, ByVal Extension As String) As String
    Dim fFilename As String
    
    If Trim(FileName) = "" Then Exit Function
    If InStr(1, FileName, "\") = 0 Then fFilename = DataPath
    fFilename = fFilename & FileName & Extension
    
    If Dir(fFilename) = "" Then
        Status "'" & FileName & "' not found."
        Exit Function
    End If
    GetFilenameL = fFilename
End Function

Public Function GetFilenameS(ByVal FileName As String, ByVal Extension As String) As String
    Dim fFilename As String
    
    If Trim(FileName) = "" Then Exit Function
    If InStr(1, FileName, "\") = 0 Then fFilename = DataPath
    fFilename = fFilename & FileName & Extension
    If Dir(fFilename) > "" Then
        X = MsgBox("Overwrite " & FileName & "?", vbYesNo, UCase(Right(Extension, Len(Extension) - 1)))
        If X = vbNo Then
            Status ""
            Exit Function
        End If
    End If

    GetFilenameS = fFilename
End Function

Public Sub DeleteDataFile(ByVal FileName As String, ByVal Extension As String)
    Dim fFilename As String
    
    If Trim(FileName) = "" Then Exit Sub
    fFilename = DataPath & FileName & Extension
    If Dir(fFilename) > "" Then
        X = MsgBox("Delete " & FileName & "?", vbYesNo)
        If X = vbNo Then
            Status ""
            Exit Sub
        End If
    Else
        Status "File not found."
        Exit Sub
    End If
    Kill fFilename
    Status "'" & FileName & "' deleted."
End Sub

Public Sub SetPageTuples()
    Dim key As String
    
'    ReDim PageTuple(0)
    InitPageTuples
    
    key = "TOC": AddPageTuple key, TOC
    key = "ARTICLES": AddPageTuple key, Articles

    key = "ADDRESS_1": AddPageTuple key, curAddress1
    key = "ADDRESS_2": AddPageTuple key, curAddress2
    key = "AUTHOR": AddPageTuple key, curAuthor
    key = "AUTHOR_EMAIL": AddPageTuple key, curAuthorEmail
    key = "AUTHOR_F_NAME": AddPageTuple key, curAuthorFName
    key = "AUTHOR_L_NAME": AddPageTuple key, curAuthorLName
'    key = "BLURB_NAME": AddPageTuple key, curBlurbName
    key = "BLURBS": AddPageTuple key, curBlurbs
'    key = "BLURB_TEXT": AddPageTuple key, curBlurbText
'    key = "BROWSER_URL": AddPageTuple key, curBrowserURL
    key = "CATEGORIES": AddPageTuple key, curCategories
    key = "CHANGE_INFO": AddPageTuple key, PadCRLF(curChangeInfo)
    key = "CITY_TOWN": AddPageTuple key, curCityTown
    key = "COMPANY": AddPageTuple key, curCompany
    key = "COMPANY_ABOUT": AddPageTuple key, curCompanyAbout
    key = "COMPANY_NAME": AddPageTuple key, curCompanyName
    key = "COMPANY_WEBSITE_URL": AddPageTuple key, curCompanyWebsiteURL
    key = "COMPANY_LOGO_URL": AddPageTuple key, curCompanyLogoURL
    key = "CONTACT_EMAIL": AddPageTuple key, curContactEMail
    key = "COMPANT_CONTACT": AddPageTuple key, curContactFName & " " & curContactLName
    key = "CONTACT_F_NAME": AddPageTuple key, curContactFName
    key = "CONTACT_L_NAME": AddPageTuple key, curContactLName
    key = "CONTACT_PHONE": AddPageTuple key, curContactPhone
    key = "COUNTRY": AddPageTuple key, curCountry
    key = "ORGANIZATION_MEMBERSHIP": AddPageTuple key, curOrganizations
    key = "DESCRIPTION_45": AddPageTuple key, PadCRLF(curDescription(0))
    key = "DESCRIPTION_80": AddPageTuple key, PadCRLF(curDescription(1))
    key = "DESCRIPTION_250": AddPageTuple key, PadCRLF(curDescription(2))
    key = "DESCRIPTION_450": AddPageTuple key, PadCRLF(curDescription(3))
    key = "DESCRIPTION_2000": AddPageTuple key, PadCRLF(curDescription(4))
    key = "DISTRIBUTION_INCLUDES": AddPageTuple key, curDistributionIncludes
    key = "INCLUDES_VB": AddPageTuple key, IIf(InStr(1, UCase(curDistributionIncludes), "VB RUNTIME") > 0, "Y", "N")
    key = "INCLUDES_JAVA": AddPageTuple key, IIf(InStr(1, UCase(curDistributionIncludes), "JAVA") > 0, "Y", "N")
    key = "INCLUDES_DIRECTX": AddPageTuple key, IIf(InStr(1, UCase(curDistributionIncludes), "DIRECTX") > 0, "Y", "N")
    key = "DISTRIBUTION_PERMISSIONS": AddPageTuple key, curDistributionPermisions
    key = "DOWNLOAD_URL_1": AddPageTuple key, curDownloadURL(0)
    key = "DOWNLOAD_URL_2": AddPageTuple key, curDownloadURL(1)
    key = "DOWNLOAD_URL_3": AddPageTuple key, curDownloadURL(2)
    key = "DOWNLOAD_URL_4": AddPageTuple key, curDownloadURL(3)
'    key = "EMAIL_ADDRESS": AddPageTuple key, curEmailAddress
    key = "EULA": AddPageTuple key, curEULA
    key = "FAX_PHONE": AddPageTuple key, curFaxPhone
    key = "FILENAME_GENERIC": AddPageTuple key, curFilenameGeneric
    key = "FILENAME_LONG": AddPageTuple key, curFilenameLong
    key = "FILENAME_PREVIOUS": AddPageTuple key, curFilenamePrevious
    key = "FILENAME_VERSIONED": AddPageTuple key, curFilenameVersioned
    key = "FILESIZE_BYTES": AddPageTuple key, Val(curFileSize)
    key = "FILESIZE_KB": AddPageTuple key, Int(Val(curFileSize) / 1024)
    key = "FILESIZE_MB": AddPageTuple key, (Int((Val(curFileSize) * 10) / 1024 / 1024)) / 10
    
    key = "DOWNLOAD_TIME_14K": AddPageTuple key, DCalc(Val(curFileSize), 14)
    key = "DOWNLOAD_TIME_28K": AddPageTuple key, DCalc(Val(curFileSize), 28)
    key = "DOWNLOAD_TIME_56K": AddPageTuple key, DCalc(Val(curFileSize), 56)
    key = "DOWNLOAD_TIME_64K": AddPageTuple key, DCalc(Val(curFileSize), 64)
    key = "DOWNLOAD_TIME_128K": AddPageTuple key, DCalc(Val(curFileSize), 128)
    key = "DOWNLOAD_TIME_T1": AddPageTuple key, DCalc(Val(curFileSize), 1544)
    
    key = "GENERAL_EMAIL": AddPageTuple key, curGeneralEmail
    key = "ICON_FILE_URL": AddPageTuple key, curIconFileURL
    key = "INFO_URL": AddPageTuple key, curInfoURL
    key = "INSTALL_SUPPORT": AddPageTuple key, curInstallSupport
    key = "KEYWORDS": AddPageTuple key, curKeywords
    key = "LANGUAGE": AddPageTuple key, curLanguage
    key = "ORDER_URL": AddPageTuple key, curOrderURL
    key = "OS_SUPPORT": AddPageTuple key, curOSSupport
'    key = "PERMISSION_NAME": AddPageTuple key, curPermissionName
    key = "PERMISSIONS": AddPageTuple key, curPermissions
    key = "PROGRAM_NAME": AddPageTuple key, curProductName
    key = "PROGRAM_NAME_BOLD": AddPageTuple key, UCase(curProductName)
    key = "PROGRAM_TYPE": AddPageTuple key, curProgramType
    key = "PROGRAM_VERSION": AddPageTuple key, curProgramVersion
    key = "REGISTRATION_COST_OTHER": AddPageTuple key, curRegistrationCostOther
    key = "REGISTRATION_COST_USD": AddPageTuple key, curRegistrationCostUSD
    
    If IsDate(curReleaseDate) Then
        key = "RELEASE_DATE": AddPageTuple key, Format(CDate(curReleaseDate), "mm/dd/yyyy")
        key = "RELEASE_MONTH": AddPageTuple key, Format(CDate(curReleaseDate), "mm")
        key = "RELEASE_DAY": AddPageTuple key, Format(CDate(curReleaseDate), "dd")
        key = "RELEASE_YEAR": AddPageTuple key, Format(CDate(curReleaseDate), "yyyy")
    Else
        key = "RELEASE_DATE": AddPageTuple key, ""
        key = "RELEASE_MONTH": AddPageTuple key, ""
        key = "RELEASE_DAY": AddPageTuple key, ""
        key = "RELEASE_YEAR": AddPageTuple key, ""
    End If
    
    key = "RELEASE_STATUS": AddPageTuple key, curReleaseStatus
    key = "SALES_EMAIL": AddPageTuple key, curSalesEmail
    key = "SALES_PHONE": AddPageTuple key, curSalesPhone
    key = "SCREENSHOT_URL": AddPageTuple key, curScreenshotURL
'    key = "SITE_LIST": AddPageTuple key, curSiteList
    key = "SPECIFIC_CATEGORY": AddPageTuple key, curSpecificCategory
    key = "STATE_PROVINCE": AddPageTuple key, curStateProvince
    key = "SUPPORT_EMAIL": AddPageTuple key, curSupportEmail
    key = "SUPPORT_PHONE": AddPageTuple key, curSupportPhone
    key = "SYSTEM_REQUIREMENTS": AddPageTuple key, PadCRLF(curSystemRequirements)
    key = "XML_FILE_URL": AddPageTuple key, curXMLfileURL
    key = "ZIP_POSTAL": AddPageTuple key, curZIPPostal
    
    key = "PROGRAM_EXPIRES": AddPageTuple key, IIf(curProgramExpires = 1, "Y", "N")
    If curProgramExpires = 1 Then
        key = "EXPIRE_OTHER_INFO": AddPageTuple key, curExpirationOtherInfo
        key = "EXPIRE_COUNT": AddPageTuple key, curExpirationCount
        key = "EXPIRE_BASED_ON": AddPageTuple key, curExpireBasedOn
        
        If IsDate(curExpirationDate) And curProgramExpires = 1 Then
            key = "EXPIRE_INFO": AddPageTuple key, "Expires on " & curExpirationDate
            key = "EXPIRE_DATE": AddPageTuple key, curExpirationDate
            key = "EXPIRE_MONTH": AddPageTuple key, Format(CDate(curExpirationDate), "mm")
            key = "EXPIRE_DAY": AddPageTuple key, Format(CDate(curExpirationDate), "dd")
            key = "EXPIRE_YEAR": AddPageTuple key, Format(CDate(curExpirationDate), "yyyy")
        Else
            key = "EXPIRE_INFO": AddPageTuple key, "Expires after " & curExpirationCount & " " & curExpireBasedOn
            key = "EXPIRE_DATE": AddPageTuple key, ""
            key = "EXPIRE_MONTH": AddPageTuple key, ""
            key = "EXPIRE_DAY": AddPageTuple key, ""
            key = "EXPIRE_YEAR": AddPageTuple key, ""
        End If
    Else
        key = "EXPIRE_OTHER_INFO": AddPageTuple key, ""
        key = "EXPIRE_COUNT": AddPageTuple key, ""
        key = "EXPIRE_BASED_ON": AddPageTuple key, ""
        key = "EXPIRE_INFO": AddPageTuple key, ""
        key = "EXPIRE_DATE": AddPageTuple key, ""
        key = "EXPIRE_MONTH": AddPageTuple key, ""
        key = "EXPIRE_DAY": AddPageTuple key, ""
        key = "EXPIRE_YEAR": AddPageTuple key, ""
    End If
    
    For X = 0 To UBound(PageTuple)
        PageTuple(X).Value = strParseTemplateText(PageTuple(X).Value)
    Next X
End Sub

Public Sub LoadBlurbOrder()
    Dim BlurbOrderFile As String
    Dim OldOrder As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    BlurbOrder = ""
    
    On Error Resume Next
    'load old order, verify contents, add new files
    BlurbOrderFile = DataPath & "blurborder.dat"
    If Dir(BlurbOrderFile) > "" Then
        Open BlurbOrderFile For Input As #FileNum
            Line Input #FileNum, OldOrder
        Close #FileNum
    End If
    
    BlurbOrder = OldOrder
End Sub

Public Sub TOCandArticles(ByVal ArticleList As String, ByVal ListOrder As String, ByVal FileDir As String, Extension As String)
    Dim Article() As FileTuple
    Dim FileList As String
    Dim eArticleList() As String
    Dim eListOrder() As String
    Dim tmpArticle As FileTuple
    Dim ArticlesInOrder As Integer
    Dim nTOC As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    ReDim Article(0)
    ReDim eListOrder(0)
    Articles = ""
    TOC = ""
    
    'verify article exists, build article array
    eArticleList = Split(ArticleList, ",")
    FileList = GetDirContents(FileDir, Extension)
    For X = UBound(eArticleList) To 0 Step -1
        If InStr(1, FileList, eArticleList(X)) > 0 Then
            y = UBound(Article) + 1
            ReDim Preserve Article(y)
            Article(y).key = eArticleList(X)
            Article(y).File = FileDir & eArticleList(X) & Extension
        End If
    Next X
    '--------------------
    
    'order articles by list order
    eListOrder = Split(ListOrder, ",")
    For X = 0 To UBound(eListOrder)
        For y = 0 To UBound(Article)
            If Article(y).key = eListOrder(X) Then
                tmpArticle = Article(y)
                Article(y) = Article(ArticlesInOrder)
                Article(ArticlesInOrder) = tmpArticle
                ArticlesInOrder = ArticlesInOrder + 1
            End If
        Next y
    Next X
    '--------------------
    
    'build TOC & Articles
    For y = 0 To UBound(Article)
        If Article(y).key > "" Then
            nTOC = nTOC + 1
            tmp = nTOC & ") " & Article(y).key
            TOC = TOC & tmp & vbCrLf
            
            Articles = Articles & tmp & vbCrLf
            Articles = Articles & String(Len(tmp), "=") & vbCrLf
            Open Article(y).File For Input As #FileNum
                Line Input #FileNum, tmp  'blurb name
                Line Input #FileNum, tmp
                    Articles = Articles & PadCRLF(tmp) & vbCrLf & vbCrLf
            Close #FileNum
        End If
    Next y
    '--------------------
End Sub

Public Function GetDirContents(ByVal FileDir As String, Extension As String) As String
    Dim tmp As String
    
    tmp = Dir(FileDir & "*" & Extension)
    If tmp = "" Then Exit Function
    Do Until tmp = ""
        If tmp = "DEFAULT" & Extension Then tmp = Extension
        tmp2 = tmp2 & Left(tmp, Len(tmp) - Len(Extension)) & ","
        tmp = Dir
    Loop
    
    GetDirContents = tmp2
End Function

