Attribute VB_Name = "BrowserFuncs"
Private Type Site
    Name As String
    URL As String
    Completed As Boolean
    Notes As String
End Type

Public SiteList() As Site
Public SubmitList() As Site
Public SubmitListIndex As Integer

Public Sub GoNextSite()
    Dim XX As Integer
    
    On Error Resume Next
    XX = SubmitListIndex + 1
    
    If DemoVersion = True And XX > 10 Then
        Status "Unregistered version limited to 10 sites."
        ReDim Preserve SubmitList(10)
    End If
    
    If XX > UBound(SubmitList) Then
        SubmitListIndex = UBound(SubmitList)
        MsgBox "End of web submission list."
    Else
        If frmMain.chkIncompleteSites.Value = 1 Then
            Do Until SubmitList(XX).Completed = False
                XX = XX + 1
                If XX > UBound(SubmitList) Then
                    MsgBox "End of web submission list."
                    XX = SubmitListIndex 'UBound(SubmitList)
                    Exit Sub
                End If
            Loop
        End If
        SubmitListIndex = XX
        Surf2ListSite SubmitListIndex
    End If
End Sub

Public Sub GoPrevSite()
    On Error Resume Next
    XX = SubmitListIndex - 1
    If XX < 1 Then
        SubmitListIndex = 1
        MsgBox "Start of web submission list."
    Else
        If frmMain.chkIncompleteSites.Value = 1 Then
            Do Until SubmitList(XX).Completed = False
                XX = XX - 1
                If XX < 1 Then
                    MsgBox "Start of web submission list."
                    XX = SubmitListIndex 'UBound(SubmitList)
                    Exit Sub
                End If
            Loop
        End If
        SubmitListIndex = XX
        Surf2ListSite SubmitListIndex
    End If
End Sub

Public Sub Surf2ListSite(ByVal Index As Integer)
    On Error Resume Next
    frmMain.WebBrowser1.Document.Clear '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    If UBound(SubmitList) = 0 Then
        MsgBox "Submit list is empty.", vbOKOnly, "Submit List"
        WebSubmiting = False
        frmMain.UpdateDisplay
        frmMain.WebBrowser1.Navigate HelpPath & "INDEX.HTML"
        Exit Sub
    End If

    SubmitListIndex = Index
    frmMain.txtBrowserURL = SubmitList(Index).URL
'    frmMain.Refresh
    frmMain.chkWebList.Value = IIf(SubmitList(Index).Completed = True, 1, 0)
    frmMain.chkWebList.Caption = SubmitList(Index).Name & " (" & Index & "/" & UBound(SubmitList) & ")"
    frmMain.WebBrowser1.Navigate SubmitList(Index).URL
    frmPostIt.Caption = SubmitList(Index).Name
    frmPostIt.txtPostIt.Text = SubmitList(Index).Notes
    If SubmitList(Index).Notes > "" Then
        frmPostIt.Show
    Else
        frmPostIt.Hide
    End If
    
    frmMain.Refresh
End Sub

Public Sub InitSubmitList()
    Dim Element() As String
    Dim tmpSite As Site
    Dim tmpSubmitlist As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReDim SubmitList(0)
    tmpSubmitlist = vbCrLf & frmMain.txtSiteList & vbCrLf
    For X = 0 To UBound(SiteList)
        If SiteList(X).Name > "" And InStr(1, tmpSubmitlist, vbCrLf & SiteList(X).Name & vbCrLf) > 0 Then
            XX = UBound(SubmitList) + 1
            ReDim Preserve SubmitList(XX)
            SubmitList(XX) = SiteList(X)
        End If
    Next X
    
    ' update submitlist info with submitlist.dat info
    Open DataPath & "submitlist.dat" For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, tmp
                eLine = Split(tmp, "};{")
                If UBound(eLine) > 1 And eLine(0) > "" Then
                    X = FindSubmitListSite(eLine(0))
                    If X > 0 Then
                        SubmitList(X).URL = eLine(1)
                        SubmitList(X).Completed = CBool(eLine(2))
                    End If
                End If
        Loop
    Close #FileNum
    
    LoadSubmitListNotes
End Sub

Public Sub SaveSubmitList()
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    Open DataPath & "submitlist.dat" For Output As #FileNum
        For X = 0 To UBound(SubmitList)
            If SubmitList(X).Name > "" Then
                Print #FileNum, SubmitList(X).Name & "};{" & SubmitList(X).URL & "};{" & SubmitList(X).Completed
            End If
        Next X
    Close #FileNum
End Sub

Public Sub LoadSubmitList()
    Dim eLine() As String
    Dim tmpSubmitlist As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    ReDim SubmitList(0)
    On Error Resume Next
    If Dir(DataPath & "submitlist.dat") > "" Then
        Open DataPath & "submitlist.dat" For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, tmp
                    eLine = Split(tmp, "};{")
                    If UBound(eLine) > 1 And eLine(0) > "" Then
                        X = UBound(SubmitList) + 1
                        ReDim Preserve SubmitList(X)
                        SubmitList(X).Name = eLine(0)
                        SubmitList(X).URL = eLine(1)
                        SubmitList(X).Completed = CBool(eLine(2))
                    End If
            Loop
        Close #FileNum
        
        For X = 0 To UBound(SubmitList)
            If SubmitList(X).Name > "" Then tmpSubmitlist = tmpSubmitlist & SubmitList(X).Name & vbCrLf
        Next X
    End If
    
    frmMain.txtSiteList = tmpSubmitlist
End Sub

Public Sub LoadSubmitListNotes()
    Dim eLine() As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    If Dir(DataPath & "submitlistnotes.dat") > "" Then
        Open DataPath & "submitlistnotes.dat" For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, tmp
                    eLine = Split(tmp, "};{")
                    If UBound(eLine) > 0 Then
                        X = FindSubmitListSite(eLine(0))
                        If X > 0 Then SubmitList(X).Notes = PadCRLF(eLine(1))
                    End If
            Loop
        Close #FileNum
    End If
End Sub

Public Function FindSubmitListSite(ByVal SiteName As String) As Integer

    SiteName = UCase(Trim(SiteName))
    If SiteName = "" Then Exit Function
    For X = 0 To UBound(SubmitList)
        If SiteName = UCase(Trim(SubmitList(X).Name)) Then y = X
    Next X
    
    FindSubmitListSite = y
End Function

Public Function FindSiteListSite(ByVal SiteName As String) As Integer

    SiteName = UCase(Trim(SiteName))
    If SiteName = "" Then Exit Function
    For X = 0 To UBound(SiteList)
        If SiteName = UCase(Trim(SiteList(X).Name)) Then y = X
    Next X
    
    FindSiteListSite = y
End Function

Public Sub SaveSubmitListNotes()
    Dim OldSubmitData() As String
    Dim FileNum As Integer
    
    FileNum = FreeFile
    On Error Resume Next
    ' open/read old file and concat new/updated info to old (preserve old data)
    ReDim OldSubmitData(0)
    If Dir(DataPath & "submitlistnotes.dat") > "" Then
        Open DataPath & "submitlistnotes.dat" For Input As #FileNum
            Do Until EOF(FileNum)
                Line Input #FileNum, tmp
                    eLine = Split(tmp, "};{")
                    If UBound(eLine) > 0 Then
                        X = FindSubmitListSite(eLine(0))
                        If X = 0 And eLine(0) > "" Then
                            X = UBound(OldSubmitData) + 1
                            ReDim Preserve OldSubmitData(X)
                            OldSubmitData(X) = tmp
                        End If
                    End If
            Loop
        Close #FileNum
    End If
    
    Open DataPath & "submitlistnotes.dat" For Output As #FileNum
        For X = 0 To UBound(SubmitList)
            If SubmitList(X).Name > "" And SubmitList(X).Notes > "" Then
                Print #FileNum, SubmitList(X).Name & "};{" & StripCRLF(SubmitList(X).Notes)
            End If
        Next X
    Close #FileNum
    
    Open DataPath & "submitlistnotes.dat" For Append As #FileNum
        For X = 0 To UBound(OldSubmitData)
            If OldSubmitData(X) > "" Then Print #FileNum, OldSubmitData(X)
        Next X
    Close #FileNum
End Sub

Public Sub LoadSiteList()
    Dim FileNum As Integer
    
    FileNum = FreeFile
    If Dir(DataPath & "sitelist.dat") = "" Then Exit Sub
    Open DataPath & "sitelist.dat" For Input As #FileNum
        Do Until EOF(FileNum)
            Line Input #FileNum, X$
            AddSite X$
        Loop
    Close #FileNum
End Sub

Public Sub SaveSiteList()
    Dim FileNum As Integer
    Dim SiteListContents As String
    
    For X = 0 To UBound(SubmitList)
        If SubmitList(X).Name > "" And SubmitList(X).URL > "" Then _
           SiteListContents = SiteListContents & SubmitList(X).Name & "};{" & SubmitList(X).URL & vbCrLf
    Next X
    
    yy = MsgBox("Include old sitelist data in new save?", vbYesNo, "Save Sitelist")
    If yy = vbYes Then
        For X = 0 To UBound(SiteList)
            y = FindSubmitListSite(SiteList(X).Name)
            If y = 0 And SiteList(X).Name > "" And SiteList(X).URL > "" Then _
               SiteListContents = SiteListContents & SiteList(X).Name & "};{" & SiteList(X).URL & vbCrLf
        Next X
    End If
    
    FileNum = FreeFile
    'on error resume next
    Open DataPath & "sitelist.dat" For Output As #FileNum
        Print #FileNum, SiteListContents;
    Close #FileNum
End Sub

Public Sub AddSite(ByVal Line As String)
    Dim Element() As String
    Dim tmpName As String
    Dim tmpURL As String
    
    Line = Line & "};{"
    Element = Split(Line, "};{")
    If UBound(Element) < 1 Then Exit Sub
    tmpName = Element(0)
    tmpURL = Element(1)
    
    If Trim(tmpName) = "" Or Trim(tmpURL) = "" Then Exit Sub
    
    For X = 0 To UBound(SiteList)
        If SiteList(X).Name = tmpName Then Exit Sub
    Next X
    
    X = UBound(SiteList) + 1
    ReDim Preserve SiteList(X)
    SiteList(X).Name = tmpName ' & " "
    SiteList(X).URL = tmpURL
End Sub


