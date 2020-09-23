Attribute VB_Name = "WordWrapper"
Public Function WordWrap(ByVal St As String, ByVal Length As Integer) As String
    Dim tmpText As String
    
    Length = Length + 1
    St = Trim(St)

    Do
        l = Len(NextLine$)
        s = InStr(St, " ")
        C = InStr(St, vbCr)

        If C Then
            If l + C <= Length Then
                tmpText = tmpText & NextLine$ & Left$(St, C)
                NextLine$ = ""
                St = Mid$(St, C + 1)
                GoTo LoopHere
            End If
        End If
        If s Then
            If l + s <= Length Then
                DoneOnce = True
                NextLine$ = NextLine$ & Left$(St, s)
                St = Mid$(St, s + 1)
            ElseIf s > Length Then
                tmpText = tmpText & vbCrLf & Left$(St, Length)
                St = Mid$(St, Length + 1)
            Else
                tmpText = tmpText & NextLine$ & vbCrLf
                NextLine$ = ""
            End If
        Else
            If l Then
                If l + Len(St) > Length Then
                    tmpText = tmpText & NextLine$ & vbCrLf & St & vbCrLf
                Else
                    tmpText = tmpText & NextLine$ & St & vbCrLf
                End If
            Else
                tmpText = tmpText & St & vbCrLf
            End If
            Exit Do
        End If
LoopHere:
    Loop

    Do Until Right(tmpText, 2) <> vbCrLf
        tmpText = Left(tmpText, Len(tmpText) - 2)
    Loop

    WordWrap = tmpText
End Function

