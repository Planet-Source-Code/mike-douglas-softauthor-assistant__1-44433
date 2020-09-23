VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form SMTPClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   945
   Icon            =   "SMTPClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock SMTPWinsock 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "SMTPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SMTP Mail send
'
'Dependencies:
' - TIMEZONE.BAS
' - REGISTRY.BAS
' - LOG.CLS
' - IPINFO.BAS
'
'----------------------------------------------
Option Explicit


'Private Response As String
'Private WithEvents SMTPWinsock As MSWinsockLib.winsock

Const dblSecond = 0.000011574
Private Boundary As String
Const vQuote = """"
Private Start As Single, tmr As Single
Private Response As String
Private LastResponse As String
Private Reply As Integer
Private Abort As Boolean                ' error, abort email
Private UserAbort As Boolean            ' error, abort email
Private DataSent As Boolean

Public ErrorLog As New Log              ' LOG.CLS instance for error reporting
Public Status As String                 ' current SMTP transfer status
Public MailServer As String
Public FromName As String
Public FromAddress As String
Public ToName As String
Public ToAddress As String
Public Subject As String
Public Message As String
Public Domain As String
Public SMTPPort As Integer
Public TimeOut As Integer               ' seconds in operation to wait before
                                        ' signaling SMTP server timeout.
Private AUTHStatement As String         ' SMTP 'AUTH=LOGIN' style response
Public RelayUser As String
Public RelayPassword As String
Public Retries As Integer
Public Attachments As String

Public Sub Send()
    Dim Addressees() As String
    Dim X As Integer
    
    DataSent = False
    Abort = False
    UserAbort = False
    AUTHStatement = ""
    Response = ""
    LastResponse = ""
    Reply = 0
    
    If MailServer = "" Or ToAddress = "" Or Message = "" Then
        UStatus "Send aborted due to incomplete messaging values"
        Exit Sub
    End If
    
    If InStr(1, ToAddress, ";") Then
        'multiple sends
        Addressees = Split(ToAddress, ";", -1)
        For X = 0 To UBound(Addressees)
            UStatus "Message " & X + 1 & " of " & (UBound(Addressees) + 1) & " (" & Addressees(X) & ")"
            If Addressees(X) > "" Then SendMessage Addressees(X)
            UStatus "====="
            If UserAbort = True Then Exit For
        Next X
    Else
        'single address
        SendMessage ToAddress
    End If
    
End Sub

Public Sub AbortSend()
    Abort = True
    UserAbort = True
End Sub

Private Sub SendMessage(EmailAddress As String)
    Dim DateNow As String
    Dim ContentType As String
    Dim RetryLoop As Integer
    Dim tmrx As Double

    ErrorLog.FileNameBase = "SMTP"                          ' set up error log
    ErrorLog.Silent = True
    IPInfo.IPInfoInit                                       ' init ipinfo stack
    
    Domain = IPInfo.Domain                                  ' configure domain for 'HELO' command
    
    If Retries < 1 Then Retries = 1
    For RetryLoop = 1 To Retries
        Boundary = "----=_Content_Boundary_Next_Part"
        UStatus ("Attempt " & RetryLoop & " of " & Retries)
        If Attachments > "" Then
            ContentType = ContentType & "multipart/mixed;" & vbCrLf & vbTab & "boundary=" & vQuote & Boundary & vQuote
        Else
            ContentType = ContentType & "text/plain;" & vbCrLf & "Charset=" & vQuote & "Windows-1252" & vQuote
        End If
        Boundary = vbCrLf & "--" & Boundary
        
            
        'setup defaults for necessary vars.
        If SMTPPort < 1 Then SMTPPort = 25
        If TimeOut < 1 Then TimeOut = 60
        
        ' Must set local port to 0 (Zero) to init socket.
        SMTPWinsock.Close
        Do Until SMTPWinsock.State = sckClosed Or Abort = True
            DoEvents
        Loop
        SMTPWinsock.LocalPort = 0
        SMTPWinsock.Close
        
            DateNow = GMTTimeString
        
            SMTPWinsock.Protocol = sckTCPProtocol      ' Set protocol for sending
            SMTPWinsock.RemoteHost = MailServer        ' Set the server address
            SMTPWinsock.RemotePort = SMTPPort
            
            SMTPWinsock.Connect                        ' Start connection
            
            UStatus "connecting"
            Start = Timer
            Do Until SMTPWinsock.State = sckConnected
                tmr = Timer - Start
                If tmr > TimeOut Then
                    UStatus "SMTP_CONNECT_ERR"
                    Abort = True
                    Exit Do
                End If
                DoEvents
            Loop
            
            WaitFor "220"
            SendData "EHLO " & Domain & vbCrLf
        
            UStatus "connected"
            WaitFor "250"
            
            'wait 2 secs for multiple '250' responses
            tmrx = Now
            Do Until (CDbl(Now) >= tmrx + (dblSecond * 2))
                DoEvents
            Loop

            If AUTHStatement > "" Then DoAUTHLogin
            
            SendData "MAIL FROM:" & FromAddress & vbCrLf
            
            UStatus "sending message(" & EmailAddress & ")"
            WaitFor "250"
            SendData "RCPT TO:" & EmailAddress & vbCrLf
            
            WaitFor "250"
            SendData "DATA" & vbCrLf
            
            WaitFor "354"
            SendData _
              "Date: " & DateNow & vbCrLf & _
              "From: " & FromName & vbCrLf & _
              "To: " & ToName & vbCrLf & _
              "Subject: " & Subject & vbCrLf & _
              "X-Mailer: Aesgard Technologies (www.aesgard.com)" & vbCrLf & _
              "MIME-Version: 1.0" & vbCrLf & _
              "Content-Type: " & ContentType & vbCrLf & vbCrLf
            
            If Attachments > "" Then
                SendData _
                  "This is a multi-part message in MIME format." & vbCrLf & vbCrLf & _
                  Boundary & vbCrLf & _
                  "Content-Type: text/plain;" & vbCrLf & _
                  "Charset = ""Windows-1252""" & vbCrLf & _
                  "Content-Transfer-Encoding: 7bit" & vbCrLf
            End If
            
            SendData vbCrLf & Message & vbCrLf & vbCrLf
            
            If Attachments > "" Then DoAttachments
            
            SendData (vbCrLf & "." & vbCrLf)
        
            WaitFor "250"
            SendData ("QUIT" & vbCrLf)
            UStatus "disconnecting"
        
            WaitFor "221"
            SMTPWinsock.Close
            
            UStatus "disconnected"
            If Abort = False Then Exit For
            If UserAbort = True Then Exit For
            Abort = False
        Next RetryLoop
End Sub

Private Sub DoAttachments()
    Dim aFilename() As String
    Dim X As Integer
    Dim XX As Long
    Dim FileContents As String
    Dim cIN As Long
    Dim tmpByte As String * 1
    Dim tmpKByte As String * 1024
    Dim cINKb As Long
    Dim cINRem As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    Dim tmr As Long
    
    If Abort = True Then Exit Sub
    
    aFilename = Split(Attachments, ";")
    For X = 0 To UBound(aFilename)
        FileContents = ""
        If Dir(aFilename(X)) > "" Then
            UStatus ("ATTACHING: " & aFilename(X) & "...")
            
            cIN = FileLen(aFilename(X))
            cINKb = cIN / 1024
            cINRem = cIN - (cINKb * 1024)
            
            UStatus cINKb & "Kb"
            
            Open aFilename(X) For Binary Access Read As #FileNum
                
                For XX = 1 To cINKb
                    Get #FileNum, , tmpKByte
                    FileContents = FileContents & tmpKByte
                Next XX
                
                For XX = 1 To cINRem
                    Get #FileNum, , tmpByte
                    FileContents = FileContents & tmpByte
                Next XX
                
            Close #FileNum
            
            tmr = Timer
            UStatus "Encoding..."
            FileContents = Encode64(FileContents)
            
            UStatus Len(FileContents)
            
            UStatus Int(Timer - tmr) & "secs"
            UStatus "Sending..."
            
            
            SendData _
              Boundary & vbCrLf & _
              "Content-Type: application/octet-stream;" & vbCrLf & _
              vbTab & "name=" & vQuote & FileNameOnly(aFilename(X)) & vQuote & vbCrLf & _
              "Content-Transfer-Encoding: base64" & vbCrLf & _
              "Content-Disposition: attachment;" & vbCrLf & _
              vbTab & "filename=" & vQuote & FileNameOnly(aFilename(X)) & vQuote & vbCrLf & vbCrLf & _
              FileContents & vbCrLf
        Else
            UStatus (aFilename(X) & " FILE NOT FOUND!")
        End If
    Next X
    
    SendData Boundary & vbCrLf
    
End Sub

Private Sub DoAUTHLogin()
    
    If RelayUser = "" Or RelayPassword = "" Or Abort = True Then Exit Sub
    
    SendData "AUTH LOGIN" & vbCrLf
    WaitFor "334"
    SendData Encode64(RelayUser) & vbCrLf
    WaitFor "334"
    SendData Encode64(RelayPassword) & vbCrLf
    WaitFor "235"
End Sub

Private Sub SendData(Data As String)
    DataSent = False
    If Abort = True Then Exit Sub
    If SMTPWinsock.State = sckConnected Then
        SMTPWinsock.SendData Data
        
        Do Until DataSent = True Or Abort = True
            DoEvents
        Loop
    End If
End Sub

Private Sub WaitFor(ResponseCode As String)
'    Dim Tmr2 As Long
    Dim tmr As Long
    Dim Start As Long
        
    If Abort = True Then Exit Sub
    
    Start = Timer
    
    Do Until InStr(1, LastResponse, ResponseCode)
        tmr = Timer - Start
        DoEvents
        
        'check for timeout
        If tmr > TimeOut Then
            UStatus "SMTP_TIMEOUT"
            Abort = True
            Exit Do
        End If
    Loop
    LastResponse = ""
End Sub

Private Sub SMTPWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim tmpResp     As String
    
    SMTPWinsock.GetData tmpResp                            ' Check for incoming response
    ParseServerResponse tmpResp
End Sub

Private Sub ParseServerResponse(ServerResponse As String)
    Dim RespCode As String
    Dim tmrx As Double
    
    LastResponse = ServerResponse
    
    RespCode = Left(LastResponse, 3)
    'fatal responses
    If InStr(1, "45-50-55", (Left(RespCode, 2))) Then
        UStatus "SMTP_ERR:" & ServerResponse
        Abort = True
        Exit Sub
    End If
    
    If InStr(1, ServerResponse, "AUTH") Then AUTHStatement = ServerResponse
End Sub

Private Sub UStatus(StatusMessage As String)
    If Abort = True Then Exit Sub
    Status = StatusMessage
    ErrorLog.WriteLog StatusMessage
    Main.Status StatusMessage
End Sub

Public Function Encode64(sz_UnEncoded As String) As String
    Dim ic_LowFill      As Integer
    Dim i_Char          As Integer
    Dim i_LowMask       As Integer
    Dim i_Ptr           As Integer
    Dim sz_Alphabet     As String * 64
    Dim sz_Alpha(64)    As String * 1
    Dim i_Counter       As Long
    Dim ic_BitShift     As Integer
    Dim ic_ChopMask     As Integer
    Dim i_HighMask      As Integer
    Dim i_Shift         As Integer
    Dim i_RollOver      As Integer
    Dim sz_Temp()       As String * 1
    Dim sz_TempCntr     As Long
    Dim X               As Long
    Dim EncodedString   As String
    Dim xxx As Integer
    
    ' Check if empty decoded string.
    ' If Empty, return NUL and Generate error 254
    If Len(sz_UnEncoded) = 0 Then Exit Function

    ' Initialize lookup dictionary and constants
    sz_Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    For X = 1 To 64
        sz_Alpha(X) = Mid$(sz_Alphabet, X, 1)
    Next X
    
    ic_BitShift = 4
    ic_ChopMask = 255
    ic_LowFill = 3

    ' Initialize Masks
'    sz_Temp = ""
    
    ReDim sz_Temp(Len(sz_UnEncoded) * 1.4)
    
    i_HighMask = &HFC
    i_LowMask = &H3
    i_Shift = &H10
    i_RollOver = 0


    ' Begin Encoding process
    For i_Counter = 1 To Len(sz_UnEncoded)

    ' Fetch ascii character in decoded string
        i_Char = Asc(Mid$(sz_UnEncoded, i_Counter, 1))

    ' Calculate Alphabet lookup pointer
        i_Ptr = ((i_Char And i_HighMask) \ (i_LowMask + 1)) Or i_RollOver

    ' Roll bit patterns
        i_RollOver = (i_Char And i_LowMask) * i_Shift

    ' Concatenate encoded character to working encoded string
'        sz_Temp = sz_Temp + Mid$(sz_Alphabet, i_Ptr + 1, 1)
        sz_TempCntr = sz_TempCntr + 1
        sz_Temp(sz_TempCntr) = sz_Alpha(i_Ptr + 1)

    ' Adjust masks
        i_HighMask = (i_HighMask * ic_BitShift) And ic_ChopMask
        i_LowMask = i_LowMask * ic_BitShift + ic_LowFill
        i_Shift = i_Shift \ ic_BitShift

    ' If last character in block, concat last RollOver and
    '   reset masks
        If i_HighMask = 0 Then
'            sz_Temp = sz_Temp + Mid$(sz_Alphabet, i_RollOver + 1, 1)
            sz_TempCntr = sz_TempCntr + 1
            sz_Temp(sz_TempCntr) = Mid$(sz_Alphabet, i_RollOver + 1, 1)
            i_RollOver = 0
            i_HighMask = &HFC
            i_LowMask = &H3
            i_Shift = &H10
'            DoEvents
        End If

    Next i_Counter

    ' If RollOver remains, concat it to the working string
    If i_Shift < &H10 Then
'        sz_Temp = sz_Temp + Mid$(sz_Alphabet, i_RollOver + 1, 1)
        sz_TempCntr = sz_TempCntr + 1
        sz_Temp(sz_TempCntr) = Mid$(sz_Alphabet, i_RollOver + 1, 1)
    End If

'    i_Ptr = (Len(sz_Temp) Mod 4)
'    If i_Ptr Then sz_Temp = sz_Temp + String$(4 - i_Ptr, "=")
    i_Ptr = (sz_TempCntr Mod 4)
    
    If i_Ptr Then
        sz_TempCntr = sz_TempCntr + 1
        If UBound(sz_Temp) < sz_TempCntr Then ReDim Preserve sz_Temp(sz_TempCntr)
        sz_Temp(sz_TempCntr) = String$(4 - i_Ptr, "=")
    End If
    
    EncodedString = Space(sz_TempCntr)
    For X = 1 To sz_TempCntr
        Mid(EncodedString, X, 1) = sz_Temp(X)
    Next X
    
    Encode64 = EncodedString
End Function

Private Function FileNameOnly(FullName As String) As String
    '-------------------
    ' return filename w/o path
    '-------------------
    Dim tmpChar As String
    Dim X As Integer
    
    If InStr(1, FullName, "\") Or InStr(1, FullName, ":") Then
        For X = Len(FullName) To 1 Step -1
            tmpChar = Mid$(FullName, X, 1)
            If tmpChar = "\" Or tmpChar = ":" Then
                If X < Len(FullName) Then FullName = Mid$(FullName, X + 1)
                Exit For
            End If
        Next X
    End If
    
    FileNameOnly = FullName
End Function

Private Sub SMTPWinsock_SendComplete()
    DataSent = True
End Sub
