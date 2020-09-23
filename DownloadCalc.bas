Attribute VB_Name = "DownloadCalc"
Public Function DCalc(ByVal BytesTransfered As Long, ByVal KiloBitsPerSecond As Long) As String
    Dim Speed As Long
    Dim tSeconds As Long
    Dim tMinutes As Long
    Dim tHours As Long
    Dim tDays As Long
    Dim tFull As String
    
    Speed = KiloBitsPerSecond * 1000
    Speed = Speed / 8 'convert bits/sec to bytes/sec
    
    tSeconds = BytesTransfered / Speed
    
    tMinutes = tSeconds \ 60
    tSeconds = tSeconds - tMinutes * 60
    
    tHours = tMinutes \ 60
    tMinutes = tMinutes - tHours * 60
    
    tDays = tHours \ 24
    tHours = tHours - tDays * 24
    
    tFull = tSeconds & " seconds"
    If tMinutes > 0 Then tFull = tMinutes & " minutes " & tFull
    If tHours > 0 Then tFull = tHours & " hours " & tFull
    If tDays > 0 Then tFull = tDays & " days " & tFull
    
    DCalc = tFull
End Function

