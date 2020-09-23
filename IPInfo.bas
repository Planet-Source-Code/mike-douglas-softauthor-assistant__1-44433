Attribute VB_Name = "IPInfo"
'IPINFO.BAS
'TCP/IP stack information routines
'
'Dependancies:
' -REGISTRY.BAS
'---------------------------------
Option Explicit

Private Initialized As Boolean          ' has IPInfo been initialized yet?
Public Domain As String                 ' IP domain ex:'yahoo.com'
Public DNSServers As New Collection     ' DNS server IP addresses
Public HostName As String               ' Machine (host) name

Public Sub IPInfoInit()
    '----------
    ' called at least once to set values
    '----------
    If Initialized = True Then Exit Sub
    Initialized = True
    IPInfoHostName
    IPInfoDomain
    LoadDNSServers
End Sub

Private Sub IPInfoHostName()
    Dim Domain As String
    'NT
    HostName = CleanString(QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "Hostname"))
    'Win9x
    If HostName = "" Then HostName = CleanString(QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\MSTCP", "HostName"))
    
End Sub

Private Sub IPInfoDomain()
    
    'NT
    Domain = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "Domain")
    If Domain = "" Then Domain = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "DhcpDomain")
    'Win9x
    If Domain = "" Then Domain = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\MSTCP", "Domain")
    If Domain = "" Then Domain = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\MSTCP", "DhcpDomain")
    'If nuthin else...
    If Domain = "" Then Domain = "aesgard.com"
    Domain = CleanString(Domain)
End Sub

Private Sub LoadDNSServers()
    '-------------
    ' Makes appropriate registry lookup for DNS
    '-------------
    Dim DNSList As String
    
    'NT
    DNSList = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\", "NameServer")
    'Win9x
    If DNSList = "" Then DNSList = QueryValue(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\VxD\MSTCP\", "NameServer")
    ParseDNS DNSList
End Sub

Private Sub ParseDNS(RegEntry As String)
    '-----------
    ' takes comma-delimited DNS registry entry and fills DNSServers
    ' collection with individual server addresses
    '-----------
    Dim LastComma As Integer
    Dim Comma As Integer
    Dim TempName As String
    
    If Len(RegEntry) < 7 Then Exit Sub  ' minimum valid entry length
    
    RegEntry = RegEntry & ","
    
    Do Until LastComma = Len(RegEntry)
        Comma = InStr(LastComma + 1, RegEntry, ",")
        TempName = (Mid(RegEntry, LastComma + 1, Comma - LastComma - 1))
        If TempName > "" Then DNSServers.Add Item:=TempName, key:=TempName
        LastComma = Comma
        DoEvents
    Loop
End Sub

Private Function CleanString(Incoming As String) As String
    '----------
    ' remove non alphamerics
    '----------
    Dim temp As String
    Dim allowable As String
    Dim X As Integer
    Dim tempchar As String
    
    allowable = "01234567890.ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Incoming = Trim(Incoming)
    For X = 1 To Len(Incoming)
        tempchar = Mid(Incoming, X, 1)
        temp = temp & IIf(InStr(1, allowable, tempchar), tempchar, "")
    Next X
    
    CleanString = temp
End Function
