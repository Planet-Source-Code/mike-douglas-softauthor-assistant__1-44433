Attribute VB_Name = "TimeZoneInfo"
Option Explicit
' Time Zone API declarations
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Type SYSTEMTIME
  wYear(1 To 2) As Byte ' VB pads integers
  wMonth(1 To 2) As Byte ' so we use bytes
  wDayOfWeek(1 To 2) As Byte
  wDay(1 To 2) As Byte
  wHour(1 To 2) As Byte
  wMinute(1 To 2) As Byte
  wSecond(1 To 2) As Byte
  wMilliseconds(1 To 2) As Byte
End Type

Private Type TIME_ZONE_INFORMATION
  bias As Long ' current offset to GMT
  StandardName(1 To 64) As Byte ' unicode string
  StandardDate As SYSTEMTIME
  StandardBias As Long
  DaylightName(1 To 64) As Byte
  DaylightDate As SYSTEMTIME
  DaylightBias As Long
End Type

Private Declare Sub GetSystemTime Lib "kernel32" _
  (lpSystemTime As SYSTEMTIME)
  
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
  (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
  
Private uTZI As TIME_ZONE_INFORMATION
Private dtGMT As Date ' current GMT time
Private lBias As Long ' current offset
Private sName As String ' current tz name
Private sDS As Boolean ' Daylight Savings?
Private sInfo As String ' scratch
  
Public Function DayLightSavings() As Boolean
    '----------
    ' returns true if daylight savings is in effect
    '----------
    SetZoneInfo
    DayLightSavings = sDS
End Function

Public Function GMTTime() As Date
    '----------
    ' returns Greenich Mean Time as date
    '----------
    SetZoneInfo
    GMTTime = dtGMT
End Function

Public Function GMTTimeString() As String
    '----------
    ' returns properly formatted GMT time as string
    '----------
    SetZoneInfo
    GMTTimeString = Format$(dtGMT, "dd-mmm-yyyy hh:mm:ss") & " GMT"
End Function

Public Function TimeZoneName() As String
    '----------
    ' returns plain english timezone area info
    '----------
    SetZoneInfo
    TimeZoneName = sName
End Function

Public Function GMTOffset() As String
    '----------
    ' returns formatted GMT hour offset in "-00:00" format
    '----------
    SetZoneInfo
    GMTOffset = IIf(lBias < 1, "+", "-") & Right$("00" & CStr(Abs(lBias) \ 60), 2) & ":" & Right$("00" & CStr(lBias Mod 60), 2)
End Function

Private Sub SetZoneInfo()
    '----------
    ' "meat" sub to calculate values in module
    '----------
    Dim X As Long ' scratch
    
    Select Case GetTimeZoneInformation(uTZI)
      ' if not daylight assume standard
      Case TIME_ZONE_ID_DAYLIGHT:
        sName = uTZI.DaylightName ' convert to string
        lBias = uTZI.bias + uTZI.DaylightBias
        sDS = True
      Case Else:
        sName = uTZI.StandardName
        lBias = uTZI.bias + uTZI.StandardBias
        sDS = False
    End Select
    
    ' name terminates with null
    X = InStr(sName, vbNullChar)
    If X > 0 Then sName = Left$(sName, X - 1)
    
    dtGMT = DateAdd("n", lBias, Now) ' calculate GMT
End Sub
