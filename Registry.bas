Attribute VB_Name = "mdlRegistry"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)

Public Sub DeleteKey(lPredefinedKey As Long, sKeyName As String)
    '   This Function will Delete a key
    '
    '   DeleteKey Location, KeyName
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
    '   KeyName is name of the key you wish to delete, it may include subkeys (example "Key1\SubKey1")
    
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    
    lRetVal = RegDeleteKey(lPredefinedKey, sKeyName)
End Sub

Public Sub DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    '   This Function will delete a value
    '
    '   DeleteValue Location, KeyName, ValueName
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
    '   KeyName is the name of the key that the value you wish to delete is in, it may include subkeys (example "Key1\SubKey1")
    '   ValueName is the name of value you wish to delete

    Dim lRetVal As Long
    Dim hKey As Long
    
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = RegDeleteValue(hKey, sValueName)
    RegCloseKey (hKey)
End Sub

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
    
End Function

Public Sub CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
    '   This Function will create a new key
    '
    '   lPredefinedKey must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
    '   KeyName is name of the key you wish to create, it may include subkeys (example "Key1\SubKey1")
    
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub

Public Sub SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    '   This Function will set the data field of a value
    '
    '   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
    '   KeyName is the key that the value is under (example: "Key1\SubKey1")
    '   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
    '   ValueSetting is what you want the value to equal
    '   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

    Dim lRetVal As Long
    Dim hKey As Long
    
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub

Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String) As Variant
    '   This Function will return the data field of a value
    '
    '   Variable = QueryValue(Location, KeyName, ValueName)
    '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS
    '   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
    '   ValueName is the name of the value you want to access (example: "link")

    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    QueryValue = vValue
End Function





