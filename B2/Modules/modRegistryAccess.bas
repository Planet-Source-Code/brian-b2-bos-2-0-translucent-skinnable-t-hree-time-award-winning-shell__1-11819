Attribute VB_Name = "modRegistryAccess"


'***************************************************************
'Windows API/Global Declarations for :strGetMachineName
'***************************************************************
    Private Const MAX_LEN = 260
'-- data types for registry entries
Const REG_NONE As Long = 0
Const REG_SZ As Long = 1
Const REG_EXPAND_SZ As Long = 2
Const REG_BINARY As Long = 3
Const REG_DWORD As Long = 4
Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Const REG_DWORD_BIG_ENDIAN As Long = 5
Const REG_LINK As Long = 6
Const REG_MULTI_SZ As Long = 7
Const REG_RESOURCE_LIST As Long = 8
Private Const ERROR_SUCCESS As Long = 0&


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
    End Type
    '-- top level registry hives
    Public Const HKEY_CLASSES_ROOT As Long = &H80000000
    Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
    Public Const HKEY_CURRENT_USER As Long = &H80000001
    Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
    Public Const HKEY_PERFORMANCE_DATA As Long = &H80000004
    Public Const HKEY_USERS As Long = &H80000003


Private Declare Function RegOpenKey _
    Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal HKEY As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


Private Declare Function RegCloseKey _
    Lib "advapi32.dll" _
    (ByVal HKEY As Long) As Long


Private Declare Function RegSetValueEx _
    Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal HKEY As Long, _
    ByVal lpValueName As String, _
    ByVal reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long


Private Declare Function RegCreateKey _
    Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal HKEY As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


Private Declare Function RegQueryValueEx _
    Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal HKEY As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long
    
    Public Function GetRegistryValue( _
    xlngTopKey As Long, _
    xstrSubKey As String, _
    xstrValueName As String) As String
    On Error Resume Next
    Const constProcName = "GetRegistryValue"
    Dim lngRtn As Long
    Dim lngKeyHwnd As Long
    Dim strRtnValue As String
    Dim abytValue() As Byte
    Dim lngRtnValue As Long
    '-- Create under the passed, top-level subkey
    lngRtn = RegOpenKey(xlngTopKey, xstrSubKey, lngKeyHwnd)


    If lngRtn <> ERROR_SUCCESS Then
        Exit Function
    Else
        '-- lngKeyHwnd is handle to the open key
    End If

    '-- Set the variables to pass
    strRtnValue = String(MAX_LEN, Chr$(0))
    lngRtnValue = Len(strRtnValue)
    '-- Set the byte array
    abytValue = strRtnValue
    '-- Get the value
    lngRtn = RegQueryValueEx(lngKeyHwnd, _
    xstrValueName, _
    0&, REG_SZ, _
    abytValue(0), _
    lngRtnValue)
    '-- Close the handle to the subkey
    lngRtn = RegCloseKey(lngKeyHwnd)
    '-- Convert byte array into string
    strRtnValue = abytValue
    '-- Convert to UniCode, and trim based on return length
    strRtnValue = Mid$(StrConv(strRtnValue, vbUnicode), 1, lngRtnValue)
    '-- Trim any trailing null
    lngRtn = InStr(strRtnValue, Chr$(0))


    If lngRtn > 0 Then
        strRtnValue = Mid$(strRtnValue, 1, lngRtn - 1)
    End If

    '-- Set the return value
    GetRegistryValue = strRtnValue
    On Error GoTo 0
End Function

Public Function GetDesktopPath() As String
    GetDesktopPath = GetRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
End Function

Public Function GetStartMenuPath() As String
    GetStartMenuPath = GetRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs")
End Function

Public Function GetFavoritesPath() As String
    GetFavoritesPath = GetRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites")
End Function

Public Function GetRecent() As String
    GetRecent = GetRegistryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Recent")
End Function
