Attribute VB_Name = "ModRegistry"
' Module to read from and write to the registry

'Thinc: yes, I copied this module from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49643&lngWId=1
' take a look at that thing if you wish

Option Explicit

Private r As Long, keyhand As Long, lResult As Long, lValueType As Long, lDataBufSize As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub savekey(Hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetRegString(Hkey As Long, strPath As String, strValue As String)
    Dim datatype As Long
    Dim strBuf As String
    Dim intZeroPos As Integer
    
    ' Open the key
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        ' Create a buffer
        strBuf = String(lDataBufSize, " ")
        ' Retrieve the key's content
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetRegString = Left$(strBuf, intZeroPos - 1)
            Else
                GetRegString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveRegString(Hkey As Long, strPath As String, strValue As String, strdata As String)
    ' Create a new key
    r = RegCreateKey(Hkey, strPath, keyhand)
    ' Save a string to the key
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    ' Close the key
    r = RegCloseKey(keyhand)
End Sub

Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    r = RegDeleteKey(Hkey, strKey)
End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function
