Attribute VB_Name = "ModRegistry"
Option Explicit

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long

Public Enum REG
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_MULTI_SZ As Long = 7

Public Const KEY_WRITE = &H20000 Or &H2& Or &H4&
Public Const KEY_READ = &H20000 Or &H1& Or &H4& Or &H10&

Dim hKey As Long
Dim rtn As Long
Dim lBuffer As Long
Dim lBufferSize As Long

Public Function SetDWORDValue(MainKeyHandle As REG, SubKey As String, Entry As String, Value As Long) As Long
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
    SetDWORDValue = rtn
    rtn = RegCloseKey(hKey)
End Function

Public Function SetStringValue(MainKeyHandle As REG, SubKey As String, Entry As String, Value As String) As Long
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    rtn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value))
    SetStringValue = rtn
    rtn = RegCloseKey(hKey)
End Function
Public Function DeleteValue(MainKeyHandle As REG, SubKey As String, Entry As String) As Long
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey)
    rtn = RegDeleteValue(hKey, Entry)
    DeleteValue = rtn
    rtn = RegCloseKey(hKey)
End Function

Public Function GetDWORDValue(MainKeyHandle As REG, SubKey As String, Entry As String)
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
    rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4)
    rtn = RegCloseKey(hKey)
    GetDWORDValue = lBuffer
End Function

Public Function GetStringValue(MainKeyHandle As REG, SubKey As String, Entry As String)
Dim sGet As String
    rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey)
    sGet = String(255, " ")
    rtn = RegQueryValueEx(hKey, Entry, 0&, REG_SZ Or REG_EXPAND_SZ Or REG_MULTI_SZ, ByVal sGet, 255)
    GetStringValue = StripNulls(sGet)
    rtn = RegCloseKey(hKey)
End Function

Public Function GetValueName(MainKeyHandle As REG, SubKey As String, Entry As String) As String
Dim MainKeyString As String
Dim sEntry As String

    Select Case MainKeyHandle
        Case &H80000000: MainKeyString = "HKEY_CLASSES_ROOT"
        Case &H80000001: MainKeyString = "HKEY_CURRENT_USER"
        Case &H80000002: MainKeyString = "HKEY_LOCAL_MACHINE"
        Case &H80000003: MainKeyString = "HKEY_USERS"
        Case &H80000004: MainKeyString = "HKEY_PERFORMANCE_DATA"
        Case &H80000005: MainKeyString = "HKEY_CURRENT_CONFIG"
        Case &H80000006: MainKeyString = "HKEY_DYN_DATA"
    End Select
    
    If Entry = "" Then
        sEntry = "(Default)"
    Else
        sEntry = Entry
    End If
        
    GetValueName = MainKeyString & "\" & SubKey & "\" & sEntry
    GetValueName = Replace$(GetValueName, "\\", "\")
End Function
