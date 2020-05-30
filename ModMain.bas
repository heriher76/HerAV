Attribute VB_Name = "ModMain"
Option Explicit

Public Enum SPT
    SmadTempDir = 1
    RegExePath = 2
    StartupDirCurUser = 3
    StartupDirAllUser = 4
End Enum

Public Const KEY_CUR_WIN As String = "Software\Microsoft\Windows\CurrentVersion"
Public Const KEY_CUR_WINNT As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Public Const VBDLL As String = "MSVBVM60.DLL"
Public Const LAST_UPDATE As String = "15-May-2016"
Public Const APP_EXT As String = "exe|vbs|dll|ocx|bat|pif|lnk|scr|cmd|com"

Public Sub MakeDef()

    If PathFileExists(GetSpecPath(SmadTempDir)) = 0 Then
        MkDir GetSpecPath(SmadTempDir)
        SetFileAttributes GetSpecPath(SmadTempDir), ATTR.tNORMAL
    End If

End Sub

Public Function GetSpecPath(ByVal SpecPathType As SPT) As String
    Select Case SpecPathType
    Case 1
        GetSpecPath = GetDir(GetSpecialfolder(DIR_PERSONAL)) & "Local Settings\Temp\HerAV_tmp"
        If PathFileExists(GetSpecPath) = 0 Then
            GetSpecPath = App.Path
        End If
    Case 2
        GetSpecPath = WindowsPath & "\Regedit.exe"
    Case 3
        GetSpecPath = GetSpecialfolder(DIR_STARTUP)
    Case 4
        GetSpecPath = Replace(GetSpecialfolder(DIR_STARTUP), Trim$(UserAktif), "All Users")
    End Select
End Function
