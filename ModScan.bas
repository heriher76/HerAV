Attribute VB_Name = "ModScan"
Option Explicit

Public FileCount As Long
Public FileDelete As Long
Public FilePathNow As String
Public DirToScan As String
Public FileToScan As String
Public ScanInfo As String
Public ScanFinish As Boolean
Public ScanIndex As Long

Public Sub FindFiles(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim RealPath As String
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If ScanFinish = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    If SubDirs = True Then
                        FindFiles RealPath & CurItem, True
                    End If
                Else
                    If WFD.nFileSizeLow > 5120 Or WFD.nFileSizeHigh > 5120 Then
                        ScanInfo = "Scan File"
                        ScanFile RealPath & CurItem
                    End If
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
    End If
End Sub

Public Sub FindFilesEx(ByVal lpFolderName As String, ByVal SubDirs As Boolean)
    Dim i As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    Dim tempDir() As String, dirCount As Long
    Dim RealPath As String
    
    dirCount = -1
    
    ScanInfo = "Scan File"
    
    If Right$(lpFolderName, 1) = "\" Then
        RealPath = lpFolderName
    Else
        RealPath = lpFolderName & "\"
    End If
    
    hSearch = FindFirstFile(RealPath & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            DoEvents
            If ScanFinish = True Then Exit Do
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                If PathIsDirectory(RealPath & CurItem) <> 0 Then
                    If SubDirs = True Then
                        dirCount = dirCount + 1
                        ReDim Preserve tempDir(dirCount) As String
                        tempDir(dirCount) = RealPath & CurItem
                    End If
                Else
                    If WFD.nFileSizeLow > 5120 Or WFD.nFileSizeHigh > 5120 Then
                        ScanFile RealPath & CurItem
                    End If
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        Loop
        FindClose hSearch
        
        If SubDirs = True Then
            If dirCount <> -1 Then
                For i = 0 To dirCount
                    FindFilesEx tempDir(i), True
                Next i
            End If
        End If
    End If
End Sub

Public Sub ScanFile(ByVal lpFileName As String)

Dim GetViriName As String
Dim i As Long, IconID As String, hDel As Long
Dim AddText As String

DoEvents
If ScanFinish = True Then Exit Sub

GetViriName = AnalyzeFile(lpFileName)
    If GetViriName <> "" Then
        If Right$(GetViriName, 8) = ".Variant" Then
            AddText = GetViriName & "|" & lpFileName & "|Executable File" & "|Virus Variant,  Please send this file to my e-mail : Zainuddin_Nafarin@Yahoo.co.id"
        Else
            AddText = GetViriName & "|" & lpFileName & "|Executable File"
        End If
        AddToLV FormScan.LVDetect, AddText, FormScan.ImgSmall, 1, , False
    End If
FileCount = FileCount + 1
FilePathNow = lpFileName

End Sub

Public Sub ScanStartUp()

ScanInfo = "Scan StartUp"

'Registry StartUp
ScanRegKey HKEY_CURRENT_USER, KEY_CUR_WIN & "\Run", False
ScanRegKey HKEY_CURRENT_USER, KEY_CUR_WIN & "\Run-", False
ScanRegKey HKEY_CURRENT_USER, KEY_CUR_WIN & "\RunOnce", False
ScanRegKey HKEY_CURRENT_USER, KEY_CUR_WIN & "\policies\Explorer\Run", False
ScanRegValue HKEY_CURRENT_USER, KEY_CUR_WINNT & "\Windows", "Load", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\Run", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\Run-", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\RunOnce", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\policies\Explorer\Run", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\RunOnce-", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\RunOnceEx", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\RunOnceServices", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\RunOnceServices-", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Run", False
ScanRegValue HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Windows", "Load", False
ScanRegValue HKEY_CURRENT_USER, "\Control Panel\Desktop", "SCRNSAVE.EXE", False
ScanRegKey HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Image File Execution Options", False

'Folder StartUp
FindFiles GetSpecPath(StartupDirAllUser), True
FindFiles GetSpecPath(StartupDirCurUser), True

End Sub

Public Sub ScanRegKey(ByVal MainKeyHandle As REG, SubKey As String, ByVal OnlyScan As Boolean, Optional ByVal ChildSubKey As Boolean = True)
Dim hKey As Long, Counter As Long
Dim sSave As String, Gets As String

RegOpenKey MainKeyHandle, SubKey, hKey
Counter = 0
ScanRegKeyEx MainKeyHandle, SubKey, OnlyScan
If ChildSubKey = True Then
    Do
        DoEvents
        If ScanFinish = True Then Exit Do
        sSave = String(255, " ")
        If RegEnumKey(hKey, Counter, sSave, 255) <> 0 Then Exit Do
        ScanRegKey MainKeyHandle, SubKey & "\" & StripNulls(sSave), ChildSubKey, OnlyScan
        Counter = Counter + 1
    Loop
End If
RegCloseKey hKey

End Sub

Public Sub ScanRegKeyEx(ByVal MainKeyHandle As REG, SubKey As String, ByVal OnlyScan As Boolean)
    Dim hKey As Long, Counter As Long, sSave As String
    Dim hType As Long
    
    RegOpenKey MainKeyHandle, SubKey, hKey
    Counter = 0
    Do
        DoEvents
        If ScanFinish = True Then Exit Sub
        sSave = String(255, " ")
        If RegEnumValue(hKey, Counter, sSave, 255, 0, hType, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        If hType = REG_SZ Then
            ScanRegValue MainKeyHandle, SubKey, StripNulls(sSave), OnlyScan
        End If
        Counter = Counter + 1
    Loop
    RegCloseKey hKey
End Sub

Public Function ScanRegValue(ByVal MainKeyHandle As REG, SubKey As String, ByVal Entry As String, ByVal OnlyScan As Boolean)
    Dim PathNow As String
    Dim sValueName As String
    Dim sTag As String
    Dim GetViriName As String

    PathNow = NoRegCommand(GetStringValue(MainKeyHandle, SubKey, Entry))
    sValueName = GetValueName(MainKeyHandle, SubKey, Entry)
    ScanInfo = sValueName
    If PathFileExists(PathNow) <> 0 Then
    FilePathNow = PathNow
    GetViriName = AnalyzeFile(PathNow)
        If GetViriName <> "" Then
            'Jika registry value hanya untuk di-scan maka lokasi value tidak perlu ditambahkan ke listview reg
            If OnlyScan = False Then
                sTag = "Delete|" & CStr(MainKeyHandle) & "|" & SubKey & "|" & Entry
                AddToLV FormScan.LVRegDet, GetFileName(sValueName) & "|" & sValueName & "|" & GetViriName & " Value", FormScan.ImgSmall, , GetSpecPath(RegExePath), , , sTag
            End If
            'Tambahkan lokasi virus ke listview
            AddToLV FormScan.LVDetect, GetViriName & "|" & PathNow & "|Executable File", FormScan.ImgSmall, 1
        End If
    End If
    
End Function

Public Sub ScanProcess()
Dim GetViriName As String, i As Long, VC As Long
Dim VirusExist() As String, AllViri As String
Dim VirusID() As Long
UpdateProcessList

ScanInfo = "Scan Process Memory"
For i = 0 To ProcessCount - 1
    DoEvents
    If PathFileExists(ProcessPath(i)) <> 0 Then
    GetViriName = AnalyzeFile(ProcessPath(i))
        If GetViriName <> "" Then
            AddToLV FormScan.LVDetect, GetViriName & "|" & ProcessPath(i) & "|Process File", FormScan.ImgSmall, , ProcessPath(i)
            AddToLV FormScan.LVProcDet, GetViriName & "|" & GetFileName(ProcessPath(i)) & "|" & ProcessId(i) & "|Killed", FormScan.ImgSmall, , ProcessPath(i)
            ReDim Preserve VirusExist(VC) As String
            ReDim Preserve VirusID(VC) As Long
            VirusID(VC) = ProcessId(i)
            VirusExist(VC) = GetViriName
            VC = VC + 1
        End If
        FilePathNow = ProcessPath(i)
    End If
Next i

If VC > 0 Then
    AllViri = VirusExist(0)
    For i = 1 To UBound(VirusExist)
        If InStr(AllViri & " , ", VirusExist(i) & " , ") = 0 Then
            AllViri = AllViri & " , " & VirusExist(i)
        End If
    Next i
    
LoopKill:
    For i = 0 To UBound(VirusID)
        KillProcessByID VirusID(i)
    Next i
    
    For i = 0 To UBound(VirusID)
        If ProcessExist(VirusID(i)) = True Then
            GoTo LoopKill
            Exit For
        End If
    Next i
    
    If MsgBox("System was infected by " & AllViri & " Virus! HerAV will scan your system", vbExclamation + vbYesNo, "System Infected") = vbYes Then
       If ScanIndex <> 2 Then ScanSystem
    End If
End If

End Sub

Public Sub FixReg()

    Dim i As Long
    Dim ShellExtScan() As String
    Dim ShellExtExe() As String
    
    Const SHELL_EXTS_EXE As String = "exefile|batfile|cmdfile|comfile|piffile|lnkfile|scrfile"
    Const SHELL_EXTS_OTHER As String = "regfile|txtfile|inffile|inifile|htmlfile|Word.Document.8"
    
    ScanInfo = "Fix Registry"
    
    FixRegString HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Winlogon", "Userinit", SystemPath & "\userinit.exe,"
    FixRegString HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Winlogon", "Shell", "Explorer.exe"
    FixRegString HKEY_LOCAL_MACHINE, KEY_CUR_WINNT & "\Winlogon", "System", ""
    FixRegString HKEY_CURRENT_USER, KEY_CUR_WINNT & "\Winlogon", "Shell", "Explorer.exe"
    
    FixRegString HKEY_CLASSES_ROOT, "exefile", "", "Application"
    FixRegString HKEY_CLASSES_ROOT, "scrfile", "", "Screen Saver"
    FixRegString HKEY_CURRENT_USER, "Control Panel\International", "s1159", ""
    FixRegString HKEY_CURRENT_USER, "Control Panel\International", "s2359", ""

    ShellExtExe = Split(SHELL_EXTS_EXE, "|")
    For i = 0 To UBound(ShellExtExe)
        FixRegString HKEY_CLASSES_ROOT, ShellExtExe(i) & "\shell\Open\Command", "", Chr$(34) & "%1" & Chr$(34) & " %*"
    Next i
    
    ShellExtScan = Split(SHELL_EXTS_EXE & "|" & SHELL_EXTS_OTHER, "|")
    For i = 0 To UBound(ShellExtScan)
        ScanRegKey HKEY_CLASSES_ROOT, ShellExtScan(i), True
    Next i
    
    FixRegString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot", "AlternateShell", "cmd.exe"
    FixRegString HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\SafeBoot", "AlternateShell", "cmd.exe"
    FixRegString HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet002\Control\SafeBoot", "AlternateShell", "cmd.exe"
    FixRegString HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet003\Control\SafeBoot", "AlternateShell", "cmd.exe"

    FixRegDWORD HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\Explorer\Advanced\Folder\HideFileExt", "UncheckedValue", 0
    FixRegDWORD HKEY_LOCAL_MACHINE, KEY_CUR_WIN & "\Explorer\Advanced\Folder\SuperHidden", "UncheckedValue", 1
    
End Sub

Public Sub FixRegDWORD(MainKeyHandle As REG, SubKey As String, Entry As String, ByVal FixDWORD As Long)
Dim lTemp As Long
Dim sValueName As String
Dim sTag As String

lTemp = GetDWORDValue(MainKeyHandle, SubKey, Entry)

If lTemp <> FixDWORD Then
    sValueName = GetValueName(MainKeyHandle, SubKey, Entry)
    sTag = "FixDWORD|" & CStr(MainKeyHandle) & "|" & SubKey & "|" & Entry & "|" & FixDWORD
    AddToLV FormScan.LVRegDet, Entry & "|" & sValueName & "|Suspected DWORD Value", FormScan.ImgSmall, , GetSpecPath(RegExePath), , , sTag
End If

End Sub

Public Sub FixRegString(MainKeyHandle As REG, SubKey As String, Entry As String, FixString As String)

Dim sTemp As String
Dim sValueName As String
Dim sTag As String

sTemp = Trim$(GetStringValue(MainKeyHandle, SubKey, Entry))

If UCase$(sTemp) <> UCase$(FixString) Then
sTag = "FixString|" & CStr(MainKeyHandle) & "|" & SubKey & "|" & Entry & "|" & FixString
    sValueName = GetValueName(MainKeyHandle, SubKey, Entry)
    AddToLV FormScan.LVRegDet, GetFileName(sValueName) & "|" & sValueName & "|Suspected String Value", FormScan.ImgSmall, , GetSpecPath(RegExePath), , , sTag
End If

End Sub

Public Function NoRegCommand(RegPath As String) As String
    Dim sTemp As String, PathCheck As String, i As Long
    sTemp = Replace(RegPath, Chr$(34), "")
    
    For i = 0 To Len(sTemp) - 1
    PathCheck = Left$(sTemp, Len(sTemp) - i)
        If IsFile(PathCheck) = True Then
            NoRegCommand = PathCheck
            If AnalyzeFile(PathCheck) <> "" Then Exit Function
        End If
    Next i
End Function

Public Sub ScanSystem()
Dim DirNow() As String
Dim i As Long
    
    ScanStartUp
    FixReg
    
    FindFiles Left$(WindowsPath, 3), False
    FindFiles WindowsPath, True
    GetContains GetDir(GetSpecialfolder(DIR_PERSONAL)), DirNow(), 1
    For i = 0 To UBound(DirNow)
        If UCase$(DirNow(i)) <> UCase$(GetSpecialfolder(DIR_PERSONAL)) Then
            FindFiles DirNow(i), True
        End If
    Next i
    
End Sub
