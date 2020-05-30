Attribute VB_Name = "ModProcess"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWND As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As Any, ReturnLength As Any) As Long

Private Type LUID
   lowpart As Long
   highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    LuidUDT As LUID
    Attributes As Long
End Type

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2

Public Const PROCESS_TERMINATE As Long = &H1
Public Const PROCESS_CREATE_THREAD As Long = &H2
Public Const PROCESS_SET_SESSIONID As Long = &H4
Public Const PROCESS_VM_OPERATION As Long = &H8
Public Const PROCESS_VM_READ As Long = &H10
Public Const PROCESS_VM_WRITE As Long = &H20
Public Const PROCESS_DUP_HANDLE As Long = &H40
Public Const PROCESS_CREATE_PROCESS As Long = &H80
Public Const PROCESS_SET_QUOTA As Long = &H100
Public Const PROCESS_SET_INFORMATION As Long = &H200
Public Const PROCESS_QUERY_INFORMATION As Long = &H400

Public ProcessCount As Long
Public ProcessPath()    As String
Public ProcessId() As Long

Public Sub GetAccessSystem()
    Dim hProcess As Long
    Dim hToken As Long
    Dim tp As TOKEN_PRIVILEGES
    
    If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
        GoTo CleanUp
    End If
        
    If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
        GoTo CleanUp
    End If
    
    tp.PrivilegeCount = 1
    tp.Attributes = SE_PRIVILEGE_ENABLED
    
    If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
        GoTo CleanUp
    End If
    
CleanUp:
    If hToken Then CloseHandle hToken
End Sub

Private Function ProcessPathByPID(PID As Long) As String
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
    ModuleName = Space(260)
    nSize = 500
    ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
    ProcessPathByPID = Left(ModuleName, ret)

    ret = CloseHandle(hProcess)
    If ProcessPathByPID = "" Then ProcessPathByPID = "SYSTEM"
End Function

Public Sub UpdateProcessList()

    Dim cb As Long, cbNeeded  As Long: cb = 8: cbNeeded = 96
    Dim ProcessIDs() As Long
    Dim PathNow As String
    Dim lRet  As Long, i As Long
    
    Do While cb <= cbNeeded
       cb = cb * 2
       ReDim ProcessIDs(cb / 4) As Long
       lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
    
    ProcessCount = 0
    For i = 1 To cb / 4
        PathNow = FixPath(ProcessPathByPID(ProcessIDs(i)))
        If PathFileExists(PathNow) <> 0 Then
            ReDim Preserve ProcessPath(ProcessCount)
            ReDim Preserve ProcessId(ProcessCount)
            ProcessId(ProcessCount) = ProcessIDs(i)
            ProcessPath(ProcessCount) = PathNow
            ProcessCount = ProcessCount + 1
        End If
    Next

End Sub

Public Function FixPath(ByVal lpTestPath As String) As String
    Dim ResPath As String
    ResPath = lpTestPath
        If InStrRev(LCase$(ResPath), "\systemroot\") <> 0 Then
            ResPath = Replace$(LCase$(ResPath), "\systemroot", WindowsPath)
        End If
        If InStrRev(LCase$(lpTestPath), "\??\") <> 0 Then
            ResPath = Replace$(ResPath, "\??\", "")
        End If
    FixPath = ResPath
End Function

Public Function KillProcessByID(ByVal PID As Long) As Long
Dim lnghProcess As Long
Dim lngReturn As Long
    GetAccessSystem
    lnghProcess = OpenProcess(1&, -1&, PID)
    lngReturn = TerminateProcess(lnghProcess, 0&)
    KillProcessByID = lngReturn
End Function

Public Function GetModule(ByVal lpProcessID As String, ByRef lbGetModule() As String, Optional KillVB As Boolean = False) As Long

Dim i As Long
Dim Modules(1 To 1024) As Long
Dim lRet  As Long
Dim hProcess As Long
Dim iModDlls As Long
Dim sModName  As String
Dim sChildModName As String
Dim ModuleName As String
Dim VBAppCount As Long
Dim ModuleCount As Long
Dim PID() As String

PID = Split(lpProcessID, "|")

For i = 0 To UBound(PID)
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID(i))
        
        lRet = EnumProcessModules(hProcess, Modules(1), 1024, 0)
        lRet = EnumProcessModules(hProcess, Modules(1), 0, 0)
        
        ModuleName = Space(MAX_PATH)
        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, Len(ModuleName))
        sModName = Left$(ModuleName, lRet)
        iModDlls = 1
        Do
            iModDlls = iModDlls + 1
            ModuleName = Space(MAX_PATH)
            If iModDlls > 1024 Then Exit Do
            lRet = GetModuleFileNameExA(hProcess, Modules(iModDlls), ModuleName, Len(ModuleName))
            sChildModName = Left$(ModuleName, lRet)
            
            If sChildModName <> "" Then
                ReDim Preserve lbGetModule(ModuleCount) As String
                If UCase$(sChildModName) = UCase$(sModName) Then
                    lbGetModule(ModuleCount) = GetFileName(sModName) & "|" & GetFileType(FixPath(sModName)) & "|" & FixPath(sModName)
                    ModuleCount = ModuleCount + 1
                    Exit Do
                Else
                    lbGetModule(ModuleCount) = GetFileName(sModName) & "|" & GetFileType(sChildModName) & "|" & sChildModName
                    ModuleCount = ModuleCount + 1
                    If KillVB = True Then
                        If UCase$(sModName) <> UCase$(MyFilePath) Then
                            If UCase$(GetFileName(sChildModName)) = VBDLL Then
                                KillProcessByID PID(i)
                                VBAppCount = VBAppCount + 1
                            End If
                        End If
                    End If
                End If
            End If
        Loop
    lRet = CloseHandle(hProcess)
Next i

If KillVB = True Then
    GetModule = VBAppCount
Else
    GetModule = ModuleCount
End If

End Function

Public Function ProcessExist(ByVal PID As Long) As Boolean
Dim i As Long
UpdateProcessList
ProcessExist = False
For i = 0 To UBound(ProcessId)
    If PID = ProcessId(i) Then
    ProcessExist = True
    Exit Function
    End If
Next i
End Function
