Attribute VB_Name = "ModFile"
Option Explicit

'Semua Fungsi/Sub API yang dipakai untuk menangani file
'==============================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hmem As Long)

Public Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
'==============================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================================

Public Type FileTime
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const SW_SHOWNORMAL = 1
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FileTime
    ftLastAccessTime    As FileTime
    ftLastWriteTime     As FileTime
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Enum BIF
    BROWSEFORCOMPUTER = &H1000
    BROWSEFORPRINTER = &H2000
    BROWSEINCLUDEFILES = &H4000
    BROWSEINCLUDEURLS = &H80
    DONTGOBELOWDOMAIN = &H2
    EDITBOX = &H10
    NEWDIALOGSTYLE = &H40
    RETURNFSANCESTORS = &H8
    RETURNONLYFSDIRS = &H1
    SHAREABLE = &H8000
    STATUSTEXT = &H4
    USENEWUI = &H40
    VALIDATE_BIF = &H20
End Enum

Enum SFolder
    DIR_DESKTOP = &H0
    DIR_PROGRAMS = &H2
    DIR_CONTROLS = &H3
    DIR_PRINTERS = &H4
    DIR_PERSONAL = &H5
    DIR_FAVORITES = &H6
    DIR_STARTUP = &H7
    DIR_RECENT = &H8
    DIR_SENDTO = &H9
    DIR_BITBUCKET = &HA
    DIR_STARTMENU = &HB
    DIR_DESKTOPDIRECTORY = &H10
    DIR_DRIVES = &H11
    DIR_NETWORK = &H12
    DIR_NETHOOD = &H13
    DIR_FONTS = &H14
    DIR_TEMPLATES = &H15
End Enum

Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Enum ATTR
    tARCHIVE = &H20
    tDIRECTORY = &H10
    tHIDDEN = &H2
    tNORMAL = &H80
    tREADONLY = &H1
    tSYSTEM = &H4
    tTEMPORARY = &H100
End Enum

Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    HWND As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Enum SHGFI
    C_LARGEICON = &H0
    C_USEFILEATTRIBUTES = &H10
    C_TYPENAME = &H400
    C_SMALLICON = &H1
    C_SYSICONINDEX = &H4000
    C_SHELLICONSIZE = &H4
    C_DISPLAYNAME = &H200
    C_EXETYPE = &H2000
    C_BASIC_FLAGS = SHGFI.C_TYPENAME Or SHGFI.C_SHELLICONSIZE Or SHGFI.C_SYSICONINDEX Or SHGFI.C_DISPLAYNAME Or SHGFI.C_EXETYPE
End Enum

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

'Mengakses file dengan default program untuk jenis file tersebut
Public Sub FastShell(ByVal lpFileName As String, ByVal ExeCommand As String)

Dim GetDir As String

GetDir = Left$(lpFileName, 3)
ShellExecute GetDesktopWindow, "Open", lpFileName, ExeCommand, GetDir, SW_SHOWNORMAL

End Sub

Public Sub ExploreDir(ByVal lpFileName As String)

ShellExecute GetDesktopWindow, "explore", lpFileName, "", "", SW_SHOWNORMAL

End Sub

'Mendapatkan string buffer dari fungsi API
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

'Mendapatkan nama file dari path
Public Function GetFileName(PathFile As String) As String
Dim i As Long
Dim CutDirString As Long
    
    For i = 1 To Len(PathFile)
        If Mid$(PathFile, i, 1) = "\" Then CutDirString = i
    Next i
    GetFileName = Right$(PathFile, Len(PathFile) - CutDirString)
End Function

'Mendapatkan directory induk dari path
Public Function GetDir(PathFile As String) As String
Dim i As Long
Dim CutDirString As Long
    
    For i = 1 To Len(PathFile)
        If Mid$(PathFile, i, 1) = "\" Then CutDirString = i
    Next i
    GetDir = Left$(PathFile, CutDirString)
End Function

'Mendapatkan isi (file atau subfolder) dari sebuah folder
Public Function GetContains(ByVal Path As String, ByRef OutFiles() As String, DirOrFile As Long) As Long

    Dim hSearch As Long, WFD As WIN32_FIND_DATA, AddAgr As Boolean
    Dim Result As Long, CurItem As String, nFiles As Long
    
    If Not Right$(Path, 1) = "\" Then Path = Path & "\"

    hSearch = FindFirstFile(Path & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
            AddAgr = False
                Select Case DirOrFile
                Case 0: If PathIsDirectory(Path & CurItem) = 0 Then AddAgr = True
                Case 1: If PathIsDirectory(Path & CurItem) <> 0 Then AddAgr = True
                Case 2: AddAgr = True
                End Select
                If AddAgr = True Then
                    ReDim Preserve OutFiles(nFiles)
                    OutFiles(nFiles) = Path & CurItem
                    nFiles = nFiles + 1
                End If
            End If
            Result = FindNextFile(hSearch, WFD)
        DoEvents
        Loop
        FindClose hSearch
    End If
GetContains = nFiles
End Function

'Mendapatkan Special Folder (Folder Khusus) yang ada di sistem komputer
Public Function GetSpecialfolder(FolderType As SFolder) As String
    Dim r As Long
    Dim lpBuffer As String
    
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, FolderType, IDL)
        lpBuffer = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal lpBuffer)
        GetSpecialfolder = StripNulls(lpBuffer)
End Function

'Mendapatkan nama user yang sedang aktif sekarang
Public Function UserAktif() As String
Dim uTemp As String
    uTemp = GetSpecialfolder(DIR_PERSONAL)
    uTemp = GetDir(uTemp)
    uTemp = Left$(uTemp, Len(uTemp) - 1)
    uTemp = GetFileName(uTemp)
    UserAktif = uTemp
End Function

'Menentukan path adalah file atau bukan
Public Function IsFile(ByVal lpFileName As String) As Boolean
    If PathFileExists(lpFileName) = 1 And PathIsDirectory(lpFileName) = 0 Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function

'Mendapatkan lokasi folder sistem
Public Function SystemPath() As String
    Dim buffer As String * 255
    Dim Temp As Long

    Temp = GetSystemDirectory(buffer, 255)
    SystemPath = Left(buffer, Temp)
End Function

'Mendapatkan lokasi folder windows
Public Function WindowsPath() As String
    Dim buffer As String * 255
    Dim Temp As Long
    
    Temp = GetWindowsDirectory(buffer, 255)
    WindowsPath = Left(buffer, Temp)
End Function

'Menampilkan common dialog untuk penentuan file oleh user
Public Function BrowseForFile(hWndOwner As Long, sFilter As String, sTitle As String, OFName As OPENFILENAME, Optional nFlags As Long = OFN_EXPLORER) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(4999, vbNullChar)
    OFName.nMaxFile = 5000
    OFName.lpstrFileTitle = String(4999, vbNullChar)
    OFName.nMaxFileTitle = 5000
    OFName.lpstrTitle = sTitle
    OFName.Flags = nFlags

    If GetOpenFileName(OFName) Then
        BrowseForFile = StripNulls(OFName.lpstrFile)
    Else
        BrowseForFile = ""
    End If

End Function

'Menampilkan browse folder dialog untuk penentuan folder oleh user
Public Function BrowseForFolder(hWndOwner As Long, sTitle As String) As String
    
Dim BInfo As BROWSEINFO
Dim lpIDList As Long

    With BInfo
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF.EDITBOX
    End With

    lpIDList = SHBrowseForFolder(BInfo)
    If lpIDList Then
        BrowseForFolder = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, BrowseForFolder
        CoTaskMemFree lpIDList
        BrowseForFolder = StripNulls(BrowseForFolder)
    End If
End Function

'Mendapatkan drive yang aktif
Public Function ExistDrive() As String
    Dim sDrive As String, i As Long
    
    sDrive = ""
    For i = 65 To 90
        sDrive = Chr(i) & ":\"
        If PathFileExists(sDrive) <> 0 Then
            ExistDrive = ExistDrive & "|" & sDrive
        End If
    Next i
    ExistDrive = Right$(ExistDrive, Len(ExistDrive) - 1)
End Function

'Mendapatkan lokasi file program ini
Public Function MyFilePath() As String
Dim AppExt() As String
Dim i As Long

AppExt = Split(APP_EXT, "|")
For i = 0 To UBound(AppExt)
    If PathFileExists(App.Path & App.EXEName & "." & AppExt(i)) <> 0 Then
        MyFilePath = App.Path & App.EXEName & "." & AppExt(i)
        Exit Function
    End If
        If PathFileExists(App.Path & "\" & App.EXEName & "." & AppExt(i)) <> 0 Then
        MyFilePath = App.Path & "\" & App.EXEName & "." & AppExt(i)
        Exit Function
    End If
Next i

End Function

Public Function GetExt(ByVal lpFileName As String)
Dim sTemp As String
Dim i As Long

sTemp = GetFileName(lpFileName)
    If InStr(lpFileName, ".") Then
        For i = 0 To Len(sTemp) - 1
            If Mid$(sTemp, Len(sTemp) - i, 1) = "." Then
                GetExt = Mid$(sTemp, Len(sTemp) - i + 1, i)
                Exit Function
            End If
        Next i
    End If
End Function

Public Sub ShowProps(FileName As String, OwnerhWnd As Long)

    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .HWND = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
     ShellExecuteEx SEI
End Sub

Public Function SystemDrive() As String
    SystemDrive = Left$(SystemPath, 3)
End Function

Public Function SuperDelete(ByVal lpPath As String) As Long
    If IsFile(lpPath) = True Then
        SetFileAttributes lpPath, ATTR.tNORMAL
        SuperDelete = DeleteFile(lpPath)
    Else
        SuperDelete = RemoveDirectory(lpPath)
    End If
End Function

Public Function GetFileType(ByVal lpFileName As String) As String
Dim hType As Long
Dim FileInfoNow As SHFILEINFO

hType = SHGetFileInfo(lpFileName, 0&, FileInfoNow, Len(FileInfoNow), SHGFI.C_TYPENAME)

If hType <> 0 Then
    GetFileType = StripNulls(FileInfoNow.szTypeName)
Else
    GetFileType = ""
End If

End Function
