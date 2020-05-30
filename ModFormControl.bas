Attribute VB_Name = "ModFormControl"
Option Explicit

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Public Const ILD_TRANSPARENT = &H1

'Windows Position Constants
Const HWND_NOTOPMOST    As Long = -2
Const HWND_TOPMOST      As Long = -1
Const SWP_NOMOVE        As Long = &H2
Const SWP_NOSIZE        As Long = &H1
Const TOPMOST_FLAGS     As Long = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExefileName As String, ByVal nIconIndex As Long) As Long

Public LV As ListItem

Public Sub MakeTop(HWND As Long, Top As Boolean)

Select Case Top
Case True
    SetWindowPos HWND, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
Case False
    SetWindowPos HWND, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Select

End Sub

Public Function MyTitle() As String
Dim AppNow As String
Dim i As Long
Dim SplitChar As String
Dim lRnd As Integer
Dim ResTitle As String

AppNow = "HerAV " & App.Major & "." & App.Minor
For i = 1 To Len(AppNow)
    Randomize
    lRnd = CInt(Rnd * 3) + 1
    Select Case lRnd
    Case 1: SplitChar = Chr$(160) & " "
    Case 2: SplitChar = " " & Chr$(160)
    Case 3: SplitChar = Chr$(160) & Chr$(160)
    Case 4: SplitChar = "  "
    End Select
    ResTitle = ResTitle & SplitChar & Mid$(AppNow, i, 1)
Next i

MyTitle = ResTitle

End Function

Public Function GetSelLV(TheLV As ListView, TheSubItem As Long, Optional StringOrLong As Long = 0) As String
    Dim Res As String
    Dim i As Long
    
    Select Case StringOrLong
    Case 0
        For i = 1 To TheLV.ListItems.Count
            If TheLV.ListItems(i).Selected = True Then Res = Res & "|" & GetLV(TheLV, i, TheSubItem + 1)
            DoEvents
        Next i
    Case 1
        For i = 1 To TheLV.ListItems.Count
            If TheLV.ListItems(i).Selected = True Then Res = Res & "|" & i
            DoEvents
        Next i
    End Select
    
    If Res <> "" Then GetSelLV = Right$(Res, Len(Res) - 1)
End Function

Public Function GetAllList(TheLV As ListView, TheSubItem As Long, Optional StringOrLong As Long = 0) As String
    Dim Res As String
    Dim i As Long
    
    Select Case StringOrLong
    Case 0
        For i = 1 To TheLV.ListItems.Count
            Res = Res & "|" & TheLV.ListItems(i).SubItems(TheSubItem)
            DoEvents
        Next i
    Case 1
        For i = 1 To TheLV.ListItems.Count
            Res = Res & "|" & i
            DoEvents
        Next i
    End Select
    
    If Res <> "" Then GetAllList = Right$(Res, Len(Res) - 1)
End Function

Public Sub AddToLV(TheLV As ListView, AllColumn As String, YourImageList As ImageList, Optional PathFile As Long, Optional PathString As String, Optional TheSel As Boolean = False, Optional clFailIfExist As Long = 2, Optional LVTag As String)

'On Error Resume Next
Dim SplitColumn() As String
Dim LItem As ListItem
Dim ImageIndex As Integer
Dim i As Long

SplitColumn = Split(AllColumn, "|")

If clFailIfExist <> 0 Then
    If LVExist(TheLV, clFailIfExist, SplitColumn(1)) = True Then Exit Sub
End If

If PathString = "" Then
    ImageIndex = DrawFileIcon(SplitColumn(PathFile), YourImageList)
Else
    ImageIndex = DrawFileIcon(PathString, YourImageList)
End If

Set LV = TheLV.ListItems.Add(, , SplitColumn(0), , ImageIndex)

If LVTag <> "" Then LV.Tag = LVTag

For i = 1 To UBound(SplitColumn)
    LV.SubItems(i) = SplitColumn(i)
    LV.Selected = TheSel
Next i

End Sub

Public Sub EditLV(TheLV As ListView, ByVal Row As Long, ByVal Column As Long, Text As String)
    Select Case Column
    Case 0
        TheLV.ListItems(Row).Text = Text
    Case Else
        TheLV.ListItems(Row).SubItems(Column) = Text
    End Select
End Sub

Public Function LVExist(TheLV As ListView, ByVal Column As Long, TextAdd As String) As Boolean
Dim i As Long

LVExist = False
If TheLV.ListItems.Count <> 0 Then
    For i = 1 To TheLV.ListItems.Count
        If UCase$(GetLV(TheLV, i, Column)) = UCase$(TextAdd) Then
            LVExist = True
            Exit Function
        End If
    Next i
End If

End Function

Public Function GetLV(TheLV As ListView, ByVal Row As Long, ByVal Column As Long) As String
    Select Case Column - 1
    Case 0
        GetLV = TheLV.ListItems(Row).Text
    Case Else
        GetLV = TheLV.ListItems(Row).SubItems(Column - 1)
    End Select
End Function

Public Function DrawFileIcon(ByVal lpFileName As String, ImageListToAdd As ImageList) As Long
Dim SmallIcon As Long
Dim FileInfoNow As SHFILEINFO
Dim NewImage As ListImage
Dim IconIndex As Integer

SmallIcon = SHGetFileInfo(lpFileName, 0&, FileInfoNow, Len(FileInfoNow), SHGFI.C_BASIC_FLAGS Or SHGFI.C_SMALLICON)

If SmallIcon <> 0 Then
    With FormMain.picBuffer
      .Picture = LoadPicture("")
      .AutoRedraw = True
      SmallIcon = ImageList_Draw(SmallIcon, FileInfoNow.iIcon, .hDC, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    
    IconIndex = ImageListToAdd.ListImages.Count + 1
    Set NewImage = ImageListToAdd.ListImages.Add(IconIndex, , FormMain.picBuffer.Image)

    DrawFileIcon = IconIndex
End If

End Function

Public Sub LVColumnClick(TheLV As ListView, ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    TheLV.SortKey = ColumnHeader.Index - 1
    If TheLV.SortOrder = 1 Then
        TheLV.SortOrder = 0
    Else
        TheLV.SortOrder = 1
    End If
End Sub
