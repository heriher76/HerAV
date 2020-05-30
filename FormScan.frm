VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormScan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6975
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7035
   Icon            =   "FormScan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox sIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1860
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   27
      Top             =   6240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   435
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Stop"
      Height          =   435
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdCleanAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Clean All"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Back"
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox PicTabMain 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   453
      TabIndex        =   11
      Top             =   2520
      Width           =   6795
      Begin ComctlLib.ProgressBar PBMain 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ListView LVDetect 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   255
         BackColor       =   16777215
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Virus Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Image Path"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Information"
            Object.Width           =   7056
         EndProperty
      End
      Begin ComctlLib.ListView LVProcDet 
         Height          =   2775
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Virus Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Process Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "PID"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Information"
            Object.Width           =   7056
         EndProperty
      End
      Begin ComctlLib.ListView LVRegDet 
         Height          =   2775
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4895
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   16711680
         BackColor       =   16777215
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value Path"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Information"
            Object.Width           =   3528
         EndProperty
      End
      Begin ComctlLib.TabStrip TabDet 
         Height          =   3435
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6059
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Detected Virus"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Registry Infected"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Virus Process"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrmCurProcess 
      BackColor       =   &H00400040&
      Caption         =   "Current Process"
      ForeColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   6795
      Begin VB.PictureBox PicFrmCurProc 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   120
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   433
         TabIndex        =   5
         Top             =   240
         Width           =   6495
         Begin VB.TextBox TxDeleted 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1230
            Width           =   2175
         End
         Begin VB.TextBox TxSpeed 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   900
            Width           =   2175
         End
         Begin VB.TextBox TxDetect 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1230
            Width           =   1995
         End
         Begin VB.TextBox TxScanned 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   900
            Width           =   1995
         End
         Begin VB.TextBox TxInfo 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   5415
         End
         Begin VB.TextBox TxFileName 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1020
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   540
            Width           =   5415
         End
         Begin VB.TextBox TxDirPath 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1020
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   0
            Width           =   5415
         End
         Begin VB.Timer TimerSpeed 
            Interval        =   400
            Left            =   5520
            Top             =   60
         End
         Begin VB.Timer TimerUp 
            Interval        =   10
            Left            =   6000
            Top             =   60
         End
         Begin VB.Label LbDirPathS 
            BackColor       =   &H00400040&
            Caption         =   "Dir Path :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   60
            Width           =   915
         End
         Begin VB.Label LbInfoS 
            BackColor       =   &H00400040&
            Caption         =   "Information :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   1560
            Width           =   915
         End
         Begin VB.Label LbSpeed 
            BackColor       =   &H00400040&
            Caption         =   "Speed :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.Label LbDetectCt 
            BackColor       =   &H00400040&
            Caption         =   "Detected :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   1260
            Width           =   795
         End
         Begin VB.Label LbScanCt 
            BackColor       =   &H00400040&
            Caption         =   "Scanned :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   8
            Top             =   960
            Width           =   795
         End
         Begin VB.Label LbFileNameS 
            BackColor       =   &H00400040&
            Caption         =   "File Name :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   600
            Width           =   795
         End
         Begin VB.Label LbDeleteCt 
            BackColor       =   &H00400040&
            Caption         =   "Deleted :"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   6
            Top             =   1260
            Width           =   675
         End
      End
   End
   Begin VB.Label LbSite 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2016 By HerAV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   195
      Left            =   4980
      MouseIcon       =   "FormScan.frx":1CCA
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   6660
      Width           =   1980
   End
   Begin ComctlLib.ImageList ImgSmall 
      Left            =   2460
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu PopDet 
      Caption         =   "&PopDet"
      Visible         =   0   'False
      Begin VB.Menu PopDetDelSel 
         Caption         =   "&Delete Selected"
      End
      Begin VB.Menu PopDetDelAll 
         Caption         =   "&Delete All"
      End
      Begin VB.Menu PopLine1 
         Caption         =   "-"
      End
      Begin VB.Menu PopDetExp 
         Caption         =   "&Open Containing Folder"
      End
      Begin VB.Menu PopLine2 
         Caption         =   "-"
      End
      Begin VB.Menu PopDetDelProp 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu PopReg 
      Caption         =   "&PopReg"
      Visible         =   0   'False
      Begin VB.Menu PopRegFixSel 
         Caption         =   "Fix\Clean &Selected"
      End
      Begin VB.Menu PopRegFixAll 
         Caption         =   "Fix\Clean &All"
      End
   End
End
Attribute VB_Name = "FormScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SEL_VIRI = 1
Private Const ALL_VIRI = 2

Dim LastTime As Long
Dim NewTime As Long
Dim AllTime As Long
Dim TimeSpeedCount As Long
Dim SpeedRate As Long
Dim ExitCode As Long
Dim ClickItem As Boolean

Private Sub cmdBack_Click()
'On Error Resume Next
    Form_Unload (0)
End Sub

Private Sub cmdCleanAll_Click()
'On Error Resume Next
    CleanRegViri ALL_VIRI
    DeleteViri ALL_VIRI
    RefreshForm
End Sub

Private Sub cmdExit_Click()
'On Error Resume Next
Unload Me
If ExitCode = vbYes Then
    FormMain.Visible = False
    End
End If
End Sub

Private Sub cmdStop_Click()
'On Error Resume Next
Dim Ask As Long
Select Case ScanFinish
Case False
    Ask = MsgBox("Are you sure you want to stop current process?", vbQuestion + vbYesNo, Me.Caption)
Case True
    Ask = vbYes
End Select

Select Case Ask
Case vbYes
    ScanFinish = True
Case vbNo
    ScanFinish = False
End Select

End Sub

Private Sub Form_Initialize()
'On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
Dim Ask As Long
Select Case ScanFinish
Case False
    Ask = MsgBox("Are you sure you want to back and stop current process?", vbQuestion + vbYesNo, Me.Caption)
Case True
    Ask = vbYes
End Select

ExitCode = Ask
    
Select Case Ask
Case vbYes
    ScanFinish = True
    LVDetect.ListItems.Clear
    Unload Me
    FormMain.Show
Case vbNo
    Cancel = 1
    Exit Sub
End Select

End Sub

Private Sub lbSite_Click()
'On Error Resume Next
   FormMain.cmdAbout_Click
End Sub

Private Sub LvDetect_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVDetect, ColumnHeader
End Sub

Private Sub LvDetect_ItemClick(ByVal Item As ComctlLib.ListItem)
'On Error Resume Next
    ClickItem = True
End Sub

Private Sub LvDetect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
    If Button = 2 Then
        If ClickItem = True Then
            If PathFileExists(LVDetect.SelectedItem.SubItems(1)) <> 0 Then
                PopupMenu PopDet, , , , PopDetDelSel
                ClickItem = False
            End If
        End If
    End If
End Sub

Private Sub LVProcDet_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVProcDet, ColumnHeader
End Sub

Private Sub LVProcDet_ItemClick(ByVal Item As ComctlLib.ListItem)
    ClickItem = True
End Sub

Private Sub LVRegDet_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVRegDet, ColumnHeader
End Sub

Private Sub LVRegDet_ItemClick(ByVal Item As ComctlLib.ListItem)
    ClickItem = True
End Sub

Private Sub LVRegDet_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
    If Button = 2 Then
        If ClickItem = True Then
            Select Case LVRegDet.SelectedItem.SubItems(3)
            Case "Cleaned", "Fixed"
            Case Else
                PopupMenu PopReg, , , , PopRegFixSel
                ClickItem = False
            End Select
        End If
    End If
End Sub

Private Sub PopDetDelAll_Click()
'On Error Resume Next
    DeleteViri ALL_VIRI
End Sub

Private Sub PopDetDelProp_Click()
'On Error Resume Next
    ShowProps LVDetect.SelectedItem.SubItems(1), Me.HWND
End Sub

Private Sub PopDetDelSel_Click()
'On Error Resume Next
    DeleteViri SEL_VIRI
End Sub

Private Sub PopDetExp_Click()
'On Error Resume Next
    ExploreDir GetDir(LVDetect.SelectedItem.SubItems(1)) & "\"
End Sub

Private Sub PopRegFixAll_Click()
'On Error Resume Next
    CleanRegViri ALL_VIRI
End Sub

Private Sub PopRegFixSel_Click()
'On Error Resume Next
    CleanRegViri SEL_VIRI
End Sub

Private Sub TabDet_Click()
'On Error Resume Next
    TabClick TabDet.SelectedItem.Index
End Sub

Private Sub TimerSpeed_Timer()
'On Error Resume Next
Dim CurSpeed As String

If ScanFinish = False Then
    NewTime = FileCount
    CurSpeed = Int((NewTime - LastTime) * 2.5)
    AllTime = AllTime + 1
    TimeSpeedCount = TimeSpeedCount + 1
    TxSpeed.Text = Abs(CurSpeed) & " File/s"
    LastTime = NewTime
Else
    LbSpeed.Caption = "Speed Rate : "
    SpeedRate = Abs(Int(FileCount / (AllTime * 0.4)))
    TxSpeed.Text = SpeedRate & " File/s"
    TimerSpeed.Enabled = False
End If

End Sub

Private Sub TimerUp_Timer()
'On Error Resume Next
Dim ToAdd As String
    If ScanFinish = True Then
        TxFileName.Text = ""
        TxDirPath.Text = ""
        TxInfo.Text = "Scan Finished"
        TxScanned.Text = FileCount & " (Finished)"
        TxDetect.Text = LVDetect.ListItems.Count & " (Finished)"
        cmdStop.Enabled = False
        TimerUp.Enabled = False
        cmdCleanAll.Enabled = True
    Else
        TxFileName.Text = GetFileName(FilePathNow)
        TxDirPath.Text = GetDir(FilePathNow)
        TxInfo.Text = ScanInfo
        TxScanned.Text = FileCount
        TxDetect.Text = LVDetect.ListItems.Count
        TxDeleted.Text = FileDelete
    End If

End Sub

Public Sub ScanNow(Optional ByVal lpTypeScan As Long = 0)
'On Error Resume Next
Dim FileNow() As String
Dim DirNow() As String
Dim i As Long

    MakeForm
    ScanProcess
    
    Select Case lpTypeScan
    Case 0
        If FormMain.CkSystemScan.Value = 1 Then
            ScanSystem
        Else
            ScanStartUp
            If FormMain.CkRegFix.Value = 1 Then FixReg
        End If
    Case 1
        ScanStartUp
        FixReg
    Case 2
        ScanSystem
    End Select
    
    If FileToScan <> "" Then
        FileNow = Split(FileToScan, "|")
        For i = 0 To UBound(FileNow)
            ScanFile FileNow(i)
        Next i
    End If
    
    If DirToScan <> "" Then
        DirNow = Split(DirToScan, "|")
        For i = 0 To UBound(DirNow)
            FindFilesEx DirNow(i), CBool(FormMain.CkSubDir.Value)
        Next i
    End If
    
    ScanFinish = True
End Sub

Private Sub TabClick(Index As Integer)
'On Error Resume Next
    LVDetect.Visible = False
    LVRegDet.Visible = False
    LVProcDet.Visible = False
    Select Case Index
    Case 1
        LVDetect.Visible = True
    Case 2
        LVRegDet.Visible = True
    Case 3
        LVProcDet.Visible = True
    End Select
    RefreshForm
End Sub

Public Sub MakeForm()
'On Error Resume Next
    FilePathNow = ""
    FileCount = 0
    FileDelete = 0
    ScanFinish = False
    
    Me.Show
    Me.iCon = FormMain.iCon
    FormMain.Hide
    Me.Caption = MyTitle
    TabClick 1
    
    TimerUp.Enabled = True
    TimerSpeed.Enabled = True
End Sub

Public Sub DeleteViri(ByVal wType As Long)
'On Error Resume Next
Dim i As Long, SL() As String
Dim LastFinish As Boolean
LVDetect.Enabled = False
LastFinish = ScanFinish
    If LastFinish = True Then ScanFinish = False
    
    'Dapatkan semua baris listview atau hanya yang dipilih saja
    Select Case wType
    Case SEL_VIRI
        SL = Split(GetSelLV(LVDetect, 1, 1), "|")
    Case ALL_VIRI
        SL = Split(GetAllList(LVDetect, 1, 1), "|")
    End Select
    
    'Delete semua virus pada baris listview yang didapatkan diatas
    PBMain.Visible = True
    For i = 0 To UBound(SL)
        DeleteViriEx SL(i)
        PBMain.Value = CInt((i + 1) * 100 / (UBound(SL) + 1))
        DoEvents
    Next i
    PBMain.Visible = False
    
LVDetect.Enabled = True
If LastFinish = True Then ScanFinish = True

End Sub

Public Sub CleanRegViri(ByVal wType As Long)
'On Error Resume Next
Dim i As Long, SL() As String
Dim LastFinish As Boolean
LVRegDet.Enabled = False
LastFinish = ScanFinish
    If LastFinish = True Then ScanFinish = False
            
    'Dapatkan semua baris listview atau hanya yang dipilih saja
    Select Case wType
    Case SEL_VIRI
        SL = Split(GetSelLV(LVRegDet, 1, 1), "|")
    Case ALL_VIRI
        SL = Split(GetAllList(LVRegDet, 1, 1), "|")
    End Select
    
    'Delete semua registry value pada baris listview yang didapatkan diatas
    For i = 0 To UBound(SL)
        CleanRegViriEx SL(i)
    Next i
    
LVRegDet.Enabled = True
If LastFinish = True Then ScanFinish = True

End Sub

Public Sub CleanRegViriEx(ByVal LVRow As Long)
'On Error Resume Next
Dim InfoTags As String
Dim InfoTag() As String
Dim i As Long

    InfoTags = LVRegDet.ListItems(LVRow).Tag
    If InfoTags <> "" Then
        InfoTag = Split(InfoTags, "|")
        Select Case InfoTag(0)
        Case "FixDWORD", "FixString"
            If InfoTag(0) = "FixDWORD" Then
                SetDWORDValue CLng(InfoTag(1)), InfoTag(2), InfoTag(3), CLng(InfoTag(4))
            ElseIf InfoTag(0) = "FixString" Then
                SetStringValue CLng(InfoTag(1)), InfoTag(2), InfoTag(3), InfoTag(4)
            End If
            EditLV LVRegDet, LVRow, 3, "Fixed"
        Case "Delete"
            DeleteValue CLng(InfoTag(1)), InfoTag(2), InfoTag(3)
            EditLV LVRegDet, LVRow, 3, "Cleaned"
        End Select
    End If
End Sub

Public Sub DeleteViriEx(ByVal LVRow As Long)
'On Error Resume Next
    Dim DelFile As String
    Dim hDel As Long, i As Long
    DelFile = GetLV(LVDetect, LVRow, 2)
    
    If PathFileExists(DelFile) <> 0 Then
        hDel = SuperDelete(DelFile)
        If hDel = 0 Then
            EditLV LVDetect, LVRow, 3, "Cannot Deleted"
        Else
            EditLV LVDetect, LVRow, 3, "Deleted"
            FileDelete = FileDelete + 1
            TxDeleted.Text = FileDelete
        End If
    Else
        EditLV LVDetect, LVRow, 3, "Deleted"
        FileDelete = FileDelete + 1
    End If
End Sub

Public Sub RefreshForm()
    If ScanFinish = True Then
        If FileDelete >= LVDetect.ListItems.Count Then
            cmdCleanAll.Enabled = False
        Else
            cmdCleanAll.Enabled = True
        End If
    Else
        cmdCleanAll.Enabled = False
    End If
End Sub
