VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "HerAV"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMain.frx":1CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerProcess 
      Interval        =   1000
      Left            =   5760
      Top             =   4320
   End
   Begin ComctlLib.StatusBar SBMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   5835
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9260
            MinWidth        =   9260
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox CkOnTop 
      BackColor       =   &H00400040&
      Caption         =   "&Always On Top"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   4560
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CommandButton cmdScan 
      BackColor       =   &H00FFFFFF&
      Caption         =   " &Scan "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame frmOptScan 
      BackColor       =   &H00400040&
      Caption         =   "Optional Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   1815
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         Begin VB.CheckBox CkRegFix 
            BackColor       =   &H00400040&
            Caption         =   "&Registry Value"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   600
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox CkSystemScan 
            BackColor       =   &H00400040&
            Caption         =   "Sys&tem Areas"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   6
            Top             =   300
            Width           =   1455
         End
         Begin VB.CheckBox CkSubDir 
            BackColor       =   &H00400040&
            Caption         =   "&Sub Directory"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   0
            Value           =   1  'Checked
            Width           =   1455
         End
      End
   End
   Begin VB.Frame frmAVEngine 
      BackColor       =   &H00400040&
      Caption         =   "AV Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   11
         Top             =   240
         Width           =   1875
         Begin VB.CommandButton cmdAbout 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&About"
            Height          =   435
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Update"
            Height          =   435
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
         Begin VB.Label LbDBCount 
            BackStyle       =   0  'Transparent
            Caption         =   "Virus DB : "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label LbVersion 
            BackStyle       =   0  'Transparent
            Caption         =   "Version : "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label LbUpdate 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   540
            Width           =   1395
         End
      End
   End
   Begin VB.PictureBox picBuffer 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4800
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   4740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   120
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   8
      Top             =   480
      Width           =   7575
      Begin ComctlLib.ProgressBar PBMain 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ListView LVDirFileScan 
         Height          =   2775
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Full Path"
            Object.Width           =   12330
         EndProperty
      End
      Begin ComctlLib.ListView LVProc 
         Height          =   2775
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Process Name"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "PID"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Path"
            Object.Width           =   12330
         EndProperty
      End
      Begin ComctlLib.ListView LVDBViri 
         Height          =   2775
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Virus Name"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Virus Checksum"
            Object.Width           =   3519
         EndProperty
      End
      Begin ComctlLib.ListView LVModule 
         Height          =   2775
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   4895
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         _Version        =   327682
         Icons           =   "ImgSmall"
         SmallIcons      =   "ImgSmall"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Module Name"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Parent Module"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   3519
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Path"
            Object.Width           =   12330
         EndProperty
      End
      Begin ComctlLib.TabStrip TabMain 
         Height          =   3495
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6165
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   4
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&File/Directory To Scan"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Virus Database"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "&Process Memory"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Process &Modules"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
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
      Left            =   5760
      MouseIcon       =   "FormMain.frx":1FD4
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label LbToScan 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00400040&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   4875
   End
   Begin ComctlLib.ImageList ImgSmall 
      Left            =   6360
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu PopProcMan 
      Caption         =   "&Process View"
      Visible         =   0   'False
      Begin VB.Menu PopProcManMod 
         Caption         =   "&Process Module"
      End
      Begin VB.Menu PopLine1 
         Caption         =   "-"
      End
      Begin VB.Menu PopProcManExp 
         Caption         =   "&Open Containing Folder"
      End
      Begin VB.Menu PopProcManKill 
         Caption         =   "&Terminate Process"
         Begin VB.Menu PopProcManKillSel 
            Caption         =   "&Selected Apps"
         End
         Begin VB.Menu PopProcManKillVB 
            Caption         =   "&VB Apps"
         End
      End
      Begin VB.Menu PopLine2 
         Caption         =   "-"
      End
      Begin VB.Menu PopProcManProp 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu PopFileView 
      Caption         =   "&File View"
      Visible         =   0   'False
      Begin VB.Menu PopFileViewAdd 
         Caption         =   "&Add"
         Begin VB.Menu PopFileViewAddFile 
            Caption         =   "&File"
         End
         Begin VB.Menu PopFileViewAddFolder 
            Caption         =   "&Directory"
         End
      End
      Begin VB.Menu PopFileViewAddExp 
         Caption         =   "&Add By Drag From Explorer"
      End
      Begin VB.Menu PopLine3 
         Caption         =   "-"
      End
      Begin VB.Menu PopFileViewRem 
         Caption         =   "&Remove"
      End
      Begin VB.Menu PopFileViewClear 
         Caption         =   "&Remove All"
      End
   End
   Begin VB.Menu PopModule 
      Caption         =   "&Module View"
      Visible         =   0   'False
      Begin VB.Menu PopModuleExp 
         Caption         =   "&Open Containing Folder"
      End
      Begin VB.Menu PopModuleProp 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu PopScan 
      Caption         =   "&Scan"
      Visible         =   0   'False
      Begin VB.Menu PopScanEx 
         Caption         =   "&Selected Directory/File"
         Index           =   0
      End
      Begin VB.Menu PopScanEx 
         Caption         =   "&Computer (Full Scan)"
         Index           =   1
      End
      Begin VB.Menu PopScanEx 
         Caption         =   "S&ystem Area"
         Index           =   2
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TabSelected As Long
Dim ModuleID As String
Dim ClickItem As Boolean
Dim TitleCode As Long

Private Sub ckOnTop_Click()
'On Error Resume Next
    Select Case CkOnTop.Value
    Case 1
        MakeTop Me.HWND, True
    Case 0
        MakeTop Me.HWND, False
    End Select
End Sub

Public Sub cmdAbout_Click()
'On Error Resume Next
    FormAbout.Show
    MakeTop Me.HWND, False
    MakeTop FormAbout.HWND, True
    Me.Enabled = False
End Sub

Private Sub AddDialog(ByVal wType As Long)

Dim GetFolder As String
Dim OFN As OPENFILENAME
Dim GetFile As String
Dim AllFiles As String
Dim sRdFiles() As String, sRdPath As String, lNrFiles As Long
Dim i As Long

Select Case wType
Case 1 'Tab File

    AllFiles = BrowseForFile(Me.HWND, "All Files" & vbNullChar & "*.*", LbToScan.Caption, OFN, OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT)

    'Jika hanya satu file yang dipilih
    If IsFile(AllFiles) = True Then
        AddPathToList AllFiles
        Exit Sub
    End If
    
    'Jika file yang dipilih lebih dari satu
    sRdPath = OFN.lpstrFile
    sRdFiles = Split(OFN.lpstrFile, Chr$(0))
    
    sRdPath = sRdFiles(0)
    For i = 1 To UBound(sRdFiles)
    If sRdFiles(i) = "" Then Exit For
        GetFile = Trim$(sRdPath & "\" & sRdFiles(i))
        AddPathToList GetFile
    Next i
    
Case 2 ' Tab Folder
    GetFolder = BrowseForFolder(Me.HWND, LbToScan.Caption)
    AddPathToList GetFolder
End Select

End Sub
Private Sub cmdScan_Click()
'On Error Resume Next

    If TabSelected = 1 And LVDirFileScan.ListItems.Count > 0 Then
        PopScanEx(0).Enabled = True
    Else
        PopScanEx(0).Enabled = False
    End If
    
    PopupMenu PopScan, 32, cmdScan.Left, cmdScan.Top
End Sub

Private Sub cmdUpdate_Click()
'On Error Resume Next
Dim hMsg As Long

    hMsg = MsgBox("Update Procedure : " & vbCrLf & _
    "1. Download HerAV File (zip)" & vbCrLf & _
    "2. Extract HerAV file (zip) to HerAV Directory" & vbCrLf & _
    "3. Replace HerAV old version with HerAV last version" & vbCrLf & vbCrLf & _
    "Do you want to download update now?", vbQuestion + vbOKCancel, "HerAV Update")
    Select Case hMsg
    Case vbOK
        ShellExecute Me.HWND, "Open", "http://www.herherplay.cf/", "", "", 1
    Case vbCancel
        Exit Sub
    End Select
End Sub

Private Sub Form_Load()
'On Error Resume Next
Dim CmdFile As String

If App.PrevInstance = True Then End

SetVariable
MakeForm
RefreshProcessList
MakeDef
If Command$ <> "" Then
    CmdFile = Replace(Command$, Chr$(34), "")
    If PathFileExists(CmdFile) <> 0 Then
        If PathIsDirectory(CmdFile) <> 0 Then
            DirToScan = CmdFile
        Else
            FileToScan = CmdFile
        End If
        FormScan.ScanNow
    End If
End If

End Sub

Private Sub MakeForm()
'On Error Resume Next

Dim i As Long
Dim IntViri() As String
Dim DateUpd As String

'Memasukkan database virus ke listview
IntViri = Split(IntViriName, "|")

For i = 0 To UBound(IntViri)
    AddToLV LVDBViri, IntViri(i) & "|" & ViriString4(i), ImgSmall, , SystemPath & "\SHELL32.DLL"
Next i

'Pengaturan form utama
Me.Caption = MyTitle
LbDBCount.Caption = LbDBCount.Caption & UBound(ViriName) + 1
LbVersion.Caption = LbVersion.Caption & App.Major & "." & App.Minor
DateUpd = DateDiff("D", LAST_UPDATE, Date, vbUseSystemDayOfWeek, vbUseSystem)
If CLng(DateUpd) = 0 Then
    DateUpd = "New Update"
Else
    DateUpd = DateUpd & " Days Old"
End If
SBMain.Panels(1).Text = "Update : " & LAST_UPDATE & " (" & DateUpd & ")"
SBMain.Panels(2).Text = App.FileDescription
TabClick 1
MakeTop Me.HWND, True

End Sub

Private Sub TabClick(Index As Integer)
'On Error Resume Next
Dim i As Long
Dim ModulePath() As String
Dim RegValueNow As String
Dim RegDataNow As String
Dim RegValue() As String
Dim LVRegTag As String

    TabSelected = Index
    
    LVDirFileScan.Visible = False
    LVDBViri.Visible = False
    LVProc.Visible = False
    LVModule.Visible = False
    CkSubDir.Enabled = False
    CkSystemScan.Enabled = False
    CkRegFix.Enabled = False
    TimerProcess.Enabled = False
    
    Select Case Index
    Case 1
        LVDirFileScan.Visible = True
        CkSubDir.Enabled = True
        CkSystemScan.Enabled = True
        CkRegFix.Enabled = True
        LbToScan.Caption = "Select File/Directory To Scan"
    Case 2
        LVDBViri.Visible = True
        LbToScan.Caption = "Internal Virus Database List"
    Case 3
        TimerProcess.Enabled = True
        LVProc.Visible = True
        LbToScan.Caption = "Kill All Suspected Virus Process"

        RefreshProcessList
    Case 4
        LVModule.Visible = True
        LbToScan.Caption = "Modules Of Process Memory"
        
        UpdateProcessList
        LVModule.ListItems.Clear
        
        PBMain.Visible = True
        TabMain.Enabled = False
        LVModule.Enabled = False
    
        If ModuleID = "" Then
            For i = 0 To UBound(ProcessId)
                ModuleID = ModuleID & "|" & ProcessId(i)
            Next i
            ModuleID = Right$(ModuleID, Len(ModuleID) - 1)
        End If
        GetModule ModuleID, ModulePath()
        For i = 0 To UBound(ModulePath)
            AddToLV LVModule, GetFileName(ModulePath(i)) & "|" & ModulePath(i), ImgSmall, 3, , , 0
            PBMain.Value = CInt((i + 1) * 100 / (UBound(ModulePath) + 1))
            DoEvents
        Next i
        
        LVModule.Enabled = True
        TabMain.Enabled = True
        PBMain.Visible = False
    End Select
End Sub

Private Sub lbSite_Click()
    cmdAbout_Click
End Sub

Private Sub LVDirFileScan_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
If Button = 2 Then
    If ClickItem = True Then
        PopFileViewRem.Enabled = True
        PopFileViewClear.Enabled = True
        PopupMenu PopFileView
        ClickItem = False
    Else
        PopFileViewRem.Enabled = False
        PopFileViewClear.Enabled = False
        PopupMenu PopFileView
    End If
End If
End Sub

Private Sub LVDirFileScan_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
Dim i As Long

    For i = 1 To Data.Files.Count
        AddPathToList Data.Files(i)
    Next i
End Sub

Private Sub AddPathToList(ByVal lpFileName As String)
    If lpFileName = "" Then Exit Sub
    If PathFileExists(lpFileName) = 0 Then
        MsgBox lpFileName & Chr$(34) & lpFileName & Chr$(34) & " not exist!!", vbExclamation, Me.Caption
    Else
        If PathIsDirectory(lpFileName) <> 0 Then
            If GetFileName(lpFileName) = "" Then
                AddToLV LVDirFileScan, lpFileName & "||Directory|" & lpFileName, ImgSmall, , lpFileName, , 4
            Else
                AddToLV LVDirFileScan, GetFileName(lpFileName) & "||" & GetFileType(lpFileName) & "|" & lpFileName, ImgSmall, , lpFileName, , 4
            End If
        Else
            AddToLV LVDirFileScan, GetFileName(lpFileName) & "|" & (CInt(FileLen(lpFileName) / 1024) + 1) & " KB |" & GetFileType(lpFileName) & "|" & lpFileName, ImgSmall, , lpFileName, , 4
        End If
    End If
End Sub

Private Sub LVModule_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
    If Button = 2 And ClickItem = True Then
        PopupMenu PopModule, , , , PopModuleExp
        ClickItem = False
    End If
End Sub

Private Sub LVProc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error Resume Next
Dim i As Long

If Button = 2 And ClickItem = True Then
    PopupMenu PopProcMan, , , , PopProcManMod
    ClickItem = False
End If
    
End Sub

Private Sub PopFileViewAddExp_Click()
'On Error Resume Next
    ExploreDir App.Path
End Sub

Private Sub PopFileViewAddFile_Click()
    AddDialog 1
End Sub

Private Sub PopFileViewAddFolder_Click()
    AddDialog 2
End Sub

Private Sub PopFileViewClear_Click()
'On Error Resume Next
    LVDirFileScan.ListItems.Clear
End Sub

Private Sub PopFileViewRem_Click()
'On Error Resume Next
    LVDirFileScan.ListItems.Remove (LVDirFileScan.SelectedItem.Index)
End Sub

Private Sub PopModuleExp_Click()
'On Error Resume Next
    ExploreDir GetDir(LVModule.SelectedItem.SubItems(3)) & "\"
End Sub

Private Sub PopModuleProp_Click()
'On Error Resume Next
    ShowProps LVModule.SelectedItem.SubItems(3), Me.HWND
End Sub

Private Sub PopProcManExp_Click()
'On Error Resume Next
    ExploreDir GetDir(LVProc.SelectedItem.SubItems(2)) & "\"
End Sub

Private Sub PopProcManKillSel_Click()
'On Error Resume Next
Dim i As Long
Dim GetSelNow As String
Dim GetNameNow As String
Dim AppToKill() As String
GetSelNow = GetSelLV(LVProc, 1)
GetNameNow = GetSelLV(LVProc, 0)
AppToKill = Split(GetSelNow, "|")

If GetSelNow = "" Then
    MsgBox "Please select one or more process to terminate!!", vbExclamation, "Nothing To Kill"
    Exit Sub
End If

If UBound(AppToKill) = 0 Then
    If MsgBox("Are you sure you want to terminate " & Chr$(34) & GetNameNow & Chr$(34) & " process", vbExclamation + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
Else
    If MsgBox("Are you sure you want to terminate these process:" & vbCrLf & "- " & Replace$(GetNameNow, "|", vbCrLf & "- "), vbExclamation + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
End If

For i = 0 To UBound(AppToKill)
    KillProcessByID CLng(AppToKill(i))
Next i

End Sub

Private Sub PopProcManKillVB_Click()
'On Error Resume Next
Dim VBAppsCount As Long
Dim AllID As String
Dim GB() As String
Dim i As Long

UpdateProcessList
For i = 0 To UBound(ProcessId)
    AllID = AllID & "|" & ProcessId(i)
Next i
AllID = Right$(AllID, Len(AllID) - 1)

VBAppsCount = GetModule(AllID, GB(), True)
If VBAppsCount = 0 Then
    MsgBox "Nothing VB Apps Process!!", vbInformation, "Nothing To Kill"
Else
    MsgBox VBAppsCount & " VB Apps Process was killed !!", vbInformation, "Killed By HerAV"
End If

End Sub

Private Sub PopProcManMod_Click()
'On Error Resume Next
    ModuleID = GetSelLV(LVProc, 1)
    TabMain.Tabs(4).Selected = True
    ModuleID = ""
End Sub

Private Sub PopProcManProp_Click()
'On Error Resume Next
    ShowProps LVProc.SelectedItem.SubItems(2), Me.HWND
End Sub

Private Sub PopScanEx_Click(Index As Integer)
'On Error Resume Next
Dim ObjectToScan() As String
Dim i As Long

FileToScan = ""
DirToScan = ""

ScanIndex = Index
Select Case Index
Case 0

    ObjectToScan = Split(GetAllList(LVDirFileScan, 3), "|")
    
    For i = 0 To UBound(ObjectToScan)
        If PathIsDirectory(ObjectToScan(i)) <> 0 Then
            DirToScan = DirToScan & "|" & ObjectToScan(i)
        Else
            FileToScan = FileToScan & "|" & ObjectToScan(i)
        End If
    Next i
    
    If DirToScan <> "" Then
        DirToScan = Right$(DirToScan, Len(DirToScan) - 1)
    End If
    
    If FileToScan <> "" Then
        FileToScan = Right$(FileToScan, Len(FileToScan) - 1)
    End If
    
Case 1
    DirToScan = ExistDrive
Case 2
    FileToScan = ""
End Select

FormScan.ScanNow (Index)
End Sub

Private Sub TabMain_Click()
'On Error Resume Next
If TabMain.SelectedItem.Index = TabSelected Then Exit Sub
    TabMain.Enabled = False
    TabClick TabMain.SelectedItem.Index
    TabMain.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
    End
End Sub

Private Sub LVProc_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVProc, ColumnHeader
End Sub

Private Sub LVProc_ItemClick(ByVal Item As ComctlLib.ListItem)
    ClickItem = True
End Sub

Private Sub LVModule_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVModule, ColumnHeader
End Sub

Private Sub LVModule_ItemClick(ByVal Item As ComctlLib.ListItem)
    ClickItem = True
End Sub

Private Sub LVDBViri_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVDBViri, ColumnHeader
End Sub

Private Sub LVDirFileScan_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'On Error Resume Next
    LVColumnClick LVDirFileScan, ColumnHeader
End Sub

Private Sub LVDirFileScan_ItemClick(ByVal Item As ComctlLib.ListItem)
    ClickItem = True
End Sub

Private Sub Form_Initialize()
'On Error Resume Next
    InitCommonControls
End Sub

Private Sub TimerProcess_Timer()
'On Error Resume Next
    DoEvents
    RefreshProcessList
End Sub

Private Sub RefreshProcessList()
'On Error Resume Next

Dim i As Long
Dim IDNow As String
Dim IDsNow As String
    UpdateProcessList
    
    For i = 0 To ProcessCount - 1
        IDsNow = IDsNow & "|" & CStr(ProcessId(i))
        AddToLV LVProc, GetFileName(ProcessPath(i)) & "|" & ProcessId(i) & "|" & ProcessPath(i), ImgSmall, 2
        DoEvents
    Next i
    IDsNow = IDsNow & "|"
    
Up:
    For i = 1 To LVProc.ListItems.Count
    IDNow = CStr(GetLV(LVProc, i, 2))
        If InStrRev(IDsNow, IDNow) = 0 Then
            LVProc.ListItems.Remove (i)
            GoTo Up
            Exit For
        End If
    Next i
End Sub
