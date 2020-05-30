VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormAbout 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HerAV"
   ClientHeight    =   3765
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5460
   Icon            =   "FormAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmConAuthor 
      BackColor       =   &H80000005&
      Height          =   2355
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4695
      Begin VB.PictureBox PicFrmConAut 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   297
         TabIndex        =   3
         Top             =   120
         Width           =   4455
         Begin VB.Label lbSN 
            BackColor       =   &H80000009&
            Caption         =   "Facebook :"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lbSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "www.facebook.com/hidden76"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   840
            MouseIcon       =   "FormAbout.frx":1CCA
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   1080
            Width           =   2220
         End
         Begin VB.Label lbSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "www.heriherplay.cf"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   840
            MouseIcon       =   "FormAbout.frx":1E1C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label lbSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "www.herherplay.hol.es"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   2
            Left            =   840
            MouseIcon       =   "FormAbout.frx":1F6E
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label lbSiteX 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Please send Your Comment to E-mail:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   120
            MouseIcon       =   "FormAbout.frx":20C0
            TabIndex        =   6
            Top             =   1560
            Width           =   2685
         End
         Begin VB.Label lbSiteX 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "Programmer's Site :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   120
            MouseIcon       =   "FormAbout.frx":2212
            TabIndex        =   5
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbMail2 
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "herhermawan007@gmail.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   840
            MouseIcon       =   "FormAbout.frx":2364
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1800
            Width           =   2160
         End
      End
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   349
      TabIndex        =   1
      Top             =   120
      Width           =   5235
      Begin VB.TextBox TxMain 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   4575
      End
      Begin ComctlLib.TabStrip TabMain 
         Height          =   2955
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5212
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "About HerAV"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "About Programmer"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Contact Programmer"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
'On Error Resume Next
    Me.iCon = FormMain.iCon
    Me.Caption = MyTitle
    TabMain.Tabs(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
    Unload Me
    FormMain.Enabled = True
End Sub

Private Sub Label1_Click()
ShellExecute Me.HWND, "Open", "mailto:" & lbMail.Caption, "", "", 1
End Sub

Private Sub lbMail_Click()
 ShellExecute Me.HWND, "Open", "mailto:" & lbMail.Caption, "", "", 1
End Sub

Private Sub lbSite_Click(Index As Integer)
'On Error Resume Next
    ShellExecute Me.HWND, "Open", lbSite(Index).Caption, "", "", 1
End Sub

Private Sub lbMail2_Click()
 ShellExecute Me.HWND, "Open", "mailto:" & lbMail2.Caption, "", "", 1
End Sub

Private Sub lbSite2_Click(Index As Integer)
' On Error Resume Next
    ShellExecute Me.HWND, "Open", lbSite2(Index).Caption, "", "", 1
End Sub

Private Sub TabMain_Click()
'On Error Resume Next
    Select Case TabMain.SelectedItem.Index
    Case 1
        frmConAuthor.Visible = False
        TxMain.Visible = True
        TxMain.Text = ReadMeText
    Case 2
        frmConAuthor.Visible = False
        TxMain.Visible = True
        TxMain.Text = AboutText
    Case 3
        frmConAuthor.Visible = True
        TxMain.Visible = False
    End Select
End Sub

Private Function ReadMeText() As String
'On Error Resume Next
ReadMeText = _
    Chr$(34) & "About HerAV" & Chr$(34) & vbCrLf & _
    "HerAV adalah Antivirus yang dikhususkan untuk mengatasi virus - virus lokal ataupun mancanegara yang menyebarluas akhir-akhir ini di Indonesia." & vbCrLf & _
    Chr$(34) & "HerAV Engine" & Chr$(34) & vbCrLf & _
    "Dengan menggunakan engine antivirus-nya, HerAV dapat melakukan scanning virus dengan cepat dan akurat." & vbCrLf & _
    Chr$(34) & "Virus Variant" & Chr$(34) & vbCrLf & _
    "Dengan menggunakan algoritma heuristic-nya, varian - varian virus (baik varian terbaru ataupun varian lama) yang belum ada di Database HerAV dapat dideteksi dengan mudah." & vbCrLf & _
    Chr$(34) & "Please Contact Me" & Chr$(34) & vbCrLf & _
    "Apabila terdapat kesalahan program/bug, saran/kritik, atau ingin bekerja sama dalam pengembangan HerAV silakan kirim pesan ke email saya." & vbCrLf & _
    Chr$(34) & "HerAV is Freeware" & Chr$(34) & vbCrLf & _
    "Anda bebas menggunakan dan menyebarluaskan HerAV selama bukan untuk kepentingan komersial. " & vbCrLf & _
    Chr$(34) & "HerAV Bug" & Chr$(34) & vbCrLf & _
    "Segala bentuk kerusakan yang mungkin diakibatkan oleh penggunaan HerAV diluar tanggung jawab programmer."
End Function

Private Function AboutText() As String
AboutText = _
    "Name : Heri Hermawan " & vbCrLf & _
    "Country: Indonesia" & vbCrLf & _
    "Province: Jawa Barat" & vbCrLf & _
    "City: Sumedang" & vbCrLf & _
    "Job: " & vbCrLf & _
    "    1. Pelajar SMK" & vbCrLf & _
    "    2. Web Tester" & vbCrLf & _
    "School: SMK Guna Dharma Nusantara" & vbCrLf & _
    "Class: X-C TKJ"
End Function

Private Sub TxMain_Change()
    TabMain.Tabs(TabMain.SelectedItem.Index).Selected = True
End Sub
