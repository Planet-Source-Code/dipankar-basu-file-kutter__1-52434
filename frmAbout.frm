VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0442
   ScaleHeight     =   2895
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   6375
      Begin FileSplit.chameleonButton cClose 
         Height          =   495
         Left            =   5280
         TabIndex        =   7
         ToolTipText     =   "Close"
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         BTYPE           =   4
         TX              =   "&OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   8421376
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAbout.frx":2092
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAuthor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   1160
         Width           =   4335
      End
      Begin VB.Label lblmail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1600
         Width           =   4335
      End
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Web:URL"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label lblSystem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "System :  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label lblcopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (c) All Rights Reserved"
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1725
         Left            =   120
         MousePointer    =   12  'No Drop
         Picture         =   "frmAbout.frx":20AE
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblAppNameVer 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Application Name and Version"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const AppName = "File Kutter"   ' Modify this line *******

Private Sub cClose_Click()
Unload Me   ' Close Dialog
End Sub

Private Sub Form_Load()
Dim OSystemVer As String, WinMajor As Integer, WinMinor As Integer, RetLong As Long, LoWord As Integer, HiWord As Integer
RetLong = GetVersion()
Call GetHiLoWord(RetLong, LoWord, HiWord)
Call GetHiLoByte(LoWord, WinMajor, WinMinor)
OSystemVer = "Microsoft Windows " & WinMajor & "." & WinMinor
' **** Modify this procedure if necessary ****
lblSystem.Caption = "System :  " & OSystemVer
lblURL.Caption = "http://www.geocities.com/basudip_in/" 'url:web
lblmail.Caption = "basudip_in@hotmail.com" 'mailto:email
lblAuthor.Caption = AppName & " is developed by Dipankar Basu" '& App.CompanyName
lblcopy.Caption = "(c) 2003  All Rights Reserved  " ' App.LegalCopyright
lblAppNameVer.Caption = AppName & " version :  " & App.Major & "." & App.Minor & "  Build :  " & App.Revision
Me.Caption = "About " & AppName
End Sub
Private Sub lblcopy_DblClick()
Warranty
End Sub
Private Sub lblmail_Click()
ShellExecute 0, "Open", "mailto:" & lblmail.Caption & "?Subject=" & App.Title, _
    vbNullString, vbNullString, vbNormal
End Sub
Private Sub lblmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmail.ForeColor = vbRed
End Sub
Private Sub lblmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmail.ForeColor = vbBlue
End Sub
Private Sub lblURL_Click()
ShellExecute 0, "Open", lblURL.Caption, vbNullString, vbNullString, vbNormal
End Sub
Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblURL.ForeColor = vbRed
End Sub
Private Sub lblURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblURL.ForeColor = vbBlue
End Sub
Private Sub GetHiLoByte(X As Integer, LoByte As Integer, HiByte As Integer)
LoByte = X And &HFF&
HiByte = X \ &H100
End Sub
Private Sub GetHiLoWord(X As Long, LoWord As Integer, HiWord As Integer)
LoWord = CInt(X And &HFFFF&)
HiWord = CInt(X \ &H10000)
End Sub
Private Sub Warranty()
Dim MsgW As String
MsgW = "This software is AS IS without warranty of any kind."
MsgW = MsgW + " While every possible care is taken, to ensure that the software is efficient and bug free."
MsgW = MsgW + " The developer does not hold himself responsible for any damage or data loss as a result of using"
MsgW = MsgW + " or distributing this software. In no event will Dipankar Basu be liable for any damages, however"
MsgW = MsgW + " caused and regardless of the theory of liability, arising out of the use of or inability to use the software."
Call MsgBox(MsgW, vbInformation + vbOKOnly, AppName)
End Sub

' Copyright (c)2003 All Rights Reserved by Dipankar Basu
