VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   Caption         =   " Please wait . . ."
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FileSplit.Linux LinuxFrame 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8493
      Caption         =   "Splitting file in progress; Please wait . . ."
      Begin FileSplit.chameleonButton cClose 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Close"
         Top             =   4080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         BTYPE           =   14
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProgress.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstProgress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   3570
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "status log"
         Top             =   360
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cClose_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cClose.Caption = "&Cancel" Then End
End Sub

Private Sub lstProgress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstProgress.ToolTipText = lstProgress.Text
End Sub
