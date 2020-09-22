VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Kutter"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin FileSplit.chameleonButton cExit 
      Height          =   615
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Quit"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "Form1.frx":0442
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FileSplit.chameleonButton cAbout 
      Height          =   615
      Left            =   5160
      TabIndex        =   8
      ToolTipText     =   "Version information"
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&About"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "Form1.frx":045E
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FileSplit.chameleonButton cSplit 
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      ToolTipText     =   "OK"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Split"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "Form1.frx":047A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FileSplit.chameleonButton cBrowse 
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "Open the file to Split"
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Browse"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "Form1.frx":0496
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmSplitSize 
      Caption         =   "Split File"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6495
      Begin MSComDlg.CommonDialog cmDlg 
         Left            =   2520
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtFileName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   6255
      End
      Begin VB.ComboBox cboSize 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":04B2
         Left            =   480
         List            =   "Form1.frx":04BF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Split Size"
         Height          =   1455
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         Height          =   750
         Left            =   120
         Picture         =   "Form1.frx":04D5
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6285
      End
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Program: File Split Application
'  Coded and Designed by: Dipankar Basu
'  Date of Update: March 17, 2004.
'  Web/URL: http://www.geocities.com/basudip_in/download/
'  Credits: Kishore@DeveloperIQ for help, thanks a lot.
'  Copyright (c)2004 Dipankar Basu

Option Explicit: Option Base 1
Dim sFileNameToSplit As String ' required in join split in bat file

Private Sub cAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub cBrowse_Click()
On Error GoTo eh:
With cmDlg
 .CancelError = True
 .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
 .DialogTitle = "Open the File to split"
 .Filter = "All Files|*.*|"
 .ShowOpen
 txtFileName.Text = Trim(.FileName)
 sFileNameToSplit = Trim(.FileTitle)
End With
Exit Sub
eh:
    If Err.Number = 32755 Then  ' Cancel Button Selected
    txtFileName.Text = vbNullString: sFileNameToSplit = vbNullString
    Exit Sub: End If
MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, Err.Source
End Sub


Private Sub cExit_Click()
Unload Me: End
End Sub

Private Sub cSplit_Click()
On Error GoTo eh:
If Trim(txtFileName) = vbNullString Or Dir(txtFileName) = "" Then
MsgBox "The file could not be found", vbCritical, "File is inaccessible"
cBrowse.SetFocus
Exit Sub
End If
If Val(txtSize) < 1 Then
MsgBox "Incorrect split size", vbCritical, "Split size"
txtSize.SetFocus
Exit Sub
End If
            frmProgress.Show vbApplicationModal, Me
Dim sFilePath As String ' Splitted files path
Dim i As Long, Ss As Double ' i=number of splitted files     ss=split file size
            frmProgress.lstProgress.AddItem "Creating Destination folder for Split files . . ."
sFilePath = Trim(IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")) & Trim("SplitFile")
    Dim Fs
    Set Fs = CreateObject("Scripting.FileSystemObject")
FolderExists:
    i = Val(i) + 1
    If Fs.FolderExists(Trim(sFilePath) & Trim(Str(i))) Then GoTo FolderExists
    sFilePath = Trim(sFilePath) & Trim(Str(i))
    Set Fs = Nothing: i = 0 ' Clear memory & prepare variable i for reuse
MkDir sFilePath ' Creates a destination directory SplitFile
            frmProgress.lstProgress.AddItem sFilePath & "\  created"
            frmProgress.lstProgress.AddItem "Determining Split file size . . ."
If cboSize.ListIndex = 2 Then  ' MB
    Ss = 1024
    Ss = Ss * 1024
ElseIf cboSize.ListIndex = 1 Then ' KB
    Ss = 1024
ElseIf cboSize.ListIndex = 0 Then ' Bytes
    Ss = 1
End If
Ss = Round(Val(txtSize) * Ss, 0)
DoEvents
Dim nLen As Double, B() As Byte ' nlen =File Size
nLen = FileLen(txtFileName)
Open txtFileName For Binary As 1 ' SourceFile
While nLen > Ss
ReDim B(Ss)
Get #1, Ss * i + 1, B() ' Get FileNumber, ReadBeginByte, StoreReadDataVariable
Open sFilePath & "\" & "SplitFile" & "." & Format(i + 1, "000") For Binary As 2
Put #2, , B() ' Store splitted data in FileNo2
Close #2
            frmProgress.lstProgress.AddItem "Split file " & Format(i + 1, "000") & "  created  " & Ss & " Bytes"
i = i + 1
nLen = nLen - Ss
DoEvents
Wend
ReDim B(nLen)  ' Get file size for the last part of the spanned file
Get #1, Ss * i + 1, B()
Open sFilePath & "\" & "SplitFile" & "." & Format(i + 1, "000") For Binary As 2
Put #2, , B()
Close #2  ' Last part of the splitted file created
Close #1 ' Close original file
DoEvents
            frmProgress.lstProgress.AddItem "Split file " & Format(i + 1, "000") & "  created  " & nLen & " Bytes"
            frmProgress.lstProgress.AddItem "Creating Batch to join splitted files"
Dim a, PutToBat As String
PutToBat = "Copy/b "
For Ss = 1 To i + 1
If PutToBat = "Copy/b " Then
PutToBat = PutToBat & "SplitFile" & "." & Format(Ss, "000")
Else
PutToBat = PutToBat & " + SplitFile" & "." & Format(Ss, "000")
End If
Next
PutToBat = PutToBat & "  " & Chr(34) & sFileNameToSplit & Chr(34) & vbCrLf & "Exit"
DoEvents
Set Fs = CreateObject("Scripting.FileSystemObject")
Set a = Fs.CreateTextFile(sFilePath & "\JoinFile.bat")
a.WriteLine (PutToBat)
a.Close
Set Fs = Nothing: Set a = Nothing
            frmProgress.lstProgress.AddItem "Split process completed successfully"
            frmProgress.cClose.Caption = "&OK": frmProgress.LinuxFrame.Caption = "Log _ status"
Exit Sub
eh:
frmProgress.lstProgress.AddItem "Error: " & Err.Number & vbCrLf & Err.Description & " at " & Err.Source
frmProgress.lstProgress.AddItem "Split Aborted"
Close: Exit Sub
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
End Sub

Private Sub txtSize_Change()
If cboSize.ListIndex = -1 Then cboSize.ListIndex = 1 ' Default to KB
End Sub
