VERSION 5.00
Begin VB.UserControl Linux 
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlContainer=   -1  'True
   PropertyPages   =   "Linux User Control.ctx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   ToolboxBitmap   =   "Linux User Control.ctx":0026
   Begin VB.Image Image9 
      Height          =   300
      Left            =   4440
      Picture         =   "Linux User Control.ctx":0338
      Top             =   0
      Width           =   315
   End
   Begin VB.Image Image8 
      Height          =   300
      Left            =   0
      Picture         =   "Linux User Control.ctx":087A
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Linux User Control.ctx":0B8C
      Top             =   240
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   120
      Left            =   4320
      Picture         =   "Linux User Control.ctx":0DD2
      Top             =   2400
      Width           =   330
   End
   Begin VB.Image Image4 
      Height          =   120
      Left            =   0
      Picture         =   "Linux User Control.ctx":1034
      Top             =   2400
      Width           =   345
   End
   Begin VB.Label lblMe 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
   Begin VB.Image Image3 
      Height          =   2250
      Left            =   0
      Picture         =   "Linux User Control.ctx":12B6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
   Begin VB.Image Image7 
      Height          =   120
      Left            =   120
      Picture         =   "Linux User Control.ctx":134C
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Image Image6 
      Height          =   2490
      Left            =   4560
      Picture         =   "Linux User Control.ctx":144E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "Linux User Control.ctx":14B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4740
   End
   Begin VB.Image Image10 
      Height          =   2175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Linux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=49771&lngWId=1

Private Sub Image9_Click()
Unload UserControl.Parent
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
With UserControl
Image5.Top = .Height - Image5.Height
Image5.Left = .Width - Image5.Width
Image7.Top = .Height - Image7.Height
Image7.Width = .Width
Image7.Left = 0
Image4.Left = 0
Image4.Top = .Height - Image4.Height
Image3.Left = 0
Image3.Top = 0
Image3.Height = .Height
Image6.Left = .Width - Image6.Width
Image6.Top = 0
Image6.Height = .Height
Image1.Top = 0
Image1.Left = 0
Image1.Width = .Width
Image2.Top = 0
Image2.Left = 0
Image9.Top = 0
Image9.Left = .Width - Image9.Width
Image8.Top = 0
Image8.Left = Image9.Left - Image8.Width
helft = .Width / 2
helft = helft - lblMe.Width / 2
Image10.Top = Image1.Top - -Image1.Height
Image10.Height = .Height - Image1.Height - Image7.Height
Image10.Left = Image3.Left - -Image3.Width
Image10.Width = .Width - Image3.Width - Image6.Width
'helft = helft - Image9.Width
'MsgBox helft
lblMe.Left = helft ' Me.ScaleWidth / 25
lblMe.Top = 0
End With
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Image10,Image10,-1,Picture
'Public Property Get Image() As Picture
'    Set Image = Image10.Picture
'End Property
'
'Public Property Set Image(ByVal New_Image As Picture)
'    Set Image10.Picture = New_Image
'    PropertyChanged "Image"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMe,lblMe,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblMe.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblMe.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Image", Nothing)
    lblMe.Caption = PropBag.ReadProperty("Caption", "")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Image", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", lblMe.Caption, "")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Image10,Image10,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Image10.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image10.Picture = New_Picture
    PropertyChanged "Picture"
End Property

