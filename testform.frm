VERSION 5.00
Object = "{DB984B8B-1DEE-11D4-9F90-A5A5A5A5A5A5}#46.0#0"; "fcs.ocx"
Begin VB.Form testform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fancy Controls Demo"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "testform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin fcs.Proggy PP2 
      Height          =   3315
      Left            =   450
      TabIndex        =   22
      Top             =   225
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5847
      BackColor       =   -2147483633
      Orientation     =   1
   End
   Begin fcs.SideLogo SideLogo2 
      Height          =   4290
      Left            =   6825
      TabIndex        =   20
      Top             =   1425
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   7567
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3D Side Logo"
      Style           =   1
   End
   Begin fcs.Proggy PP1 
      Height          =   315
      Left            =   1125
      TabIndex        =   19
      Top             =   1125
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      BackColor       =   -2147483633
      FillColor       =   255
      BackStyle       =   0
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   465
      Left            =   1350
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   18
      Top             =   4875
      Width           =   465
   End
   Begin fcs.IconBrowser IB1 
      Left            =   2700
      Top             =   5325
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdIcon 
      Caption         =   "Select Icon"
      Height          =   465
      Left            =   1800
      TabIndex        =   17
      Top             =   4275
      Width           =   1215
   End
   Begin VB.PictureBox picLarge 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   450
      ScaleHeight     =   540
      ScaleWidth      =   615
      TabIndex        =   16
      Top             =   4200
      Width           =   615
   End
   Begin VB.PictureBox picSmall 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   1200
      ScaleHeight     =   390
      ScaleWidth      =   465
      TabIndex        =   15
      Top             =   4275
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   600
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   390
      Left            =   2550
      Picture         =   "testform.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   375
      Width           =   390
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   600
      TabIndex        =   13
      Top             =   375
      Width           =   1965
   End
   Begin fcs.FolderBrowser Fb1 
      Left            =   2175
      Top             =   5325
      _ExtentX        =   847
      _ExtentY        =   847
      DONTGOBELOWDOMAIN=   -1  'True
   End
   Begin fcs.Palette Palette1 
      Height          =   4365
      Left            =   3600
      TabIndex        =   11
      ToolTipText     =   "Click here to change the color of 3d logo"
      Top             =   1425
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   7699
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3225
      Top             =   5475
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Height          =   330
      Left            =   630
      TabIndex        =   8
      Top             =   1050
      Width           =   435
   End
   Begin fcs.ColorSel256 ColorSel2 
      Height          =   540
      Left            =   735
      TabIndex        =   6
      Top             =   2730
      Visible         =   0   'False
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   953
   End
   Begin fcs.ColorSel256 ColorSel1 
      Height          =   330
      Left            =   2850
      TabIndex        =   5
      ToolTipText     =   "Tooltip here"
      Top             =   3870
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change logo color"
      Height          =   540
      Left            =   735
      TabIndex        =   4
      ToolTipText     =   "Click me and select a color"
      Top             =   2730
      Width           =   1275
   End
   Begin fcs.UrlLabel UrlLabel1 
      Height          =   285
      Left            =   675
      Top             =   3450
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverTextColor  =   16711680
      URL             =   "mailto:andemg@hotmail.com"
      Caption         =   "Mail me thru this URL label"
   End
   Begin fcs.SideLogo SideLogo1 
      Align           =   3  'Align Left
      Height          =   5835
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   10292
      EndColor        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "FancyControls demo, By Venu G."
   End
   Begin fcs.ColorSelector ColorSelector1 
      Height          =   330
      Left            =   4095
      TabIndex        =   1
      Top             =   720
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Palette"
      Height          =   240
      Left            =   3600
      TabIndex        =   21
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   4425
      TabIndex        =   10
      Top             =   75
      Width           =   1365
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Value"
      Height          =   165
      Left            =   3450
      TabIndex        =   9
      Top             =   75
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Change my Background color"
      Height          =   330
      Left            =   630
      TabIndex        =   7
      Top             =   3870
      Width           =   2325
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pick a Color"
      Height          =   225
      Left            =   4200
      TabIndex        =   2
      Top             =   405
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a folder"
      Height          =   225
      Left            =   630
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "testform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim runn, runn1 As Integer

Private Sub cmdIcon_Click()
    IB1.ShowIconDialog
    picLarge.Picture = IB1.getIconLargeImage
    picSmall.Picture = IB1.getIconSmallImage
    Picture1.Cls
    IB1.DrawIconTohDC Picture1.hDC, Large
    Picture1.Refresh
    Picture2.Cls
    IB1.DrawIconTohDC Picture2.hDC, Small, 5, 5
    Picture2.Refresh
End Sub

Private Sub ColorSel1_Click()
'selColor is the color selected by the user
    Me.BackColor = ColorSel1.selColor
End Sub

Private Sub ColorSel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Trackcolor is the current color under the mouse
    Label5.ForeColor = ColorSel1.TrackColor
    Label5 = ColorSel1.TrackColor
End Sub

Private Sub ColorSel2_Click()
'selColor is the color selected by the user
    SideLogo1.FontColor = ColorSel2.selColor
End Sub

Private Sub ColorSel2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Trackcolor is the current color under the mouse
    Label5.ForeColor = ColorSel2.TrackColor
    Label5 = ColorSel2.TrackColor
End Sub

Private Sub ColorSelector1_Click()
'SelectedColor is the color selected by the user
    Me.BackColor = ColorSelector1.SelectedColor
End Sub

Private Sub Command1_Click()
'This event pops up the ColorSelection box whenever you want
    ColorSel2.showPopupNow
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command3_Click()
    Text1.SetFocus
'shows the browse for folder dialog
    Fb1.ShowFolderDialog
    Text1 = Fb1.FolderPath
End Sub

Private Sub Form_Activate()
    runn = 3
    runn1 = 4
    
End Sub

Private Sub Palette1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'selColor is the color selected by the user
    SideLogo2.StartColor = Palette1.selColor
End Sub

Private Sub Timer1_Timer()
    DoEvents
    If PP1.Value >= PP1.Max - 1 Then runn = -runn
    If PP1.Value <= 0 Then runn = -runn
    PP1.Value = PP1.Value + runn
    
    DoEvents
    If PP2.Value >= PP2.Max - 1 Then runn1 = -runn1
    If PP2.Value <= 0 Then runn1 = -runn1
    PP2.Value = PP2.Value + runn1
    
End Sub
