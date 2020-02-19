VERSION 5.00
Object = "*\ATLLCCTRL.vbp"
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11190
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   3  'Windows Default
   Begin TLLCCTRL.TabLabels TabLabels2 
      Height          =   4845
      Left            =   150
      TabIndex        =   13
      Top             =   930
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   8546
      ForeColorEnter  =   12632256
      ForeColor       =   0
      BackColor       =   8421504
      Caption         =   "Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorActive =   16777215
      LineColor       =   16777215
      LineVisible     =   -1  'True
      AnimationType   =   1
      PositionTab     =   2
   End
   Begin TLLCCTRL.TabLabels TabLabels1 
      Height          =   450
      Left            =   1995
      TabIndex        =   14
      Top             =   420
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   794
      ForeColorEnter  =   12632256
      ForeColor       =   12582912
      BackColor       =   8421504
      Caption         =   "Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorActive =   0
      LineColor       =   16777215
      LineVisible     =   -1  'True
      AnimationType   =   1
   End
   Begin TLLCCTRL.LayerContainer LayerContainer1 
      Height          =   4935
      Left            =   2055
      TabIndex        =   0
      Top             =   885
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LayerCount      =   10
      LayerEnabled0   =   0   'False
      LayerVisible0   =   0   'False
      Item(0).Control(0)=   "Label1"
      Item(0).Control(1)=   "Frame1"
      Item(0).ControlCount=   2
      Item(1).Control(0)=   "Label2"
      Item(1).Control(1)=   "List1"
      Item(1).Control(2)=   "Frame2"
      Item(1).ControlCount=   3
      Item(2).Control(0)=   "Label3"
      Item(2).Control(1)=   "Command1"
      Item(2).ControlCount=   2
      Item(3).Control(0)=   "Label4"
      Item(3).Control(1)=   "Command2"
      Item(3).ControlCount=   2
      Item(4).Control(0)=   "Label5"
      Item(4).Control(1)=   "Image1"
      Item(4).ControlCount=   2
      Item(5).Control(0)=   "Label6"
      Item(5).Control(1)=   "Picture1"
      Item(5).ControlCount=   2
      Item(6).Control(0)=   "Label7"
      Item(6).ControlCount=   1
      Item(7).Control(0)=   "Label8"
      Item(7).ControlCount=   1
      Item(9).Control(0)=   "Label9"
      Item(9).ControlCount=   1
      ItemMax         =   9
      BorderColor     =   4210752
      GradientStart   =   8421504
      GradientStyle   =   2
      Begin VB.PictureBox Picture1 
         BackColor       =   &H008080FF&
         FillColor       =   &H008080FF&
         Height          =   1215
         Left            =   -68500
         ScaleHeight     =   1155
         ScaleWidth      =   5205
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2070
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change Caption"
         Height          =   360
         Left            =   -68680
         TabIndex        =   11
         Top             =   345
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   -64780
         TabIndex        =   10
         Top             =   225
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   870
         Left            =   -69835
         TabIndex        =   9
         Top             =   3960
         Visible         =   0   'False
         Width           =   6540
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4545
         Left            =   4980
         TabIndex        =   8
         Top             =   225
         Width           =   1680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change Caption"
         Height          =   360
         Left            =   -69070
         TabIndex        =   7
         Top             =   420
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   195
         Left            =   -69760
         TabIndex        =   19
         Top             =   330
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         Height          =   195
         Left            =   -69850
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label7"
         Height          =   195
         Left            =   -69775
         TabIndex        =   17
         Top             =   195
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3150
         Left            =   -66760
         Picture         =   "Form2.frx":0000
         Top             =   1575
         Visible         =   0   'False
         Width           =   3555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         Height          =   195
         Left            =   -69820
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         Height          =   195
         Left            =   -69865
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   195
         Left            =   -69805
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   195
         Left            =   -69835
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   195
         Left            =   -69805
         TabIndex        =   2
         Top             =   135
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   135
         Width           =   90
      End
   End
   Begin TLLCCTRL.TabLabels TabLabels3 
      Height          =   4890
      Left            =   8985
      TabIndex        =   15
      Top             =   930
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   8625
      ForeColorEnter  =   12632256
      ForeColor       =   16777215
      BackColor       =   8421504
      Caption         =   "Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorActive =   8388608
      LineColor       =   8388608
      LineVisible     =   -1  'True
      AnimationType   =   1
      PositionTab     =   3
   End
   Begin TLLCCTRL.TabLabels TabLabels4 
      Height          =   450
      Left            =   2040
      TabIndex        =   16
      Top             =   5865
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   794
      ForeColorEnter  =   12632256
      ForeColor       =   16777215
      BackColor       =   8421504
      Caption         =   "Tab"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorActive =   16777215
      LineColor       =   16777215
      LineVisible     =   -1  'True
      AnimationType   =   1
      PositionTab     =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub LoadTab()
  TabLabels1.AddTab "General", "General", "Configuración general"
  TabLabels1.AddTab "Video", "Video", "Prueba de video"
  TabLabels1.AddTab "Audio", "Audio", "Prueba de audio"
  TabLabels1.AddTab "HUD", "HUD", "Prueba de HUD"
  TabLabels1.AddTab "Export", "Export", "Exportar archivo a disco"
  TabLabels1.AddTab "Teclado", "Tag", "Teclas de acceso rápido"
  TabLabels1.AddTab "GUID", "GUID"
  TabLabels1.AddTab "Exported Func", "ExportF"
  TabLabels1.AddTab "Tagging1", "Taged1"
  TabLabels1.AddTab "Tagging2", "Taged2"
  
  TabLabels2.AddTab "General", "General"
  TabLabels2.AddTab "Video", "Video"
  TabLabels2.AddTab "Audio", "Audio"
  TabLabels2.AddTab "HUD", "HUD"
  TabLabels2.AddTab "Export", "Export"
  TabLabels2.AddTab "Configuración", "Tag"
  
  TabLabels3.AddTab "Summary", "General"
  TabLabels3.AddTab "Details", "Video"
  TabLabels3.AddTab "Linkedin", "Audio"
  TabLabels3.AddTab "CompanyProfile", "HUD"
  TabLabels3.AddTab "Related", "Export"
  
  TabLabels4.AddTab "General", "General"
  TabLabels4.AddTab "Video", "Video"
  TabLabels4.AddTab "Audio", "Audio"
  TabLabels4.AddTab "HUD", "HUD"
  TabLabels4.AddTab "Export", "Export"
  TabLabels4.AddTab "Teclado", "Tag"

End Sub


Private Sub Command1_Click()
TabLabels1.SetOptions 2, "Test Options"
End Sub

Private Sub Command2_Click()
TabLabels1.SetOptions 3, "Caption Long Sized Test"
End Sub

Private Sub Form_Initialize()
  LoadTab
  TabLabels4.TabActive = 3
End Sub

Private Sub TabLabels1_Click(TabIndex As Integer)
LayerContainer1.ActiveLayer = TabIndex
End Sub

