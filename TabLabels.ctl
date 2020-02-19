VERSION 5.00
Begin VB.UserControl TabLabels 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ToolboxBitmap   =   "TabLabels.ctx":0000
   Begin VB.PictureBox SliderBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5640
      ScaleHeight     =   285
      ScaleWidth      =   495
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   495
      Begin VB.PictureBox picSliderLeft 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         Picture         =   "TabLabels.ctx":0532
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
      Begin VB.PictureBox picSliderRight 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   255
         Picture         =   "TabLabels.ctx":0808
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   210
      End
   End
   Begin VB.Timer tmrSlideTabs 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1530
      Top             =   90
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   1
      Left            =   1080
      Top             =   60
   End
   Begin VB.Line LineTap 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   2
      X2              =   57
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label cTab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   315
   End
   Begin VB.Shape shpTapLine 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   0
      Top             =   330
      Visible         =   0   'False
      Width           =   6000
   End
End
Attribute VB_Name = "TabLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************************
'* TabLabels 19.x - Tab Label user control                                                          *
'* -------------------------------------------------------------------------------------------------*
'* Author   : Alberto Miñano Villavicencio                                                          *
'* Company  : AMV Solutions Group S.A.C.                                                            *
'* Date     : 18/01/2020                                                                            *
'*                                                                                                  *
'*                                                                                                  *
'* DESCRIPTION                                                                                      *
'* ------------------------------------------------------------------------------------------------ *
'* BootStrap Tab Labels with unicode on NT4/2000/XP+; ANSI fallback on 95/98/ME                     *
'*                                                                                                  *
'*                                                                                                  *
'* LICENSE                                                                                          *
'* ------------------------------------------------------------------------------------------------ *
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en                                         *
'*                                                                                                  *
'* Terms: 1) If you make your own version, share using this same license.                           *
'*        2) When used in a program, mention my name in the program's credits.                      *
'*        3) May not be used as a part of commercial controls suite.                                *
'*        4) Free for any other commercial and non-commercial usage.                                *
'*        5) Use at your own risk. No support guaranteed.                                           *
'*                                                                                                  *
'*                                                                                                  *
'* DEPENNDECY                                                                                       *
'* ------------------------------------------------------------------------------------------------ *
'*   No Dependency                                                                                  *
'*                                                                                                  *
'* VERSION HISTORY                                                                                  *
'* ------------------------------------------------------------------------------------------------ *
'*   ver. 19.1.24 [2020-01-18]                                                                      *
'*   Initial release                                                                                *
'*                                                                                                  *
'*                                                                                                  *
'* CREDITS                                                                                          *
'* ------------------------------------------------------------------------------------------------ *
'* - Albertomi : Initial release, UC Author                                                         *
'* - Yacosta   : Animations, MouseHand Cursor implementation                                        *
'* - AxioUK    : Corrections, Modifications, Add Animations (Forked Version)                        *
'* -                                                                                                *
'****************************************************************************************************
'* v19.1.25 MouseHand Icon, Animación: Yacosta :-(
'* v19.1.26 Correcciones : AxioUK ;)
'* v19.1.27 Correcciones : Albertomi
'* v19.1.28 [AxioUK] Implementa Animaciones, Version simple se elimina Unilabel UC.
'* v19.1.29 [AxioUK] Implementa Posiciones, rediseño.
'* v19.1.30 [AxioUK] Implementa Slider para desplazar Tabs, corrección de Bugs.

Option Explicit

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'---
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Const IDC_HAND                  As Long = 32649

Public Enum eAnimationType
    ShortAnimation = 0
    LongAnimation = 1
End Enum

Public Enum ePosition
  TopTabs = 0
  DownTabs = 1
  LeftTabs = 2
  RightTabs = 3
End Enum

Public Enum sSliderST
        ssLeft = 0
        ssRight = 1
End Enum

'Default Property Values:
Private Const m_def_lngUserControlWidth = 28
Private Const m_def_TabActive = 0
Private Const m_def_LineVisible = False
Private Const m_def_Autosize = False
Private Const m_def_ForeColorEnter = &HFFFFFF
Private Const m_def_ForeColor = &HE0E0E0
Private Const m_def_ActiveColor = &HC0C0C0

'Property Variables:
Private m_TabActive As Long
Private m_LineVisible As Boolean
Private m_SeparationTab As Long
Private m_AutoSize As Boolean
Private m_ActiveColor As OLE_COLOR
Private m_ForeColorEnter As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_LineColor As OLE_COLOR
Private m_AnimationType As eAnimationType
Private m_PositionTab As ePosition
Private m_lngUserControlWidth As Long
Private m_lngIndex As Long
Private m_lngCurrentX As Long
Private m_lngStep  As Long
Private hCur As Long
Private cIndex As Integer
Private pIndex As Integer
Private m_SliderStatus As sSliderST
Private lPosX As Long, lPosXstep As Long
Private iCountStep As Integer
Private i As Integer

'Event Declarations:
Public Event Click(TabIndex As Integer)
Public Event MouseUp(TabIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(TabIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(TabIndex As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Function AddTab(ByVal Caption As String, Optional ByVal Key As String, Optional ByVal ToolTipText As String) As Variant
  If m_lngIndex = 0 Then
    m_lngIndex = 0
    With cTab(m_lngIndex)
      .Caption = Caption
      .Tag = Key
      .ToolTipText = ToolTipText
      .ZOrder 0
      
      Select Case m_PositionTab
        Case Is = TopTabs
            .Alignment = 0
            .Left = 2
            .Top = 2
        
        Case Is = DownTabs
            .Alignment = 0
            .Left = 2
            .Top = 8
        
        Case Is = LeftTabs
            .Alignment = 1
            .Left = UserControl.ScaleWidth - (.Width + 12)
            .Top = 5
        
        Case Is = RightTabs
            .Alignment = 0
            .Left = 12
            .Top = 5
      
      End Select
      
      m_lngUserControlWidth = .Width
    End With
    
    m_lngIndex = m_lngIndex + 1
  Else
    m_lngIndex = cTab.Count
    
    Load cTab(m_lngIndex)
    With cTab(m_lngIndex)
      .Caption = Caption
      .Tag = Key
      .ToolTipText = ToolTipText
      .ZOrder 0
      Select Case m_PositionTab
        Case Is = TopTabs, DownTabs
          .Top = cTab(0).Top
          .Left = cTab(m_lngIndex - 1).Left + cTab(m_lngIndex - 1).Width + m_SeparationTab
          
        Case Is = LeftTabs
          .Alignment = 1
          .Top = cTab(m_lngIndex - 1).Top + cTab(m_lngIndex - 1).Height + m_SeparationTab
          .Left = UserControl.ScaleWidth - (.Width + 12)
          
        Case Is = RightTabs
          .Top = cTab(m_lngIndex - 1).Top + cTab(m_lngIndex - 1).Height + m_SeparationTab
          .Left = cTab(0).Left
          
      End Select
      
      m_lngUserControlWidth = cTab(m_lngIndex).Left + cTab(m_lngIndex).Width + m_SeparationTab
      .Visible = True
 
    End With
  End If
  
  If m_AutoSize Then UserControl.Width = m_lngUserControlWidth
    
  ReOrderTabs
  Refresh
  
End Function

Public Sub SetOptions(ByVal TabIndex As Integer, ByVal Caption As String, Optional ByVal Key As String, Optional ByVal ToolTipText As String)
  With cTab(TabIndex)
  .Caption = Caption
  If Key <> vbNullString Then .Tag = Key
  If ToolTipText <> vbNullString Then .ToolTipText = ToolTipText
    
    Select Case m_PositionTab
      Case Is = TopTabs, DownTabs
        For i = TabIndex + 1 To cTab.UBound
          cTab(i).Left = cTab(i - 1).Left + cTab(i - 1).Width + m_SeparationTab
        Next i
    
        LineTap.X1 = .Left
        LineTap.X2 = .Left + .Width
    End Select
  End With

ReOrderTabs
End Sub

Private Sub Animation(ByVal X As Integer)
  If LineTap.X1 = cTab(X).Left Then Exit Sub
  
Select Case m_AnimationType
  Case Is = ShortAnimation
    Select Case m_PositionTab
      Case Is = TopTabs, DownTabs
          m_lngCurrentX = cTab(X).Left
          
          If LineTap.X1 < cTab(X).Left Then
            m_lngStep = (cTab(X).Left - LineTap.X1) / 10
          Else
            m_lngStep = (LineTap.X1 - cTab(X).Left) / 10
          End If
          
      Case Is = LeftTabs, RightTabs
           m_lngCurrentX = cTab(X).Top - 1
           
           If LineTap.Y1 < cTab(X).Top Then
             m_lngStep = (cTab(X).Top - LineTap.Y1) / 10
           Else
             m_lngStep = (LineTap.Y1 - cTab(X).Top) / 10
           End If
    End Select
    
  Case Is = LongAnimation
    Dim lInitPos As Integer
    
    Select Case m_PositionTab
        Case Is = TopTabs, DownTabs
            If X > pIndex Then
              lInitPos = cTab(X).Left + cTab(X).Width
              m_lngCurrentX = cTab(X).Left + cTab(X).Width
              
              If LineTap.X2 < lInitPos Then
                m_lngStep = (lInitPos - LineTap.X2) / 10
              Else
                m_lngStep = (LineTap.X2 - lInitPos) / 10
              End If
              
            Else
              m_lngCurrentX = cTab(X).Left
              
              If LineTap.X1 < cTab(X).Left Then
                m_lngStep = (cTab(X).Left - LineTap.X1) / 10
              Else
                m_lngStep = (LineTap.X1 - cTab(X).Left) / 10
              End If
            End If
        Case Is = LeftTabs, RightTabs
            If X > pIndex Then
              lInitPos = cTab(X).Top + cTab(X).Height
              m_lngCurrentX = cTab(X).Top + cTab(X).Height
              
              If LineTap.Y2 < lInitPos Then
                m_lngStep = (lInitPos - LineTap.Y2) / 10
              Else
                m_lngStep = (LineTap.Y2 - lInitPos) / 10
              End If
              
            Else
              m_lngCurrentX = cTab(X).Top
              
              If LineTap.Y1 < cTab(X).Top Then
                m_lngStep = (cTab(X).Top - LineTap.Y1) / 10
              Else
                m_lngStep = (LineTap.Y1 - cTab(X).Top) / 10
              End If
            End If
    
    End Select
End Select

  tmrAnimation.Enabled = True
  Debug.Print "m_lngCurrentX:" & m_lngCurrentX
End Sub

Private Sub fSliderBox(ByVal Visible As Boolean)
Dim sldOldColor As Long

  sldOldColor = GetPixel(picSliderLeft.hdc, 1, 1)
  ReplaceColor picSliderLeft, sldOldColor, m_BackColor
  ReplaceColor picSliderRight, sldOldColor, m_BackColor
  
  sldOldColor = GetPixel(picSliderLeft.hdc, 6, 6)
  ReplaceColor picSliderLeft, sldOldColor, m_ForeColor
  ReplaceColor picSliderRight, sldOldColor, m_ForeColor
  
  With SliderBox
      .Top = 2
      .Left = UserControl.ScaleWidth - (.Width + 2)
      .BackColor = m_BackColor
      .Visible = Visible
      .ZOrder 0
  End With
  
End Sub

Private Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, IPic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
 
    Set GetSystemHandCursor = IPic
    
End Function

Private Sub MousePointerHands(ByVal Index As Integer, ByVal SetCursonHand As Boolean)     'YACO
    If SetCursonHand Then
        If Ambient.UserMode Then
            cTab(Index).MousePointer = vbCustom
            cTab(Index).MouseIcon = GetSystemHandCursor
        End If
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        cTab(Index).MousePointer = vbDefault
        cTab(Index).MouseIcon = Nothing
    End If
End Sub

Private Sub Refresh()
  With cTab(0)
    Select Case m_PositionTab
      Case Is = TopTabs
          LineTap.X1 = .Left
          LineTap.X2 = .Left + .Width
          LineTap.Y1 = .Height + .Top + 6
          LineTap.Y2 = .Height + .Top + 6
          shpTapLine.Top = .Height + .Top + 7
          shpTapLine.Height = 1
      
      Case Is = DownTabs
          LineTap.X1 = .Left
          LineTap.X2 = .Left + .Width
          LineTap.Y1 = .Top - 6
          LineTap.Y2 = .Top - 6
          shpTapLine.Top = .Top - 7
          shpTapLine.Height = 1
          
      Case Is = LeftTabs
          LineTap.X1 = UserControl.ScaleWidth - 4
          LineTap.X2 = UserControl.ScaleWidth - 4
          LineTap.Y1 = .Top
          LineTap.Y2 = .Top + .Height
          shpTapLine.Top = 0
          shpTapLine.Width = 1
          shpTapLine.Left = UserControl.ScaleWidth - 3
          
      Case Is = RightTabs
          LineTap.X1 = 4
          LineTap.X2 = 4
          LineTap.Y1 = .Top
          LineTap.Y2 = .Top + .Height
          shpTapLine.Top = 0
          shpTapLine.Width = 1
          shpTapLine.Left = 3
            
    End Select
  End With
  
  SetActiveTabColor
  LineTap.BorderColor = m_LineColor
  shpTapLine.BorderColor = m_LineColor
  shpTapLine.Visible = m_LineVisible
  
  UserControl.Refresh
End Sub

Private Sub ReOrderTabs()
Dim m_lngUserControl As Long
  m_lngUserControl = UserControl.ScaleWidth - SliderBox.Width
  
If m_PositionTab = DownTabs Or m_PositionTab = TopTabs Then
  For i = 0 To cTab.UBound
    lPosX = cTab(i).Left
    If lPosX > m_lngUserControl - cTab(i).Width Then
      cTab(i).Visible = False
      Call fSliderBox(True)
    ElseIf lPosX < 0 Then
      cTab(i).Visible = False
      Call fSliderBox(True)
    Else
      cTab(i).Visible = True
    End If
    Debug.Print "cTab_" & i & " : " & cTab(i).Visible
  Next i
End If

End Sub

Private Sub ReplaceColor(ByVal PictureBox As Object, ByVal FromColor As Long, ByVal ToColor As Long)
  If PictureBox.Picture Is Nothing Then Err.Raise Number:=1, Description:="Picture not set"
  If PictureBox.Picture.Handle = 0 Then Err.Raise Number:=2, Description:="Picture handle is null"
  Dim WinFromColor As Long, WinToColor As Long, MemAutoRedraw As Boolean
  WinFromColor = WinColor(FromColor)
  WinToColor = WinColor(ToColor)
  With PictureBox
    MemAutoRedraw = .AutoRedraw
    .AutoRedraw = True
    Dim X As Long, Y As Long
    For X = 0 To CInt(.ScaleX(.Picture.Width, vbHimetric, vbPixels))
        For Y = 0 To CInt(.ScaleY(.Picture.Height, vbHimetric, vbPixels))
            If GetPixel(.hdc, X, Y) = WinFromColor Then SetPixel .hdc, X, Y, WinToColor
        Next Y
    Next X
    .Refresh
    .Picture = .Image
    .AutoRedraw = MemAutoRedraw
  End With
End Sub

Private Sub SetActiveTabColor()
  For i = 0 To cTab.UBound
   cTab(i).ForeColor = m_ForeColor
   If i = m_TabActive Then cTab(m_TabActive).ForeColor = m_ActiveColor
   'Debug.Print m_TabActive; i
  Next i

End Sub

Private Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function


Private Sub cTab_Click(Index As Integer)
  pIndex = m_TabActive
  cIndex = Index
  m_TabActive = Index
  Animation Index
  SetActiveTabColor
  
  RaiseEvent Click(Index)
  Debug.Print "Left:" & cTab(Index).Left & " + Width:" & cTab(Index).Width & " = m_lngCurrentX:" & cTab(Index).Left + cTab(Index).Width
End Sub

Private Sub cTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  cTab(Index).ForeColor = m_ActiveColor
  RaiseEvent MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub cTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  cIndex = Index
  cTab(Index).ForeColor = m_ForeColorEnter
  MousePointerHands Index, True
  RaiseEvent MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub cTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub picSliderLeft_Click()
Dim sldOldColor As Long
  sldOldColor = GetPixel(picSliderLeft.hdc, 6, 6)
  ReplaceColor picSliderLeft, sldOldColor, m_ActiveColor
  '--------
  For i = 0 To cTab.UBound
      If cTab(0).Visible = True Then Exit Sub
      If cTab(i).Visible = True Then
        lPosX = cTab(i - 1).Width
      End If
  Next i
  '--------
  iCountStep = 0
  m_SliderStatus = ssLeft
  lPosXstep = lPosX / 10
  tmrSlideTabs.Enabled = True
End Sub

Private Sub picSliderLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sldOldColor As Long
  sldOldColor = GetPixel(picSliderLeft.hdc, 6, 6)
  ReplaceColor picSliderLeft, sldOldColor, m_ForeColorEnter
  ReplaceColor picSliderRight, sldOldColor, m_ForeColor
End Sub

Private Sub picSliderRight_Click()
Dim sldOldColor As Long
  sldOldColor = GetPixel(picSliderRight.hdc, 6, 6)
  ReplaceColor picSliderRight, sldOldColor, m_ActiveColor
  '--------
  For i = cTab.Count - 1 To 0 Step -1
      If cTab(cTab.UBound).Visible = True Then Exit Sub
      If cTab(i).Visible = True Then
        lPosX = cTab(i + 1).Width
      End If
  Next i
  '--------
  iCountStep = 0
  m_SliderStatus = ssRight
  lPosXstep = lPosX / 10
  tmrSlideTabs.Enabled = True
End Sub

Private Sub picSliderRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sldOldColor As Long
  sldOldColor = GetPixel(picSliderRight.hdc, 6, 6)
  ReplaceColor picSliderRight, sldOldColor, m_ForeColorEnter
  ReplaceColor picSliderLeft, sldOldColor, m_ForeColor
End Sub

Private Sub SliderBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReplaceColor picSliderLeft, m_ForeColorEnter, m_ForeColor
  ReplaceColor picSliderRight, m_ForeColorEnter, m_ForeColor

End Sub

Private Sub tmrAnimation_Timer()
Select Case m_AnimationType
  Case Is = ShortAnimation
      Select Case m_PositionTab
          Case Is = TopTabs, DownTabs
              If LineTap.X1 < m_lngCurrentX Then
              'to Right
                LineTap.X1 = LineTap.X1 + m_lngStep
                LineTap.X2 = LineTap.X1 + cTab(m_TabActive).Width
                If LineTap.X1 > m_lngCurrentX Then
                  LineTap.X1 = cTab(m_TabActive).Left
                  LineTap.X2 = cTab(m_TabActive).Left + cTab(m_TabActive).Width
                  tmrAnimation.Enabled = False
                End If
              Else
              'to Left
                LineTap.X1 = LineTap.X1 - m_lngStep
                LineTap.X2 = LineTap.X1 + cTab(m_TabActive).Width
                If LineTap.X1 < m_lngCurrentX Then
                  LineTap.X2 = cTab(m_TabActive).Left + cTab(m_TabActive).Width
                  tmrAnimation.Enabled = False
                End If
              End If
              
          Case Is = LeftTabs, RightTabs
              If LineTap.Y1 < m_lngCurrentX Then
              'to Down
                LineTap.Y1 = LineTap.Y1 + m_lngStep
                LineTap.Y2 = LineTap.Y1 + cTab(m_TabActive).Height
                If LineTap.Y1 > m_lngCurrentX Then
                  LineTap.Y1 = cTab(m_TabActive).Top
                  LineTap.Y2 = cTab(m_TabActive).Top + cTab(m_TabActive).Height
                  tmrAnimation.Enabled = False
                End If
              Else
              'to Up
                LineTap.Y1 = LineTap.Y1 - m_lngStep
                LineTap.Y2 = LineTap.Y1 + cTab(m_TabActive).Height
                If LineTap.Y1 < m_lngCurrentX Then
                  LineTap.Y2 = cTab(m_TabActive).Top + cTab(m_TabActive).Height
                  tmrAnimation.Enabled = False
                End If
              End If
          
      End Select
      
  Case Is = LongAnimation
      Select Case m_PositionTab
          Case Is = TopTabs, DownTabs
              If LineTap.X2 <= m_lngCurrentX Then
              'to Right
                  LineTap.X2 = LineTap.X2 + m_lngStep
                If LineTap.X2 > m_lngCurrentX Then
                  LineTap.X2 = m_lngCurrentX
                  LineTap.X1 = LineTap.X1 + m_lngStep
                  If LineTap.X2 = m_lngCurrentX And LineTap.X1 >= cTab(m_TabActive).Left Then
                    LineTap.X1 = cTab(m_TabActive).Left
                    tmrAnimation.Enabled = False
                  End If
                End If
              End If
        
              If LineTap.X1 >= m_lngCurrentX Then
              'to Left
                LineTap.X1 = LineTap.X1 - m_lngStep
                If LineTap.X1 < m_lngCurrentX Then
                  LineTap.X1 = m_lngCurrentX
                  LineTap.X2 = LineTap.X2 - m_lngStep
                  If LineTap.X1 = m_lngCurrentX And LineTap.X2 <= m_lngCurrentX + cTab(m_TabActive).Width Then
                    LineTap.X2 = m_lngCurrentX + cTab(m_TabActive).Width
                    tmrAnimation.Enabled = False
                  End If
                End If
              End If
              
          Case Is = LeftTabs, RightTabs
              If LineTap.Y2 <= m_lngCurrentX Then
              'to Down
                  LineTap.Y2 = LineTap.Y2 + m_lngStep
                If LineTap.Y2 > m_lngCurrentX Then
                  LineTap.Y2 = m_lngCurrentX
                  LineTap.Y1 = LineTap.Y1 + m_lngStep
                  If LineTap.Y2 = m_lngCurrentX And LineTap.Y1 >= cTab(m_TabActive).Top Then
                    LineTap.Y1 = cTab(m_TabActive).Top
                    tmrAnimation.Enabled = False
                  End If
                End If
              End If
        
              If LineTap.Y1 >= m_lngCurrentX Then
              'to Up
                LineTap.Y1 = LineTap.Y1 - m_lngStep
                If LineTap.Y1 < m_lngCurrentX Then
                  LineTap.Y1 = m_lngCurrentX
                  LineTap.Y2 = LineTap.Y2 - m_lngStep
                  If LineTap.Y1 = m_lngCurrentX And LineTap.Y2 <= m_lngCurrentX + cTab(m_TabActive).Height Then
                    LineTap.Y2 = m_lngCurrentX + cTab(m_TabActive).Height
                    tmrAnimation.Enabled = False
                  End If
                End If
              End If
          
      End Select
End Select

End Sub

Private Sub tmrSlideTabs_Timer()
iCountStep = iCountStep + 1
Select Case m_SliderStatus
  Case Is = ssRight
        For i = 0 To cTab.UBound
            cTab(i).Left = cTab(i).Left - lPosXstep
        Next i
      
  Case Is = ssLeft
        For i = cTab.Count - 1 To 0 Step -1
            cTab(i).Left = cTab(i).Left + lPosXstep
        Next i
End Select

  LineTap.X1 = cTab(m_TabActive).Left
  LineTap.X2 = cTab(m_TabActive).Left + cTab(m_TabActive).Width

  ReOrderTabs
  If iCountStep = 10 Then tmrSlideTabs.Enabled = False
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  m_lngUserControlWidth = m_def_lngUserControlWidth * Screen.TwipsPerPixelX
  m_TabActive = m_def_TabActive
  m_LineVisible = m_def_LineVisible
  m_SeparationTab = 12
  m_AutoSize = m_def_Autosize
  m_ForeColorEnter = m_def_ForeColorEnter
  m_ForeColor = m_def_ForeColor
  m_AnimationType = ShortAnimation
  m_PositionTab = TopTabs
  m_ActiveColor = m_def_ActiveColor
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 SetActiveTabColor
 ReplaceColor picSliderLeft, m_ForeColorEnter, m_ForeColor
 ReplaceColor picSliderRight, m_ForeColorEnter, m_ForeColor
 MousePointerHands cIndex, False
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  m_ForeColorEnter = .ReadProperty("ForeColorEnter", m_def_ForeColorEnter)
  m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
  m_ActiveColor = .ReadProperty("ForeColorActive", m_def_ActiveColor)
  m_LineColor = .ReadProperty("LineColor", &HFF&)
  m_BackColor = .ReadProperty("BackColor", &H404040)
  m_TabActive = .ReadProperty("TabActive", m_def_TabActive)
  m_LineVisible = .ReadProperty("LineVisible", m_def_LineVisible)
  m_SeparationTab = .ReadProperty("SeparationTab", 12)
  m_AutoSize = .ReadProperty("Autosize", m_def_Autosize)
  m_AnimationType = .ReadProperty("AnimationType", 0)
  m_PositionTab = .ReadProperty("PositionTab", 0)
  cTab(0).Caption = .ReadProperty("Caption", "Gengral")
  Set cTab(0).Font = .ReadProperty("Font", Ambient.Font)
  
  UserControl.BackColor = .ReadProperty("BackColor", &H404040)
  UserControl.Enabled = .ReadProperty("Enabled", True)
End With
  
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  
  If m_AutoSize Then UserControl.Width = m_lngUserControlWidth
 
  Select Case m_PositionTab
    Case Is = TopTabs, DownTabs
        shpTapLine.Width = UserControl.ScaleWidth
        
    Case Is = LeftTabs, RightTabs
        shpTapLine.Height = UserControl.ScaleHeight
  End Select
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ForeColorEnter", m_ForeColorEnter, m_def_ForeColorEnter)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H404040)
  Call PropBag.WriteProperty("Caption", cTab(0).Caption, "Caption")
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Font", cTab(0).Font, Ambient.Font)
  Call PropBag.WriteProperty("TabActive", m_TabActive, m_def_TabActive)
  Call PropBag.WriteProperty("ForeColorActive", m_ActiveColor, m_def_ActiveColor)
  Call PropBag.WriteProperty("LineColor", m_LineColor, &HFF&)
  Call PropBag.WriteProperty("LineVisible", m_LineVisible, m_def_LineVisible)
  Call PropBag.WriteProperty("SeparationTab", m_SeparationTab, 12)
  Call PropBag.WriteProperty("Autosize", m_AutoSize, m_def_Autosize)
  Call PropBag.WriteProperty("AnimationType", m_AnimationType, 0)
  Call PropBag.WriteProperty("PositionTab", m_PositionTab, 0)
End Sub

Public Property Get AnimationType() As eAnimationType
  AnimationType = m_AnimationType
End Property

Public Property Let AnimationType(ByVal New_AnimationType As eAnimationType)
  m_AnimationType = New_AnimationType
  PropertyChanged "AnimationType"
End Property

'// Properties
Public Property Get AutoSize() As Boolean
  AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_Autosize As Boolean)
 m_AutoSize = New_Autosize
  If m_AutoSize Then UserControl.Width = m_lngUserControlWidth
  PropertyChanged "Autosize"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  cTab(0).BackColor() = New_BackColor
  UserControl.BackColor = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
  Set Font = cTab(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set cTab(0).Font = New_Font
  cTab(0).AutoSize = True
  UserControl.Refresh
  PropertyChanged "Font"
End Property

Public Property Get ForeColorActive() As OLE_COLOR
Attribute ForeColorActive.VB_Description = "Devuelve o establece el color usado para rellenar formas, círculos y cuadros."
  ForeColorActive = m_ActiveColor
End Property

Public Property Let ForeColorActive(ByVal New_ForeColorActive As OLE_COLOR)
  m_ActiveColor = New_ForeColorActive
  PropertyChanged "ForeColorActive"
  Refresh
End Property

Public Property Get ForeColorEnter() As OLE_COLOR
  ForeColorEnter = m_ForeColorEnter
End Property

Public Property Let ForeColorEnter(ByVal New_ForeColorEnter As OLE_COLOR)
  m_ForeColorEnter = New_ForeColorEnter
  PropertyChanged "ForeColorEnter"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  cTab(0).ForeColor = m_ForeColor
  PropertyChanged "ForeColor"
End Property

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Devuelve o establece el color usado para rellenar formas, círculos y cuadros."
  LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
  m_LineColor = New_LineColor
  LineTap.BorderColor = m_LineColor
  shpTapLine.BorderColor = m_LineColor
  PropertyChanged "LineColor"
End Property

Public Property Get LineVisible() As Boolean
  LineVisible = m_LineVisible
End Property

Public Property Let LineVisible(ByVal New_LineVisible As Boolean)
  m_LineVisible = New_LineVisible
  shpTapLine.Visible = m_LineVisible
  PropertyChanged "LineVisible"
End Property
'm_PositionTab
Public Property Get PositionTab() As ePosition
  PositionTab = m_PositionTab
End Property

Public Property Let PositionTab(ByVal New_PositionTab As ePosition)
  m_PositionTab = New_PositionTab
  PropertyChanged "PositionTab"
  Refresh
End Property

Public Property Get SeparationTab() As Long
  SeparationTab = m_SeparationTab
End Property

Public Property Let SeparationTab(ByVal New_SeparationTab As Long)
  m_SeparationTab = New_SeparationTab
  PropertyChanged "SeparationTab"
  Refresh
End Property

Public Property Get TabActive() As Long
  TabActive = m_TabActive
End Property

Public Property Let TabActive(ByVal New_TabActive As Long)
  m_TabActive = New_TabActive
  cTab_Click (New_TabActive)
  PropertyChanged "TabActive"
End Property

