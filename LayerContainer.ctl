VERSION 5.00
Begin VB.UserControl LayerContainer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "LayerContainer.ctx":0000
   Begin VB.Timer TimerCheckMouseOut 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   60
      Top             =   60
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      Height          =   1650
      Left            =   600
      Top             =   810
      Width           =   2175
   End
End
Attribute VB_Name = "LayerContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************************
'* LayerContainer 7.1 - Layer Container UserControl                                                 *
'* -------------------------------------------------------------------------------------------------*
'* Autor Original  : Alberto Miñano Villavicencio - Albertomi                                       *
'* Autor Edición   : David Rojas Arraño - AxioUK                                                    *
'* Date Orig: 06/14/2006                                                                            *
'* Date Edit: 17/02/2020                                                                            *
'* ------------------------------------------------------------------------------------------------ *
'* DESCRIPTION                                                                                      *
'* MultiLayer Container                                                                             *
'* ------------------------------------------------------------------------------------------------ *
'*                                                                                                  *
'* LICENSE                                                                                          *
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en                                         *
'*                                                                                                  *
'* Terms: 1) If you make your own version, share using this same license.                           *
'*        2) When used in a program, mention my name in the program's credits.                      *
'*        3) May not be used as a part of commercial controls suite.                                *
'*        4) Free for any other commercial and non-commercial usage.                                *
'*        5) Use at your own risk. No support guaranteed.                                           *
'*                                                                                                  *
'* ------------------------------------------------------------------------------------------------ *
'* DEPENNDECY                                                                                       *
'*   Nothing                                                                                        *
'*                                                                                                  *
'* ------------------------------------------------------------------------------------------------ *
'* VERSION HISTORY                                                                                  *
'*                                                                                                  *
'*   01Apr06 - Initial UserControl Build                                                            *
'*   02Apr06 - Fixed Separator Alignment Bug                                                        *
'*           - Fixed Row Offset for new controls if they exceed the maximum width of the            *
'*             the control surface.                                                                 *
'*   11May06 - Fixed Tab Control Offset Bug by adding ActiveLayer = 0 to Usercontrol Terminate Event*
'*           - Added Font Property to control the Tab label font style                              *
'*           - Fixed bug in the Refresh sub which painted the control recursively                   *
'*   12May06 - Added DrawCGradient, DrawDGradient, DrawHGradient, and DrawVGradient                 *
'*             methods and associated routines for gradient background painting                     *
'*           - Added GradientStart, GradientEnd properties                                          *
'*           - Added hWnd and hDC properties                                                        *
'*   13May06 - Added bInitGradient flag to prevent unwanted redrawing of the gradient               *
'*             on each refresh to improve performance.                                              *
'*   14May06 - Fixed Drawing bug with DrawRGradient which resulted in the final rectangle           *
'*             not being drawn when X=X2 and/or Y=Y2 so the delta = 0.                              *
'*           - Added all public events expected in a UserControl plus some custom ones...           *
'*           - Added Accelerator Keys functionality                                                 *
'*           - Fixed bug with the Font property which did not "set" the m_Font private variable.    *
'*   19May06 - Added All API method for DrawCGradient                                               *
'*   11Jun06 - Added Additional In-Line comments for clarity                                        *
'*   14Jun06 - Fixed bug in Circular Gradient Methid where the we were not reselecting the old      *
'*             pen back into the DC....this caused a GDI Leak!!                                     *
'*                                                                                                  *
'*   15Feb20 - Rearmed UserControl implemented different method for distribute controls in Layers   *
'*   16Feb20 - Implement function to reorder Layers                                                 *
'*   17Feb20 - Fixed minor bugs...                                                                  *
'*                                                                                                  *
'* ------------------------------------------------------------------------------------------------ *
'* CREDITS                                                                                          *
'* - Albertomi (Author)(2006)                                                                       *
'* - AxioUK (Edition and improvements)(2020)                                                        *
'*                                                                                                  *
'*                                                                                                  *
'****************************************************************************************************
Option Explicit

Private Type POINT
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'---
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'---
Private Const DT_VCENTER = &H4&
Private Const DT_NOPREFIX = &H800

Private Type tLayerInfo
    sTag  As String
    sKey As String
    ItemData As Long
    Left As Long
    Width As Long
    Image As Long
    Alignment As Long
    Enabled As Boolean
    Visible As Boolean
    Active As Boolean
    Top As Long
    Height As Long
    Right As Long
    Bottom As Long
End Type

Private Type tControlInfo
    ctlName As String
    iLayerIndex As Long
End Type

Private Enum tsGradientDirectionEnum
  [tsNWSE] = &H0
  [tsSWNE] = &H1
End Enum

Public Enum tsGradientStyleEnum
  [tsNoGradient] = &H0
  [tsCircular] = &H1
  [tsDiagonalNWSE] = &H2
  [tsDiagonalSWNE] = &H3
  [tsHorizontal] = &H4
  [tsVertical] = &H6
End Enum

'---------------------------------
Public Event BeforeLayerChange(ByVal LastLayer As Long, ByRef NewLayer As Long)
Public Event LayerOrderChanged(ByVal LastLayer As Long, ByRef NewLayer As Long)


Private m_Layers() As tLayerInfo
Private m_LayerCount As Long
Private m_LayerOrder() As Long
Private m_Ctls() As tControlInfo
Private m_ctlCount As Long

Private mhDC As Long, hBmp As Long, hBmpOld As Long
Private lScrollX As Long, lScrollWidth As Long
Private lLayerIndex As Long
Private cx As Long, cy As Long
Private lLayerDragging As Boolean
Private lLayerDragged As Long
Private lLayerHover As Long
Private mFont As IFont
Private m_SelIndexLayer As Long
Private bInFocus As Boolean
Private m_AllowReorder As Boolean
Private bInitGradient As Boolean
Private m_GradientStyle As tsGradientStyleEnum
Private m_Border As Boolean

'Default Property Values:
Const m_def_BorderColor = &H808080
Const m_def_BackColor = vbWhite
Const m_def_ForeColor = vbBlack

'Property Variables:
Dim m_BorderColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_GradientEnd As OLE_COLOR
Dim m_GradientStart As OLE_COLOR



Public Function AddLayer(ByVal sKey As String, Optional ByVal lWidth As Long = 64, Optional ByVal ItemData As Long, Optional ByVal lImage As Long) As Long
ReDim Preserve m_Layers(m_LayerCount)
    With m_Layers(m_LayerCount)
      .sKey = sKey
      .Width = lWidth
      .ItemData = ItemData
      .Image = lImage
      .Visible = True
      .Enabled = True
    End With
    
    ReDim Preserve m_LayerOrder(m_LayerCount)
    m_LayerOrder(m_LayerCount) = m_LayerCount
    m_LayerCount = m_LayerCount + 1
    AddLayer = m_LayerCount
            
End Function

Public Function FindLayerByKey(ByVal sKey As String) As Long
Dim c As Long
    FindLayerByKey = -1
        For c = 0 To m_LayerCount - 1
            If m_Layers(c).sKey = sKey Then
                FindLayerByKey = c
                Exit For
            End If
        Next
End Function

Public Function ItemCount() As Long
    ItemCount = m_LayerCount
End Function

Public Sub Refresh()
  '   This Sub refreshes the controls with the correct backcolor,, forecolor, borderstyle and font
  With UserControl
    shpBorder.BorderColor = m_BorderColor
    shpBorder.Visible = m_Border
    UserControl.BackColor = m_BackColor
    
    '   See if the Gradient has been built, if not then do it....
    '   This prevents unwanted painting of the control which when
    '   the control is large can be slow...
    'If bInitGradient = False Then
      '   Speed things up by locking the window
      LockWindowUpdate .hwnd
      '   Clear the control surface
      .Cls
      '   Paint the correct gradient
      Select Case m_GradientStyle
        Case tsNoGradient:
          'Nothing
          
        Case tsCircular:
          Call DrawCGradient(m_GradientStart, m_GradientEnd, .ScaleWidth)
        
        Case tsDiagonalNWSE:
          Call DrawDGradient(m_GradientEnd, m_GradientStart, 0, 0, .ScaleWidth, .ScaleHeight, tsNWSE)
        
        Case tsDiagonalSWNE:
          Call DrawDGradient(m_GradientEnd, m_GradientStart, 0, 0, .ScaleWidth, .ScaleHeight, tsSWNE)
        
        Case tsHorizontal:
          Call DrawHGradient(m_GradientStart, m_GradientEnd, 0, 0, .ScaleWidth, .ScaleHeight)
        
        Case tsVertical:
          Call DrawVGradient(m_GradientStart, m_GradientEnd, 0, 0, .ScaleWidth, .ScaleHeight)
      
      End Select
      '   Now unlock things for the update to take effect
      LockWindowUpdate 0&
      
      '   Make sure to mark that we have built the gradient
      '   to prevent unwanted drawing of the control....
      'bInitGradient = True
    'End If
  End With
       
End Sub

Public Sub RemoveLayer(ByVal nIndex As Long)
Dim j As Long
Dim itemOrder As Long
If nIndex < m_LayerCount Then
        itemOrder = m_LayerOrder(nIndex)
   '// Reset m_Layers
   For j = m_LayerOrder(nIndex) To m_LayerCount - 2
      m_Layers(j) = m_Layers(j + 1)
   Next
   '// Adjust m_LayerOrder
   For j = nIndex To m_LayerCount - 2
      m_LayerOrder(j) = m_LayerOrder(j + 1)
   Next
   '// Validate Indexes for Items after deleted Item
   For j = 0 To m_LayerCount - 1
      If m_LayerOrder(j) > itemOrder Then
         m_LayerOrder(j) = m_LayerOrder(j) - 1
      End If
   Next

m_LayerCount = m_LayerCount - 1
    ReDim Preserve m_Layers(m_LayerCount)
    ReDim Preserve m_LayerOrder(m_LayerCount)
        If m_SelIndexLayer > m_LayerCount - 1 Then
            ActiveLayer = m_LayerCount - 1
        Else
            Refresh
        End If
End If
End Sub

Public Sub SetLayerEnabled(ByVal nIndex As Long, ByVal Newval As Boolean)
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
    m_Layers(m_LayerOrder(nIndex)).Enabled = Newval
    Refresh
End If
End Sub

Private Sub AddLayerControls(ByVal lIndex As Long, ByVal ctlName As String)
    ReDim Preserve m_Ctls(m_ctlCount)
    m_Ctls(m_ctlCount).ctlName = ctlName
    m_Ctls(m_ctlCount).iLayerIndex = lIndex
    m_ctlCount = m_ctlCount + 1
End Sub

Private Sub APILine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
    '   Use the API LineTo for Fast Drawing
    On Error GoTo APILine_Error
    
    Dim Pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(UserControl.hdc, hPen)
    MoveToEx UserControl.hdc, X1, Y1, Pt
    LineTo UserControl.hdc, X2, Y2
    SelectObject UserControl.hdc, hPenOld
    DeleteObject hPen
    Exit Sub
    
APILine_Error:
End Sub

Private Function APIRectangle(ByVal X As Long, ByVal Y As Long, ByVal W As Long, _
    ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
    
    '   Use the API Rectangle for Fast Drawing
    On Error GoTo APIRectangle_Error
    
    Dim hPen As Long, hPenOld As Long
    Dim Pt As POINT
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hdc, hPen)
    Rectangle UserControl.hdc, X, Y, W, H
    SelectObject UserControl.hdc, hPenOld
    DeleteObject hPen
    Exit Function
    
APIRectangle_Error:
End Function

Private Sub DrawCGradient(ByVal StartColor As Long, ByVal EndColor As Long, Optional ByVal numSteps As Integer = 256, Optional ByVal XCenter As Single = -1, Optional ByVal YCenter As Single = -1)
  '// Draw a Circular Gradient
  Dim StartRed As Integer, StartGreen As Integer, StartBlue As Integer
  Dim DeltaRed As Integer, DeltaGreen As Integer, DeltaBlue As Integer
  Dim stp As Long, hPen As Long, hPenOld As Long, lColor As Long
  Dim X As Long, Y As Long, X2 As Long, Y2 As Long

  With UserControl
      ' Evaluate the coordinates off the center if omitted.
      If XCenter = -1 And YCenter = -1 Then
          XCenter = .ScaleWidth / 3
          YCenter = .ScaleHeight / 3
      End If
              
      ' Split the start color into its RGB components
      StartRed = StartColor And &HFF
      StartGreen = (StartColor And &HFF00&) \ 256
      StartBlue = (StartColor And &HFF0000) \ 65536
      ' Split the end color into its RGB components
      DeltaRed = (EndColor And &HFF&) - StartRed
      DeltaGreen = (EndColor And &HFF00&) \ 256 - StartGreen
      DeltaBlue = (EndColor And &HFF0000) \ 65536 - StartBlue

      ' Draw all circles, going from the outside in.
      For stp = 0 To numSteps - 1
          lColor = RGB(StartRed + (DeltaRed * stp) \ numSteps, _
              StartGreen + (DeltaGreen * stp) \ numSteps, _
              StartBlue + (DeltaBlue * stp) \ numSteps)
          X = XCenter - numSteps + stp
          Y = YCenter - numSteps + stp
          X2 = XCenter + numSteps - stp
          Y2 = YCenter + numSteps - stp
          hPen = CreatePen(0, 2, lColor)
          hPenOld = SelectObject(UserControl.hdc, hPen)
          Ellipse .hdc, X, Y, X2, Y2
          SelectObject UserControl.hdc, hPenOld
          DeleteObject hPen
      Next
      
  End With
End Sub

Private Sub DrawDGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Direction As tsGradientDirectionEnum)
    
    '   Draw a Diagonal Gradient in the current HDC
    On Error GoTo DrawDGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long, lColor As Long, hPen As Long, hPenOld As Long
    
    lh = Y2 - Y
    lw = X2 - X
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If lh > lw Then
        dR = (sR - eR) / (lh)
        dG = (sG - eG) / (lh)
        dB = (sB - eB) / (lh)
        For ni = 0 To lh + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If Direction = tsNWSE Then
                '   NWSE (Move Only the UL corner towards the LR)
                Call APIRectangle(X + ni, Y + ni, X2, Y2, lColor)
            Else
                '   SWNE (Move Only the LL corner towards the UR)
                Call APIRectangle(X + ni, Y, X2, Y2 - ni, lColor)
            End If
        Next 'ni
    Else
        dR = (sR - eR) / (lw)
        dG = (sG - eG) / (lw)
        dB = (sB - eB) / (lw)
        For ni = 0 To lw + 1
            lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
            If Direction = tsNWSE Then
                '   NWSE (Move Only the UL corner towards the LR)
                Call APIRectangle(X + ni, Y + ni, X2, Y2, lColor)
            Else
                '   SWNE (Move Only the LL corner towards the UR)
                Call APIRectangle(X + ni, Y, X2, Y2 - ni, lColor)
            End If
        Next 'ni
    End If
    Exit Sub
    
DrawDGradient_Error:
End Sub

Private Sub DrawHGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    '   Draw a Horizontal Gradient in the current HDC
    On Error GoTo DrawHGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long
    lh = Y2 - Y
    lw = X2 - X
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw
    
    For ni = 0 To lw
        APILine X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawHGradient_Error:
End Sub

Private Sub DrawVGradient(ByVal lStartColor As Long, ByVal lEndColor As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal X2 As Long, ByVal Y2 As Long)
    
    '   Draw a Vertical Gradient in the current HDC
    On Error GoTo DrawVGradient_Error
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawVGradient_Error:
End Sub

Private Function getLayerOrder(ByVal Index As Long) As Long
Dim c As Long
    For c = 0 To m_LayerCount - 1
        If m_LayerOrder(c) = Index Then
            getLayerOrder = c
            Exit For
        End If
    Next
End Function
 
Private Sub HandleControls(ByVal LastIndex As Long, ByVal nIndex As Long)
Dim mCTL As Control, ctlName As String
Dim i As Long, z As Long, j As Long

  If m_LayerCount > 0 Then
      LastIndex = getLayerOrder(LastIndex)
      nIndex = getLayerOrder(nIndex)
  End If
            
  For Each mCTL In UserControl.ContainedControls
              ctlName = pGetControlId(mCTL)
      If IsInLayerControls(ctlName, nIndex) <> -1 Then
          If mCTL.Left < -35000 Then
              mCTL.Visible = True
              mCTL.Left = mCTL.Left + 70000
          End If
      Else
          If mCTL.Left > -35000 Then
              If IsInLayerControls(ctlName, LastIndex) = -1 Then
                  AddLayerControls LastIndex, ctlName
              End If
                  mCTL.Left = mCTL.Left - 70000
                  mCTL.Visible = False
          End If
      End If
  Next
  
  For j = 0 To m_ctlCount - 1
      If m_Ctls(j).iLayerIndex > (m_LayerCount - 1) Then
          For Each mCTL In UserControl.ContainedControls
                  ctlName = pGetControlId(mCTL)
              If Trim$(m_Ctls(j).ctlName) = ctlName Then
                  If mCTL.Left > -35000 Then
                      mCTL.Left = mCTL.Left - 70000
                      mCTL.Visible = False
                  End If
                      Exit For
              End If
          Next
      End If
  Next
End Sub

Private Function HasIndex(ByVal Ctl As Control) As Boolean
    'determine if it's a control array
    HasIndex = Not Ctl.Parent.Controls(Ctl.Name) Is Ctl
End Function

Private Function hitTest(ByVal lX As Long, ByVal lY As Long) As Long
Dim i As Long, c As Long, rc As RECT, lLayer As Long
  lLayer = -1
      
  For i = 0 To m_LayerCount - 1
    c = m_LayerOrder(i)
    With m_Layers(c)
      If .Enabled = True And .Visible = True Then
        rc.Left = .Left: rc.Top = .Top: rc.Right = .Right: rc.Bottom = .Bottom
        If lX >= rc.Left And lX <= rc.Right - 10 And lY >= rc.Top And lY <= rc.Bottom Then
          lLayer = i ' C
          Exit For
        End If
      End If
    End With
  Next
  
  hitTest = lLayer
End Function

Private Function IsInLayerControls(ByVal ctlName As String, ByVal lIndex As Long) As Long
Dim j As Long
  IsInLayerControls = -1
  For j = 0 To m_ctlCount - 1
      If Trim$(m_Ctls(j).ctlName) = ctlName And m_Ctls(j).iLayerIndex = lIndex Then
          IsInLayerControls = j
          Exit For
      End If
  Next
End Function

Private Sub MoveLayer(ByVal nLayer As Long, ByVal toLayer As Long)
Dim c As Long, j As Long, tempIndex As Long
Dim nInfo As tLayerInfo, nIndex As Long, toIndex As Long
'****** Sorting Controls **************
Dim NoMoreSwaps As Boolean, NumberOfItems As Long, Temp As tControlInfo, bDirection As Boolean
    bDirection = True
        NumberOfItems = UBound(m_Ctls)
    Do Until NoMoreSwaps = True
            NoMoreSwaps = True
         For c = 0 To (NumberOfItems - 1)
            If bDirection = True Then 'Ascending
                 If m_Ctls(c).iLayerIndex > m_Ctls(c + 1).iLayerIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            Else
                 If m_Ctls(c).iLayerIndex < m_Ctls(c + 1).iLayerIndex Then
                     NoMoreSwaps = False
                     Temp = m_Ctls(c)
                     m_Ctls(c) = m_Ctls(c + 1)
                     m_Ctls(c + 1) = Temp
                 End If
            End If
         Next
            NumberOfItems = NumberOfItems - 1
    Loop
'********* End Sorting *************************************
nIndex = m_LayerOrder(nLayer)
toIndex = m_LayerOrder(toLayer)
If toLayer > nLayer Then
   For c = nLayer To toLayer - 1
      m_LayerOrder(c) = m_LayerOrder(c + 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iLayerIndex > nLayer And m_Ctls(j).iLayerIndex <= toLayer Then
                m_Ctls(j).iLayerIndex = m_Ctls(j).iLayerIndex - 1
            ElseIf m_Ctls(j).iLayerIndex = nLayer Then
                m_Ctls(j).iLayerIndex = toLayer
            End If
        Next

Else
   For c = nLayer To toLayer + 1 Step -1
      m_LayerOrder(c) = m_LayerOrder(c - 1)
   Next
        For j = 0 To m_ctlCount - 1
            If m_Ctls(j).iLayerIndex >= toLayer And m_Ctls(j).iLayerIndex < nLayer Then
                m_Ctls(j).iLayerIndex = m_Ctls(j).iLayerIndex + 1
            ElseIf m_Ctls(j).iLayerIndex = nLayer Then
                m_Ctls(j).iLayerIndex = toLayer
            End If
        Next
End If
    
    m_LayerOrder(toLayer) = nIndex

End Sub

Private Function pGetControlId(ByVal oCtl As Control) As String
Dim sCtlName As String
Dim iCtlIndex As Integer
        iCtlIndex = -1
    sCtlName = oCtl.Name
On Error Resume Next
If HasIndex(oCtl) Then
    iCtlIndex = oCtl.Index
End If
    pGetControlId = sCtlName & IIf(iCtlIndex <> -1, "(" & iCtlIndex & ")", "")
End Function

Private Sub TimerCheckMouseOut_Timer()
    Dim Pos As POINT
    Dim WFP As Long
    
    GetCursorPos Pos
    WFP = WindowFromPoint(Pos.X, Pos.Y)
    
    If WFP <> Me.hwnd Then
        UserControl_MouseMove -1, 0, -1, -1
        TimerCheckMouseOut.Enabled = False 'kill that timer at once
    End If
End Sub

Private Sub UserControl_EnterFocus()
    bInFocus = True
    Refresh
End Sub

Private Sub UserControl_ExitFocus()
    bInFocus = False
    Refresh
End Sub

Private Sub UserControl_InitProperties()
Dim c As Long
  
  m_LayerCount = 1
  
  For c = 0 To m_LayerCount
      AddLayer c
  Next
  
  m_Border = True
  m_GradientStyle = tsNoGradient
  m_GradientEnd = &HFFFFFF
  m_GradientStart = &HFFC0C0
  m_BackColor = m_def_BackColor
  m_ForeColor = m_def_ForeColor
  m_BorderColor = m_def_BorderColor

  Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lLayer As Long, lCount As Long
     lLayer = getLayerOrder(m_SelIndexLayer)
NextLayer:
    lCount = lCount + 1
    
    Select Case KeyCode
        Case vbKeyLeft
            If lCount > 2 Then
                lLayer = m_LayerCount ' + 1
            End If
            If lLayer > 0 Then
                lLayer = lLayer - 1
            End If
    Case vbKeyRight
            If lCount > 2 Then
                lLayer = -1 '0
            End If
            If lLayer < m_LayerCount - 1 Then
                lLayer = lLayer + 1
            End If
    End Select

    If lLayer >= 0 And lLayer < m_LayerCount And lCount < m_LayerCount Then
        If m_Layers(m_LayerOrder(lLayer)).Enabled = False Or m_Layers(m_LayerOrder(lLayer)).Visible = False Then
            GoTo NextLayer:
        End If
    End If
    lLayer = m_LayerOrder(lLayer)
    
    If m_SelIndexLayer <> lLayer Then
        RaiseEvent BeforeLayerChange(m_SelIndexLayer, lLayer)
        ActiveLayer = lLayer
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> -1 Then
      cx = ScaleX(X, UserControl.ScaleMode, 3)
      cy = ScaleY(Y, UserControl.ScaleMode, 3)
  Else
      cx = X
      cy = Y
  End If
    
Dim L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lLayer As Long

    If Shift <> -1 Then
        lX = ScaleX(X, UserControl.ScaleMode, 3)
        lY = ScaleY(Y, UserControl.ScaleMode, 3)
    Else
        lX = X
        lY = Y
    End If
            
    lLayer = hitTest(lX, lY)
    lLayerIndex = lLayer
    
    If lLayer <> -1 Then
            lLayer = m_LayerOrder(lLayer)
        If m_SelIndexLayer <> lLayer Then
            RaiseEvent BeforeLayerChange(m_SelIndexLayer, lLayer)
            ActiveLayer = lLayer
        End If
    End If
    lLayerDragging = False

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Long, L As Long, lX As Long, lY As Long, W As Long, H As Long
Dim lLayer As Long
        
  If Shift <> -1 Then
      lX = ScaleX(X, UserControl.ScaleMode, 3)
      lY = ScaleY(Y, UserControl.ScaleMode, 3)
  Else
      lX = X
      lY = Y
  End If
  If m_AllowReorder = True And lLayerDragging = False And Button = 1 And lLayerIndex <> -1 And Abs(cx - lX) > 4 Then
      lLayerDragged = lLayerIndex
      lLayerDragging = True
  End If
  If Button = 0 And lLayerDragging = False Then
     TimerCheckMouseOut.Enabled = True
  End If
  lLayer = -1
  If lLayerDragging = True Then
  lLayer = hitTest(lX, lY)
  End If
  If lLayerDragging = False And Screen.MousePointer <> 0 Then
      Screen.MousePointer = 0
  End If
      lLayerHover = lLayer
  If lLayerDragging = True Then
  Screen.MousePointer = 5
  If lLayerHover <> lLayerIndex Then
      lLayerIndex = lLayer
  End If
      Refresh
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lLayerDragging = True Then
      Screen.MousePointer = 0
      lLayerHover = -1
      lLayerDragging = False
      If lLayerIndex <> -1 And lLayerIndex <> lLayerDragged Then
          lLayerIndex = m_LayerOrder(lLayerIndex)
          RaiseEvent LayerOrderChanged(m_LayerOrder(lLayerDragged), lLayerIndex)
          lLayerIndex = getLayerOrder(lLayerIndex)
          MoveLayer lLayerDragged, lLayerIndex
          ActiveLayer = m_LayerOrder(lLayerIndex)
      Else
          Refresh
      End If
      lLayerDragged = -1
      lLayerIndex = -1
      Exit Sub
  End If
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = mFont
    m_AllowReorder = PropBag.ReadProperty("AllowReorder", False)
    m_SelIndexLayer = PropBag.ReadProperty("ActiveLayer", 0)
    m_LayerCount = PropBag.ReadProperty("LayerCount", 4)
   
    ReDim m_Layers(m_LayerCount - 1)
    ReDim m_LayerOrder(m_LayerCount - 1)

    ReDim ctlLst(m_LayerCount - 1)
        m_ctlCount = 0
    ReDim m_Ctls(m_ctlCount)
        
    Dim i As Long, z As Long
    Dim mCCount As Long, ctlName As String, ItemMax As Long
    For i = 0 To m_LayerCount - 1
        m_LayerOrder(i) = PropBag.ReadProperty("LayerOrder" & i, i)
        m_Layers(i).Image = PropBag.ReadProperty("LayerIcon" & i, 0)
        m_Layers(i).Enabled = PropBag.ReadProperty("LayerEnabled" & i, True)
        m_Layers(i).sKey = PropBag.ReadProperty("Key" & i, "")
        m_Layers(i).sTag = PropBag.ReadProperty("LayerTag" & i, "")
        m_Layers(i).Visible = PropBag.ReadProperty("LayerVisible" & i, True)
        
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
                
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iLayerIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
        
    Next i
        
    ItemMax = PropBag.ReadProperty("ItemMax", 0)
    
    For i = m_LayerCount To ItemMax
        mCCount = PropBag.ReadProperty("Item(" & i & ").ControlCount", 0)
        For z = 0 To mCCount - 1
                ctlName = PropBag.ReadProperty("Item(" & i & ").Control(" & z & ")", "")
            If ctlName <> "" Then
                ReDim Preserve m_Ctls(m_ctlCount)
                   m_Ctls(m_ctlCount).ctlName = ctlName
                   m_Ctls(m_ctlCount).iLayerIndex = i
                m_ctlCount = m_ctlCount + 1
            End If
        Next z
    Next
    
    If m_SelIndexLayer > m_LayerCount - 1 Then
        m_SelIndexLayer = m_LayerCount - 1
    End If
    ActiveLayer = m_SelIndexLayer
            
  m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
  m_Border = PropBag.ReadProperty("Border", True)
  m_GradientEnd = PropBag.ReadProperty("GradientEnd", &HFFFFFF)
  m_GradientStart = PropBag.ReadProperty("GradientStart", &HFFC0C0)
  m_GradientStyle = PropBag.ReadProperty("GradientStyle", tsNoGradient)

End Sub

Private Sub UserControl_Resize()
shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Refresh
End Sub

Private Sub UserControl_Show()
ActiveLayer = 0
Refresh
End Sub

Private Sub UserControl_Terminate()
Erase m_Layers
Erase m_LayerOrder
    m_LayerCount = 0
Erase m_Ctls
    m_ctlCount = 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    Call PropBag.WriteProperty("AllowReorder", m_AllowReorder, False)
    Call PropBag.WriteProperty("ActiveLayer", m_SelIndexLayer, 0)
    Call PropBag.WriteProperty("LayerCount", m_LayerCount, 4)
    
    Dim i As Long, z As Long, c As Long, MaxIndex As Long
    For i = 0 To m_LayerCount - 1
        PropBag.WriteProperty "LayerOrder" & i, m_LayerOrder(i), i
        PropBag.WriteProperty "LayerIcon" & i, m_Layers(i).Image, 0
        PropBag.WriteProperty "LayerEnabled" & i, m_Layers(i).Enabled, True
        PropBag.WriteProperty "LayerTag" & i, m_Layers(i).sTag, ""
        PropBag.WriteProperty "LayerKey" & i, m_Layers(i).sKey, ""
        PropBag.WriteProperty "LayerVisible" & i, m_Layers(i).Visible, True
        c = 0
        For z = 0 To m_ctlCount - 1
            If m_Ctls(z).iLayerIndex = i Then
                Call PropBag.WriteProperty("Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, "")
                c = c + 1
            End If
            If MaxIndex < m_Ctls(z).iLayerIndex Then
                MaxIndex = m_Ctls(z).iLayerIndex
            End If
        Next z
        Call PropBag.WriteProperty("Item(" & i & ").ControlCount", c, 0)
    Next i
    
    Call PropBag.WriteProperty("ItemMax", MaxIndex, 0)
    
    For i = m_LayerCount To MaxIndex
            c = 0
        For z = 0 To m_ctlCount - 1
            If m_Ctls(z).iLayerIndex = i Then
                PropBag.WriteProperty "Item(" & i & ").Control(" & c & ")", m_Ctls(z).ctlName, ""
                c = c + 1
            End If
        Next z
        PropBag.WriteProperty "Item(" & i & ").ControlCount", c, 0
    Next
    
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("Border", m_Border, True)
    Call PropBag.WriteProperty("GradientEnd", m_GradientEnd, &HFFFFFF)
    Call PropBag.WriteProperty("GradientStart", m_GradientStart, &HFFC0C0)
    Call PropBag.WriteProperty("GradientStyle", m_GradientStyle, tsNoGradient)
  
End Sub

Public Property Get ActiveLayer() As Long
    ActiveLayer = m_SelIndexLayer
End Property

Public Property Let ActiveLayer(ByVal nLayerIndex As Long)
    If nLayerIndex < 0 Or nLayerIndex > m_LayerCount Then
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If
    If nLayerIndex > m_LayerCount - 1 Then
        nLayerIndex = m_LayerCount - 1
    End If

    HandleControls m_SelIndexLayer, nLayerIndex
    m_SelIndexLayer = nLayerIndex
    PropertyChanged "ActiveLayer"
    
    Refresh
End Property

Public Property Get AllowReorder() As Boolean
    AllowReorder = m_AllowReorder
End Property

Public Property Let AllowReorder(ByVal nV As Boolean)
  m_AllowReorder = nV
  PropertyChanged "AllowReorder"
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  PropertyChanged "BackColor"
  Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get Border() As Boolean
  Border = m_Border
End Property

Public Property Let Border(ByVal hasBorder As Boolean)
  m_Border = hasBorder
  PropertyChanged "Border"
  Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(ByVal nV As StdFont)
    Set mFont = nV
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
  Refresh
End Property

Public Property Get GradientEnd() As OLE_COLOR
    GradientEnd = m_GradientEnd
End Property

Public Property Let GradientEnd(ByVal lNewColor As OLE_COLOR)
    m_GradientEnd = lNewColor
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    PropertyChanged "GradientEnd"
    Refresh
End Property

Public Property Get GradientStart() As OLE_COLOR
    GradientStart = m_GradientStart
End Property

Public Property Let GradientStart(ByVal lNewColor As OLE_COLOR)
    m_GradientStart = lNewColor
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    PropertyChanged "GradientStart"
    Refresh
End Property

Public Property Get GradientStyle() As tsGradientStyleEnum
    GradientStyle = m_GradientStyle
End Property

Public Property Let GradientStyle(ByVal NewStyle As tsGradientStyleEnum)
    m_GradientStyle = NewStyle
    '   Allow the Gradient to be redrawn
    bInitGradient = False
    PropertyChanged "GradientStyle"
    Refresh
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Get LayerCount() As Long
    LayerCount = m_LayerCount
End Property

Public Property Let LayerCount(ByVal nLayer As Long)
Dim i As Long
If nLayer > m_LayerCount Then
    For i = m_LayerCount + 1 To nLayer
        AddLayer (i)
    Next
    Refresh
Else
    If nLayer > 0 Then
        Dim T As Long
            T = m_LayerCount - 1
        For i = T To nLayer Step -1
            RemoveLayer i
        Next
    End If
End If
    PropertyChanged "LayerCount"
End Property

Public Property Get LayerEnabled() As Boolean
Dim nIndex As Long
       nIndex = m_SelIndexLayer
If nIndex > -1 And nIndex < m_LayerCount Then
    LayerEnabled = m_Layers(nIndex).Enabled
End If
End Property

Public Property Let LayerEnabled(ByVal Newval As Boolean)
Dim nIndex As Long
       nIndex = m_SelIndexLayer
If nIndex > -1 And nIndex < m_LayerCount Then
    m_Layers(nIndex).Enabled = Newval
    PropertyChanged "LayerEnabled"
    Refresh
End If
End Property

Public Property Get LayerImage(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
    LayerImage = m_Layers(nIndex).Image
End If
End Property

Public Property Let LayerImage(ByVal nIndex As Long, ByVal lImage As Long)
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
     m_Layers(nIndex).Image = lImage

End If
End Property

Public Property Get LayerKey(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
    LayerKey = m_Layers(nIndex).sKey
End If
End Property

Public Property Get LayerTag(ByVal nIndex As Long) As String
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
    LayerTag = m_Layers(nIndex).sTag
End If
End Property

Public Property Let LayerTag(ByVal nIndex As Long, ByVal sTag As String)
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
     m_Layers(nIndex).sTag = sTag
End If
End Property

Public Property Get LayerWidth(ByVal nIndex As Long) As Long
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
    LayerWidth = m_Layers(nIndex).Width
End If
End Property

Public Property Let LayerWidth(ByVal nIndex As Long, ByVal lWidth As Long)
If nIndex > -1 And nIndex < m_LayerCount And m_LayerCount > 0 Then
     m_Layers(nIndex).Width = lWidth
End If
End Property

Public Property Get MoveItem() As Long
    MoveItem = getLayerOrder(m_SelIndexLayer)
End Property

Public Property Let MoveItem(ByVal nLayerIndex As Long)
    If nLayerIndex < 0 Or nLayerIndex > m_LayerCount - 1 Then
        MsgBox "invalid property value", vbCritical
        Exit Property
    End If

    MoveLayer getLayerOrder(m_SelIndexLayer), nLayerIndex
    ActiveLayer = m_LayerOrder(nLayerIndex)
End Property


