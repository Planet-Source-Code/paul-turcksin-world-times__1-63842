VERSION 5.00
Begin VB.UserControl DiagButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "DiagButton.ctx":0000
End
Attribute VB_Name = "DiagButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+
'   World Times - Graphical Representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          DiagButton
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul Turcksin
'
'   Credits:
'   LaVolpe's lvButtons
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=45279&lngWId=1
'   Extracted the rectangular and diagonal XP style
'   Also only the Click event was kept

'   Legal Copyright & Trademarks:
'       Copyright © 2005, by Paul Turcksin, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul Turcksin, All Rights Reserved Worldwide
'
'   You are free to use this code within your own applications, but you
'   are expressly forbidden from selling or otherwise distributing this
'   source code without prior written consent.
'
'   Redistributions of source code must include this list of conditions,
'   and the following acknowledgment:
'
'   This code was developed by Paul Turcksin.
'   Source code, written in Visual Basic, is freely available for non-
'   commercial, non-profit use.
'   Redistributions in binary form, as part of a larger project, must
'   include the above acknowledgment in the end-user documentation.
'   Alternatively, the above acknowledgment may appear in the software
'   itself, if and where such third-party acknowledgments normally appear.
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul Turcksin shall not be liable for any
'       incidental or consequential damages suffered by any use of this  software.

'       Many thanks to my friend Paul R. Territo Ph.D (TerriTop) for his careful review, suggestions,
'       and support of this program prior to public release. In addtion, I wish to
'       thank the numerous open source authors who provide code and inspiration to
'       make such work possible.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: paul_turcksin@Hotmail.com
'_________________________________________________________________________________
'
' Documentation:
'
' La Volpe's original comments:
' Note on use. The biggest question about this button control is its size.
' This control is NOT intended to be compiled into your application,
' it is intended to be compiled as a separate OCX and packaged with your app.

' This button control was inspired by Gonchuki's Chameleon Button v1.x. The
' 1st three versions of this control were based off of his control but
' eventually scrapped because of memory leaks, faulty logic and some buggy
' code; mostly on my part. Any code in this version that is similar to Gonchuki's
' control are coincidence with the exception of the following routines...
' The ShadeColor and the Step calculations in DrawButtonBackground routines
' for XP colors are formulas found in Gonchuki's v1.x and credit goes to
' him & Ghuran Kartal for the formula.

' last update: 2003Sep12. Memory leak when creating fonts. The button fonts were created too
'                                      early and subsequently overwritten without being deleted first.

'/////// Public Events sent back to the parent container
'-
Option Explicit
Option Compare Text

Public Event Click()
Attribute Click.VB_MemberFlags = "200"

' GDI32 Function Calls
' =====================================================================
' DC manipulation
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
' Region Forming functions
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_DIFF = 4
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
' Other drawing functions
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

' KERNEL32 Function Calls
' =====================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

' USER32 Function Calls
' =====================================================================
' General Windows related functions
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Standard TYPE Declarations used
' =====================================================================
Private Type POINTAPI                ' general use. Typically used for cursor location
    x As Long
    y As Long
End Type
Private Type RECT                    ' used to set/ref boundaries of a rectangle
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type BITMAP                  ' used to determine if an image is a bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type ICONINFO                ' used to determine if image is an icon
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type LOGFONT               ' used to create fonts
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type

' Custom TYPE Declarations used
' =====================================================================
Private Type ButtonDCInfo    ' used to manage the drawing DC
    hdc As Long              ' the temporary DC handle
    OldBitmap As Long        ' the original bitmap of the DC
    OldPen As Long           ' the original pen of the DC
    OldBrush As Long         ' the original brush of the DC
    ClipRgn As Long          ' used for circular button borders
    ClipBorder As Long         ' used for shaped buttons with borders
    OldFont As Long          ' the original font of the DC
End Type
Private Type ButtonProperties
    bCaption As String                           ' button caption
    bCaptionAlign As AlignmentConstants          ' caption alignment (3 options)
    bStatus As Integer                           ' 0=Up, 1=Focus, 2=Down, 4=Hover
    bShape As ButtonStyleConstants               ' shape of button (rect, diagonal, circle)
    bSegPts As POINTAPI                          ' left/right offsets for diagonal button
    bRect As RECT                                ' cached caption's bounding rectangle
    bShowFocus As Boolean                        ' flag to display/hide focus rectangle
    bBackHover As Long                           ' button back color when mouse hovers
    bForeHover As Long                           ' button text color when mouse hovers
    bLockHover As HoverLockConstants             ' allows/restricts hover colors same as normal button colors (4 options)
    bGradient As GradientConstants               ' 4 gradient directions
    bGradientColor As Long                       ' Gradient color to use
End Type
Private Type ImageProperties
    Image As StdPicture                          ' button image
    TransImage As Long
    TransSize As POINTAPI
    Align As ImagePlacementConstants             ' image alignment (6 options)
    Size As Integer                              ' image size (5 options)
    iRect As RECT                                ' cached image's bounding rectangle
    SourceSize As POINTAPI                       ' cached source image dimensions
    Type As Long                                 ' cached source image type (bmp/ico)
End Type

' Standard CONSTANTS as Constants or Enumerators
' =====================================================================
Private Const WHITENESS = &HFF0062
Private Const CI_BITMAP = &H0
Private Const CI_ICON = &H1
Private Const WM_KEYDOWN As Long = &H100
' //////////// Custom Colors \\\\\\\\\\\\\\\\\
Private Const vbGray = 8421504
' //////////// DrawText API Constants \\\\\\\\\\\\\\
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
' ///////////////// PROJECT-WIDE VARIABLES \\\\\\\\\\\\\\
Private ButtonDC As ButtonDCInfo       ' menu DC for drawing menu items
Private myProps As ButtonProperties    ' cached button properties
Private myImage As ImageProperties     ' cached image properties
Private bNoRefresh As Boolean          ' flag to prevent drawing when multiple properties are changing
Private curBackColor As Long           ' control's back color
Private adjBackColorUp As Long         ' cached backcolor in UP state
Private adjBackColorDn As Long         ' cached backcolor in DOWN state
Private adjHoverColor As Long          ' cached hover backcolor
Private mButton As Integer             ' used to prevent right click firing a click event
Private bTimerActive As Boolean        ' indication mouse is over a button
Private cParentBC As Long              ' parent control's back color
Private cCheckBox As Long              ' cached color for checkbox
Private bKeyDown As Boolean            ' used to properly display checkboxes

' Custom CONSTANTS as Constants or Enumerators
' =====================================================================
' ////////////// Used to set/reset HDC objects \\\\\\\\\\\\\\
Private Enum ColorObjects
    cObj_Brush = 0
    cObj_Pen = 1
    cObj_Text = 2
End Enum
' ////////////// Button Properties \\\\\\\\\\\\\\\
Public Enum ImagePlacementConstants ' image alignment
    lv_LeftEdge = 0
    lv_LeftOfCaption = 1
    lv_RightEdge = 2
    lv_RightOfCaption = 3
    lv_TopCenter = 4
    lv_BottomCenter = 5
End Enum
Public Enum ImageSizeConstants      ' image sizes
    lv_16x16 = 0
    lv_24x24 = 1
    lv_32x32 = 2
    lv_Fill_Stretch = 3
    lv_Fill_ScaleUpDown = 4
End Enum
Public Enum ButtonStyleConstants    ' button shapes
    lv_Rectangular = 0
    lv_LeftDiagonal = 1
    lv_RightDiagonal = 2
    lv_FullDiagonal = 3
End Enum
Public Enum HoverLockConstants      ' hover lock options
    lv_LockTextandBackColor = 0
    lv_LockTextColorOnly = 1
    lv_LockBackColorOnly = 2
    lv_NoLocks = 3
End Enum
Public Enum GradientConstants       ' gradient directions
    lv_NoGradient = 0
    lv_Left2Right = 1
    lv_Right2Left = 2
    lv_Top2Bottom = 3
    lv_Bottom2Top = 4
End Enum
Public Enum FontStyles
    lv_PlainStyle = 0
    lv_Bold = 2
    lv_Italic = 4
    lv_Underline = 8
    lv_BoldItalic = 2 Or 4
    lv_BoldUnderline = 2 Or 8
    lv_ItalicUnderline = 4 Or 8
    lv_BoldItalicUnderline = 2 Or 4 Or 8
End Enum

Public Property Let Caption(sCaption As String)
Attribute Caption.VB_Description = "The caption of the button. Double pipe (||) is a line break."
Attribute Caption.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
' Sets the button caption & hot key for the control
   Dim i As Integer
   Dim j As Integer
   
' We look from right to left. VB uses this logic & so do I
   i = InStrRev(sCaption, "&")
   Do While i
      If Mid$(sCaption, i, 2) = "&&" Then
         i = InStrRev(i - 1, sCaption, "&")
      Else
         j = i + 1: i = 0
     End If
   Loop
   
' if found, we use the next character as a hot key
   If j Then
      AccessKeys = Mid$(sCaption, j, 1)
   End If
   myProps.bCaption = sCaption                     ' cache the caption
   CalculateBoundingRects False                          ' recalculate button text/image bounding rects
   RedrawButton
   PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
   Caption = myProps.bCaption
End Property

Public Property Let CaptionAlign(nAlign As AlignmentConstants)
Attribute CaptionAlign.VB_Description = "Horizontal alignment of caption on the button."
Attribute CaptionAlign.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Caption options: Left, Right or Center Justified
   If nAlign < vbLeftJustify Or nAlign > vbCenter Then Exit Property
   
' also prevent left/right justifying captions when image is centered in caption
   If myImage.Align > lv_RightOfCaption And nAlign < vbCenter And (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
      If UserControl.Ambient.UserMode = False Then
         ' if not in user mode, then explain whey it is prevented
         MsgBox "When button images are aligned top/bottom center, " & vbCrLf & "button captions can only be center aligned", vbOKOnly + vbInformation
      End If
      Exit Property
   End If
   myProps.bCaptionAlign = nAlign
   CalculateBoundingRects False              ' recalculate text/image bounding rects
   RedrawButton
   PropertyChanged "CapAlign"
End Property

Public Property Get CaptionAlign() As AlignmentConstants
   CaptionAlign = myProps.bCaptionAlign
End Property

Public Property Let ButtonShape(nShape As ButtonStyleConstants)
Attribute ButtonShape.VB_Description = "Rectangular or various diagonal shapes"
Attribute ButtonShape.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Sets the button's shape (rectangular, or diagonal)


   myProps.bShape = nShape
   If myProps.bCaptionAlign <> vbCenter Then
      myProps.bCaptionAlign = vbCenter
   End If
   Call UserControl_Resize
   myProps.bCaptionAlign = Me.CaptionAlign
   DelayDrawing False
   PropertyChanged "Shape"
End Property

Public Property Get ButtonShape() As ButtonStyleConstants
   ButtonShape = myProps.bShape
End Property

Public Property Set Picture(xPic As StdPicture)
Attribute Picture.VB_Description = "The image used to display on the button."
' Sets the button image which to display
   Set myImage.Image = xPic
   If myImage.Size = 0 Then myImage.Size = 16
   GetGDIMetrics "Picture"
      CalculateBoundingRects True              ' recalculate button's text/image bounding rects
      RedrawButton
   PropertyChanged "Image"
End Property

Public Property Get Picture() As StdPicture
   Set Picture = myImage.Image
End Property

Public Property Let PictureAlign(ImgAlign As ImagePlacementConstants)
Attribute PictureAlign.VB_Description = "Alignment of the button image in relation to the caption and/or button."
' Image alignment options for button (6 different positions)
   If ImgAlign < lv_LeftEdge Or ImgAlign > lv_BottomCenter Then Exit Property
   myImage.Align = ImgAlign
   If ImgAlign = lv_BottomCenter Or ImgAlign = lv_TopCenter Then CaptionAlign = vbCenter
   CalculateBoundingRects False             ' recalculate button's text/image bounding rects
   RedrawButton
   PropertyChanged "ImgAlign"
End Property

Public Property Get PictureAlign() As ImagePlacementConstants
   PictureAlign = myImage.Align
End Property

Public Property Let Enabled(bEnabled As Boolean)
Attribute Enabled.VB_Description = "Determines if events are fired for this button."
Attribute Enabled.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
' Enables or disables the button
   If bEnabled = UserControl.Enabled Then Exit Property
   UserControl.Enabled = bEnabled
   RedrawButton
   PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let ShowFocusRect(bShow As Boolean)
Attribute ShowFocusRect.VB_Description = "Allows or prevents a focus rectangle from being displayed. In design mode, this may always be displayed for button set as Default."
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Shows/hides the focus rectangle when button comes into focus
   myProps.bShowFocus = bShow
   If ((myProps.bStatus And 1) = 1) Then
      ' if currently has the focus, then we take it off
      If Ambient.UserMode Then
         myProps.bStatus = myProps.bStatus And Not 1
         RedrawButton
      Else
         ' however, we don't if it is the default button
          MsgBox "The focus rectangle may appear on default buttons ONLY while in design mode, " & vbCrLf & _
                 "but will not appear when the form is running.", vbInformation + vbOKOnly
      End If
   Else
      RedrawButton
  End If
   PropertyChanged "Focus"
End Property

Public Property Get ShowFocusRect() As Boolean
   ShowFocusRect = myProps.bShowFocus
End Property

Public Property Let PictureSize(nSize As ImageSizeConstants)
Attribute PictureSize.VB_Description = "Various sizes for images used on buttons. Last 2 options center image automatically."
' Sets up to 5 picture sizes
   If PictureSize < lv_16x16 Or PictureSize > lv_Fill_ScaleUpDown Then Exit Property
   myImage.Size = (nSize + 2) * 8      ' I just want the size as pixel x pixel
   CalculateBoundingRects True         ' recalculate text/image bounding rects
   RedrawButton
   PropertyChanged "ImgSize"
End Property

Public Property Get PictureSize() As ImageSizeConstants
' parameters are 0,1,2,3,4 & 5, but we store them as 16,24,32,40, & 44
   If myImage.Size = 0 Then myImage.Size = 16
   PictureSize = Choose(myImage.Size / 8 - 1, lv_16x16, lv_24x24, lv_32x32, lv_Fill_Stretch, lv_Fill_ScaleUpDown)
End Property

Public Property Let MousePointer(nPointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Various optional mouse pointers to use when mouse is over the button"
' Sets the mouse pointer for the button
   UserControl.MousePointer = nPointer
   PropertyChanged "mPointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
   MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(nIcon As StdPicture)
Attribute MouseIcon.VB_Description = "Icon or cursor used to display when mouse is over the button. MousePointer must be set to Custom."
' Sets the mouse icon for the button, MousePointer must be vbCustom
   On Error GoTo ShowPropertyError
   Set UserControl.MouseIcon = nIcon
   If Not nIcon Is Nothing Then
      Me.MousePointer = vbCustom
      PropertyChanged "mIcon"
   End If
   Exit Property
   
ShowPropertyError:
   If Ambient.UserMode = False Then
      MsgBox Err.Description, vbInformation + vbOKOnly, "Select .ico Or .cur Files Only"
   End If
End Property

Public Property Get MouseIcon() As StdPicture
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set Font(nFont As StdFont)
Attribute Font.VB_Description = "Font used to display the caption."
Attribute Font.VB_ProcData.VB_Invoke_PropertyPutRef = ";Font"
' Sets the control's font & also the logical font to use on off-screen DC
   Set UserControl.Font = nFont
   GetGDIMetrics "Font"
   CalculateBoundingRects False          ' recalculate caption's text/image bounding rects
   RedrawButton
   PropertyChanged "Font"
End Property

Public Property Get Font() As StdFont
   Set Font = UserControl.Font
End Property

Public Property Let FontStyle(nStyle As FontStyles)
Attribute FontStyle.VB_Description = "Various font attributes that can be changed directly."
' Allows direct changes to font attributes
   With UserControl.Font
      .Bold = ((nStyle And lv_Bold) = lv_Bold)
      .Italic = ((nStyle And lv_Italic) = lv_Italic)
      .Underline = ((nStyle And lv_Underline) = lv_Underline)
   End With
   GetGDIMetrics "Font"
   CalculateBoundingRects False
   PropertyChanged "Font"
   RedrawButton
End Property
Public Property Get FontStyle() As FontStyles
   Dim nStyle As Integer
   
   nStyle = nStyle Or Abs(UserControl.Font.Bold) * 2
   nStyle = nStyle Or Abs(UserControl.Font.Italic) * 4
   nStyle = nStyle Or Abs(UserControl.Font.Underline) * 8
   FontStyle = nStyle
End Property

Public Property Let ForeColor(nColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "The color of the caption's font ."
Attribute ForeColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Sets the caption text color
   If nColor = UserControl.ForeColor Then Exit Property
   UserControl.ForeColor = nColor
  If myProps.bLockHover = lv_LockTextandBackColor Or myProps.bLockHover = lv_LockTextColorOnly Then
     Me.HoverForeColor = UserControl.ForeColor
   End If
   bNoRefresh = False
   RedrawButton
   PropertyChanged "cFore"
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property

Public Property Let BackColor(nColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Button back color. See also ResetDefaultColors"
Attribute BackColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Sets the backcolor of the button
   curBackColor = nColor
   If myProps.bLockHover = lv_LockBackColorOnly Or myProps.bLockHover = lv_LockTextandBackColor Then
       If myProps.bGradient Then
          Me.HoverBackColor = myProps.bGradientColor
       Else
          Me.HoverBackColor = nColor
       End If
   End If
   GetGDIMetrics "BackColor"
   RedrawButton
   PropertyChanged "cBack"
End Property
Public Property Get BackColor() As OLE_COLOR
   BackColor = curBackColor
End Property

Public Property Let GradientColor(nColor As OLE_COLOR)
Attribute GradientColor.VB_Description = "Secondary color used for gradient shades. The BackColor property is the primary color."
' Sets the gradient color. Gradients are used this way...
' Shade from BackColor to GradientColor
' GradientMode must be set

   If (myProps.bLockHover = lv_LockTextandBackColor _
   Or myProps.bLockHover = lv_LockBackColorOnly) _
   And myProps.bGradient > lv_NoGradient Then
      myProps.bBackHover = nColor
      myProps.bBackHover = Me.HoverBackColor
   End If
   
   myProps.bGradientColor = nColor
   GetGDIMetrics "BackColor"
   If myProps.bGradient Then RedrawButton
   PropertyChanged "cGradient"
End Property
Public Property Get GradientColor() As OLE_COLOR
   GradientColor = myProps.bGradientColor
End Property

Public Property Let GradientMode(nOpt As GradientConstants)
Attribute GradientMode.VB_Description = "Various directions to draw the gradient shading."
' Sets the direction of gradient shading

   If nOpt < lv_NoGradient Or nOpt > lv_Bottom2Top Then Exit Property
   myProps.bGradient = nOpt
   If myProps.bLockHover = lv_LockBackColorOnly _
   Or myProps.bLockHover = lv_LockTextandBackColor Then
      If nOpt > lv_NoGradient Then
         myProps.bBackHover = myProps.bGradientColor
      Else
         myProps.bBackHover = curBackColor
      End If
      myProps.bBackHover = Me.HoverBackColor
      GetGDIMetrics "BackColor"
   End If
   
   RedrawButton
   PropertyChanged "Gradient"
End Property
Public Property Get GradientMode() As GradientConstants
   GradientMode = myProps.bGradient
End Property

Public Property Let ResetDefaultColors(nDefault As Boolean)
Attribute ResetDefaultColors.VB_Description = "Resets button's back color and text color to Window's standard. The hover properties are also reset."
Attribute ResetDefaultColors.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Resets the BackColor, ForeColor, GradientColor,
' HoverBackColor & HoverForeColor to defaults

   If Ambient.UserMode Or nDefault = False Then Exit Property
   DelayDrawing True
   curBackColor = vbButtonFace
   Me.ForeColor = vbButtonText
   Me.GradientColor = vbButtonFace
   Me.GradientMode = lv_NoGradient
   Me.HoverColorLocks = lv_LockTextandBackColor
   myProps.bGradientColor = Me.GradientColor
   GetGDIMetrics "BackColor"
   DelayDrawing False
   PropertyChanged "cGradient"
   PropertyChanged "cBack"
End Property
Public Property Get ResetDefaultColors() As Boolean
   ResetDefaultColors = False
End Property

Public Property Let HoverColorLocks(nLock As HoverLockConstants)
Attribute HoverColorLocks.VB_Description = "Can ensure the hover colors match the caption and back colors. Click for more options."
Attribute HoverColorLocks.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
' Has two purposes.
' 1. If the lock wasn't set but is now set, then setting it will
' force HoverForeColor=ForeColor & HoverBackColor=Backcolor
' If gradeints in use, then HoverBackColor=GradientColor
' 2. If the lock was already set, then changing BackColor
' will force HoverBackColor to match. If gradients are used then
' it will force HoverBackColor to match GradientColor
' It will also force HoverForeColor to match ForeColor.

' After the locks have been set, manually changing the
' HoverForeColor, HoverBackColor will adjust/remove the lock

   myProps.bLockHover = nLock
   If myProps.bLockHover = lv_LockTextandBackColor _
   Or myProps.bLockHover = lv_LockBackColorOnly Then
      If myProps.bGradient Then
         myProps.bBackHover = myProps.bGradientColor
      Else
         myProps.bBackHover = curBackColor
      End If
      PropertyChanged "cBHover"
   End If
   
   If myProps.bLockHover = lv_LockTextandBackColor _
   Or myProps.bLockHover = lv_LockTextColorOnly Then
      myProps.bForeHover = UserControl.ForeColor
      PropertyChanged "cFHover"
   End If
   myProps.bBackHover = Me.HoverBackColor
   myProps.bForeHover = Me.HoverForeColor
   GetGDIMetrics "BackColor"
   PropertyChanged "LockHover"
End Property
Public Property Get HoverColorLocks() As HoverLockConstants
   HoverColorLocks = myProps.bLockHover
End Property

Public Property Let HoverForeColor(nColor As OLE_COLOR)
Attribute HoverForeColor.VB_Description = "Color of button caption's text when mouse is hovering over it. Affects the HoverLockColors property."
Attribute HoverForeColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Changes the text color when mouse is over the button
' Changing this property will affect the type of HoverLock

   If myProps.bForeHover = nColor Then Exit Property
   myProps.bForeHover = nColor
   PropertyChanged "cFHover"
   If nColor <> UserControl.ForeColor Then
      If myProps.bLockHover = lv_LockTextandBackColor Then
         myProps.bLockHover = lv_LockBackColorOnly
      Else
         If myProps.bLockHover = lv_LockTextColorOnly Then myProps.bLockHover = lv_NoLocks
      End If
   End If
   myProps.bLockHover = Me.HoverColorLocks
   PropertyChanged "cFHover"
End Property
Public Property Get HoverForeColor() As OLE_COLOR
   HoverForeColor = myProps.bForeHover
End Property

Public Property Let HoverBackColor(nColor As OLE_COLOR)
Attribute HoverBackColor.VB_Description = "Color of button background when mouse is hovering over it. Affects the HoverLockColors property."
Attribute HoverBackColor.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
' Changes the backcolor when mouse is over the button
' Changing this property will affect the type of HoverLock

   If myProps.bBackHover = nColor Then Exit Property
   myProps.bBackHover = nColor
   If nColor <> curBackColor Then
      If myProps.bLockHover = lv_LockTextandBackColor Then
         myProps.bLockHover = lv_LockTextColorOnly
      Else
         If myProps.bLockHover = lv_LockBackColorOnly Then
            myProps.bLockHover = lv_NoLocks
         End If
      End If
   End If
   myProps.bLockHover = Me.HoverColorLocks
   GetGDIMetrics "BackColor"
   PropertyChanged "cBHover"
End Property
Public Property Get HoverBackColor() As OLE_COLOR
HoverBackColor = myProps.bBackHover
End Property

Public Property Get hdc() As Long

' Makes the control's hDC availabe at runtime

hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long

' Makes the control's hWnd available at runtime

hwnd = UserControl.hwnd
End Property

' //////////////////// GENERAL FUNCTIONS, PUBLIC \\\\\\\\\\\\\\\\\\\\\
Public Sub Refresh()

' Refreshes the button & can be called from any form/module
bNoRefresh = False
RedrawButton
End Sub

Public Sub DelayDrawing(bDelay As Boolean)

' Used to prevent redrawing button until all properties are set.
' Should you want to set multiple properties of the control during runtime
' call this function first with a TRUE parameter. Set your button
' attributes and then call it again with a FALSE property to update the
' button.   IMPORTANT: If called with a TRUE parameter you must
' also release it with a call and a FALSE parameter

' NOTE: this function will prevent flicker when several properties
' are being changed at once during run time. It is similar to
' the BeginPaint & EndPaint API functionality
bNoRefresh = bDelay
If bDelay = False Then Refresh
End Sub

Private Sub RedrawButton()
' ==================================================
' Main switchboard routine for redrawing a button
' ==================================================
If bNoRefresh = True Then Exit Sub

Dim polyPts(0 To 15) As POINTAPI, polyColors(1 To 12) As Long
Dim ActiveStatus As Integer, ActiveClipRgn As Integer
    DrawButton_WinXP polyPts(), polyColors(), ActiveStatus, ActiveClipRgn
    
Dim FocusColor As Long
FocusColor = polyColors(12)
Erase polyPts()
Erase polyColors()
If ActiveClipRgn Then
    GetSetOffDC False    ' copy the offscreen DC onto the control
    ' to help preventing unnecessary border drawing for round/custom buttons
    ' we returned the current clipping region & will draw the focus
    ' rectangles directly on the HDC for these type buttons only
    If ActiveClipRgn > 1 Then
        Dim tRgn As Long
        If tRgn Then
            Dim hBrush As Long
            hBrush = CreateSolidBrush(FocusColor)
            FrameRgn UserControl.hdc, tRgn, hBrush, 1, 1
            If tRgn <> ButtonDC.ClipBorder And tRgn <> ButtonDC.ClipRgn Then DeleteObject tRgn
            DeleteObject hBrush
        End If
    End If
    UserControl.Refresh
End If
End Sub

Friend Sub TimerUpdate(lvTimerID As Long)

' pretty good way to determine when cursor moves outside of any shape region
' especially useful for my diagonal/round buttons since they are not your typical
' rectangular shape.

Dim mousePt As POINTAPI, cRect As RECT
GetCursorPos mousePt
If WindowFromPoint(mousePt.x, mousePt.y) <> UserControl.hwnd Then
    ' when exits button area, kill the timer
    KillTimer UserControl.hwnd, lvTimerID
    myProps.bStatus = myProps.bStatus And Not 4
    bTimerActive = False
    bNoRefresh = False
'    RaiseEvent MouseOnButton(False)
    RedrawButton
End If
End Sub

Private Sub CalculateBoundingRects(bNormalizeImage As Boolean)

' Routine measures and places the rectangles to draw
' the caption and image on the control. The results
' are cached so this routine doesn't need to run
' every time the button is redrawn/painted

Dim cRect As RECT, tRect As RECT, iRect As RECT
Dim imgOffset As RECT, bImgWidthAdj As Boolean, bImgHeightAdj As Boolean
Dim rEdge As Long, lEdge As Long, adjWidth As Long

' calculations needed for diagonal buttons
Select Case myProps.bShape
Case lv_RightDiagonal
    rEdge = myProps.bSegPts.y + ((ScaleWidth - myProps.bSegPts.y) \ 3)
    adjWidth = rEdge
Case lv_LeftDiagonal
    lEdge = myProps.bSegPts.x - (myProps.bSegPts.x \ 3) + 3
    rEdge = ScaleWidth
    adjWidth = ScaleWidth - lEdge
Case lv_FullDiagonal
    lEdge = myProps.bSegPts.x - (myProps.bSegPts.x \ 3) + 3
    rEdge = myProps.bSegPts.y + ((ScaleWidth - myProps.bSegPts.y) \ 3)
    adjWidth = rEdge - lEdge
Case Else
    adjWidth = myProps.bSegPts.y
    rEdge = ScaleWidth
End Select

If (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
    ' image in use, calculations for image rectangle
    If myImage.Size < 33 Then
        Select Case myImage.Align
        Case lv_LeftEdge, lv_LeftOfCaption
            imgOffset.Left = myImage.Size
            bImgWidthAdj = True
        Case lv_RightEdge, lv_RightOfCaption
            imgOffset.Right = myImage.Size
            bImgWidthAdj = True
        Case lv_TopCenter
            imgOffset.Top = myImage.Size
            bImgHeightAdj = True
        Case lv_BottomCenter
            imgOffset.Bottom = myImage.Size
            bImgHeightAdj = True
        End Select
    End If
End If


If Len(myProps.bCaption) Then
    Dim sCaption As String  ' note: Replace$ not compatible with VB5
    sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
    ' calculate total available button width available for text
    cRect.Right = adjWidth - 8 - (myImage.Size * Abs(CInt(bImgWidthAdj)))
    cRect.Bottom = ScaleHeight - 8 - (myImage.Size * Abs(CInt(bImgHeightAdj = True And myImage.Align > lv_RightOfCaption)))
    ' calculate size of rectangle to hold that text, using multiline flag
    DrawText ButtonDC.hdc, sCaption, Len(sCaption), cRect, DT_CALCRECT Or DT_WORDBREAK
End If

' now calculate the position of the text rectangle
If Len(myProps.bCaption) Then
    tRect = cRect
    Select Case myProps.bCaptionAlign
    Case vbLeftJustify
        OffsetRect tRect, imgOffset.Left + lEdge + 4 + (Abs(CInt(imgOffset.Left > 0) * 4)), 0
    Case vbRightJustify
        OffsetRect tRect, rEdge - imgOffset.Right - 4 - cRect.Right - (Abs(CInt(imgOffset.Right > 0) * 4)), 0
    Case vbCenter
        If imgOffset.Left > 0 And myImage.Align = lv_LeftOfCaption Then
            OffsetRect tRect, (adjWidth - (imgOffset.Left + cRect.Right + 4)) \ 2 + lEdge + 4 + imgOffset.Left, 0
        Else
            If imgOffset.Right > 0 And myImage.Align = lv_RightOfCaption Then
                OffsetRect tRect, (adjWidth - (imgOffset.Right + cRect.Right + 4)) \ 2 + lEdge, 0
            Else
                OffsetRect tRect, ((adjWidth - (imgOffset.Left + imgOffset.Right)) - cRect.Right) \ 2 + lEdge + imgOffset.Left, 0
            End If
        End If
    End Select
End If
If (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
    ' finalize image rectangle position
    Select Case myImage.Align
    Case lv_LeftEdge
        iRect.Left = lEdge + 4
    Case lv_LeftOfCaption
        If Len(myProps.bCaption) Then
            iRect.Left = tRect.Left - 4 - imgOffset.Left
        Else
            iRect.Left = lEdge + 4
        End If
    Case lv_RightOfCaption
        If Len(myProps.bCaption) Then
            iRect.Left = tRect.Right + 4
        Else
            iRect.Left = rEdge - 4 - imgOffset.Right
        End If
    Case lv_RightEdge
        iRect.Left = rEdge - 4 - imgOffset.Right
    Case lv_TopCenter
        iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Top)) \ 2
        OffsetRect tRect, 0, iRect.Top + 2 + imgOffset.Top
    Case lv_BottomCenter
        iRect.Top = (ScaleHeight - (cRect.Bottom + imgOffset.Bottom)) \ 2 + cRect.Bottom + 4
        OffsetRect tRect, 0, iRect.Top - 2 - cRect.Bottom
    End Select
    If myImage.Align < lv_TopCenter Then
        OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
        iRect.Top = (ScaleHeight - myImage.Size) \ 2
    Else
        iRect.Left = (adjWidth - myImage.Size) \ 2 + lEdge
    End If
    iRect.Right = iRect.Left + myImage.Size
    iRect.Bottom = iRect.Top + myImage.Size
Else
    OffsetRect tRect, 0, (ScaleHeight - cRect.Bottom) \ 2
End If
' sanity checks
If tRect.Top < 4 Then tRect.Top = 4
If tRect.Left < 4 + lEdge Then tRect.Left = 4 + lEdge
If tRect.Right > rEdge - 4 Then tRect.Right = rEdge - 4
If tRect.Bottom > ScaleHeight - 5 Then tRect.Bottom = ScaleHeight - 5
myProps.bRect = tRect
Select Case myImage.Size
Case Is < 33
    If iRect.Top < 4 Then iRect.Top = 4
    If iRect.Left < 4 + lEdge Then iRect.Left = 4 + lEdge
    If iRect.Right > rEdge - 4 Then iRect.Right = rEdge - 4
    If iRect.Bottom > ScaleHeight - 5 Then iRect.Bottom = ScaleHeight - 5
Case 40 ' stretch
        SetRect iRect, 3, 3, ScaleWidth - 3, ScaleHeight - 3
    bNormalizeImage = True
Case Else   ' scale
        If (myImage.SourceSize.x + myImage.SourceSize.y) > 0 Then
            ScaleImage adjWidth - 12, ScaleHeight - 12, cRect.Right, cRect.Bottom
            iRect.Left = (adjWidth - cRect.Right) \ 2 + lEdge
            iRect.Top = (ScaleHeight - cRect.Bottom) \ 2
            iRect.Right = iRect.Left + cRect.Right
            iRect.Bottom = iRect.Top + cRect.Bottom
            bNormalizeImage = True
        End If
End Select
myImage.iRect = iRect
If bNormalizeImage Then NormalizeImage iRect.Right - iRect.Left, iRect.Bottom - iRect.Top, 0
End Sub

Private Sub GetSetOffDC(bSet As Boolean)

' This sets up our off screen DC & pastes results onto our control.

If bSet = True Then
    If ButtonDC.hdc = 0 Then
        ButtonDC.hdc = CreateCompatibleDC(UserControl.hdc)
        SetBkMode ButtonDC.hdc, 3&
        ' by pulling these objects now, we ensure no memory leaks &
        ' changing the objects as needed can be done in 1 line of code
        ' in the SetButtonColors routine
        ButtonDC.OldBrush = SelectObject(ButtonDC.hdc, CreateSolidBrush(0&))
        ButtonDC.OldPen = SelectObject(ButtonDC.hdc, CreatePen(0&, 1&, 0&))
    End If
    GetGDIMetrics "Font"
    If ButtonDC.OldBitmap = 0 Then
        Dim hBmp As Long
        hBmp = CreateCompatibleBitmap(UserControl.hdc, ScaleWidth, ScaleHeight)
        ButtonDC.OldBitmap = SelectObject(ButtonDC.hdc, hBmp)
    End If
Else
    BitBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, ButtonDC.hdc, 0, 0, vbSrcCopy
End If
End Sub

Private Sub DrawRect(m_hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
                                   ByVal X2 As Long, ByVal Y2 As Long, _
                                   tColor As Long, Optional pColor As Long = -1, _
                                   Optional PenWidth As Long = 0, Optional PenStyle As Long = 0)

' Simple routine to draw a rectangle

If pColor <> -1 Then SetButtonColors True, m_hDC, cObj_Pen, pColor, , PenWidth, , PenStyle
SetButtonColors True, m_hDC, cObj_Brush, tColor, (pColor = -1)
Call Rectangle(m_hDC, X1, Y1, X2, Y2)
End Sub


Private Sub SetButtonColors(bSet As Boolean, m_hDC As Long, TypeObject As ColorObjects, lColor As Long, _
    Optional bSamePenColor As Boolean = True, Optional PenWidth As Long = 1, _
    Optional bSwapPens As Boolean = False, Optional PenStyle As Long = 0)

' This is the basic routine that sets a DC's pen, brush or font color

' here we store the most recent "sets" so we can reset when needed
Dim tBrush As Long, tPen As Long
If bSet Then    ' changing a DC's setting
    Select Case TypeObject
    Case cObj_Brush         ' brush is being changed
        DeleteObject SelectObject(ButtonDC.hdc, CreateSolidBrush(lColor))
        If bSamePenColor Then   ' if the pen color will be the same
            DeleteObject SelectObject(ButtonDC.hdc, CreatePen(PenStyle, PenWidth, lColor))
        End If
    Case cObj_Pen   ' pen is being changed (mostly for drawing lines)
        DeleteObject SelectObject(ButtonDC.hdc, CreatePen(PenStyle, PenWidth, lColor))
    Case cObj_Text  ' text color is changing
        SetTextColor m_hDC, ConvertColor(lColor)
    End Select
Else            ' resetting the DC back to the way it was
    DeleteObject SelectObject(ButtonDC.hdc, ButtonDC.OldBrush)
    DeleteObject SelectObject(ButtonDC.hdc, ButtonDC.OldPen)
End If
End Sub

Private Function ConvertColor(tColor As Long) As Long

' Converts VB color constants to real color values

If tColor < 0 Then
    ConvertColor = GetSysColor(tColor And &HFF&)
Else
    ConvertColor = tColor
End If
End Function

Private Sub CreateButtonRegion()

' this function creates the regions for the specific type of button style

Dim rgnA As Long, rgn2Use As Long, i As Long
Dim lRatio As Single, lEdge As Long, rEdge As Long, Wd As Long
Dim ptTRI(0 To 9) As POINTAPI

myProps.bSegPts.x = 0
myProps.bSegPts.y = ScaleWidth

SelectClipRgn ButtonDC.hdc, 0
If ButtonDC.ClipRgn Then
    ' this was set for round buttons
    DeleteObject ButtonDC.ClipRgn
    ButtonDC.ClipRgn = 0
End If
If ButtonDC.ClipBorder Then
    DeleteObject ButtonDC.ClipBorder
    ButtonDC.ClipBorder = 0
End If
 ' diagonals
    ' here is my trick for ensuring a sharp edge on diagonal buttons.
    ' Basically a bastardized carpenters formula for right angles
    ' (i.e., 3+4=5 < the hypoteneus). Here I want a 60 degree angle,
    ' and not a 45 degree angle. The difference is sharp or choppy.
    ' Based off of the button height, I need to figure how much of
    ' the opposite end I need to cutoff for the diagonal edge
    lRatio = (ScaleHeight + 1) / 4
    Wd = ScaleWidth
    lEdge = (4 * lRatio)
    ' here we ensure a width of at least 5 pixels wide
    Do While Wd - lEdge < 5
        Wd = Wd + 5
    Loop
    If Wd <> ScaleWidth Then
        ' resize the control if necessary
        DelayDrawing True
        If (TypeOf Parent Is MDIForm) Then
            UserControl.Width = ScaleX(Wd, vbPixels, vbTwips)
        Else
            UserControl.Width = ScaleX(Wd, vbPixels, Parent.ScaleMode)
        End If
        myProps.bSegPts.y = ScaleWidth
        bNoRefresh = False
    End If
    rEdge = ScaleWidth - lEdge
    ' initial dimensions of our rectangle
    ptTRI(0).x = 0: ptTRI(0).y = 0
    ptTRI(1).x = 0
    ptTRI(1).y = ScaleHeight + 1
    ptTRI(2).x = ScaleWidth + 1
    ptTRI(2).y = ScaleHeight + 1
    ptTRI(3).x = ScaleWidth + 1
    ptTRI(3).y = 0
    ' now modify the left/right side as needed
    If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_LeftDiagonal Then
        ptTRI(1).x = lEdge  ' left portion
        myProps.bSegPts.x = lEdge
    End If
    If myProps.bShape = lv_FullDiagonal Or myProps.bShape = lv_RightDiagonal Then
        ptTRI(3).x = rEdge + 1        ' bottom right
        myProps.bSegPts.y = rEdge
    End If
    ' for rounded corner buttons, we'll take of the corner pixels where appropriate when the
    ' diagonal button is not a fully-segmeneted type. Diagonal edge corners are always sharp,
    ' never rounded.
    rgn2Use = CreatePolygonRgn(ptTRI(0), 4, 2)
        If myProps.bShape = lv_RightDiagonal Then GoSub LopOffCorners3
        If myProps.bShape = lv_LeftDiagonal Then GoSub LopOffCorners4

Erase ptTRI
If rgnA Then DeleteObject rgnA
SetWindowRgn UserControl.hwnd, rgn2Use, True
If myProps.bSegPts.y = 0 Then myProps.bSegPts.y = ScaleWidth
ExitRegionCreator:
Exit Sub

LopOffCorners1: ' left side top/bottom corners (Java/Win3.x)
        rgnA = CreateRectRgn(0, 0, 1, 1)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
LopOffCorners2: ' right side top/bottom corners (Java/Win3.x)
        rgnA = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth - 1, ScaleHeight - 1)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    rgnA = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    Return
LopOffCorners3: ' left side top/bottom corners (XP/Mac)
    ptTRI(0).x = 0: ptTRI(0).y = 0
    ptTRI(1).x = 2: ptTRI(1).y = 0
    ptTRI(2).x = 0: ptTRI(2).y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).x = 0: ptTRI(0).y = ScaleHeight
    ptTRI(1).x = 3: ptTRI(1).y = ScaleHeight
    ptTRI(2).x = 0: ptTRI(2).y = ScaleHeight - 3
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
Return
LopOffCorners4: ' right side top/bottom corners (XP/Mac)
    ptTRI(0).x = ScaleWidth: ptTRI(0).y = 0
    ptTRI(1).x = ScaleWidth - 2: ptTRI(1).y = 0
    ptTRI(2).x = ScaleWidth: ptTRI(2).y = 2
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
    ptTRI(0).x = ScaleWidth: ptTRI(0).y = ScaleHeight
    ptTRI(1).x = ScaleWidth - 3: ptTRI(1).y = ScaleHeight
    ptTRI(2).x = ScaleWidth: ptTRI(2).y = ScaleHeight - 3
    rgnA = CreatePolygonRgn(ptTRI(0), 3, 2)
    CombineRgn rgn2Use, rgn2Use, rgnA, RGN_DIFF
    DeleteObject rgnA
Return
End Sub

Private Sub ScaleImage(SizeX As Long, SizeY As Long, ImgX As Long, ImgY As Long)
' helper function for resizing images to scale
Dim Ratio(0 To 1) As Double
Ratio(0) = SizeX / myImage.SourceSize.x
Ratio(1) = SizeY / myImage.SourceSize.y
If Ratio(1) < Ratio(0) Then Ratio(0) = Ratio(1)
ImgX = myImage.SourceSize.x * Ratio(0)
ImgY = myImage.SourceSize.y * Ratio(0)
Erase Ratio
End Sub
Private Sub NormalizeImage(newSizeX As Long, newSizeY As Long, rtnRgn As Long)
If myImage.Image Is Nothing Then Exit Sub
If myImage.TransImage Then DeleteObject myImage.TransImage
If myImage.Type Then Exit Sub
' pain in the tush.
' In order to make a bitmap transparent, we need to decide which color will be the transparent color
' Well, the API CopyImage is used to resize images to fit the buttons. The downside is that this
' API has a habit of changing pixel colors very slightly. Even a single RGB value changed by a
' value of one can prevent the transparency routines from making the image transparent.

' This routine cleans up an image to help ensure it can be made transparent.
' Note: This routine is called each time a button is resized
' So even though this routine can be time consuming, it is called normally during IDE or initial form load.

' Last but not least, the routine also builds the non-rectangular regions for
' custom button shapes & returns the regions back to the CreateButtonRegion routine

Dim cTrans As Long, lImage As Long
Dim valGreen As Long, valRed As Long, valBlue As Long
Dim tGreen As Long, tRed As Long, tBlue As Long
Dim x As Long, y As Long, cPixel As Long
Dim oldBMP As Long, newDC As Long
' these are used only for creating custom button regions
Dim rgnA As Long, xySet As Long, bAdjRegion As Boolean, tRect As RECT

' can't use ButtonDC.hDC -- need to create another DC 'cause if a clipping
' region is active (shaped/circular buttons), selecting image into DC may fail
newDC = CreateCompatibleDC(UserControl.hdc)
' get the image into a DC so we can clean it up
If myImage.Type Then    ' icons
    myImage.TransImage = CreateCompatibleBitmap(UserControl.hdc, newSizeX, newSizeY)
    oldBMP = SelectObject(newDC, myImage.TransImage)
    DrawIconEx newDC, 0, 0, myImage.Image.Handle, newSizeX, newSizeY, 0&, 0&, &H3
Else    ' bitmaps
    myImage.TransImage = CopyImage(myImage.Image.Handle, myImage.Type, newSizeX, newSizeY, ByVal 0&)
    oldBMP = SelectObject(newDC, myImage.TransImage)
End If
' determine the mask color (top left corner pixel)
cTrans = GetPixel(newDC, 0, 0)
' get the RGB values for that pixel
valRed = (cTrans \ (&H100 ^ 0) And &HFF)
valGreen = (cTrans \ (&H100 ^ 1) And &HFF)
valBlue = (cTrans \ (&H100 ^ 2) And &HFF)
If rtnRgn Then
    ButtonDC.ClipBorder = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    ButtonDC.ClipRgn = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
End If
' now loop thru each pixel & clean up any that were changed by the CopyImage API
For y = 0 To newSizeY
    xySet = -1      ' custom regions only. flag indicating rectangle not started
    For x = 0 To newSizeX
        cPixel = GetPixel(newDC, x, y)                      ' current pixel
        tRed = (cPixel \ (&H100 ^ 0) And &HFF)          ' RGB values for current pixel
        tGreen = (cPixel \ (&H100 ^ 1) And &HFF)
        tBlue = (cPixel \ (&H100 ^ 2) And &HFF)
        ' Test to see if the current pixel is real close to the transparent color used & change it if so
        If tRed >= valRed - 3 And tRed <= valRed + 3 And _
            tBlue >= valBlue - 3 And tBlue <= valBlue + 3 And _
                tGreen >= valGreen - 3 And tGreen <= valGreen + 3 Then
            SetPixel newDC, x, y, cTrans
            ' custom regions only...
            ' if it is a transparent pixel, set the start of the rectangle if needed
            If xySet = -1 Then xySet = x
            If x = newSizeX Then bAdjRegion = True
        Else
            If xySet > -1 Then bAdjRegion = True
        End If
        ' custom regions only
        If bAdjRegion And rtnRgn <> 0 Then
            ' not a transparent pixel, so we need to close the rectangle and remove it from the regions
            ' set up the rectangle to remove from the core region & remove it
            SetRect tRect, xySet, y, x, y + 1
            rgnA = CreateRectRgn(tRect.Left, tRect.Top, tRect.Right, tRect.Bottom)
            CombineRgn rtnRgn, rtnRgn, rgnA, RGN_DIFF
            DeleteObject rgnA
            ' create a 1 pixel inner border region between outer edge & image
            ' we do this by expanding the rectangle to be removed thereby creating a
            ' smaller overall region
            InflateRect tRect, 1, 1
            rgnA = CreateRectRgn(tRect.Left, tRect.Top, tRect.Right, tRect.Bottom)
            CombineRgn ButtonDC.ClipBorder, ButtonDC.ClipBorder, rgnA, RGN_DIFF
            DeleteObject rgnA
            ' used if shaped button has a border or used as check box/option button
            ' create another 1 clipping region for the actual image/background
            InflateRect tRect, 1, 1
            rgnA = CreateRectRgn(tRect.Left, tRect.Top, tRect.Right, tRect.Bottom)
            CombineRgn ButtonDC.ClipRgn, ButtonDC.ClipRgn, rgnA, RGN_DIFF
            DeleteObject rgnA
            xySet = -1      ' reset flag to look for another rectangle to be removed
            bAdjRegion = False
        End If
    Next
Next
' Pull the image out of the DC & use it for all other image routines
SelectObject newDC, oldBMP
DeleteDC newDC
myImage.TransSize.x = newSizeX
myImage.TransSize.y = newSizeY
End Sub

Private Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, _
                                                    lBMPsource As Long, bmpRect As RECT, _
                                                    Optional lMaskColor As Long = -1, _
                                                    Optional lNewBmpCx As Long, _
                                                    Optional lNewBmpCy As Long)
Const DSna = &H220326 '0x00220326
' =====================================================================
' A pretty good transparent bitmap maker I use in several projects
' Modified here to remove stuff I wont use (i.e., Flipping/Rotating images)
' =====================================================================

    Dim lMask2Use As Long 'COLORREF
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    Dim lHDCMem As Long, lHDCscreen As Long, lHDCsrc As Long, lHDCMask As Long, lHDCcolor As Long
    Dim x As Long, y As Long, srcX As Long, srcY As Long
    Dim lRatio(0 To 1) As Single
    Dim hPalOld As Long, hPalMem As Long
    
    lHDCscreen = GetDC(0&)
    lHDCsrc = CreateCompatibleDC(lHDCscreen)     'Create a temporary HDC compatible to the Destination HDC
    SelectObject lHDCsrc, lBMPsource             'Select the bitmap

        srcX = myImage.TransSize.x ' lNewBmpCx                  'Get width of bitmap
        srcY = myImage.TransSize.y ' lNewBmpCy                'Get height of bitmap
        
        If bmpRect.Right = 0 Then bmpRect.Right = srcX Else srcX = bmpRect.Right - bmpRect.Left
        If bmpRect.Bottom = 0 Then bmpRect.Bottom = srcY Else srcY = bmpRect.Bottom - bmpRect.Top
        
        If (destRect.Right) = 0 Then x = lNewBmpCx Else x = (destRect.Right - destRect.Left)
        If (destRect.Bottom) = 0 Then y = lNewBmpCy Else y = (destRect.Bottom - destRect.Top)
        If lNewBmpCx > x Or lNewBmpCy > y Then
            lRatio(0) = (x / lNewBmpCx)
            lRatio(1) = (y / lNewBmpCy)
            If lRatio(1) < lRatio(0) Then lRatio(0) = lRatio(1)
            lNewBmpCx = lRatio(0) * lNewBmpCx
            lNewBmpCy = lRatio(0) * lNewBmpCy
            Erase lRatio
        End If
    
    lMask2Use = ConvertColor(GetPixel(lHDCsrc, 0, 0))
    
    'Create some DCs & bitmaps
    lHDCMask = CreateCompatibleDC(lHDCscreen)
    lHDCMem = CreateCompatibleDC(lHDCscreen)
    lHDCcolor = CreateCompatibleDC(lHDCscreen)
    
    lBmColor = CreateCompatibleBitmap(lHDCscreen, srcX, srcY)
    lBmAndMem = CreateCompatibleBitmap(lHDCscreen, x, y)
    lBmMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)
    
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ReleaseDC 0&, lHDCscreen
    
' ====================== Start working here ======================
    
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lHDCMem, 0, True)
    RealizePalette lHDCMem
    
    BitBlt lHDCMem, 0&, 0&, x, y, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy
    
    
    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy
    
    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite
    
    BitBlt lHDCMask, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lHDCMask, 0, 0, DSna

    StretchBlt lHDCMem, 0, 0, lNewBmpCx, lNewBmpCy, lHDCMask, 0&, 0&, srcX, srcY, vbSrcAnd
    
    StretchBlt lHDCMem, 0&, 0&, lNewBmpCx, lNewBmpCy, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint
    
    BitBlt lHDCdest, destRect.Left, destRect.Top, x, y, lHDCMem, 0&, 0&, vbSrcCopy
    
    'Delete memory bitmaps & DCs
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc
End Sub

Private Sub DrawButtonIcon(iRect As RECT)

' Routine will draw the button image

If (myImage.SourceSize.x + myImage.SourceSize.y) = 0 Then Exit Sub
If myImage.TransImage = 0 Then NormalizeImage iRect.Right - iRect.Left, iRect.Bottom - iRect.Top, 0

Dim imgWidth As Long, imgHeight As Long
Dim rcImage As RECT, dRect As RECT
Const MAGICROP = &HB8074A

imgWidth = iRect.Right - iRect.Left
imgHeight = iRect.Bottom - iRect.Top
' destination rectangle for drawing on the DC
dRect = iRect

Dim hMemDC As Long
If UserControl.Enabled Then
    hMemDC = ButtonDC.hdc
Else
    Dim hBitmap As Long, hOldBitmap As Long
    Dim hOldBrush As Long
    Dim hOldBackColor As Long, hbrShadow As Long, hbrHilite As Long
    
    ' Create a temporary DC and bitmap to hold the image
    hMemDC = CreateCompatibleDC(ButtonDC.hdc)
    hBitmap = CreateCompatibleBitmap(ButtonDC.hdc, imgWidth, imgHeight)
    hOldBitmap = SelectObject(hMemDC, hBitmap)
    PatBlt hMemDC, 0, 0, imgWidth, imgHeight, WHITENESS
    OffsetRect dRect, -dRect.Left, -dRect.Top
End If
    
    If myImage.Type = CI_ICON Then
'        ' draw icon directly onto the temporary DC
'        ' for icons, we can draw directly on the destination DC
        DrawIconEx hMemDC, dRect.Left, dRect.Top, myImage.Image.Handle, imgWidth, imgHeight, 0, 0, &H3
    Else
        ' draw transparent bitmap onto the temporary DC
        DrawTransparentBitmap hMemDC, dRect, myImage.TransImage, rcImage, , imgWidth, imgHeight
    End If
  
If UserControl.Enabled = False Then
    hOldBackColor = SetBkColor(ButtonDC.hdc, vbWhite)
    hbrShadow = CreateSolidBrush(vbGray)
    hOldBrush = SelectObject(ButtonDC.hdc, hbrShadow)
    BitBlt ButtonDC.hdc, iRect.Left, iRect.Top, imgWidth, imgHeight, hMemDC, 0, 0, MAGICROP
  
    SetBkColor ButtonDC.hdc, hOldBackColor
    SelectObject ButtonDC.hdc, hOldBrush
    SelectObject hMemDC, hOldBitmap
    DeleteObject hbrShadow
    DeleteObject hBitmap
    DeleteDC hMemDC
End If
End Sub

Private Function ShadeColor(lColor As Long, shadeOffset As Integer, lessBlue As Boolean, _
    Optional bFocusRect As Boolean, Optional bInvert As Boolean) As Long

' Basically supply a value between -255 and +255. Positive numbers make
' the passed color lighter and negative numbers make the color darker

Dim valRGB(0 To 2) As Integer, i As Integer

CalcNewColor:
valRGB(0) = (lColor And &HFF) + shadeOffset
valRGB(1) = ((lColor And &HFF00&) / 255&) + shadeOffset
If lessBlue Then
    valRGB(2) = (lColor And &HFF0000) / &HFF00&
    valRGB(2) = valRGB(2) + ((valRGB(2) * CLng(shadeOffset)) \ &HC0)
Else
    valRGB(2) = (lColor And &HFF0000) / &HFF00& + shadeOffset
End If

For i = 0 To 2
    If valRGB(i) > 255 Then valRGB(i) = 255
    If valRGB(i) < 0 Then valRGB(i) = 0
    If bInvert = True Then valRGB(i) = Abs(255 - valRGB(i))
Next
ShadeColor = valRGB(0) + 256& * valRGB(1) + 65536 * valRGB(2)
Erase valRGB

If bFocusRect = True And (ShadeColor = vbBlack Or ShadeColor = vbWhite) Then
    shadeOffset = -shadeOffset
    If shadeOffset = 0 Then shadeOffset = 64
    GoTo CalcNewColor
End If
End Function

Private Sub GetGDIMetrics(sObject As String)

' This routine caches information we don't want to keep gathering every time a button is redrawn.

Select Case sObject
Case "Font"
    ' called when font is changed or control is initialized
    Dim newFont As LOGFONT
    newFont.lfCharSet = 1
    newFont.lfFaceName = UserControl.Font.Name & Chr$(0)
    newFont.lfHeight = (UserControl.Font.Size * -20) / Screen.TwipsPerPixelY
    newFont.lfWeight = UserControl.Font.Weight
    newFont.lfItalic = Abs(CInt(UserControl.Font.Italic))
    newFont.lfStrikeOut = Abs(CInt(UserControl.Font.Strikethrough))
    newFont.lfUnderline = Abs(CInt(UserControl.Font.Underline))
    If ButtonDC.OldFont Then
        DeleteObject SelectObject(ButtonDC.hdc, CreateFontIndirect(newFont))
    Else
        ButtonDC.OldFont = SelectObject(ButtonDC.hdc, CreateFontIndirect(newFont))
    End If
Case "Picture"
    ' get key image information
    Dim bmpInfo As BITMAP, icoInfo As ICONINFO
    If myImage.Image Is Nothing Then
        If myImage.TransImage Then DeleteObject myImage.TransImage
        myImage.SourceSize.x = 0
        myImage.SourceSize.y = 0
    Else
        GetGDIObject myImage.Image.Handle, LenB(bmpInfo), bmpInfo
        If bmpInfo.bmBits = 0 Then
            GetIconInfo myImage.Image.Handle, icoInfo
            If icoInfo.hbmColor <> 0 Then
                ' downside... API creates 2 bitmaps that we need to destroy since they aren't used in this
                ' routine & are not destroyed automatically. To prevent memory leak, we destroy them here
                GetGDIObject icoInfo.hbmColor, LenB(bmpInfo), bmpInfo
                DeleteObject icoInfo.hbmColor
                If icoInfo.hbmMask <> 0 Then DeleteObject icoInfo.hbmMask
                myImage.Type = CI_ICON        ' flag indicating image is an icon
            End If
        Else
            myImage.Type = CI_BITMAP     ' flag indicating image is a bitmap
        End If
        myImage.SourceSize.x = bmpInfo.bmWidth
        myImage.SourceSize.y = bmpInfo.bmHeight
    End If
Case "BackColor"
    adjBackColorUp = ConvertColor(curBackColor)
    adjBackColorDn = adjBackColorUp
    adjHoverColor = ConvertColor(myProps.bBackHover)
    
        adjBackColorUp = ShadeColor(adjBackColorUp, &H30, True)
        adjHoverColor = ShadeColor(adjHoverColor, &H30, True)
        adjBackColorDn = ShadeColor(adjBackColorUp, -&H20, True)
        cCheckBox = ShadeColor(vbWhite, -&H20, True)
End Select
End Sub

Private Sub DrawButtonBackground(bColor As Long, ActiveStatus As Integer, ActiveRegion As Integer, _
                Optional bGradientColor As Long = -1, Optional bHoverColor As Long = -1)
                
' Fill the button with the appropriate backcolor

Call DrawCustomBorders(ActiveRegion)
If ActiveRegion = 0 Then Exit Sub

Dim i As Integer, bColor2Use As Long, rtnVal As Integer
Dim focusOffset As Byte, isDown As Byte

focusOffset = Abs(((myProps.bStatus And 1) = 1))
isDown = Abs((myProps.bStatus And 6) = 6)
If isDown Then ActiveStatus = 2 Else ActiveStatus = focusOffset
                            
If bHoverColor < 0 Then bHoverColor = bColor
If bTimerActive And isDown = 0 Then
    bColor2Use = bHoverColor
Else
    bColor2Use = bColor
End If
If myProps.bGradient Then
    If bTimerActive = True And ((myProps.bStatus And 6) = 6) = False And _
        (myProps.bGradientColor <> myProps.bBackHover) Then
        DrawRect ButtonDC.hdc, 0, 0, ScaleWidth, ScaleHeight, bHoverColor
    Else
        If bGradientColor < 0 Then bGradientColor = bColor
        DrawGradient bColor, bGradientColor
    End If
Else
    If (UserControl.Enabled = True) Then
        For i = 0 To ScaleHeight
            DrawRect ButtonDC.hdc, 0, i, ScaleWidth, i + 1, ShadeColor(bColor2Use, -(25 / ScaleHeight) * i, True)
        Next
    Else
        DrawRect ButtonDC.hdc, 0, 0, ScaleWidth, ScaleHeight, bColor2Use
    End If
End If

End Sub

Private Sub DrawCustomBorders(ActiveRegion As Integer)
' This routine gets more complicated as each new type of button is added
' With custom shapes & round shapes, we have to use clipping regions since
' the window shape is not rectangular. Without clipping regions, we end up
' drawing over the button borders. This entire routine is just to determine
' which clipping region will be used & drawing the border (if any)

' backstyle of 5 = Hover buttons. This type button also complicates things
' as it doesn't have a border until a mouse is over it & then loses its
' border when the mouse leaves

ActiveRegion = 1
If bTimerActive = True Or ((myProps.bStatus And 6) = 6) Then
    Dim tRegion As Long, tBrush As Long, i As Integer
        ' this little trick gives us a good edge to our round button
        ' Too simple really--draw over the entire button a gradient background
        ' then set a clipping region excluding the border size and draw the
        ' rest of the button. Text/images can't overlap the border this way
        i = myProps.bGradient
        ' hover buttons have no border, so we need to create it on demand
        SelectClipRgn ButtonDC.hdc, 0
        If ((myProps.bStatus And 6) = 6) Then
            SelectClipRgn ButtonDC.hdc, ButtonDC.ClipBorder
            myProps.bGradient = lv_Top2Bottom
            DrawGradient vbGray, vbWhite
            SelectClipRgn ButtonDC.hdc, ButtonDC.ClipRgn
            ActiveRegion = 4
        Else
           SelectClipRgn ButtonDC.hdc, ButtonDC.ClipBorder
           ActiveRegion = 3
        End If
        myProps.bGradient = i
End If
End Sub

Private Sub DrawButtonBorder(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer, _
                    Optional OuterBorderStyle As Long = -1)

' This routine draws the border depending on the button style

Dim i As Integer, j As Integer, xColorRef As Integer
Dim lBorderStyle As Long, lastColor As Long
Dim polyOffset As POINTAPI

' need to run special calculations for diagonal buttons
If myProps.bShape > lv_Rectangular Then
        polyOffset.x = Abs(CInt(myProps.bShape <> lv_RightDiagonal))
        polyOffset.y = Abs(CInt(myProps.bShape <> lv_LeftDiagonal))
End If
' calculate X,Y points for all three levels of borders
polyPts(0).x = 2 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(0).y = ScaleHeight - 3
polyPts(1).x = 2 + polyOffset.x * 2: polyPts(1).y = 2
polyPts(2).x = myProps.bSegPts.y - 3 + polyOffset.y * 3: polyPts(2).y = 2
polyPts(3).x = ScaleWidth - 3 - polyOffset.y * 3: polyPts(3).y = ScaleHeight - 3
polyPts(4).x = 1 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(4).y = ScaleHeight - 3
For i = 5 To 9
    polyPts(i).x = polyPts(i - 5).x + Choose(i - 4, polyOffset.x - 1, -1 - polyOffset.x, 1 - polyOffset.y, 1 + polyOffset.y, -1, -1)
    polyPts(i).y = polyPts(i - 5).y + Choose(i - 4, 1, -1, -1, 1, 1, 1)
Next
polyPts(10).x = myProps.bSegPts.x - polyOffset.x: polyPts(10).y = ScaleHeight - 1 + polyOffset.x
polyPts(11).x = 0: polyPts(11).y = 0
polyPts(12).x = myProps.bSegPts.y - 1 + polyOffset.y: polyPts(12).y = 0
polyPts(13).x = ScaleWidth - 1: polyPts(13).y = ScaleHeight - 1 + polyOffset.y
polyPts(14).x = myProps.bSegPts.x - 1 - polyOffset.x * 2: polyPts(14).y = ScaleHeight - 1
lastColor = -1

For i = 0 To 13
    Select Case i
        Case Is < 4: xColorRef = i + 1
        Case Is > 8:
            xColorRef = i - 1   ' next line used for dashed borders
            If OuterBorderStyle > -1 Then lBorderStyle = OuterBorderStyle
        Case Else: xColorRef = i
    End Select
    If (i <> 4 And i <> 9) Then
        ' if -1 is the color, we skip that level
        If polyColors(xColorRef) > -1 Then
            ' change the pen color if needed
            If lastColor <> polyColors(xColorRef) Then SetButtonColors True, ButtonDC.hdc, cObj_Pen, polyColors(xColorRef), , , , lBorderStyle
            Polyline ButtonDC.hdc, polyPts(i), 2
            lastColor = polyColors(xColorRef)
        End If
    End If
Next
If polyOffset.y <> ScaleWidth Then
    ' tweak to ensure bottom, outer border draws correctly on diagonal buttons
    polyPts(15).x = ScaleWidth - 1: polyPts(15).y = ScaleHeight - 1
    Polyline ButtonDC.hdc, polyPts(14), 2
End If
End Sub

Private Sub DrawFocusRectangle(fColor As Long, bSolid As Boolean, _
    bOnText As Boolean, polyPts() As POINTAPI)

' Draws focus rectangles for the button style & button mode

Dim tRgn As Long, hBrush As Long
Dim focusOffset As Byte, bDownOffset As Byte

If ((myProps.bStatus And 1) <> 1) = True Then Exit Sub

If myProps.bShape > lv_FullDiagonal Then Exit Sub

If myProps.bShape > lv_Rectangular Then
    ' diagonal buttons
    Dim polyOffset As POINTAPI
    If myProps.bSegPts.x Then polyOffset.x = 1
    If myProps.bSegPts.y < ScaleWidth Then polyOffset.y = 1
    polyPts(0).x = 4 + myProps.bSegPts.x - polyOffset.x * 6: polyPts(0).y = ScaleHeight - 5
    polyPts(1).x = 4 + polyOffset.x * 4: polyPts(1).y = 4
    polyPts(2).x = myProps.bSegPts.y - 5 + polyOffset.y * 4: polyPts(2).y = 4
    polyPts(3).x = ScaleWidth - 5 - polyOffset.y * 6: polyPts(3).y = ScaleHeight - 5
    polyPts(4).x = 3 + myProps.bSegPts.x - polyOffset.x * 4: polyPts(4).y = ScaleHeight - 5
    SetButtonColors True, ButtonDC.hdc, cObj_Pen, fColor
    Polyline ButtonDC.hdc, polyPts(0), 5
Else
  Dim fRect As RECT
    If fColor < 0 Then fColor = 0
    SetButtonColors True, ButtonDC.hdc, cObj_Pen, fColor
    If bOnText = True Then
        If Len(myProps.bCaption) Then
            fRect = myProps.bRect
        Else
            fRect = myImage.iRect
            If fRect.Bottom > ScaleHeight - 4 Then fRect.Bottom = ScaleHeight - 4
            If fRect.Right > ScaleWidth - 4 Then fRect.Right = ScaleWidth - 4
        End If
        If myProps.bRect.Left > 4 + myProps.bSegPts.x And myProps.bRect.Right < myProps.bSegPts.y - 4 Then focusOffset = 2 Else focusOffset = 1
        bDownOffset = Abs((myProps.bStatus And 6) = 6)
        OffsetRect fRect, -focusOffset + bDownOffset * Abs(myProps.bShape), -focusOffset + bDownOffset
        fRect.Right = fRect.Right + focusOffset * 2 + bDownOffset * Abs(myProps.bShape)
        fRect.Bottom = fRect.Bottom + focusOffset * 2 + bDownOffset
        If bSolid Then   ' for now, only used on Java buttons & round buttons
            polyPts(0).x = fRect.Left: polyPts(0).y = fRect.Top
            polyPts(1).x = fRect.Right - 1: polyPts(1).y = fRect.Top
            polyPts(2).x = fRect.Right - 1: polyPts(2).y = fRect.Bottom - 1
            polyPts(3).x = fRect.Left: polyPts(3).y = fRect.Bottom - 1
            polyPts(4).x = fRect.Left: polyPts(4).y = fRect.Top
            Polyline ButtonDC.hdc, polyPts(0), 5
        Else            ' for now, only used on Macintosh buttons
            DrawFocusRect ButtonDC.hdc, fRect
        End If
    Else
        SetRect fRect, 0, 0, myProps.bSegPts.y - (myProps.bSegPts.x + 8), ScaleHeight - 8
        OffsetRect fRect, 4 + myProps.bSegPts.x, 4
        If bSolid Then ' used when option buttons/checkboxes have focus if Value=True
            polyPts(0).x = fRect.Left: polyPts(0).y = fRect.Bottom
            polyPts(1).x = fRect.Left: polyPts(1).y = fRect.Top
            polyPts(2).x = fRect.Right: polyPts(2).y = fRect.Top
            polyPts(3).x = fRect.Right: polyPts(3).y = fRect.Bottom
            polyPts(4).x = fRect.Left: polyPts(4).y = fRect.Bottom
            SetButtonColors True, ButtonDC.hdc, cObj_Pen, fColor
            Polyline ButtonDC.hdc, polyPts(0), 5
        Else
            DrawFocusRect ButtonDC.hdc, fRect
        End If
    End If
End If
End Sub

Private Sub DrawCaptionIcon(bColor As Long, Optional tColorDisabled As Long = -1, _
            Optional bOffsetTextDown As Boolean = False, _
            Optional bSingleDisableColor As Boolean = False)

' Routine draws the caption & calls the DrawButtonIcon routine

Dim tRect As RECT, iRect As RECT
Dim lColor As Long

' set these rectangles & they may be adjusted a little later
tRect = myProps.bRect
iRect = myImage.iRect
' if the button is in a down position, we'll offset the image/text rects by 1
If (((myProps.bStatus And 6) = 6) Or bOffsetTextDown) Then
    OffsetRect tRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
    OffsetRect iRect, 1 + Int(myProps.bShape > lv_FullDiagonal), 1
End If

DrawButtonIcon iRect
If Len(myProps.bCaption) = 0 Then Exit Sub

Dim sCaption As String  ' note Replace$ not compatible with VB5
sCaption = Replace$(myProps.bCaption, "||", vbNewLine)
' Setting text colors and offsets
If UserControl.Enabled = False Then
    If tColorDisabled > -1 And myProps.bGradient = lv_NoGradient Then
        lColor = tColorDisabled
    Else
        lColor = vbWhite
        OffsetRect tRect, 1, 1
        bSingleDisableColor = False
    End If
Else
    ' get the right forecolor to use
    If bTimerActive = True And ((myProps.bStatus And 6) = 6) = False Then
        lColor = ConvertColor(myProps.bForeHover)
    Else
        If myProps.bGradient Then
            lColor = ConvertColor(UserControl.ForeColor)
        Else
                lColor = ConvertColor(UserControl.ForeColor)
        End If
    End If
End If
SetButtonColors True, ButtonDC.hdc, cObj_Text, lColor
DrawText ButtonDC.hdc, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
If UserControl.Enabled = False And bSingleDisableColor = False Then
    ' finish drawing the disabled caption
    SetButtonColors True, ButtonDC.hdc, cObj_Text, vbGray
    OffsetRect tRect, -1, -1
    DrawText ButtonDC.hdc, sCaption, Len(sCaption), tRect, DT_WORDBREAK Or Choose(myProps.bCaptionAlign + 1, DT_LEFT, DT_RIGHT, DT_CENTER)
End If

End Sub

Private Sub DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long)
Dim mRect As RECT
Dim i As Long, rctOffset As Integer
Dim PixelStep As Long, rIndex As Long
Dim Colors() As Long

' The gist is to draw 1 pixel rectangles of various colors to create
' the gradient effect. If the size of the rectangle is greater than a
' quarter of the screen size, we'll step it up to 2 pixel rectangles
' to speed things up a bit


On Error Resume Next
mRect.Right = ScaleWidth
mRect.Bottom = ScaleHeight
rctOffset = 1
If myProps.bGradient < 3 Then
        If (Screen.Width \ Screen.TwipsPerPixelX) \ ScaleWidth < 4 Then
            PixelStep = ScaleWidth \ 2
            rctOffset = 2
        Else
            PixelStep = ScaleWidth
        End If
Else
    If (Screen.Height \ Screen.TwipsPerPixelY) \ ScaleHeight < 4 Then
        PixelStep = ScaleHeight \ 2
        rctOffset = 2
    Else
        PixelStep = ScaleHeight
    End If
End If
ReDim Colors(0 To PixelStep - 1) As Long
LoadGradientColors Colors(), Color1, Color2
If myProps.bGradient > 2 Then mRect.Bottom = rctOffset Else mRect.Right = rctOffset
For i = 0 To PixelStep - 1
    If myProps.bGradient Mod 2 Then rIndex = i Else rIndex = PixelStep - i - 1
    DrawRect ButtonDC.hdc, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, Colors(rIndex)
    If myProps.bGradient > 2 Then
        OffsetRect mRect, 0, rctOffset
    Else
        OffsetRect mRect, rctOffset, 0
    End If
Next
End Sub

Private Sub LoadGradientColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
Dim i As Integer, j As Integer
Dim sBase(0 To 2) As Single
Dim xBase(0 To 2) As Long
Dim lRatio(0 To 2) As Single

' routine adds/removes colors between a range of two colors
' Used by the DrawGradient routine. A variation of the ShadeColor routine


sBase(0) = (Color1 And &HFF)
sBase(1) = (Color1 And &HFF00&) / 255&
sBase(2) = (Color1 And &HFF0000) / &HFF00&
xBase(0) = (Color2 And &HFF)
xBase(1) = (Color2 And &HFF00&) / 255&
xBase(2) = (Color2 And &HFF0000) / &HFF00&

For j = 0 To 2
    lRatio(j) = (xBase(j) - sBase(j)) / UBound(Colors)
Next
Colors(0) = Color1
For j = 1 To UBound(Colors)
    For i = 0 To 2
        sBase(i) = sBase(i) + lRatio(i)
        If sBase(i) > 255 Then sBase(i) = 255
        If sBase(i) < 0 Then sBase(i) = 0
    Next
    Colors(j) = Int(sBase(0)) + 256& * Int(sBase(1)) + 65536 * Int(sBase(2))
Next

Erase sBase
Erase xBase
Erase lRatio
End Sub


' //////////////////// USER CONTROL EVENTS  \\\\\\\\\\\\\\\\\\\\\\\\

Private Sub UserControl_AmbientChanged(PropertyName As String)

' something on the parent container changed

On Error GoTo AbortCheck
    Select Case PropertyName
    Case "DisplayAsDefault" 'changing focus
        If Ambient.DisplayAsDefault = True And (myProps.bShowFocus = True Or Not Ambient.UserMode) Then
            myProps.bStatus = myProps.bStatus Or 1
        Else
            myProps.bStatus = myProps.bStatus And Not 1
        End If
       If Not Ambient.UserMode Then bTimerActive = ((myProps.bStatus And 1) = 1)
        RedrawButton
        If Not Ambient.UserMode Then bTimerActive = False
    Case "BackColor"
        cParentBC = ConvertColor(Ambient.BackColor)
    End Select
AbortCheck:
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

' Only allow left clicks to fire a click event

' key variable... this tells our mouse routines & the click event
' whether or not the left button is doing the clicking
mButton = Button
If Button = vbLeftButton Then
    bKeyDown = True
    myProps.bStatus = myProps.bStatus Or 2      ' simulate a "down" state
    bNoRefresh = False
    RedrawButton
    ' we need this in case the user clicks & drags mouse off of the control
    ' Without it, we may never get the mouse up event
    SetCapture UserControl.hwnd
End If
'RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

' Here we may fire 2 events: MouseMove & MouseOnButton

On Error GoTo RaiseTheMouseEvent
' if we are already over the button, simply fire the MouseMove event
If bTimerActive Then GoTo RaiseTheMouseEvent
' if we are outside of the mouse we fire the MouseMove event only.
' Note. We don't use SetCapture/ReleaseCapture, except in one special
' case, because it affects the actual control (not my button control)
' that should rightfully have the focus. However, should the mouse
' be down on this button & the use drags mouse off of the button,
' we will continue to get mouse move events & will fire them accordingly;
' and this is the only appropriate exception for using SetCapture

Dim mousePt As POINTAPI
GetCursorPos mousePt
If WindowFromPoint(mousePt.x, mousePt.y) <> UserControl.hwnd Then GoTo RaiseTheMouseEvent

'If x < 0 Or y < 0 Or x > ScaleWidth Or y > ScaleHeight Then GoTo RaiseTheMouseEvent

' An improvement over most other button routines out there....
' A soft timer. No timer control needed. The trick is to get the timer
' to fire back to this instance of the button control. We do that by
' setting a reference to this instance in our Window properties. When
' the timer routine (see modLvTimer) gets an event, the hWnd is passed
' along & with that, the timer routine can retrieve the property we set.
' All this allows the timer routine to positively identify this instance.

myProps.bStatus = myProps.bStatus Or 4      ' set a mouse hover state
'RaiseEvent MouseOnButton(True)              ' fire this event
' The MouseOnButton event allows users to change the properties of the
' button while the mouse is over it. For instance, you can supply a different
' image/font/etc & replace it when the mouse leaves the button area

SetProp UserControl.hwnd, "lv_ClassID", ObjPtr(Me)
' the next line is used with expandability in mind. May use multiple timers in future upgrade
SetProp UserControl.hwnd, "lv_TimerID", 237
SetTimer UserControl.hwnd, 237, 50, AddressOf lv_TimerCallBack
bTimerActive = True                         ' flag used for drawing
bNoRefresh = False                         ' ensure flag is reset
If Button = vbLeftButton Then bKeyDown = True
RedrawButton

RaiseTheMouseEvent:
'RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

' The only tweak here is to trigger a fake click event if user
' double clicked on this button

bKeyDown = False
If Button = vbLeftButton Then
    ReleaseCapture
    myProps.bStatus = myProps.bStatus And Not 2     ' "normal" state
    bNoRefresh = False                              ' ensure flag is reset
    RedrawButton
End If
'RaiseEvent MouseUp(Button, Shift, X, Y)            ' fire event
' key flag. Update
mButton = Button
End Sub

Private Sub UserControl_Click()


' Again, only allow left mouse button to fire click events. Keyboard
' actions may set the mButton variable to ensure event is fired

If mButton = vbLeftButton Then
    RaiseEvent Click
End If
End Sub


Private Sub UserControl_Paint()

' this routine typically called by Windows when another window covering
' this button is removed, or when the parent is moved/minimized/etc.

bNoRefresh = False
RedrawButton
End Sub

Private Sub UserControl_InitProperties()

' Initial properties for a new button

Tag = UserControl.ContainerHwnd
With myProps
    .bCaption = Ambient.DisplayName
    .bCaptionAlign = vbCenter
    .bShowFocus = True
    .bForeHover = vbButtonText
    .bBackHover = vbButtonFace
End With

If Not (TypeOf Parent Is MDIForm) Then Set UserControl.Font = Parent.Font
cParentBC = ConvertColor(Ambient.BackColor)
curBackColor = vbButtonFace         ' this will be the button's initial backcolor
GetGDIMetrics "BackColor"
PropertyChanged "Caption"
PropertyChanged "CapAlign"
PropertyChanged "Focus"
PropertyChanged "cFHover"
PropertyChanged "cBHover"
PropertyChanged "Font"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

' Read properties

Tag = UserControl.ContainerHwnd
cParentBC = ConvertColor(Ambient.BackColor)
DelayDrawing True
With PropBag
    myProps.bCaption = .ReadProperty("Caption", "")
    myProps.bCaptionAlign = .ReadProperty("CapAlign", 2)
    myProps.bShape = .ReadProperty("Shape", 0)
    myProps.bGradient = .ReadProperty("Gradient", 0)
    myProps.bGradientColor = .ReadProperty("cGradient", vbButtonFace)
    UserControl.ForeColor = .ReadProperty("cFore", vbButtonText)
    Set UserControl.Font = .ReadProperty("Font")
'    If UserControl.FontName = "Webdings" Then Stop
    myProps.bShowFocus = .ReadProperty("Focus", True)
    Set myImage.Image = .ReadProperty("Image", Nothing)
    myImage.Size = .ReadProperty("ImgSize", 16)
    myImage.Align = .ReadProperty("ImgAlign", 0)
    myProps.bForeHover = .ReadProperty("cFHover", vbButtonText)
    UserControl.Enabled = .ReadProperty("Enabled", True)
    curBackColor = .ReadProperty("cBack", Parent.BackColor)
    myProps.bBackHover = .ReadProperty("cBHover", curBackColor)
    myProps.bLockHover = .ReadProperty("LockHover", 0)
    Set Me.MouseIcon = .ReadProperty("mIcon", Nothing)
    Me.MousePointer = .ReadProperty("mPointer", 0)
End With
On Error Resume Next
GetGDIMetrics "Picture"
GetGDIMetrics "BackColor"
Me.Caption = myProps.bCaption      ' sets the hot key if needed
bNoRefresh = False
Call UserControl_Resize
End Sub

Private Sub UserControl_Show()
' interesting, NT won't send the DisplayAsDefault (while in IDE) until after the button is shown
' Win98 fires this regardless. So fix is to put the test here also.
If Ambient.UserMode = False Then
    If Ambient.DisplayAsDefault = True And myProps.bShowFocus = True Then
        myProps.bStatus = myProps.bStatus Or 1
        RedrawButton
    End If
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

' Store Properties

With PropBag
    .WriteProperty "Caption", myProps.bCaption, ""
    .WriteProperty "CapAlign", myProps.bCaptionAlign, 0
'    .WriteProperty "BackStyle", myProps.bBackStyle, 0
    .WriteProperty "Shape", myProps.bShape, 0
    .WriteProperty "Font", UserControl.Font, Nothing
    .WriteProperty "cFore", UserControl.ForeColor, vbButtonText
    .WriteProperty "cFHover", myProps.bForeHover, vbButtonText
    .WriteProperty "cBhover", myProps.bBackHover, curBackColor
    .WriteProperty "Focus", myProps.bShowFocus, True
    .WriteProperty "LockHover", myProps.bLockHover, 0
    .WriteProperty "cGradient", myProps.bGradientColor, vbButtonFace
    .WriteProperty "Gradient", myProps.bGradient, 0
    .WriteProperty "ImgAlign", myImage.Align, 0
    .WriteProperty "Image", myImage.Image, Nothing
    .WriteProperty "ImgSize", myImage.Size, 16
    .WriteProperty "Enabled", UserControl.Enabled, True
    .WriteProperty "cBack", curBackColor
    .WriteProperty "mPointer", UserControl.MousePointer, 0
    .WriteProperty "mIcon", UserControl.MouseIcon, Nothing
End With
End Sub

Private Sub UserControl_Resize()

' since we are using a separate DC for drawing, we need to resize the
' bitmap in that DC each time the control resizes

If ButtonDC.hdc Then
    DeleteObject SelectObject(ButtonDC.hdc, ButtonDC.OldBitmap)
    ButtonDC.OldBitmap = 0  ' this will force a new bitmap for existing DC
End If
GetSetOffDC True
If Not bNoRefresh Then
    CreateButtonRegion
    CalculateBoundingRects False
    RedrawButton
End If
End Sub

Private Sub UserControl_Terminate()
' Button is ending, let's clean up

' should never happen that we have a timer left over; but just in case
If bTimerActive Then KillTimer UserControl.hwnd, 1
' circular/custom buttons have clipping regions, kill them too
If ButtonDC.ClipRgn Then DeleteObject ButtonDC.ClipRgn
If ButtonDC.ClipBorder Then DeleteObject ButtonDC.ClipBorder
If ButtonDC.hdc Then
    ' get rid of left over pen & brush
    SetButtonColors False, ButtonDC.hdc, cObj_Pen, 0
    ' get rid of logical font
    DeleteObject SelectObject(ButtonDC.hdc, ButtonDC.OldFont)
    ' destroy the separate Bitmap & select original back into DC
    DeleteObject SelectObject(ButtonDC.hdc, ButtonDC.OldBitmap)
    ' destroy the temporary DC
    DeleteDC ButtonDC.hdc
End If
' kill image used for transparencies when selected button pic is a bitmap
If myImage.TransImage Then DeleteObject myImage.TransImage
End Sub

Private Sub DrawButton_WinXP(polyPts() As POINTAPI, polyColors() As Long, ActiveStatus As Integer, lastClipRgn As Integer)
'==========================================================================
' If not used in your project, replace this entire routine from the
' Dim statements to the last line before the End Sub with
' a simple Exit Sub
'==========================================================================

Dim backShade As Long, darkShade As Long, liteShade As Long, midShade As Long
Dim i As Integer, lColor As Long
Dim cDisabled As Long, lGradientColor As Long


    lColor = adjBackColorUp
    If ((myProps.bStatus And 6) = 6) And myProps.bGradient = lv_NoGradient Then
        backShade = adjBackColorDn
    Else
        If UserControl.Enabled = False Then
            backShade = ShadeColor(lColor, -&H18, True)
        Else
            backShade = lColor
        End If
    End If

If Not UserControl.Enabled Then cDisabled = ShadeColor(backShade, -&H68, True)
    If myProps.bGradient Then
        lGradientColor = ShadeColor(ConvertColor(myProps.bGradientColor), &H30, True)
    Else
        lGradientColor = backShade
    End If
DrawButtonBackground backShade, ActiveStatus, lastClipRgn, lGradientColor, adjHoverColor
If lastClipRgn = 0 Then Exit Sub
    
DrawCaptionIcon backShade, cDisabled, True, True
    
    If UserControl.Enabled Then
        If ((myProps.bStatus And 4) = 4) And ((myProps.bStatus And 2) <> 2) Then
             ActiveStatus = 3
        Else
            If ActiveStatus = 1 And myProps.bShowFocus = False Then ActiveStatus = 0
        End If
        liteShade = lColor
        ' inner Rectangle
        polyColors(1) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&HA, True), &HF0D1B5, ShadeColor(liteShade, -&H16, True), &H6BCBFF)
        polyColors(2) = Choose(ActiveStatus + 1, ShadeColor(lColor, &HA, True), &HF7D7BD, ShadeColor(liteShade, -&H18, True), &H8CDBFF)
        polyColors(3) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H18, True), &HF0D1B5, lColor, &H6BCBFF)
        polyColors(4) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H20, True), &HE7AE8C, ShadeColor(liteShade, &HA, True), &H31B2FF)
        ' middle Rectangle
        polyColors(5) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H5, True), &HE7AE8C, ShadeColor(liteShade, -&H20, True), &H31B2FF)
        polyColors(6) = Choose(ActiveStatus + 1, ShadeColor(lColor, &H10, True), &HFFDFBF, ShadeColor(liteShade, -&H20, True), &HA6E9FF)
        polyColors(7) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H24, True), &HE7AE8C, ShadeColor(liteShade, &H5, True), &H31B2FF)
        polyColors(8) = Choose(ActiveStatus + 1, ShadeColor(lColor, -&H30, True), &HEF826B, ShadeColor(liteShade, &H10, True), &H96E7&)
        lColor = &H733C00
    Else
        For i = 1 To 8
            polyColors(i) = -1
        Next
        lColor = ShadeColor(lColor, -&H54, True)
    End If
    For i = 9 To 12
        polyColors(i) = lColor
    Next
    
    DrawButtonBorder polyPts(), polyColors(), ActiveStatus
    
    If myProps.bSegPts.x = 0 Then
        SetPixel ButtonDC.hdc, 1, ScaleHeight - 2, lColor
        SetPixel ButtonDC.hdc, 1, 1, lColor
    End If
    If myProps.bSegPts.y = ScaleWidth Then
        SetPixel ButtonDC.hdc, ScaleWidth - 2, ScaleHeight - 2, lColor
        SetPixel ButtonDC.hdc, ScaleWidth - 2, 1, lColor
    End If

End Sub
