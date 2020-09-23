Attribute VB_Name = "modGradient"
'+
'   World Times - Graphical Representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          modGradient
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul Turcksin
'
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
' Description:
'
' Fills a Device Context with a gradient using API's
'
' Public Sub subShowGradient(destDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, sOrientation As Boolean, sInitialColor As Long, sFinalColor As Long)
'   destDC : destination DC
'   x, y, nWidth,nHeight: starting position and size
'   sOrientation : Horizontal (True) or Vertical (False)
'   sInitialColor , sFinalColor: Start and ending color of gradient
'__________________________________________________________________________________
'-
Option Explicit

Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Const GRADIENT_FILL_RECT_H = 0
Private Const GRADIENT_FILL_RECT_V = 1

Dim arVert(1) As TRIVERTEX
Dim gRect As GRADIENT_RECT

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Sub subShowGradient(destDC As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, sOrientation As Boolean, sInitialColor As Long, sFinalColor As Long)
'   destDC : destination DC
'   x, y, nWidth,nHeight
'   sOrientation : Horizontal (True) or Vertical (False)
'   sInitialColor , sFinalColor

   Dim arByteClr(3) As Byte   ' used to convert (long) color to its components
   Dim arByteVert(7) As Byte   ' used to init color part of vertices array
   Dim iOrientation As Long
   
    On Local Error Resume Next
    
' init vertices : position, size and direction
      arVert(0).x = x: arVert(1).x = x + nWidth
      arVert(0).y = y: arVert(1).y = y + nHeight
   
' init vertices :colors, initial
   CopyMemory arByteClr(0), sInitialColor, 4
   arByteVert(1) = arByteClr(0)   ' red
   arByteVert(3) = arByteClr(1)   ' green
   arByteVert(5) = arByteClr(2)   ' blue
   CopyMemory arVert(0).Red, arByteVert(0), 8

' init vertices :colors, final
   CopyMemory arByteClr(0), sFinalColor, 4
   arByteVert(1) = arByteClr(0)   ' red
   arByteVert(3) = arByteClr(1)   ' green
   arByteVert(5) = arByteClr(2)   ' blue
   CopyMemory arVert(1).Red, arByteVert(0), 8

' init gradient rect
   gRect.UpperLeft = 0
   gRect.LowerRight = 1
    
   iOrientation = IIf(sOrientation, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
   
   GradientFill destDC, arVert(0), 2, gRect, 1, iOrientation
    On Error GoTo 0
End Sub
