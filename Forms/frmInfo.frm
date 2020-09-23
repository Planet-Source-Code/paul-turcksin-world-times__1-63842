VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInfo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Time - Info"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8625
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF 
      Height          =   5895
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10398
      _Version        =   393217
      BackColor       =   12648447
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\My Code\Submit\WorldTimes - v1.5\Documents\Worldtimes.rtf"
      TextRTF         =   $"frmInfo.frx":33E2
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+
'   World Times - Graphical representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          frmInfo
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
'-

Option Explicit
' registry
Private mProgramSettings As New cProgramSettings

Private Sub Form_Load()
' Get previous position from registry
  With mProgramSettings
      .Program = "TimeZones"
      Me.Move CSng(.ReadEntry("InfoLeft", "50")), CSng(.ReadEntry("InfoTop", "50"))
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Save current position of the form on the screen
   mProgramSettings.WriteEntry "InfoLeft", Format(Me.Left)
   mProgramSettings.WriteEntry "InfoTop", Format(Me.Top)
' now we can leave
   Set mProgramSettings = Nothing
   Set frmInfo = Nothing
End Sub
