VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Times"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   4680
   End
   Begin MSComDlg.CommonDialog cdColor 
      Left            =   1920
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Cose color"
   End
   Begin WorldTimes.DiagButton cmdAction 
      Height          =   345
      Index           =   0
      Left            =   6360
      TabIndex        =   17
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "Remove"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin WorldTimes.DiagButton cmdInfo 
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Info"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.TextBox fMyName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      Top             =   3480
      Width           =   1200
   End
   Begin VB.ListBox lbDisplayNames 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Timer tmrTime 
      Interval        =   30000
      Left            =   480
      Top             =   4680
   End
   Begin VB.PictureBox picTimebar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   360
      MousePointer    =   2  'Cross
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   300
      Width           =   10800
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "22:35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   7560
         TabIndex        =   26
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "22:35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   6960
         TabIndex        =   25
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "22:35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   6240
         TabIndex        =   24
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "22:35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   5640
         TabIndex        =   23
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "22:35"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   5040
         TabIndex        =   13
         Top             =   120
         Width           =   420
      End
      Begin VB.Line lnTime 
         BorderColor     =   &H0000FF00&
         Index           =   4
         Visible         =   0   'False
         X1              =   300
         X2              =   300
         Y1              =   0
         Y2              =   50
      End
      Begin VB.Line lnTime 
         BorderColor     =   &H00FFFF00&
         Index           =   3
         Visible         =   0   'False
         X1              =   250
         X2              =   250
         Y1              =   0
         Y2              =   50
      End
      Begin VB.Line lnTime 
         BorderColor     =   &H00FF00FF&
         Index           =   2
         Visible         =   0   'False
         X1              =   200
         X2              =   200
         Y1              =   0
         Y2              =   50
      End
      Begin VB.Line lnTime 
         BorderColor     =   &H00FFFF00&
         Index           =   1
         Visible         =   0   'False
         X1              =   150
         X2              =   150
         Y1              =   0
         Y2              =   50
      End
      Begin VB.Line lnTime 
         BorderColor     =   &H000000FF&
         Index           =   0
         X1              =   100
         X2              =   100
         Y1              =   0
         Y2              =   50
      End
   End
   Begin WorldTimes.DiagButton cmdSettings 
      Height          =   375
      Left            =   8640
      TabIndex        =   15
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Settings"
      CapAlign        =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin WorldTimes.DiagButton cmdExit 
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Exit"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   18
      Top             =   0
      Width           =   0
   End
   Begin WorldTimes.DiagButton cmdAction 
      Height          =   345
      Index           =   1
      Left            =   7440
      TabIndex        =   19
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "Update"
      CapAlign        =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin WorldTimes.DiagButton cmdAction 
      Height          =   345
      Index           =   2
      Left            =   8520
      TabIndex        =   20
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      Caption         =   "Add"
      CapAlign        =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin WorldTimes.DiagButton cmdAction 
      Height          =   345
      Index           =   3
      Left            =   9600
      TabIndex        =   21
      Top             =   4080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Caption         =   "Close settings"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.CheckBox chkTimeType 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "24hr Clock"
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   9900
      TabIndex        =   22
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click on colored name above to change color setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1995
      Width           =   4935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   345
      Left            =   9810
      Shape           =   4  'Rounded Rectangle
      Top             =   3450
      Width           =   1245
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   2040
      Left            =   315
      Top             =   2370
      Width           =   5130
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   345
      Left            =   5835
      Shape           =   4  'Rounded Rectangle
      Top             =   2715
      Width           =   4170
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   345
      Left            =   5835
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   5130
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   345
      Left            =   5850
      Shape           =   4  'Rounded Rectangle
      Top             =   3450
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   405
      Index           =   1
      Left            =   6330
      Shape           =   4  'Rounded Rectangle
      Top             =   4050
      Width           =   4740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00C00000&
      Height          =   435
      Index           =   0
      Left            =   7530
      Shape           =   4  'Rounded Rectangle
      Top             =   1290
      Width           =   3555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on time zone description below to select  new setting"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Left click on colored name above to update or remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a name for the selection"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblTimeZone 
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label lblDisplayName 
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label lblSetting 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   6
      ToolTipText     =   "Click to update"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSetting 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Click to update"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSetting 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Click to update"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSetting 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Click to update"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblSetting 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+
'   World Times - Graphical Representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          frmMain
'
'   Compatibility:
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
' Important notice: as the timebar is 720 pixels wide and a day is 1440 minutes long
' the time shown in the label will always be even. In other words: this is not
' a precision clock.
'-
Option Explicit

Private arTimes(4) As Long           ' time in minutes for each time zone
                                     ' (0) : local
                                     ' 1 to 4) settings ; -1 if no setting
Private arTimeDiff(4) As Long        ' only settings (entry 0 not used)
                                     ' time difference with local
Private arTemp(4) As Long            ' used in picTimeBar Mouse_Down and _Up events
Private arLblWidth(4) As Integer     ' width of each label (settings, lblSetting)
Private FlagMoveTimezones As Boolean ' controls moving time bars
Private oldPosX As Single            ' X value at start of move
Private IndexSetting As Integer      ' index lblSetting Click event

' enumerate cmdAction index
Private Enum cmdActionConstants
   Remove = 0
   Update = 1
   Add = 2
   CloseSetting = 3
End Enum

Private mProgramSettings As New cProgramSettings ' registry

Private Sub chkTimeType_Click()
' 24/12
    With Me
        .Cls
        Call subUpdateTimeValues
    End With
End Sub

Private Sub cmdAction_Click(Index As Integer)
' Remove, Update,Add command buttons
   Dim i As Integer
   Dim lTime As Long
   
   Select Case Index
   
      Case Remove
         If MsgBox("Remove " & fMyName.Text & " (" & lblDisplayName.Caption & _
                    " setting?" & ")", vbExclamation Or vbYesNo, "Confirmation") = vbYes Then
            ' Remove the name and set lines and labels
            fMyName.Text = ""
            lnTime(IndexSetting).Visible = False
            lblSetting(IndexSetting).Caption = ""
            lblSetting(IndexSetting).Visible = False
            ' update our array
            arTimes(IndexSetting) = -1
            ' set the values to the registry
            mProgramSettings.WriteEntry Format(IndexSetting), ""
         End If
      
      Case Update
          ' validate name given to setting
         If fMyName = "" Then
            fMyName.BackColor = vbRed
            Beep
         Else
            lTime = fncGetTimeZoneTime(Trim(lblTimeZone))
            If lTime <> -1 Then
               ' valid time zone
               lblSetting(IndexSetting).Caption = fMyName.Text
               arTimes(IndexSetting) = lTime
               arTimeDiff(IndexSetting) = arTimes(0) - arTimes(IndexSetting)
               mProgramSettings.WriteEntry Format(IndexSetting), fMyName.Text & "~" & Trim(lblTimeZone.Caption)
               subShowTime IndexSetting
               lnTime(IndexSetting).Visible = True
               cmdAction(Add).Enabled = False
            End If
         End If
     
      Case Add
         ' validate name given to setting
         If fMyName = "" Then
            fMyName.BackColor = vbRed
            Beep
         Else
            'scan for empty 'given name' label
            lTime = -1
            For i = 1 To 4
               If lblSetting(i).Caption = "" Then
                  lTime = fncGetTimeZoneTime(Trim(lblTimeZone))
                  Exit For
               End If
            Next i
            If lTime <> -1 Then
               ' valid time zone
               lblSetting(i).Caption = fMyName.Text
               lblSetting(i).Visible = True
               lnTime(i).BorderColor = lblSetting(i).ForeColor
               arTimes(i) = lTime
               arTimeDiff(i) = arTimes(0) - arTimes(i)
               mProgramSettings.WriteEntry Format(i), fMyName.Text & "~" & Trim(lblTimeZone.Caption)
               arLblWidth(i) = Me.TextWidth(lblSetting(i).Caption)
               subShowTime i
               lnTime(i).Visible = True
               cmdAction_Click CloseSetting
            End If
         End If
         
      Case CloseSetting
         Me.Height = 2220
         IndexSetting = 0
      
      
   End Select
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdInfo_Click()
   frmInfo.Show
End Sub

Private Sub cmdSettings_Click()
' Settings command button
   lblDisplayName = ""
   fMyName = ""
   cmdAction(Remove).Enabled = False
   cmdAction(Update).Enabled = False
   cmdAction(Add).Enabled = False
   Me.Height = 4965
End Sub

Private Sub fMyName_GotFocus()
' highlight textbox
   With fMyName
      .BackColor = vbWhite
      .SelStart = 0
      .SelLength = 99
   End With
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim l As Long
   Dim x As Single
   Dim Y1 As Single
   Dim Y2 As Single
   Dim ws As String
   
' position the form on the screen at the same place it was at the end of last
' time's run
   With mProgramSettings
      .Program = "TimeZones"
      Me.Move CSng(.ReadEntry("Left", "500")), CSng(.ReadEntry("Top", "500"))
' get 24/12 hr setting
   chkTimeType.Value = CInt(.ReadEntry("TimeType", "1"))
   End With
   
   Me.Height = 2220
   lblTime(0).Caption = ""
   ' retrieve Time Zone names from registry and add then to a listbox
   subGetTimeZones lbDisplayNames

' form' caption
  Me.Caption = "World Times    v" & App.Major & "." & App.Minor & "." & App.Revision
  
' colorise time bar
   subShowGradient picTimebar.hdc, 0, 0, 150, 50, True, &H400000, &H800000      ' 0 - 5
   subShowGradient picTimebar.hdc, 150, 0, 60, 50, True, &H800000, &HC0C0C0      ' 5 - 7
   subShowGradient picTimebar.hdc, 210, 0, 150, 50, True, &HC0C0C0, &H80FFFF              ' 7 - 12
   subShowGradient picTimebar.hdc, 360, 0, 210, 50, True, &H80FFFF, &HC0C0C0        ' 12 -19
   subShowGradient picTimebar.hdc, 540, 0, 30, 50, True, &HC0C0C0, &H800000        ' 19 - 21
   subShowGradient picTimebar.hdc, 570, 0, 150, 50, True, &H800000, &H400000    ' 21 - 224
' draw some stars in the time bar
   picTimebar.PSet (20, 14), vbWhite
   picTimebar.PSet (25, 20), vbWhite
   picTimebar.PSet (30, 10), vbWhite
   picTimebar.PSet (34, 17), vbWhite
   picTimebar.PSet (39, 20), vbWhite
   picTimebar.PSet (53, 24), vbWhite
   picTimebar.PSet (650, 35), vbWhite
   picTimebar.PSet (669, 24), vbWhite
   picTimebar.PSet (694, 40), vbWhite
   picTimebar.PSet (677, 36), vbWhite
   picTimebar.PSet (683, 40), vbWhite
   picTimebar.PSet (657, 38), vbWhite
   
   subUpdateTimeValues
   
' reset font
   Me.FontName = "MS Sans Serif"
   Me.FontSize = 8
   
' init local time
   arLblWidth(0) = Me.TextWidth("Local")
   arTimes(0) = (Hour(Time) * 60 + Minute(Time))
   l = CLng(mProgramSettings.ReadEntry("C0", "-1"))
   If l <> -1 Then
      lblSetting(0).ForeColor = l
      lnTime(0).BorderColor = l
   End If
   subShowTime 0
   
' get settings from registry
   With mProgramSettings
      .Program = "TimeZones"
      For i = 1 To 4
         ' time zones
         ws = .ReadEntry(Format(i), "")
         If ws <> "" Then
            l = InStr(1, ws, "~")
            lblSetting(i) = Left$(ws, l - 1)
            lblSetting(i).Visible = True
            arLblWidth(i) = Me.TextWidth(lblSetting(i).Caption)
            arTimes(i) = fncGetTimeZoneTime(Mid$(ws, l + 1))
            arTimeDiff(i) = arTimes(0) - arTimes(i)
            subShowTime i
            lnTime(i).Visible = True
            lblTime(i) = ""
         Else
            lblSetting(i).Caption = ""
            lblTime(i) = ""
            arTimes(i) = -1
         End If
         ' color labels setting and timebars
         l = CLng(.ReadEntry("C" & Format(i), "-1"))
         If l <> -1 Then
            lblSetting(i).ForeColor = l
            lnTime(i).BorderColor = l
          End If
      Next i
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Save current position of the form on the screen
   mProgramSettings.WriteEntry "Left", Format(Me.Left)
   mProgramSettings.WriteEntry "Top", Format(Me.Top)
' save timetype setting
   mProgramSettings.WriteEntry "TimeType", Format(chkTimeType.Value)
' now we can leave
   Set mProgramSettings = Nothing
   Unload frmInfo
   Set frmMain = Nothing
End Sub

Private Sub lbDisplayNames_Click()
' this initiates an add or is an update if cmdAction(1) 'Update" is enabled
   If cmdAction(Update).Enabled = False Then
     cmdAction(2).Enabled = True
  End If
' show selection
   lblDisplayName = lbDisplayNames.Text
   lblTimeZone = arZoneName(lbDisplayNames.ItemData(lbDisplayNames.ListIndex))
   fMyName.SetFocus
End Sub

Private Sub lblSetting_Mouseup(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' left click initiates remove/update process
' right click enables the user to change the color of the label and the timebar
   Dim ws As String
   Dim l As Long
   
   If Button = vbLeftButton Then
      ' Can't process if "local"
      If Index = 0 Then
         MsgBox "Please use standard Windows procedures to change your PC's time zone."
         Exit Sub
      End If
   
      ' get time zone from registry
      ws = mProgramSettings.ReadEntry(Format(Index), "")
      l = InStr(1, ws, "~")
      lblTimeZone.Caption = Mid$(ws, l + 1)
      ' get Display name
      lblDisplayName = fncGetDisplayName(lblTimeZone.Caption)
      ' setting name
      fMyName.Text = lblSetting(Index).Caption
      ' misc
      cmdAction(Remove).Enabled = True
      cmdAction(Update).Enabled = True
      Me.Height = 4965
      fMyName.SetFocus
      ' preserve index
      IndexSetting = Index
      
   Else ' right button
      With cdColor
         .Flags = cdlCCFullOpen Or cdlCCRGBInit
         .Color = lblSetting(Index).ForeColor
         .ShowColor
         lblSetting(Index).ForeColor = .Color
         lnTime(Index).BorderColor = .Color
         mProgramSettings.WriteEntry "C" & Format(Index), Format(.Color)
      End With
   End If
   
End Sub

Private Sub picTimebar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Mouse down in time bar shows time labels for time bars
' if mousedown position is over a timebar prepare for moving all time bars
'  -> disable timer, save current position, compute/show new position flash time bars

   Dim i As Integer
  
' show time labels
   For i = 0 To 4
      If arTimes(i) <> -1 Then
         subShowTimeLabel i
      End If
   Next i

' is mouse position over a time bar?
   FlagMoveTimezones = False
   For i = 0 To 4
      If lnTime(i).Visible _
      And lnTime(i).X1 = x Then
         FlagMoveTimezones = True
         Exit For
      End If
   Next i
      
' it is over a time bar
   If FlagMoveTimezones Then
      tmrTime.Enabled = False
      oldPosX = x
      For i = 0 To 4
         arTemp(i) = arTimes(i)
      Next i
      tmrFlash.Enabled = True
   End If
End Sub

Private Sub picTimebar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' if the mouse is down over a time bar all time bars move
' else show the time at position of mouse

   Dim PosX As Integer
   Dim OffsetX As Integer
   Dim i As Integer
   
   If FlagMoveTimezones Then
      OffsetX = (oldPosX - x) * 2
      For i = 0 To 4
         If arTemp(i) <> -1 Then
            arTimes(i) = arTemp(i) - OffsetX
            subShowTime i
            subShowTimeLabel i
         End If
      Next i
   End If
      
End Sub

Private Sub picTimebar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' restore "regular" display of time bars

   Dim i As Integer

' remove time labels
   For i = 0 To 4
      lblTime(i).Caption = ""
   Next i
   
' only after a move of time bars
   If FlagMoveTimezones Then
      For i = 0 To 4
         arTimes(i) = arTemp(i)
         If arTimes(i) <> -1 Then
            subShowTime i
         End If
      Next i
      FlagMoveTimezones = False
      tmrFlash.Enabled = False
      ' ensure time bars are visible
      If lnTime(0).Visible = False Then
         For i = 0 To 4
            If arTimes(i) <> -1 Then
               lnTime(i).Visible = True
            End If
         Next i
      End If
      tmrTime.Enabled = True
   End If
End Sub

Private Sub tmrFlash_Timer()
   Dim i As Integer
   
   For i = 0 To 4
      If arTimes(i) <> -1 Then
         lnTime(i).Visible = Not lnTime(i).Visible
      End If
   Next i
End Sub

Private Sub tmrTime_Timer()
   Dim CurrentTime As Long
   Dim i As Integer
   
   CurrentTime = (Hour(Time) * 60 + Minute(Time))
   If CurrentTime <> arTimes(0) Then
      arTimes(0) = CurrentTime
      subShowTime 0
      For i = 1 To 4
         If arTimes(i) <> -1 Then
            arTimes(i) = arTimes(0) - arTimeDiff(i)
            subShowTime i
         End If
      Next i
   End If
   
End Sub

'===================================================================================
'
'                                     LOCAL PROCEDURES
'__________________________________________________________________________________


Private Sub subShowTime(Indx As Integer)
   Dim lTime As Long
   
   lTime = arTimes(Indx)
' adjust time (previous/next day)
   If lTime < 0 Then
      lTime = lTime + 1440
      lnTime(Indx).BorderStyle = vbBSDot
   ElseIf lTime > 1440 Then
      lTime = lTime - 1440
      lnTime(Indx).BorderStyle = vbBSDashDot
   Else
      lnTime(Indx).BorderStyle = vbBSSolid
   End If
   
' position the time line and it's associated label
   With lnTime(Indx)
      .X1 = lTime \ 2
      .X2 = lTime \ 2
   End With
   
   lblSetting(Indx).Left = picTimebar.Left + lTime \ 2 - arLblWidth(Indx) \ 2
End Sub

Private Sub subShowTimeLabel(Indx As Integer)
' show time label
   Dim PosX As Integer
   Dim t As Integer
   Dim i As Integer
 
   PosX = lnTime(Indx).X1
' This allows for 24hr and 12/12 clocks to be displayed
   If chkTimeType.Value Then
      lblTime(Indx) = Format(PosX \ 30) & ":" & Format(PosX * 2 Mod 60, "00")
   Else
      t = PosX \ 30
      If PosX > 390 Then
         t = t - 12
      End If
      lblTime(Indx) = Format(t) & ":" & Format(PosX * 2 Mod 60, "00")
   End If
   
' position label next to time bar
   If PosX < 690 Then
      PosX = PosX + 5
   Else
      PosX = PosX - 30
   End If
   
' ensure visibility of label when it is in a dark or light part of the time bar
   lblTime(Indx).Left = PosX
   If PosX < 195 Or PosX > 540 Then
      lblTime(Indx).ForeColor = vbWhite
   Else
      lblTime(Indx).ForeColor = vbBlack
   End If
   

End Sub

Private Sub subUpdateTimeValues()
   Dim i As Integer
   Dim l As Long
   Dim x As Single
   Dim Y1 As Single
   Dim Y2 As Single
   Dim ws As String
   Dim t As Integer
   
   With Me
      With .picTimebar
        Y1 = .Top - 5
        Y2 = .Top
      End With
      .FontName = "Small Fonts"
      .FontSize = 6
      .ForeColor = vbWhite
      For i = 0 To 24
         x = picTimebar.Left + (i * 30)
         Me.Line (x, Y1)-(x, Y2), vbWhite
         t = i
         If .chkTimeType.Value = vbChecked Then
            ws = Format(t)
         Else
            If i >= 13 Then t = t + 1
            ws = Format((t Mod 13))
         End If
         .CurrentX = x - (.TextWidth(ws)) \ 2
         .CurrentY = 4
         Me.Print ws
      Next i
   End With
End Sub

