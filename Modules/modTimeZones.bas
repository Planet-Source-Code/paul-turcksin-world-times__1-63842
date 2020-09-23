Attribute VB_Name = "modTimeZones"
'+
'   World Times - Graphical Representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          modTimeZones
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul Turcksin
'
'   Credits:
'   With the help of MSDN :
'   HOWTO: Change Time Zone Information Using Visual Basic
'   ID: Q221542

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
' Procedures:
'
'Public Function fncGetDisplayName(ByVal strTimeZone As String) As String
'   strTimeZone: Registry Key of time zone
'   returns the Display name, looks like "(GMT + 5) City1, city2, ...
'
'Public Function fncGetTimeZoneTime(strTimeZone As String) As Long
'   strTimeZone: Registry Key of time zone
'   returns time in time zone in minutes
'
' Public Sub subGetTimeZones(List1 As ListBox)
'   Returns in a listbox  Registry Key's of all time zones for Win9x and WinNT
'   registry structures.
'__________________________________________________________________________________
'-
Option Explicit

' Operating System version information declares

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128 ' Maintenance string
End Type

Private Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' Time Zone information declares

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type REGTIMEZONEINFORMATION
   Bias As Long
   StandardBias As Long
   DaylightBias As Long
   StandardDate As SYSTEMTIME
   DaylightDate As SYSTEMTIME
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte  ' used to accommodate Unicode strings
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte   ' used to accommodate Unicode strings
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function SetTimeZoneInformation Lib "kernel32" _
(lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' Registry information declares
Private Const REG_SZ As Long = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD As Long = 4

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_SUCCESS = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
   Alias "RegOpenKeyExA" ( _
   ByVal hKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   phkResult As Long) _
As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
   Alias "RegQueryValueExA" ( _
   ByVal hKey As Long, _
   ByVal lpszValueName As String, _
   ByVal lpdwReserved As Long, _
   lpdwType As Long, _
   lpData As Any, _
   lpcbData As Long) _
As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" _
   Alias "RegQueryValueExA" ( _
   ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   ByVal lpData As String, _
   lpcbData As Long) _
As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" _
   Alias "RegEnumKeyA" ( _
   ByVal hKey As Long, _
   ByVal dwIndex As Long, _
   ByVal lpName As String, _
   ByVal cbName As Long) _
As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
   ByVal hKey As Long) _
As Long

' Module level declares
Private Const SKEY_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
Private Const SKEY_9X = "SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones"

Dim SubKey As String

Public arZoneName() As String

Public Function fncGetDisplayName(ByVal strTimeZone As String) As String
'   strTimeZone: Key time zone
'
'   returns the Display name, looks like "(GMT + 5) City1, city2, ...

   Dim lRetVal As Long
   Dim hKeyResult As Long
   Dim lDataLen As Long
   Dim strDisplayName As String

' assume function fails
   fncGetDisplayName = "Not found"
   
' get key of time zone
   lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & strTimeZone, _
                           0, KEY_ALL_ACCESS, hKeyResult)
   If lRetVal = ERROR_SUCCESS Then
      ' get Display name
      lDataLen = 64
      strDisplayName = Space$(64)
      lRetVal = RegQueryValueExString(hKeyResult, "Display", _
           0&, REG_SZ, strDisplayName, lDataLen)
      If lRetVal = ERROR_SUCCESS Then
         fncGetDisplayName = Trim(strDisplayName)
      End If
   End If
End Function

Public Function fncGetTimeZoneTime(strTimeZone As String) As Long
'   strTimeZone: Key time zone
'
'   returns time in time zone in minutes

' This function does it the easy but smart way. It retrieves info of the time zone,
' preserves the current time settings and then sets the system time to the desired
' time zone. After collecting the time for this time zone the old setting is restored.

   Dim TZ As TIME_ZONE_INFORMATION     ' information specific to the time zone
   Dim oldTZ As TIME_ZONE_INFORMATION
   Dim rTZI As REGTIMEZONEINFORMATION  ' dayligth saving info
   Dim lRetVal As Long
   Dim hKeyResult As Long
   Dim ThisDay As Date

' assume function fails
   fncGetTimeZoneTime = -1
   
' get key of time zone
   lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & strTimeZone, _
                           0, KEY_ALL_ACCESS, hKeyResult)
   If lRetVal = ERROR_SUCCESS Then
      ' get time zone info and preserve
      lRetVal = RegQueryValueEx(hKeyResult, "TZI", 0&, ByVal 0&, _
                                  rTZI, Len(rTZI))
      If lRetVal = ERROR_SUCCESS Then
         TZ.Bias = rTZI.Bias
         TZ.StandardBias = rTZI.StandardBias
         TZ.DaylightBias = rTZI.DaylightBias
         TZ.StandardDate = rTZI.StandardDate
         TZ.DaylightDate = rTZI.DaylightDate
         ' Standard name of time zone and Daylight name are not used and can be empty
         TZ.StandardName(0) = 0
         TZ.DaylightName(0) = 0
         RegCloseKey hKeyResult
         End If
   Else
      MsgBox "Unable to retrieve time zone information for: " & strTimeZone
      RegCloseKey hKeyResult
      Exit Function
   End If
      
' preserve current settings
ThisDay = Date
   lRetVal = GetTimeZoneInformation(oldTZ)
   If lRetVal = TIME_ZONE_ID_INVALID Then
      MsgBox "Error getting current TimeZone Info"
      Exit Function
   Else
      lRetVal = SetTimeZoneInformation(TZ)
      If lRetVal = TIME_ZONE_ID_INVALID Then
         MsgBox "Error setting desired TimeZone Info"
         Exit Function
      Else
         ' got it!
         ' check if same day
         If ThisDay < Date Then
            fncGetTimeZoneTime = (Hour(Time) * 60 + Minute(Time)) + 1440
         ElseIf ThisDay > Date Then
            fncGetTimeZoneTime = (Hour(Time) * 60 + Minute(Time)) - 1440
         Else
            fncGetTimeZoneTime = (Hour(Time) * 60 + Minute(Time))
         End If
        ' restore original setting
         lRetVal = SetTimeZoneInformation(oldTZ)
         If lRetVal = TIME_ZONE_ID_INVALID Then
            MsgBox "Error restoring original TimeZone Info!" & vbCrLf & _
                   "Please reset manually.", vbCritical Or vbOKOnly
         End If
      End If
   End If

End Function


Public Sub subGetTimeZones(List1 As ListBox)
   Dim lRetVal As Long
   Dim lResult As Long
   Dim lCurIdx As Long
   Dim lDataLen As Long
   Dim lValueLen As Long
   Dim hKeyResult As Long
   Dim hKeyResult2 As Long
   Dim strValue As String
   Dim strDisplayName As String
   Dim osV As OSVERSIONINFO
   Dim l As Long

' Win9x and WinNT have a slightly different registry structure. Determine
' the operating system and set a module variable to the correct subkey.

   osV.dwOSVersionInfoSize = Len(osV)
   Call GetVersionEx(osV)
   If osV.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      SubKey = SKEY_NT
   Else
      SubKey = SKEY_9X
   End If

' Preserve "Display names" ([GMT+4]...)of time zones in a listbox
' The general time zone name (key name) is kept in an array arZoneName

' get key entry of all time zones
   lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey, 0, KEY_ALL_ACCESS, hKeyResult)
   If lRetVal = ERROR_SUCCESS Then
      lCurIdx = 0
      lDataLen = 32
      lValueLen = 32
 ' enumarate subkeys and retrieve "Display" name
      Do
         strValue = String(lValueLen, 0)
         lResult = RegEnumKey(hKeyResult, lCurIdx, strValue, lValueLen)
         l = InStr(1, strValue, Chr(0))
         strValue = Left(strValue, l - 1)
         If lResult = ERROR_SUCCESS Then
            ' open enumarated key
            lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, SubKey & "\" & strValue, _
                                    0, KEY_ALL_ACCESS, hKeyResult2)
            If lRetVal = ERROR_SUCCESS Then
               ' retrieve "Display" name
               lDataLen = 64
               strDisplayName = Space$(64)
               lRetVal = RegQueryValueExString(hKeyResult2, "Display", _
                     0&, REG_SZ, strDisplayName, lDataLen)
               If lRetVal = ERROR_SUCCESS Then
                  
                   List1.AddItem Left$(strDisplayName, lDataLen)
                   RegCloseKey hKeyResult2
               ' no "display" name found, use key name of time zone
               Else
                  List1.AddItem strValue
               End If
            End If
         End If
         ' save zone name in table
         ReDim Preserve arZoneName(lCurIdx)
         arZoneName(lCurIdx) = strValue
         ' and keep index in ItemData
         List1.ItemData(List1.NewIndex) = lCurIdx
         lCurIdx = lCurIdx + 1
      Loop While lResult = ERROR_SUCCESS

      RegCloseKey hKeyResult
   Else
      List1.AddItem "Could not open registry key"
   End If
End Sub


