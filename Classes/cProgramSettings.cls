VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProgramSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+
'   World Times - Graphical representation of Daylight Saving Times
'
'   Application Name:     WorldTimes
'   Module name:          cProgramSettings
'
'   Compatibility:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software adapted by:
'       Paul Turcksin
'
'   Based on :
'   Simplified registry access routines used for saving program settings
'   Source      : Melvin Tucker
'   modified  : 03 May 2002

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
'   This code was developed by Melvin Tucker and modified by Paul Turcksin.
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
' Param names and default settings :
'    .RootKey = psrHKEY_LOCAL_MACHINE
'    .MainBranch = "SOFTWARE"
'    .RegBase = "KISS"
'    .Program  (no default)
'    .Section = "Settings"
'
' RootKey           \MainBranch\RegBase\Program\Settings
' HKEY_LOCAL_MACHINE\SOFTWARE   \KISS  \...    \Settings
'
'A typical set of options for a particular program might be:
'
'  'In declarations
'      Private mProgramSettings As CProgramSettings
'
'  ' In Form_Load
'     Set mProgramSettings = New CProgramSettings
'     mProgramSettings.Program = "Super Word Processor"
'...
'
' Modified Nov 2003
' to allow for shorter keys : Program and Setting can be null strings
'-

Option Explicit


' Declarations for Windows API calls
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' Constants for Windows API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_QUERY_VALUE = &H1

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

' Public property enumerated constants
Public Enum EnumProgramSettingsRoot
  psrHKEY_CLASSES_ROOT = &H80000000
  psrHKEY_CURRENT_USER = &H80000001
  psrHKEY_LOCAL_MACHINE = &H80000002
  psrHKEY_USERS = &H80000003
End Enum

' Local variables to hold Public Property values
Private m_lngRootKey As EnumProgramSettingsRoot
Private m_strMainBranch As String
Private m_strRegBase As String
Private m_strProgram As String
Private m_strSection As String

Private Sub Class_Initialize()
  ' Set initial values to defaults which may be overridden
  ' with property settings
  ' Source: Melvin Tucker

  m_lngRootKey = psrHKEY_LOCAL_MACHINE
  m_strMainBranch = "SOFTWARE"
  m_strRegBase = "KISS"
  m_strSection = "Settings"

End Sub

Public Property Get MainBranch() As String
  ' Returns: the current value of MainBranch
  ' Source: Melvin Tucker

  MainBranch = m_strMainBranch

End Property

Public Property Let MainBranch(strValue As String)
  m_strMainBranch = strValue
End Property

Public Property Get Program() As String
  Program = m_strProgram
End Property

Public Property Let Program(strValue As String)
  m_strProgram = strValue
End Property

Public Property Get RegBase() As String
  RegBase = m_strRegBase
End Property

Public Property Let RegBase(strValue As String)
  m_strRegBase = strValue
End Property

Public Property Get RootKey() As EnumProgramSettingsRoot
  RootKey = m_lngRootKey
End Property

Public Property Let RootKey(eValue As EnumProgramSettingsRoot)
  m_lngRootKey = eValue
End Property

Public Property Get Section() As String
  Section = m_strSection
End Property

Public Property Let Section(strValue As String)
  m_strSection = strValue
End Property

Public Function ReadEntry(strEntry As String, _
                          strDefault As String) As String
  ' Comments  : Reads a string value from the location in the
  '             registry specified by the class properties
  ' Parameters: strEntry - The value to retrieve
  '             strDefault - The value to return if the entry
  '             is not found
  ' Returns   : Either the registry value, or the default value
  '
  Dim strValue As String
  Dim strSearch As String
  
  On Error GoTo PROC_ERR
  
  strSearch = m_strMainBranch & "\" & m_strRegBase
  If m_strProgram <> "" Then strSearch = strSearch & "\" & m_strProgram
  If m_strSection <> "" Then strSearch = strSearch & "\" & m_strSection
  
  strValue = GetKeyValue(m_lngRootKey, strSearch, strEntry)
  
  If strValue = "" Then
    ReadEntry = strDefault
  Else
    ReadEntry = strValue
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReadEntry"
  Resume PROC_EXIT

End Function

Public Sub WriteEntry(strEntry As String, strValue As String)
  ' Comments  : Writes a string entry into the registry
  '             at the location specified by the class properties
  ' Parameters: strEntry - the string entry key value
  '             strValue - the value of the entry
  '
  ' Returns   : Nothing
  '
  Dim strSearch As String
  
  On Error GoTo PROC_ERR
  
  strSearch = m_strMainBranch & "\" & m_strRegBase
  If m_strProgram <> "" Then strSearch = strSearch & "\" & m_strProgram
  If m_strSection <> "" Then strSearch = strSearch & "\" & m_strSection
    
  SetKeyValue m_lngRootKey, strSearch, strEntry, strValue, REG_SZ

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WriteEntry"
  Resume PROC_EXIT

End Sub

Private Function GetKeyValue(eValue As EnumProgramSettingsRoot, _
                             strKeyName As String, _
                             strValueName As String) As Variant
  ' Comments  : Retrieves the specified key value
  ' Parameters: eValue - value indicating the root key value
  '             strKeyName - The name of the key to open
  '             strValueName - the value to open; vbNullString
  '             opens the default value
  ' Returns   : The value, if found, otherwise null
  '
  Dim r As Long      'The result of the API functions
  Dim lngHKey As Long         'The handle of opened key
  Dim varValue As Variant    'The setting of queried value
  Dim strValueData As String 'Used if the value is a string data type
  Dim lngValueData As Long   'Used if the value is a long data type
  Dim lngValueType As Long   'The data type of the value
  Dim lngDataSize As Long    'The size of the data

  On Error GoTo PROC_ERR

' Initialize return variable
   varValue = Empty
  
   r = RegOpenKeyEx(eValue, strKeyName, 0&, KEY_ALL_ACCESS, lngHKey)
   If r = ERROR_NONE Then
      r = RegQueryValueExNULL(lngHKey, strValueName, 0&, lngValueType, 0&, lngDataSize)
      If r = ERROR_NONE Then
         Select Case lngValueType
            Case REG_SZ:  ' Strings type
               strValueData = String(lngDataSize, 0)
               r = RegQueryValueExString(lngHKey, strValueName, 0&, _
                                        lngValueType, strValueData, _
                                        lngDataSize)
               If r = ERROR_NONE Then
                  If Len(strValueData) Then
                     If Mid$(strValueData, lngDataSize, 1) = vbNullChar Then
                        lngDataSize = lngDataSize - 1
                        End If
                     varValue = Left$(strValueData, lngDataSize)
                  Else
                     varValue = ""
                     End If
            
               Else
                  varValue = Empty
                  End If
        
        ' Long type
        Case REG_DWORD:
          r = RegQueryValueExLong( _
            lngHKey, _
            strValueName, _
            0&, _
            lngValueType, _
            lngValueData, _
            lngDataSize)
          If ERROR_NONE = r Then
            varValue = lngValueData
          End If
          
        Case Else
          'No other data types supported
          r = True
      End Select
        
    End If
        
    RegCloseKey (lngHKey)
    
  End If
    
  'Return varValue
  GetKeyValue = varValue

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetKeyValue"
  Resume PROC_EXIT

End Function

Private Sub SetKeyValue(eValue As EnumProgramSettingsRoot, _
                        strKeyName As String, _
                        strValueName As String, _
                        varValue As Variant, _
                        lngValueType As Long)
  ' Comments  : Sets a registry value Value. Will create key
  '             if it doesn't exist
  ' Parameters: eValue - The root key value
  '             strKeyName - The name of the key to open
  '             strValueName - The value to open, vbNullString will open the default value
  '             varValue - The data to assign to the value
  '             lngValueType - The data type of the value
  ' Returns   : Nothing
  '
  Dim lngReturnValue As Long
  Dim lngHKey As Long
        
  On Error GoTo PROC_ERR
  
  ' Open the specified key; if it does not exist then create it
  lngReturnValue = RegCreateKeyEx( _
    eValue, _
    strKeyName, _
    0&, _
    vbNullString, _
    REG_OPTION_NON_VOLATILE, _
    KEY_ALL_ACCESS, _
    0&, _
    lngHKey, _
    0&)
  
  'Determine the data type of the key
  Select Case lngValueType
    Case REG_SZ
      varValue = varValue & vbNullChar
      
      lngReturnValue = RegSetValueExString( _
        lngHKey, _
        strValueName, _
        0&, _
        lngValueType, _
        varValue, _
        Len(varValue))
    
    Case REG_DWORD
      lngReturnValue = RegSetValueExLong( _
        lngHKey, _
        strValueName, _
        0&, _
        lngValueType, _
        CLng(varValue), _
        REG_DWORD)
  End Select
  
  RegCloseKey (lngHKey)
  

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SetKeyValue"
  Resume PROC_EXIT

End Sub



