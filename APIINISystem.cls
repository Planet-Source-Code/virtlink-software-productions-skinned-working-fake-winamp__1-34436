VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIINISystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** INI-reader ***'
'Private m_cIni As CIniFile
Private strInI As String

Private Declare Function WritePrivateProfileString _
  Lib "kernel32" Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString _
  Lib "kernel32" Alias "GetPrivateProfileStringA" _
 (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long


Private Function INI_MakePath(ByVal strDrv As String, ByVal strDir As String) As String
 
 '// Makes an INI file: Guarantees a sub dir
 Do While Right$(strDrv, 1) = "\"
   strDrv = Left$(strDrv, Len(strDrv) - 1)
 Loop

 Do While Left$(strDir, 1) = "\"
   strDir = Mid$(strDir, 2)
 Loop

 '// Return the path
 INI_MakePath = strDrv & "\" & strDir
End Function

Public Sub INI_Create(strDrv As String, strDir As String)


 '// Make a new ini file
 strInI = INI_MakePath(strDrv, strDir)
End Sub

Public Sub INI_Write(strSection As String, strKey As String, strValue As String)
 
 '// Write to strINI
 WritePrivateProfileString strSection, _
   strKey, strValue, strInI
End Sub

Public Function INI_Read(strSection As String, strKey As String, Optional strDefault As String) As String

  Dim strTmp As String
  Dim lngRet As String

  strTmp = String$(100, 0)
  lngRet = GetPrivateProfileString(strSection, strKey, strDefault, strTmp, Len(strTmp), strInI)
    
  INI_Read = Left(strTmp, InStr(1, strTmp, Chr(0)) - 1)
 
End Function

Public Property Let INIFile(ByVal New_IniPath As String)

 '// Sets the new ini path
 strInI = New_IniPath
 
End Property

Public Property Get INIFile() As String

 '// Returns the current ini path
 INIFile = strInI
 
End Property


