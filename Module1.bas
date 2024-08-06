Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetConfigValue(ByVal Section As String, ByVal Key As String, ByVal IniFile As String) As String
    Dim sBuffer As String * 256
    Dim lRet As Long
    
    lRet = GetPrivateProfileString(Section, Key, "", sBuffer, Len(sBuffer), IniFile)
    GetConfigValue = Left(sBuffer, lRet)
End Function

