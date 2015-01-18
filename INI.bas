Attribute VB_Name = "INI"
Public INIFILE As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function INIRead(Session As String, KeyValue As String, INIFILE As String) As String

    Dim s As String * 1024

    Dim ReturnValue As Long

    ReturnValue = GetPrivateProfileString(Session, KeyValue, "", s, 1024, INIFILE)

    INIRead = Left(s, InStr(s, Chr(0)) - 1)

End Function

Public Function INIWrite(Session As String, KeyValue As String, DataValue As String, INIFILE As String) As String

    Dim ReturnValue As Long

    ReturnValue = WritePrivateProfileString(Session, KeyValue, DataValue, INIFILE)

End Function

