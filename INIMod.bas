Attribute VB_Name = "INIMod"
'This two Functions can read and write INIs
'Very useful for saving options
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetProfile(ByVal File As String, ByVal Section As String, ByVal KeyName As String, ByVal Default As Variant) As Variant
    '------------Read INI------------
    'Some Variables
    Dim Filename As String, RetString As String
    Dim DefSt As String
    Dim Size, RetSize
    'Set the Filename
    Filename = File
    DefSt = CStr(Default)
    'A buffer for the API
    RetString = Space(128)
    Size = Len(RetString)
    'Now call the GetAPI
    RetSize = GetPrivateProfileString(Section, KeyName, DefSt, RetString, Size, Filename)
    GetProfile = Left(RetString, RetSize)
End Function

Public Sub SaveProfile(ByVal File As String, ByVal Section As String, ByVal KeyName As String, ByVal Value As Variant)
    '----------Save INI------------
    Dim Filename As String
    Dim ValSt As String, Valid
    'Set the Filename
    Filename = File
    ValSt = CStr(Value)
    'Now call the WriteAPI
    Valid = WritePrivateProfileString(Section, KeyName, ValSt, Filename)
End Sub
