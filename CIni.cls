VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Private m_iniPath As String

Public Property Let path(ByVal pVal As String)
    m_iniPath = pVal
End Property

Public Property Get path() As String
    path = m_iniPath
End Property

Public Function checkIni() As Boolean
    If Len(Dir(m_iniPath)) > 0 Then
        checkIni = True
    Else
        checkIni = False
    End If
End Function

Public Function ReadString(ByVal section As String, ByVal key As String, Optional ByVal default As String, Optional ByVal Size As Long = 512) As String
    If Not checkIni Then Exit Function
    Dim ReturnStr As String, ReturnLng As Long
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(section, key, default, ReturnStr, Size, m_iniPath)
    ReadString = Left(ReturnStr, ReturnLng)
End Function

Public Function ReadNumber(ByVal section As String, ByVal key As String, Optional ByVal default As String, Optional ByVal Size As Long = 32) As Double
    If Not checkIni Then Exit Function
    Dim ReturnStr As String, ReturnLng As Long, ret As String
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(section, key, default, ReturnStr, Size, m_iniPath)
    ret = Left(ReturnStr, ReturnLng)
    If IsNumeric(ret) Then ReadNumber = Val(ret)
End Function

Public Function ReadDate(ByVal section As String, ByVal key As String, Optional ByVal default As String, Optional ByVal Size As Long = 32) As Date
    If Not checkIni Then Exit Function
    Dim ReturnStr As String, ReturnLng As Long, ret As String
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(section, key, default, ReturnStr, Size, m_iniPath)
    ret = Left(ReturnStr, ReturnLng)
    If IsDate(ret) Then ReadDate = CDate(ret)
End Function

Public Sub WriteString(ByVal section As String, ByVal key As String, ByVal value As String)
    If Not checkIni Then Exit Sub
    WritePrivateProfileString section, key, value, m_iniPath
End Sub


Private Sub Class_Initialize()
    m_iniPath = App.path & "\" & mDefine.INI_FILENAME
End Sub
