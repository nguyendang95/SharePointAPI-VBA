VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objWShell As IWshRuntimeLibrary.WshShell
Public Enum KeyType
    REG_SZ = 0
    REG_DWORD = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
End Enum
Public Enum HKEY_TYPE
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Private Sub Class_Initialize()
    Set objWShell = New IWshRuntimeLibrary.WshShell
End Sub

Private Sub Class_Terminate()
    Set objWShell = Nothing
End Sub

Public Function ReadRegValue(Value As String) As Variant
    ReadRegValue = objWShell.RegRead(Value)
End Function

Public Sub DeleteRegValue(Value As String)
    objWShell.RegDelete Value
End Sub

Public Sub WriteRegValue(Key As String, Value As Variant, KeyType As KeyType)
    Dim strKeyType As String
    Select Case KeyType
        Case REG_SZ: strKeyType = "REG_SZ"
        Case REG_DWORD: strKeyType = "REG_DWORD"
        Case REG_EXPAND_SZ: strKeyType = "REG_EXPAND_SZ"
        Case REG_BINARY: strKeyType = "REG_BINARY"
    End Select
    objWShell.RegWrite Key, Value, strKeyType
End Sub

Public Function RegValueExists(Value As String) As Boolean
    On Error Resume Next
    Dim varKeyValue As Variant
    varKeyValue = objWShell.RegRead(Value)
    If Err.Number <> 0 Then
        Err.Clear
        RegValueExists = False
    Else: RegValueExists = True
    End If
End Function

Public Sub DeleteRegKey(HKeyRoot As HKEY_TYPE, KeyPath As String)
    Dim objRegKey As Object
    Set objRegKey = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    objRegKey.DeleteKey HKeyRoot, KeyPath
End Sub

Public Function RegKeyExists(HKeyRoot As HKEY_TYPE, KeyPath As String) As Boolean
    Dim objRegKey As Object
    Set objRegKey = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If objRegKey.EnumKey(HKeyRoot, KeyPath, "") = 0 Then
        RegKeyExists = True
    Else: RegKeyExists = False
    End If
End Function
