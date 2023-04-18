VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Login To Microsoft Services"
   ClientHeight    =   11445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17775
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ClientId As String
Public ApplicationName As String
Public Mode As String

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim objReg As RegistryUtility
    Set objReg = New RegistryUtility
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    If Mode = "Login" Then
        If InStr(1, WebBrowser1.LocationURL, "error_description=", vbTextCompare) > 0 Then
            Err.Raise vbObjectError, , objWebUtilities.URLDecode(GetErrorDescription(WebBrowser1.LocationURL))
        End If
        objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode", GetStateCode(WebBrowser1.LocationURL), REG_SZ
        If GetAuthorizationCode(WebBrowser1.LocationURL) <> vbNullString Then
            objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode", GetAuthorizationCode(WebBrowser1.LocationURL), REG_SZ
        End If
    End If
End Sub

Private Function GetStateCode(ByVal RedirectURL As String) As String
    Dim intPos1 As Integer, intPos2 As Integer
    Dim strCode As String
    intPos1 = InStr(1, RedirectURL, "state=", vbTextCompare)
    intPos2 = InStr(1, RedirectURL, "session_state=", vbTextCompare)
    If intPos1 > 0 And intPos2 > 0 Then strCode = Mid(RedirectURL, intPos1 + 6, intPos2 - intPos1 - 7)
    GetStateCode = strCode
End Function

Private Function GetAuthorizationCode(ByVal RedirectURL As String) As String
    Dim intPos1 As Integer, intPos2 As Integer
    Dim strCode As String
    intPos1 = InStr(1, RedirectURL, "code=", vbTextCompare)
    intPos2 = InStr(1, RedirectURL, "state=", vbTextCompare)
    If intPos1 > 0 And intPos2 > 0 Then strCode = Mid(RedirectURL, intPos1 + 5, intPos2 - intPos1 - 6)
    GetAuthorizationCode = strCode
End Function

Private Function GetErrorDescription(ByVal RedirectURL As String) As String
    Dim intPos1 As Integer, intPos2 As Integer
    Dim strCode As String
    intPos1 = InStr(1, RedirectURL, "error_description=", vbTextCompare)
    intPos2 = InStr(1, RedirectURL, "state=", vbTextCompare)
    If intPos1 > 0 And intPos2 > 0 Then strCode = Mid(RedirectURL, intPos1 + 18, intPos2 - intPos1 - 19)
    GetErrorDescription = strCode
End Function
