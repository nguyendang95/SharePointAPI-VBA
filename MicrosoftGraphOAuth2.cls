VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MicrosoftGraphOAuth2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TenantType
    Common
    Organizations
    Consumers
End Enum

Private m_ApplicationName As String
Private m_ClientId As String
Private m_ClientSecret As String
Private m_Scope As Variant
Private m_AccessToken As String
Private m_Tenant As TenantType
Private m_SiteResource As SiteResource
Private m_ListResource As ListResource
Private m_FilesResource As FilesResource
Attribute m_FilesResource.VB_VarHelpID = -1
Private m_PermissionResource As PermissionResource
Private m_ListItemResource As ListItemResource
Private m_ColumnDefinitionResource As ColumnDefinitionResource
Private m_DocumentSetVersionResource As DocumentSetVersionResource
Private m_ContentTypeResource As ContentTypeResource
Private m_TaxonomyResource As TaxonomyResource

Public Property Get TaxonomyResource() As TaxonomyResource
   Set TaxonomyResource = m_TaxonomyResource
End Property

Public Property Set TaxonomyResource(Value As TaxonomyResource)
   Set m_TaxonomyResource = Value
End Property

Public Property Get ContentTypeResource() As ContentTypeResource
   Set ContentTypeResource = m_ContentTypeResource
End Property

Public Property Set ContentTypeResource(Value As ContentTypeResource)
   Set m_ContentTypeResource = Value
End Property

Public Property Get DocumentSetVersionResource() As DocumentSetVersionResource
   Set DocumentSetVersionResource = m_DocumentSetVersionResource
End Property

Public Property Set DocumentSetVersionResource(Value As DocumentSetVersionResource)
   Set m_DocumentSetVersionResource = Value
End Property

Public Property Get ColumnDefinitionResource() As ColumnDefinitionResource
   Set ColumnDefinitionResource = m_ColumnDefinitionResource
End Property

Public Property Set ColumnDefinitionResource(Value As ColumnDefinitionResource)
   Set m_ColumnDefinitionResource = Value
End Property

Public Property Get ListItemResource() As ListItemResource
   Set ListItemResource = m_ListItemResource
End Property

Public Property Set ListItemResource(Value As ListItemResource)
   Set m_ListItemResource = Value
End Property

Public Property Get PermissionResource() As PermissionResource
   Set PermissionResource = m_PermissionResource
End Property

Public Property Set PermissionResource(Value As PermissionResource)
   Set m_PermissionResource = Value
End Property

Public Property Get FilesResource() As FilesResource
   Set FilesResource = m_FilesResource
End Property

Public Property Set FilesResource(Value As FilesResource)
   Set m_FilesResource = Value
End Property

Public Property Get ListResource() As ListResource
   Set ListResource = m_ListResource
End Property

Public Property Set ListResource(Value As ListResource)
   Set m_ListResource = Value
End Property

Public Property Get SiteResource() As SiteResource
   Set SiteResource = m_SiteResource
End Property

Public Property Set SiteResource(Value As SiteResource)
   Set m_SiteResource = Value
End Property

Public Property Get ApplicationName() As String
    ApplicationName = m_ApplicationName
End Property

Public Property Let ApplicationName(Value As String)
    m_ApplicationName = Value
End Property

Public Property Get ClientId() As String
    ClientId = m_ClientId
End Property

Public Property Let ClientId(Value As String)
    m_ClientId = Value
End Property

Public Property Get ClientSecret() As String
    ClientSecret = m_ClientSecret
End Property

Public Property Let ClientSecret(Value As String)
    m_ClientSecret = Value
End Property

Public Property Get Scope() As Variant
    Scope = m_Scope
End Property

Public Property Let Scope(Value As Variant)
    m_Scope = Value
End Property

Public Property Get AccessToken() As String
    AccessToken = m_AccessToken
End Property

Public Property Let AccessToken(Value As String)
    m_AccessToken = Value
End Property

Public Property Get Tenant() As TenantType
    Tenant = m_Tenant
End Property

Public Property Let Tenant(Value As TenantType)
    m_Tenant = Value
End Property

Public Sub AuthorizeOAuth2()
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim strRequestBody As String
    Dim strUrl As String
    Dim objReg As RegistryUtility
    Dim strAuthorizationCode As String
    Dim objFSO As Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.TextStream
    Dim strFolder As String
    Dim strFile As String
    Dim strAccessToken As String
    Dim objJson As Scripting.Dictionary
    Dim dteExpTime As Date
    Dim strRefreshToken As String
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    strFile = strFolder & "\client_id_" & ClientId & ".token-response"
    Set objReg = New RegistryUtility
    Set objFSO = New Scripting.FileSystemObject
    If ApplicationName = vbNullString Then
        Err.Raise vbObjectError + 10, , "Please specify a name for your application"
    End If
    If objFSO.FolderExists(strFolder) Then
        Set objFolder = objFSO.GetFolder(strFolder)
        If objFSO.FileExists(strFile) Then
            If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime") Then
                dteExpTime = CDate(objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime"))
                If DateDiff("s", Now, dteExpTime) <= 0 Then
                    Set objFile = objFSO.OpenTextFile(strFile, ForReading, False)
                    Set objJson = JsonConverter.ParseJson(objFile.ReadAll)
                    objFile.Close
                    strRefreshToken = objJson.Item("refresh_token")
                    AccessToken = RefreshAccessToken(strRefreshToken)
                Else
                    Set objFile = objFSO.OpenTextFile(strFile, ForReading, False)
                    Set objJson = JsonConverter.ParseJson(objFile.ReadAll)
                    objFile.Close
                    AccessToken = objJson.Item("access_token")
                End If
            Else: AccessToken = GetAccessToken
            End If
        Else: AccessToken = GetAccessToken
        End If
    Else
        Set objFolder = objFSO.CreateFolder(strFolder)
        AccessToken = GetAccessToken
    End If
    Dim objSiteResource As SiteResource
    Set objSiteResource = New SiteResource
    objSiteResource.AccessToken = AccessToken
    Set SiteResource = objSiteResource
    Dim objListResource As ListResource
    Set objListResource = New ListResource
    objListResource.AccessToken = AccessToken
    Set ListResource = objListResource
    Set m_FilesResource = New FilesResource
    m_FilesResource.AccessToken = AccessToken
    Set FilesResource = m_FilesResource
    Dim objPermissionResource As PermissionResource
    Set objPermissionResource = New PermissionResource
    objPermissionResource.AccessToken = AccessToken
    Set PermissionResource = objPermissionResource
    Dim objListItemResource As ListItemResource
    Set objListItemResource = New ListItemResource
    objListItemResource.AccessToken = AccessToken
    Set ListItemResource = objListItemResource
    Dim objColumnDefinitionResource As ColumnDefinitionResource
    Set objColumnDefinitionResource = New ColumnDefinitionResource
    objColumnDefinitionResource.AccessToken = AccessToken
    Set ColumnDefinitionResource = objColumnDefinitionResource
    Dim objDocumentSetVersionResource As DocumentSetVersionResource
    Set objDocumentSetVersionResource = New DocumentSetVersionResource
    objDocumentSetVersionResource.AccessToken = AccessToken
    Set DocumentSetVersionResource = objDocumentSetVersionResource
    Dim objContentTypeResource As ContentTypeResource
    Set objContentTypeResource = New ContentTypeResource
    objContentTypeResource.AccessToken = AccessToken
    Set ContentTypeResource = objContentTypeResource
    Dim objTaxonomyResource As TaxonomyResource
    Set objTaxonomyResource = New TaxonomyResource
    objTaxonomyResource.AccessToken = AccessToken
    Set TaxonomyResource = objTaxonomyResource
End Sub

Private Function GetAccessToken() As String
    Dim objBrowser As UserForm1
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim strUrl As String
    Dim strRequestBody As String
    Dim strAuthorizationCode As String
    Dim strFile As String
    Dim strFolder As String
    Dim strAccessToken As String
    Dim objReg As RegistryUtility
    Dim objJson As Scripting.Dictionary
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Set objFSO = New Scripting.FileSystemObject
    Set objReg = New RegistryUtility
    Set objBrowser = New UserForm1
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    strFile = strFolder & "\client_id_" & ClientId & ".token-response"
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strState As String
    strState = GenerateRandomString(128)
    Dim strTenant As String
    Select Case Tenant
        Case Common
            strTenant = "common"
        Case Organizations
            strTenant = "organizations"
        Case Consumers
            strTenant = "consumers"
        Case Else: Err.Raise vbObjectError + 14, , "Wrong value or empty tenant value!"
    End Select
    If Not IsEmpty(Scope) Then
        strUrl = "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/authorize?" & "client_id=" & ClientId & "&scope=" & objWebUtilities.URLEncode("offline_access ") & objWebUtilities.URLEncode(Join(Scope, " ")) & "&response_type=code&redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient") & "&state=" & strState & "&response_mode=query"
    Else
        Err.Raise vbObjectError + 13, , "Scope cannot be null. Please specify a valid scope and try again!"
    End If
    objBrowser.ClientId = ClientId
    objBrowser.ApplicationName = ApplicationName
    objBrowser.Mode = "Login"
    objBrowser.WebBrowser1.Navigate strUrl
    objBrowser.Show vbModal
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode") Then
        If objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode") = vbNullString Then Exit Function
        If strState <> objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode") Then
            objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode"
            objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode"
            Err.Raise vbObjectError, , "Malformed authorization response!"
        End If
        objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\ServerSideStateCode"
    End If
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode") Then
        strAuthorizationCode = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode")
        objReg.DeleteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AuthorizationCode"
    End If
    If strAuthorizationCode <> vbNullString Then
        strRequestBody = "code=" & strAuthorizationCode & "&" & _
                        "state=" & strState & "&" & _
                        "client_id=" & ClientId & "&" & _
                        "redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient") & "&" & _
                        "grant_type=authorization_code"
        Set objWinHttp = New WinHttp.WinHttpRequest
        With objWinHttp
            .Open "POST", "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/token", False
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetRequestHeader "Accept", "application/json'"
            .Send strRequestBody
            If .Status = 200 Then
                Set objJson = JsonConverter.ParseJson(.ResponseText)
                strAccessToken = objJson.Item("access_token")
                If strAccessToken <> vbNullString Then
                    GetAccessToken = strAccessToken
                    Set objFile = objFSO.CreateTextFile(strFile, True)
                    objFile.Write .ResponseText
                    objFile.Close
                    objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime", CStr(DateAdd("s", CDbl(objJson.Item("expires_in")), Now)), REG_SZ
                    objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant", LCase(strTenant), REG_SZ
                Else: Err.Raise vbObjectError + 2, , "Failed to get access code"
                End If
            End If
        End With
    Else
        Err.Raise vbObjectError + 1, , "Failed to obtain the authorization code."
    End If
End Function

Private Function RefreshAccessToken(ByVal RefreshToken As String) As String
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objJson As Scripting.Dictionary
    Dim strRequestBody As String
    Dim strFolder As String
    Dim strFile As String
    Dim objReg As RegistryUtility
    Set objReg = New RegistryUtility
    Dim strTenant As String
    strTenant = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant")
    strRequestBody = "client_id=" & ClientId & "&" & _
                    "refresh_token=" & RefreshToken & "&" & _
                    "grant_type=refresh_token"
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/token"
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .SetRequestHeader "Accept", "application/json'"
        .Send strRequestBody
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            strFolder = Environ$("APPDATA") & "\" & ApplicationName
            strFile = strFolder & "\client_id_" & ClientId & ".token-response"
            Call SaveNewAccessToken(strFile, objJson.Item("access_token"))
            RefreshAccessToken = objJson.Item("access_token")
            objReg.WriteRegValue "HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\AccessTokenExpirationTime", CStr(DateAdd("s", CDbl(objJson.Item("expires_in")), Now)), REG_SZ
        End If
    End With
End Function

Private Sub SaveNewAccessToken(ByVal AccessTokenFile As String, ByVal NewAccessToken As String)
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Dim strText As String, strOldText As String, strNewText As String
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.FileExists(AccessTokenFile) Then
        strNewText = Chr(34) & "access_token" & Chr(34) & ":" & Chr(34) & NewAccessToken & Chr(34)
        Set objFile = objFSO.OpenTextFile(AccessTokenFile)
        strText = objFile.ReadAll
        objFile.Close
        strOldText = Chr(34) & "access_token" & Chr(34) & ":" & Chr(34) & JsonConverter.ParseJson(strText).Item("access_token") & Chr(34)
        strText = VBA.Replace(strText, strOldText, strNewText, , , vbTextCompare)
        Set objFile = objFSO.OpenTextFile(AccessTokenFile, ForWriting)
        objFile.Write strText
        objFile.Close
    Else
        Call GetAccessToken
    End If
End Sub

Public Sub LogOut()
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objFSO As Scripting.FileSystemObject
    Dim strFolder As String
    Dim objReg As RegistryUtility
    Dim objBrowser As UserForm1
    Set objBrowser = New UserForm1
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strTenant As String
    Set objReg = New RegistryUtility
    If objReg.RegValueExists("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant") Then
        strTenant = objReg.ReadRegValue("HKEY_CURRENT_USER\Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId & "\Tenant")
    Else: Exit Sub
    End If
    If strTenant <> vbNullString Then objBrowser.WebBrowser1.Navigate "https://login.microsoftonline.com/" & strTenant & "/oauth2/v2.0/logout?post_logout_redirect_uri=" & objWebUtilities.URLEncode("https://login.microsoftonline.com/common/oauth2/nativeclient")
    objBrowser.Show vbModal
    Set objFSO = New Scripting.FileSystemObject
    strFolder = Environ$("APPDATA") & "\" & ApplicationName
    If objFSO.FileExists(strFolder & "\client_id_" & ClientId & ".token-response") Then objFSO.GetFile(strFolder & "\client_id_" & ClientId & ".token-response").Delete
    If objReg.RegKeyExists(HKEY_CURRENT_USER, "Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId) Then
        objReg.DeleteRegKey HKEY_CURRENT_USER, "Software\MicrosoftOAuth2VBA\" & ApplicationName & "\" & ClientId
    End If
End Sub

Private Function GenerateRandomString(ByVal Length As Long) As String
    Dim i As Long
    Dim strText As String
    If Length <= 0 Then Err.Raise vbObjectError, , "Length of string cannot below zero."
    Dim arrSpecialCharacters As Variant
    arrSpecialCharacters = Array("-", ".", "_", "~")
    Randomize
    For i = 1 To Length
        If RandBetween(0, 3) = 0 Then
            strText = strText & Chr(RandBetween(65, 90))
        ElseIf RandBetween(0, 3) = 1 Then strText = strText & Chr(RandBetween(48, 57))
        ElseIf RandBetween(0, 3) = 2 Then strText = strText & arrSpecialCharacters(RandBetween(0, 3))
        Else: strText = strText & Chr(RandBetween(97, 122))
        End If
    Next
    GenerateRandomString = strText
End Function

Private Function RandBetween(ByVal Num1 As Long, ByVal Num2 As Long) As Long
    RandBetween = CLng((Num2 - Num1) * Rnd + Num1)
End Function
