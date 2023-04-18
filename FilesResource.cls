VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilesResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Destination
    DestinationDrive
    DestinationGroup
    DestinationSignedInUser
    DestinationSite
    DestinationOtherUser
    DestinationShare
End Enum

Public Enum SpecialFolderNames
    SpecialFolderNamesDocuments
    SpecialFolderNamesPhotos
    SpecialFolderNamesCameraRoll
    SpecialFolderNamesAppRoot
    SpecialFolderNamesMusic
End Enum

Private m_AccessToken As String
Public Event ProgressChanged(ByVal BytesSent As Variant, ByVal File As Object)

Friend Property Get AccessToken() As String
   AccessToken = m_AccessToken
End Property

Friend Property Let AccessToken(Value As String)
   m_AccessToken = Value
End Property

Public Function SimpleUpload(ByVal FileName As String, Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/content"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/content"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/content"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/content"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/content"
            End If
    End Select
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.FileExists(FileName) = False Then Err.Raise vbObjectError, , "File does not exist!"
    Dim objFile As Scripting.File
    Set objFile = objFSO.GetFile(FileName)
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objStream As ADODB.Stream
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PUT", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", objWebUtilities.GetMimeType(objFile.Name)
        .SetRequestHeader "Content-Length", CStr(objFile.Size)
        Set objStream = New ADODB.Stream
        With objStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile objFile.Path
        End With
        .Send objStream.Read
        .WaitForResponse
        objStream.Close
        RaiseEvent ProgressChanged(objFile.Size, objFile)
        If .Status = 413 Then Err.Raise vbObjectError + 413, , "Request too large"
        If .Status <> 201 Then
            Call HandleError(.ResponseText)
        End If
        If .Status = 201 Then
            Set SimpleUpload = FetchDriveItemInfo(.ResponseText)
        End If
    End With
End Function

Public Function ResumableUpload(ByVal FileName As String, Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal UploadableParameters As UploadableParameters, Optional ByVal IfMatch As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createUploadSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/createUploadSession"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createUploadSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/createUploadSession"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createUploadSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/createUploadSession"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createUploadSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/createUploadSession"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/createUploadSession"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/createUploadSession"
            End If
    End Select
    Dim objXmlHttp As MSXML2.XMLHTTP60
    Dim objJson As Scripting.Dictionary
    Dim strResumableUploadURI As String
    Dim objRandom As Random
    Dim decChunkSize As Variant, decCurrentPos As Variant
    Set objWebUtilities = New WebUtilities
    Set objXmlHttp = New MSXML2.XMLHTTP60
    With objXmlHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        If Not UploadableParameters Is Nothing Then
            .Send JsonConverter.ConvertToJson(UploadableParameters.ToJson)
        Else: .Send
        End If
        While .readyState <> 4
            DoEvents
        Wend
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status = 200 Then
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            strResumableUploadURI = objJson.Item("uploadUrl")
        Else
            Call HandleError(.ResponseText)
        End If
    End With
    Set objXmlHttp = Nothing
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.FileExists(FileName) = False Then Err.Raise vbObjectError, , "File does not exists!"
    Dim objFile As Scripting.File
    Set objFile = objFSO.GetFile(FileName)
    decCurrentPos = 0
    decChunkSize = 11200000
    While objFile.Size - decCurrentPos > 0
        Set objXmlHttp = New MSXML2.XMLHTTP60
        With objXmlHttp
            .Open "PUT", strResumableUploadURI, True
            .SetRequestHeader "Content-Length", CStr(decChunkSize)
            If objFile.Size - decCurrentPos >= decChunkSize Then
                .SetRequestHeader "Content-Range", "bytes " & CStr(decCurrentPos) & "-" & CStr(decCurrentPos + decChunkSize - 1) & "/" & CStr(objFile.Size)
            Else: .SetRequestHeader "Content-Range", "bytes " & CStr(decCurrentPos) & "-" & CStr(objFile.Size - 1) & "/" & CStr(objFile.Size)
            End If
            Set objRandom = New Random
            objRandom.OpenFile objFile.Path
            objRandom.SeekRelative decCurrentPos
            If objFile.Size - decCurrentPos >= decChunkSize Then
                .Send objRandom.ReadBytes(decChunkSize)
            Else: .Send objRandom.ReadBytes(objFile.Size - decCurrentPos)
            End If
            While .readyState <> 4
                DoEvents
            Wend
            objRandom.CloseFile
            Set objRandom = Nothing
            RaiseEvent ProgressChanged(decCurrentPos, objFile)
            If .Status = 201 Then
                Set ResumableUpload = FetchDriveItemInfo(.ResponseText)
            End If
            If .Status >= 400 Then Call HandleError(.ResponseText)
        End With
        Set objXmlHttp = Nothing
        decCurrentPos = decCurrentPos + decChunkSize
    Wend
End Function

Public Sub Download(ByVal DestinationFolder As String, Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfNoneMatch As String)
    Dim objWinHttp As WinHttp.WinHttpRequest
    Dim objRandom As Random
    Dim decCurrentPos As Variant, decChunkSize As Variant, decTotalSize As Variant
    Dim objJson As Scripting.Dictionary
    Dim objFSO As New Scripting.FileSystemObject
    Dim objDriveItem As DriveItem
    Set objDriveItem = GetDriveItem(Destination, ItemId, ItemPath, SiteId, DriveId, UserId, GroupId, IfNoneMatch, "select=size,@microsoft.graph.downloadUrl")
    decTotalSize = CDec(objDriveItem.Size)
    Set objWinHttp = New WinHttp.WinHttpRequest
    decChunkSize = CDec(decChunkSize)
    decChunkSize = 10485760
    decCurrentPos = CDec(decCurrentPos)
    decCurrentPos = 0
    If objFSO.FileExists(DestinationFolder) Then objFSO.DeleteFile DestinationFolder
    While decTotalSize - decCurrentPos > 0
        Set objWinHttp = New WinHttp.WinHttpRequest
        With objWinHttp
            .Open "GET", objDriveItem.DownloadUrl, True
            If decTotalSize - decCurrentPos > 0 Then
                .SetRequestHeader "Range", "bytes=" & CStr(decCurrentPos) & "-" & CStr(decCurrentPos + decChunkSize - 1)
            Else: .SetRequestHeader "Range", "bytes=" & CStr(decCurrentPos) & "-" & CStr(decTotalSize - 1)
            End If
            .Send
            .WaitForResponse
            If .Status = 206 Then
                Set objRandom = New Random
                objRandom.OpenFile DestinationFolder, ForAppending
                objRandom.WriteBytes .ResponseBody
                objRandom.CloseFile
                Set objRandom = Nothing
                RaiseEvent ProgressChanged(decCurrentPos, objDriveItem)
                decCurrentPos = decCurrentPos + decChunkSize
            Else: Call HandleError(.ResponseText)
            End If
        End With
        Set objWinHttp = Nothing
    Wend
End Sub

Public Function CreateFolder(Destination As Destination, ByVal DriveItem As DriveItem, Optional ByVal ParentItemId As String, Optional ByVal ParentItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ParentItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/root:/" & objWebUtilities.URLEncode(ParentItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ParentItemId & "/children"
            End If
        Case DestinationGroup
            If ParentItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/root:/" & objWebUtilities.URLEncode(ParentItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ParentItemId & "/children"
            End If
        Case DestinationOtherUser
            If ParentItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/root:/" & objWebUtilities.URLEncode(ParentItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ParentItemId & "/children"
            End If
        Case DestinationSignedInUser
            If ParentItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/root:/" & objWebUtilities.URLEncode(ParentItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ParentItemId & "/children"
            End If
        Case DestinationSite
            If ParentItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/root:/" & objWebUtilities.URLEncode(ParentItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ParentItemId & "/children"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(DriveItem.ToJson)
        .WaitForResponse
        If .Status = 201 Then
            Set CreateFolder = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function SearchItems(ByVal SearchText As String, Destination As Destination, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root/search(q='" & SearchText & "')"
        Case DestinationGroup
            strUrl = "https://graph.microsoft.com/v1.0/groups/drive" & GroupId & "/root/search(q='" & SearchText & "')"
        Case DestinationOtherUser
            strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root/search(q='" & SearchText & "')"
        Case DestinationSignedInUser
            strUrl = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='" & SearchText & "')"
        Case DestinationSite
            strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root/search(q='" & SearchText & "')"
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim colPagedSearchResults As New Collection
            colPagedSearchResults.Add ListDriveItems(.ResponseText)
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objWinHttpPagedRequest As WinHttp.WinHttpRequest
            Do While objJson.Exists("@odata.nextLink")
                Set objWinHttpPagedRequest = New WinHttp.WinHttpRequest
                With objWinHttp
                    .Open "GET", objJson.Item("@odata.nextLink"), True
                    .SetRequestHeader "Authorization", "Bearer " & AccessToken
                    .SetRequestHeader "Accept", "application/json"
                    .Send
                    .WaitForResponse
                    If .Status = 200 Then
                        colPagedSearchResults.Add ListDriveItems(.ResponseText)
                    End If
                End With
            Loop
            Set SearchItems = colPagedSearchResults
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetDocumentLibrary(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Drive
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetDocumentLibrary = GetDrive(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetDriveById(ByVal SiteId As String, DriveId As String, Optional ByVal ODataQuery As String) As Drive
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drives/" & DriveId
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetDriveById = GetDrive(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetUserDrive(SelectUserType As UserType, Optional ByVal IdOrUserPrincipalName As String, Optional ByVal ODataQuery As String) As Drive
    Dim strUrl As String
    Select Case SelectUserType
        Case UserTypeCurrentUser
            strUrl = "https://graph.microsoft.com/v1.0/me/drive"
        Case UserTypeOtherUser
            strUrl = "https://graph.microsoft.com/v1.0/users/" & IdOrUserPrincipalName & "/drive"
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetUserDrive = GetDrive(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetGroupDocumentLibrary(ByVal GroupId As String, Optional ByVal ODataQuery As String) As Drive
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetGroupDocumentLibrary = GetDrive(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListSiteDrives(ByVal SiteId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drives"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListSiteDrives = GetDrives(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListUserDrives(SelectUserType As UserType, Optional ByVal IdOrUserPrincipalName As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    Select Case SelectUserType
        Case UserTypeCurrentUser
            strUrl = "https://graph.microsoft.com/v1.0/me/drives"
        Case UserTypeOtherUser
            strUrl = "https://graph.microsoft.com/v1.0/users/" & IdOrUserPrincipalName & "/drives"
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListUserDrives = GetDrives(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListGroupDrives(ByVal GroupId As String, Optional ByVal ODataQuery As String) As Collection
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drives"
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListGroupDrives = GetDrives(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Sub HandleError(ByVal Source As String)
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(Source)
    Err.Raise vbObjectError, objJson.Item("error")("code"), objJson.Item("error")("message")
End Sub

Private Function GetDrives(ByVal ResponseText As String) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objDrive As Drive
    Dim objIdentity As Identity
    Dim objCreatedBy As IdentitySet, objLastModifiedBy As IdentitySet
    Dim objOwner As IdentitySet
    Dim objQuota As Quota
    Set objDrive = New Drive
    Dim i As Long
    Dim colDrives As New Collection
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            With objDrive
                .Id = objJson.Item("value")(i)("id")
                .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                .Description = objJson.Item("value")(i)("description")
                .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                .Name = objJson.Item("value")(i)("name")
                .WebUrl = objJson.Item("value")(i)("webUrl")
                .DriveType = objJson.Item("value")(i)("driveType")
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Item("value")(i).Exists("createdBy") Then
                    If objJson.Item("value")(i)("createdBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("createdBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("user")("id")
                        Set objCreatedBy.User = objIdentity
                    End If
                    If objJson.Item("value")(i)("createdBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("createdBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("application")("id")
                        Set objCreatedBy.Application = objIdentity
                    End If
                    If objJson.Item("value")(i)("createdBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("createdBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("device")("id")
                        Set objCreatedBy.Device = objIdentity
                    End If
                End If
                Set .CreatedBy = objCreatedBy
                If objJson.Item("value")(i).Exists("lastModifiedBy") Then
                    Set objLastModifiedBy = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("user")("id")
                        Set objLastModifiedBy.User = objIdentity
                    End If
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("application")("id")
                        Set objLastModifiedBy.Application = objIdentity
                    End If
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("device")("id")
                        Set objLastModifiedBy.Device = objIdentity
                    End If
                End If
                Set .LastModifiedBy = objLastModifiedBy
                If objJson.Item("value")(i).Exists("owner") Then
                    Set objOwner = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson.Item("value")(i)("owner").Exists("user") And IsEmpty(objJson.Item("value")(i)("owner")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("owner")("user")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("owner")("user")("id")
                        Set objOwner.User = objIdentity
                    End If
                    If objJson.Item("value")(i)("owner").Exists("application") And IsEmpty(objJson.Item("value")(i)("owner")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("owner")("application")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("owner")("application")("id")
                        Set objOwner.Application = objIdentity
                    End If
                    If objJson.Item("value")(i)("owner").Exists("device") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("owner")("device")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("owner")("device")("id")
                        Set objOwner.Device = objIdentity
                    End If
                End If
                Set .Owner = objOwner
                If objJson.Item("value")(i).Exists("quota") Then
                    Set objQuota = New Quota
                    With objQuota
                        .Deleted = objJson.Item("value")(i)("quota")("deleted")
                        .Remaining = objJson.Item("value")(i)("quota")("remaining")
                        .State = objJson.Item("value")(i)("quota")("state")
                        .Total = objJson.Item("value")(i)("quota")("total")
                        .Used = objJson.Item("value")(i)("quota")("used")
                    End With
                End If
                Set .Quota = objQuota
            End With
            colDrives.Add objDrive
        Next
        Set GetDrives = colDrives
    End If
End Function

Private Function FetchDriveItemInfo(ByVal ResponseText As String) As DriveItem
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objDriveItem As DriveItem
    Set objDriveItem = New DriveItem
    Dim objIdentity As Identity
    Dim objCreatedBy As IdentitySet, objLastModifiedBy As IdentitySet
    Dim objOwner As IdentitySet
    Dim objItemReference As ItemReference
    Dim objDriveFolder As DriveFolder
    Dim objSharepointIds As SharepointIds
    Dim objFileSystemInfo As FileSystemInfo
    Dim objSpecialFolder As SpecialFolder
    Dim objPackage As Package
    Dim objDriveFile As DriveFile
    Dim objHashes As Hashes
    Dim objItemShared As ItemShared
    Dim objSharedBy As IdentitySet
    Dim objImage As Image
    Dim objPhoto As Photo
    Dim objVideo As Video
    Dim objPublicationFacet As PublicationFacet
    Dim objMalware As Malware
    Dim objSearchResult As SearchResult
    With objDriveItem
        .DownloadUrl = objJson.Item("@microsoft.graph.downloadUrl")
        .Id = objJson.Item("id")
        .CreatedDateTime = objJson.Item("createdDateTime")
        .Description = objJson.Item("description")
        .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
        .Name = objJson.Item("name")
        .WebUrl = objJson.Item("webUrl")
        .Size = objJson.Item("size")
        .CTag = objJson.Item("cTag")
        .ETag = objJson.Item("eTag")
        .WebDavUrl = objJson.Item("webDavUrl")
        If objJson.Exists("folder") Then
            Set objDriveFolder = New DriveFolder
            If IsEmpty(objJson.Item("folder")) = False Then objDriveFolder.ChildCount = objJson.Item("folder")("childCount")
            Set .Folder = objDriveFolder
        End If
        Set objCreatedBy = New IdentitySet
        Set objIdentity = New Identity
        If objJson.Exists("createdBy") Then
            If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("user")("id")
                Set objCreatedBy.User = objIdentity
            End If
            If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("application")("id")
                Set objCreatedBy.Application = objIdentity
            End If
            If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("device")("id")
                Set objCreatedBy.Device = objIdentity
            End If
        End If
        Set .CreatedBy = objCreatedBy
        If objJson.Exists("lastModifiedBy") Then
            Set objLastModifiedBy = New IdentitySet
            Set objIdentity = New Identity
            If objJson.Item("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("lastModifiedBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("user")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("user")("id")
                Set objLastModifiedBy.User = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("lastModifiedBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("application")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("application")("id")
                Set objLastModifiedBy.Application = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("device")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("device")("id")
                Set objLastModifiedBy.Device = objIdentity
            End If
            Set .LastModifiedBy = objLastModifiedBy
        End If
        If objJson.Exists("parentReference") Then
            Set objItemReference = New ItemReference
            With objItemReference
                .DriveId = objJson.Item("parentReference")("driveId")
                .DriveType = objJson.Item("parentReference")("driveType")
                .Id = objJson.Item("parentReference")("id")
                .Path = objJson.Item("parentReference")("path")
                If objJson.Item("parentReference").Exists("sharepointIds") Then
                    Set objSharepointIds = New SharepointIds
                    With objSharepointIds
                        .ListId = objJson.Item("parentReference")("sharepointIds")("listId")
                        .ListItemUniqueId = objJson.Item("parentReference")("sharepointIds")("listItemUniqueId")
                        .SiteId = objJson.Item("parentReference")("sharepointIds")("siteId")
                        .SiteUrl = objJson.Item("parentReference")("sharepointIds")("siteUrl")
                        .TenantId = objJson.Item("parentReference")("sharepointIds")("tenantId")
                        .WebId = objJson.Item("parentReference")("sharepointIds")("webId")
                    End With
                    Set .SharepointIds = objSharepointIds
                End If
            End With
            Set .ParentReference = objItemReference
        End If
        If objJson.Exists("fileSystemInfo") And IsEmpty(objJson.Item("fileSystemInfo")) = False Then
            Set objFileSystemInfo = New FileSystemInfo
            objFileSystemInfo.CreatedDateTime = objJson.Item("fileSystemInfo")("createdDateTime")
            objFileSystemInfo.LastModifiedDateTime = objJson.Item("fileSystemInfo")("lastModifiedDateTime")
        End If
        If objJson.Exists("specialFolder") And IsEmpty(objJson.Item("specialFolder")) = False Then
            Set objSpecialFolder = New SpecialFolder
            objSpecialFolder.Name = objJson.Item("specialFolder")("name")
            Set .SpecialFolder = objSpecialFolder
        End If
        If objJson.Exists("package") And IsEmpty(objJson.Item("package")) = False Then
            Set objPackage = New Package
            objPackage.PackageType = objJson.Item("package")("type")
            Set .Package = objPackage
        End If
        If objJson.Exists("file") And IsEmpty(objJson.Item("file")) = False Then
            Set objDriveFile = New DriveFile
            With objDriveFile
                If objJson.Item("file").Exists("hashes") Then
                    Set objHashes = New Hashes
                    If IsEmpty(objJson.Item("file")("hashes")("quickXorHas")) = False Then objHashes.QuickXorHash = objJson.Item("file")("hashes")("quickXorHas")
                    If IsEmpty(objJson.Item("file")("hashes")("crc32Hash")) = False Then objHashes.Crc32Hash = objJson.Item("file")("hashes")("crc32Hash")
                    If IsEmpty(objJson.Item("file")("hashes")("sha1Hash")) = False Then objHashes.Sha1Hash = objJson.Item("file")("hashes")("sha1Hash")
                    If IsEmpty(objJson.Item("file")("hashes")("sha256Hash")) = False Then objHashes.Sha256Hash = objJson.Item("file")("hashes")("sha256Hash")
                    Set .Hashes = objHashes
                    .MimeType = objJson.Item("file")("mimeType")
                End If
            End With
            Set .File = objDriveFile
        End If
        If objJson.Exists("remoteItem") Then
            Set .RemoteItem = GetRemoteItem(objJson.Item("remoteItem"))
        End If
        If objJson.Exists("shared") Then
            Set objItemShared = New ItemShared
            Set objOwner = New IdentitySet
            Set objIdentity = New Identity
            If objJson("shared").Exists("owner") Then
                If objJson.Item("shared")("owner").Exists("user") And IsEmpty(objJson.Item("shared")("owner")("user")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("owner")("user")("displayName")
                    objIdentity.Id = objJson.Item("shared")("owner")("user")("id")
                    Set objOwner.User = objIdentity
                End If
                If objJson.Item("shared")("owner").Exists("application") And IsEmpty(objJson.Item("shared")("owner")("application")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("owner")("application")("displayName")
                    objIdentity.Id = objJson.Item("shared")("owner")("application")("id")
                    Set objOwner.Application = objIdentity
                End If
                If objJson.Item("shared")("owner").Exists("device") And IsEmpty(objJson.Item("shared")("owner")("device")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("owner")("device")("displayName")
                    objIdentity.Id = objJson.Item("shared")("owner")("device")("id")
                    Set objOwner.Device = objIdentity
                End If
            End If
            Set objItemShared.Owner = objOwner
            Set objSharedBy = New IdentitySet
            If objJson("shared").Exists("sharedBy") Then
                If objJson.Item("shared")("sharedBy").Exists("user") And IsEmpty(objJson.Item("shared")("sharedBy")("user")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("sharedBy")("user")("displayName")
                    objIdentity.Id = objJson.Item("shared")("sharedBy")("user")("id")
                    Set objSharedBy.User = objIdentity
                End If
                If objJson.Item("shared")("sharedBy").Exists("application") And IsEmpty(objJson.Item("shared")("sharedBy")("application")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("sharedBy")("application")("displayName")
                    objIdentity.Id = objJson.Item("shared")("sharedBy")("application")("id")
                    Set objSharedBy.Application = objIdentity
                End If
                If objJson.Item("shared")("sharedBy").Exists("device") And IsEmpty(objJson.Item("shared")("sharedBy")("device")) = False Then
                    objIdentity.DisplayName = objJson.Item("shared")("sharedBy")("device")("displayName")
                    objIdentity.Id = objJson.Item("shared")("sharedBy")("device")("id")
                    Set objSharedBy.Device = objIdentity
                End If
            End If
            Set objItemShared.SharedBy = objSharedBy
            objItemShared.Scope = objJson.Item("shared")("scope")
            objItemShared.SharedDateTime = objJson.Item("shared")("sharedDateTime")
            Set .ItemShared = objItemShared
        End If
        If objJson.Exists("photo") Then
            Set objPhoto = New Photo
            If IsEmpty(objJson.Item("photo")) = False Then
                With objPhoto
                    .CameraMake = objJson.Item("photo")("cameraMake")
                    .CameraModel = objJson.Item("photo")("cameraModel")
                    .ExposureDenominator = objJson.Item("photo")("exposureDenominator")
                    .ExposureNumerator = objJson.Item("photo")("exposureNumerator")
                    .FNumber = objJson.Item("photo")("fNumber")
                    .FocalLength = objJson.Item("photo")("focalLength")
                    .Iso = objJson.Item("photo")("iso")
                    .Orientation = objJson.Item("photo")("orientation")
                    .TakenDateTime = objJson.Item("photo")("takenDateTime")
                End With
            End If
            Set .Photo = objPhoto
        End If
        If objJson.Exists("video") Then
            Set objVideo = New Video
            If IsEmpty(objJson.Item("video")) = False Then
                With objVideo
                    .AudioBitsPerSample = objJson.Item("video")("audioBitsPerSample")
                    .AudioChannels = objJson.Item("video")("audioChannels")
                    .AudioFormat = objJson.Item("video")("audioFormat")
                    .AudioSamplesPerSecond = objJson.Item("video")("audioSamplesPerSecond")
                    .Bitrate = objJson.Item("video")("bitrate")
                    .Duration = objJson.Item("video")("duration")
                    .FourCC = objJson.Item("video")("fourCC")
                    .FrameRate = objJson.Item("video")("frameRate")
                    .Height = objJson.Item("video")("height")
                    .Width = objJson.Item("video")("width")
                End With
            End If
            Set .Video = objVideo
        End If
        If objJson.Exists("image") Then
            Set objImage = New Image
            If IsEmpty(objJson.Item("image")) = False Then
                objImage.Height = objJson.Item("image")("height")
                objImage.Width = objJson.Item("image")("width")
                Set .Image = objImage
            End If
        End If
        If objJson.Exists("publicationFacet") Then
            Set objPublicationFacet = New PublicationFacet
            objPublicationFacet.Level = objJson.Item("publicationFacet")("level")
            objPublicationFacet.VersionId = objJson.Item("publicationFacet")("versionId")
            Set .PublicationFacet = objPublicationFacet
        End If
        If objJson.Exists("malware") Then
            Set objMalware = New Malware
            objMalware.Description = objJson.Item("malware")("description")
        End If
        If objJson.Exists("searchResult") Then
            Set objSearchResult = New SearchResult
            objSearchResult.OnClickTelemetryUrl = objJson.Item("searchResult")("onClickTelemetryUrl")
            Set .SearchResult = objSearchResult
        End If
    End With
    Set FetchDriveItemInfo = objDriveItem
End Function

Private Function ListDriveItems(ByVal ResponseText) As Collection
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objDriveItem As DriveItem
    Dim colDriveItems As New Collection
    Dim i As Long
    Dim objCreatedBy As IdentitySet, objLastModifiedBy As IdentitySet
    Dim objIdentity As Identity
    Dim objDriveFolder As DriveFolder
    Dim objRoot As Root
    Dim objItemReference As ItemReference
    Dim objSharepointIds As SharepointIds
    Dim objFileSystemInfo As FileSystemInfo
    Dim objSpecialFolder As SpecialFolder
    Dim objPackage As Package
    Dim objDriveFile As DriveFile
    Dim objHashes As Hashes
    Dim objOwner As IdentitySet
    Dim objItemShared As ItemShared
    Dim objSharedBy As IdentitySet
    Dim objImage As Image
    Dim objPhoto As Photo
    Dim objVideo As Video
    Dim objPublicationFacet As PublicationFacet
    Dim objMalware As Malware
    Dim objSearchResult As SearchResult
    If objJson.Item("value").Count > 0 Then
        For i = 1 To objJson.Item("value").Count
            Set objDriveItem = New DriveItem
            With objDriveItem
                .DownloadUrl = objJson.Item("value")(i)("@microsoft.graph.downloadUrl")
                .Id = objJson.Item("value")(i)("id")
                .Name = objJson("value")(i)("name")
                .WebUrl = objJson("value")(i)("webUrl")
                .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                .Description = objJson.Item("value")(i)("description")
                .CreatedDateTime = objJson.Item("value")(i)("createdDateTime")
                .Size = objJson.Item("value")(i)("size")
                .CTag = objJson.Item("value")(i)("cTag")
                .ETag = objJson.Item("value")(i)("eTag")
                .WebDavUrl = objJson.Item("value")(i)("webDavUrl")
                If objJson("value")(i).Exists("folder") Then
                    Set objDriveFolder = New DriveFolder
                    If IsEmpty(objJson.Item("value")(i)("folder")) = False Then objDriveFolder.ChildCount = objJson.Item("value")(i)("folder")("childCount")
                    Set .Folder = objDriveFolder
                End If
                If objJson.Item("value")(i).Exists("root") Then
                    Set objRoot = New Root
                    Set .Root = objRoot
                End If
                Set objCreatedBy = New IdentitySet
                Set objIdentity = New Identity
                If objJson.Item("value")(i).Exists("createdBy") Then
                    If objJson.Item("value")(i)("createdBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("createdBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("user")("id")
                        Set objCreatedBy.User = objIdentity
                    End If
                    If objJson.Item("value")(i)("createdBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("createdBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("application")("id")
                        Set objCreatedBy.Application = objIdentity
                    End If
                    If objJson.Item("value")(i)("createdBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("createdBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("createdBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("createdBy")("device")("id")
                        Set objCreatedBy.Device = objIdentity
                    End If
                End If
                Set .CreatedBy = objCreatedBy
                If objJson.Item("value")(i).Exists("lastModifiedBy") Then
                    Set objLastModifiedBy = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("user")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("user")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("user")("id")
                        Set objLastModifiedBy.User = objIdentity
                    End If
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("application")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("application")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("application")("id")
                        Set objLastModifiedBy.Application = objIdentity
                    End If
                    If objJson.Item("value")(i)("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("device")) = False Then
                        objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("device")("displayName")
                        objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("device")("id")
                        Set objLastModifiedBy.Device = objIdentity
                    End If
                    Set .LastModifiedBy = objLastModifiedBy
                End If
                If objJson("value")(i).Exists("parentReference") And IsEmpty(objJson.Item("value")(i)("parentReference")) = False Then
                    Set objItemReference = New ItemReference
                    With objItemReference
                        .DriveId = objJson.Item("value")(i)("parentReference")("driveId")
                        .DriveType = objJson.Item("value")(i)("parentReference")("driveType")
                        .Id = objJson.Item("value")(i)("parentReference")("id")
                        .Path = objJson.Item("value")(i)("parentReference")("path")
                        If objJson.Item("value")(i)("parentReference").Exists("sharepointIds") And IsEmpty(objJson.Item("value")(i)("parentReference")("sharepointIds")) = False Then
                            Set objSharepointIds = New SharepointIds
                            With objSharepointIds
                                .ListId = objJson.Item("value")(i)("parentReference")("sharepointIds")("listId")
                                .ListItemUniqueId = objJson.Item("value")(i)("parentReference")("sharepointIds")("listItemUniqueId")
                                .SiteId = objJson.Item("value")(i)("parentReference")("sharepointIds")("siteId")
                                .SiteUrl = objJson.Item("value")(i)("parentReference")("sharepointIds")("siteUrl")
                                .TenantId = objJson.Item("value")(i)("parentReference")("sharepointIds")("tenantId")
                                .WebId = objJson.Item("value")(i)("parentReference")("sharepointIds")("webId")
                            End With
                            Set .SharepointIds = objSharepointIds
                        End If
                    End With
                    Set .ParentReference = objItemReference
                End If
                If objJson("value")(i).Exists("fileSystemInfo") And IsEmpty(objJson.Item("value")(i)("folder")) = False Then
                    Set objFileSystemInfo = New FileSystemInfo
                    objFileSystemInfo.CreatedDateTime = objJson.Item("value")(i)("fileSystemInfo")("createdDateTime")
                    objFileSystemInfo.LastModifiedDateTime = objJson.Item("value")(i)("fileSystemInfo")("lastModifiedDateTime")
                End If
                If objJson("value")(i).Exists("specialFolder") And IsEmpty(objJson.Item("value")(i)("specialFolder")) = False Then
                    Set objSpecialFolder = New SpecialFolder
                    objSpecialFolder.Name = objJson.Item("value")(i)("specialFolder")("name")
                    Set .SpecialFolder = objSpecialFolder
                End If
                If objJson.Item("value")(i).Exists("package") And IsEmpty(objJson.Item("value")(i)("package")) = False Then
                    Set objPackage = New Package
                    objPackage.PackageType = objJson.Item("value")(i)("package")("type")
                    Set .Package = objPackage
                End If
                If objJson.Exists("file") And IsEmpty(objJson.Item("value")(i)("file")) = False Then
                    Set objDriveFile = New DriveFile
                    With objDriveFile
                        If objJson.Item("value")(i)("file").Exists("hashes") Then
                            Set objHashes = New Hashes
                            If IsEmpty(objJson.Item("value")(i)("file")("hashes")("quickXorHas")) = False Then objHashes.QuickXorHash = objJson.Item("value")(i)("file")("hashes")("quickXorHas")
                            If IsEmpty(objJson.Item("value")(i)("file")("hashes")("crc32Hash")) = False Then objHashes.Crc32Hash = objJson.Item("value")(i)("file")("hashes")("crc32Hash")
                            If IsEmpty(objJson.Item("value")(i)("file")("hashes")("sha1Hash")) = False Then objHashes.Sha1Hash = objJson.Item("value")(i)("file")("hashes")("sha1Hash")
                            If IsEmpty(objJson.Item("value")(i)("file")("hashes")("sha256Hash")) = False Then objHashes.Sha256Hash = objJson.Item("value")(i)("file")("hashes")("sha256Hash")
                            Set .Hashes = objHashes
                            .MimeType = objJson.Item("file")("mimeType")
                        End If
                    End With
                End If
                If objJson.Item("value")(i).Exists("remoteItem") Then
                    Set .RemoteItem = GetRemoteItem(objJson.Item("value")(i)("remoteItem"))
                End If
                If objJson("value")(i).Exists("shared") Then
                    Set objItemShared = New ItemShared
                    Set objOwner = New IdentitySet
                    Set objIdentity = New Identity
                    If objJson("value")(i)("shared").Exists("owner") Then
                        If objJson.Item("value")(i)("shared")("owner").Exists("user") And IsEmpty(objJson.Item("value")(i)("shared")("owner")("user")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("owner")("user")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("owner")("user")("id")
                            Set objOwner.User = objIdentity
                        End If
                        If objJson.Item("value")(i)("shared")("owner").Exists("application") And IsEmpty(objJson.Item("value")(i)("shared")("owner")("application")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("owner")("application")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("owner")("application")("id")
                            Set objOwner.Application = objIdentity
                        End If
                        If objJson.Item("value")(i)("shared")("owner").Exists("device") And IsEmpty(objJson.Item("value")(i)("shared")("owner")("device")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("owner")("device")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("owner")("device")("id")
                            Set objOwner.Device = objIdentity
                        End If
                    End If
                    Set objItemShared.Owner = objOwner
                    Set objSharedBy = New IdentitySet
                    If objJson("value")(i)("shared").Exists("sharedBy") Then
                        If objJson.Item("value")(i)("shared")("sharedBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("shared")("sharedBy")("user")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("sharedBy")("user")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("sharedBy")("user")("id")
                            Set objSharedBy.User = objIdentity
                        End If
                        If objJson.Item("value")(i)("shared")("sharedBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("shared")("sharedBy")("application")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("sharedBy")("application")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("sharedBy")("application")("id")
                            Set objSharedBy.Application = objIdentity
                        End If
                        If objJson.Item("value")(i)("shared")("sharedBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("shared")("sharedBy")("device")) = False Then
                            objIdentity.DisplayName = objJson.Item("value")(i)("shared")("sharedBy")("device")("displayName")
                            objIdentity.Id = objJson.Item("value")(i)("shared")("sharedBy")("device")("id")
                            Set objSharedBy.Device = objIdentity
                        End If
                    End If
                    Set objItemShared.SharedBy = objSharedBy
                    objItemShared.Scope = objJson.Item("value")(i)("shared")("scope")
                    objItemShared.SharedDateTime = objJson.Item("value")(i)("shared")("sharedDateTime")
                    Set .ItemShared = objItemShared
                End If
                If objJson("value")(i).Exists("photo") Then
                    Set objPhoto = New Photo
                    If IsEmpty(objJson.Item("value")(i)("photo")) = False Then
                        With objPhoto
                            .CameraMake = objJson.Item("value")(i)("photo")("cameraMake")
                            .CameraModel = objJson.Item("value")(i)("photo")("cameraModel")
                            .ExposureDenominator = objJson.Item("value")(i)("photo")("exposureDenominator")
                            .ExposureNumerator = objJson.Item("value")(i)("photo")("exposureNumerator")
                            .FNumber = objJson.Item("value")(i)("photo")("fNumber")
                            .FocalLength = objJson.Item("value")(i)("photo")("focalLength")
                            .Iso = objJson.Item("value")(i)("photo")("iso")
                            .Orientation = objJson.Item("value")(i)("photo")("orientation")
                            .TakenDateTime = objJson.Item("value")(i)("photo")("takenDateTime")
                        End With
                    End If
                    Set .Photo = objPhoto
                End If
                If objJson("value")(i).Exists("video") Then
                    Set objVideo = New Video
                    If IsEmpty(objJson.Item("value")(i)("video")) = False Then
                        With objVideo
                            .AudioBitsPerSample = objJson.Item("value")(i)("video")("audioBitsPerSample")
                            .AudioChannels = objJson.Item("value")(i)("video")("audioChannels")
                            .AudioFormat = objJson.Item("value")(i)("video")("audioFormat")
                            .AudioSamplesPerSecond = objJson.Item("value")(i)("video")("audioSamplesPerSecond")
                            .Bitrate = objJson.Item("value")(i)("video")("bitrate")
                            .Duration = objJson.Item("value")(i)("video")("duration")
                            .FourCC = objJson.Item("value")(i)("video")("fourCC")
                            .FrameRate = objJson.Item("value")(i)("video")("frameRate")
                            .Height = objJson.Item("value")(i)("video")("height")
                            .Width = objJson.Item("value")(i)("video")("width")
                        End With
                    End If
                    Set .Video = objVideo
                End If
                If objJson("value")(i).Exists("image") Then
                    Set objImage = New Image
                    If IsEmpty(objJson.Item("value")(i)("image")) = False Then
                        objImage.Height = objJson.Item("value")(i)("image")("height")
                        objImage.Width = objJson.Item("value")(i)("image")("width")
                        Set .Image = objImage
                    End If
                End If
                If objJson("value")(i).Exists("publicationFacet") Then
                    Set objPublicationFacet = New PublicationFacet
                    objPublicationFacet.Level = objJson.Item("value")(i)("publicationFacet")("level")
                    objPublicationFacet.VersionId = objJson.Item("value")(i)("publicationFacet")("versionId")
                    Set .PublicationFacet = objPublicationFacet
                End If
                If objJson("value")(i).Exists("malware") Then
                    Set objMalware = New Malware
                    objMalware.Description = objJson.Item("value")(i)("malware")("description")
                End If
                If objJson("value")(i).Exists("searchResult") Then
                    Set objSearchResult = New SearchResult
                    objSearchResult.OnClickTelemetryUrl = objJson.Item("value")(i)("searchResult")("onClickTelemetryUrl")
                    Set .SearchResult = objSearchResult
                End If
            End With
            colDriveItems.Add objDriveItem
        Next
    End If
    Set ListDriveItems = colDriveItems
End Function

Private Function GetDrive(ByVal ResponseText As String) As Drive
    Dim objJson As Scripting.Dictionary
    Set objJson = JsonConverter.ParseJson(ResponseText)
    Dim objDrive As Drive
    Dim objIdentity As Identity
    Dim objCreatedBy As IdentitySet, objLastModifiedBy As IdentitySet
    Dim objOwner As IdentitySet
    Dim objQuota As Quota
    Set objDrive = New Drive
    With objDrive
        .Id = objJson.Item("id")
        .CreatedDateTime = objJson.Item("createdDateTime")
        .Description = objJson.Item("description")
        .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
        .Name = objJson.Item("name")
        .WebUrl = objJson.Item("webUrl")
        .DriveType = objJson.Item("driveType")
        Set objCreatedBy = New IdentitySet
        Set objIdentity = New Identity
        If objJson.Exists("createdBy") Then
            If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("user")("id")
                Set objCreatedBy.User = objIdentity
            End If
            If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("application")("id")
                Set objCreatedBy.Application = objIdentity
            End If
            If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("device")("id")
                Set objCreatedBy.Device = objIdentity
            End If
        End If
        Set .CreatedBy = objCreatedBy
        If objJson.Exists("lastModifiedBy") Then
            Set objLastModifiedBy = New IdentitySet
            Set objIdentity = New Identity
            If objJson.Item("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("lastModifiedBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("user")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("user")("id")
                Set objLastModifiedBy.User = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("lastModifiedBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("application")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("application")("id")
                Set objLastModifiedBy.Application = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("device")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("device")("id")
                Set objLastModifiedBy.Device = objIdentity
            End If
        End If
        Set .LastModifiedBy = objLastModifiedBy
        If objJson.Exists("owner") Then
            Set objOwner = New IdentitySet
            Set objIdentity = New Identity
            If objJson.Item("owner").Exists("user") And IsEmpty(objJson.Item("owner")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("user")("displayName")
                objIdentity.Id = objJson.Item("owner")("user")("id")
                Set objOwner.User = objIdentity
            End If
            If objJson.Item("owner").Exists("application") And IsEmpty(objJson.Item("owner")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("application")("displayName")
                objIdentity.Id = objJson.Item("owner")("application")("id")
                Set objOwner.Application = objIdentity
            End If
            If objJson.Item("owner").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("owner")("device")("displayName")
                objIdentity.Id = objJson.Item("owner")("device")("id")
                Set objOwner.Device = objIdentity
            End If
        End If
        Set .Owner = objOwner
        If objJson.Exists("quota") Then
            Set objQuota = New Quota
            With objQuota
                .Deleted = objJson.Item("quota")("deleted")
                .Remaining = objJson.Item("quota")("remaining")
                .State = objJson.Item("quota")("state")
                .Total = objJson.Item("quota")("total")
                .Used = objJson.Item("quota")("used")
            End With
        End If
        Set .Quota = objQuota
    End With
Set GetDrive = objDrive
End Function

Public Function ListChildren(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal SiteId As String, Optional ByVal IfNoneMatch As String, Optional ByVal ODataQuery As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/children"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/children"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/children"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/children"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/children"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/children"
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If IfNoneMatch <> vbNullString Then .SetRequestHeader "if-none-match", IfNoneMatch
        .Send
        .WaitForResponse
        If .Status = 304 Then
            Call HandleError(.ResponseText)
        End If
        If .Status = 200 Then
            Dim colPagedResults As New Collection
            colPagedResults.Add ListDriveItems(.ResponseText)
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objWinHttpPagedRequest As WinHttp.WinHttpRequest
            Do While objJson.Exists("@odata.nextLink")
                Set objWinHttpPagedRequest = New WinHttp.WinHttpRequest
                With objWinHttp
                    .Open "GET", objJson.Item("@odata.nextLink"), True
                    .SetRequestHeader "Authorization", "Bearer " & AccessToken
                    .SetRequestHeader "Accept", "application/json"
                    .Send
                    .WaitForResponse
                    If .Status = 200 Then
                        colPagedResults.Add ListDriveItems(.ResponseText)
                    End If
                End With
            Loop
            Set ListChildren = colPagedResults
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Private Function GetRemoteItem(ResponseText As Variant) As RemoteItem
    Dim objJson As Scripting.Dictionary
    If TypeName(ResponseText) = "Dictionary" Then
        Set objJson = ResponseText
    ElseIf TypeName(ResponseText) = "String" Then
        Set objJson = JsonConverter.ParseJson(ResponseText)
    Else: Exit Function
    End If
    Dim objDriveItem As RemoteItem
    Set objDriveItem = New RemoteItem
    Dim objIdentity As Identity
    Dim objCreatedBy As IdentitySet, objLastModifiedBy As IdentitySet
    Dim objOwner As IdentitySet
    Dim objItemReference As ItemReference
    Dim objDriveFolder As DriveFolder
    Dim objSharepointIds As SharepointIds
    Dim objFileSystemInfo As FileSystemInfo
    Dim objSpecialFolder As SpecialFolder
    Dim objPackage As Package
    Dim objDriveFile As DriveFile
    Dim objHashes As Hashes
    With objDriveItem
        .Id = objJson.Item("id")
        .CreatedDateTime = objJson.Item("createdDateTime")
        .LastModifiedDateTime = objJson.Item("lastModifiedDateTime")
        .Name = objJson.Item("name")
        .WebUrl = objJson.Item("webUrl")
        .Size = objJson.Item("size")
        .WebDavUrl = objJson.Item("webDavUrl")
        If objJson.Exists("folder") Then
            Set objDriveFolder = New DriveFolder
            If IsEmpty(objJson.Item("folder")) = False Then objDriveFolder.ChildCount = objJson.Item("folder")("childCount")
            Set .Folder = objDriveFolder
        End If
        Set objCreatedBy = New IdentitySet
        Set objIdentity = New Identity
        If objJson.Exists("createdBy") Then
            If objJson.Item("createdBy").Exists("user") And IsEmpty(objJson.Item("createdBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("user")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("user")("id")
                Set objCreatedBy.User = objIdentity
            End If
            If objJson.Item("createdBy").Exists("application") And IsEmpty(objJson.Item("createdBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("application")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("application")("id")
                Set objCreatedBy.Application = objIdentity
            End If
            If objJson.Item("createdBy").Exists("device") And IsEmpty(objJson.Item("createdBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("createdBy")("device")("displayName")
                objIdentity.Id = objJson.Item("createdBy")("device")("id")
                Set objCreatedBy.Device = objIdentity
            End If
        End If
        Set .CreatedBy = objCreatedBy
        If objJson.Exists("lastModifiedBy") Then
            Set objLastModifiedBy = New IdentitySet
            Set objIdentity = New Identity
            If objJson.Item("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("lastModifiedBy")("user")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("user")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("user")("id")
                Set objLastModifiedBy.User = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("lastModifiedBy")("application")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("application")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("application")("id")
                Set objLastModifiedBy.Application = objIdentity
            End If
            If objJson.Item("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("lastModifiedBy")("device")) = False Then
                objIdentity.DisplayName = objJson.Item("lastModifiedBy")("device")("displayName")
                objIdentity.Id = objJson.Item("lastModifiedBy")("device")("id")
                Set objLastModifiedBy.Device = objIdentity
            End If
            Set .LastModifiedBy = objLastModifiedBy
        End If
        If objJson.Exists("parentReference") And IsEmpty(objJson.Item("parentReference")) = False Then
            Set objItemReference = New ItemReference
            With objItemReference
                .DriveId = objJson.Item("parentReference")("driveId")
                .DriveType = objJson.Item("parentReference")("driveType")
                .Id = objJson.Item("parentReference")("id")
                .Path = objJson.Item("parentReference")("path")
                If objJson.Item("parentReference").Exists("sharepointIds") And IsEmpty(objJson.Item("parentReference")("sharepointIds")) = False Then
                    Set objSharepointIds = New SharepointIds
                    With objSharepointIds
                        .ListId = objJson.Item("parentReference")("sharepointIds")("listId")
                        .ListItemUniqueId = objJson.Item("parentReference")("sharepointIds")("listItemUniqueId")
                        .SiteId = objJson.Item("parentReference")("sharepointIds")("siteId")
                        .SiteUrl = objJson.Item("parentReference")("sharepointIds")("siteUrl")
                        .TenantId = objJson.Item("parentReference")("sharepointIds")("tenantId")
                        .WebId = objJson.Item("parentReference")("sharepointIds")("webId")
                    End With
                    Set .SharepointIds = objSharepointIds
                End If
            End With
            Set .ParentReference = objItemReference
        End If
        If objJson.Exists("fileSystemInfo") And IsEmpty(objJson.Item("fileSystemInfo")) = False Then
            Set objFileSystemInfo = New FileSystemInfo
            objFileSystemInfo.CreatedDateTime = objJson.Item("fileSystemInfo")("createdDateTime")
            objFileSystemInfo.LastModifiedDateTime = objJson.Item("fileSystemInfo")("lastModifiedDateTime")
        End If
        If objJson.Exists("specialFolder") And IsEmpty(objJson.Item("specialFolder")) = False Then
            Set objSpecialFolder = New SpecialFolder
            objSpecialFolder.Name = objJson.Item("specialFolder")("name")
            Set .SpecialFolder = objSpecialFolder
        End If
        If objJson.Exists("package") And IsEmpty(objJson.Item("package")) = False Then
            Set objPackage = New Package
            objPackage.PackageType = objJson.Item("package")("type")
            Set .Package = objPackage
        End If
        If objJson.Exists("file") And IsEmpty(objJson.Item("file")) = False Then
            Set objDriveFile = New DriveFile
            With objDriveFile
                If objJson.Item("file").Exists("hashes") Then
                    Set objHashes = New Hashes
                    If IsEmpty(objJson.Item("file")("hashes")("quickXorHas")) = False Then objHashes.QuickXorHash = objJson.Item("file")("hashes")("quickXorHas")
                    If IsEmpty(objJson.Item("file")("hashes")("crc32Hash")) = False Then objHashes.Crc32Hash = objJson.Item("file")("hashes")("crc32Hash")
                    If IsEmpty(objJson.Item("file")("hashes")("sha1Hash")) = False Then objHashes.Sha1Hash = objJson.Item("file")("hashes")("sha1Hash")
                    If IsEmpty(objJson.Item("file")("hashes")("sha256Hash")) = False Then objHashes.Sha256Hash = objJson.Item("file")("hashes")("sha256Hash")
                    Set .Hashes = objHashes
                    .MimeType = objJson.Item("file")("mimeType")
                End If
            End With
            Set .File = objDriveFile
        End If
    End With
    Set GetRemoteItem = objDriveItem
End Function

Public Function ListRecentFiles() As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/me/drive/recent", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListRecentFiles = ListDriveItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function
 
Public Function ListSharedFiles() As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/me/drive/sharedWithMe", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListSharedFiles = ListDriveItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetSpecialFolder(SpecialFolderName As SpecialFolderNames, Optional ByVal ODataQuery As String) As DriveItem
    Dim strSpecialFolderName As String
    Select Case SpecialFolderName
        Case SpecialFolderNamesAppRoot
            strSpecialFolderName = "approot"
        Case SpecialFolderNamesCameraRoll
            strSpecialFolderName = "cameraroll"
        Case SpecialFolderNamesDocuments
            strSpecialFolderName = "documents"
        Case SpecialFolderNamesMusic
            strSpecialFolderName = "music"
        Case SpecialFolderNamesPhotos
            strSpecialFolderName = "photos"
    End Select
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/me/drive/special/" & strSpecialFolderName
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetSpecialFolder = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetSpecialFolderChildren(ByVal SpecialFolderName As SpecialFolderNames) As Collection
    Dim strSpecialFolderName As String
    Select Case SpecialFolderName
        Case SpecialFolderNamesAppRoot
            strSpecialFolderName = "approot"
        Case SpecialFolderNamesCameraRoll
            strSpecialFolderName = "cameraroll"
        Case SpecialFolderNamesDocuments
            strSpecialFolderName = "documents"
        Case SpecialFolderNamesMusic
            strSpecialFolderName = "music"
        Case SpecialFolderNamesPhotos
            strSpecialFolderName = "photos"
    End Select
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/me/drive/special/" & strSpecialFolderName & "/children"
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set GetSpecialFolderChildren = ListDriveItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function ListFollowedItems() As Collection
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", "https://graph.microsoft.com/v1.0/me/drive/following", True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set ListFollowedItems = ListDriveItems(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetDriveItem(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfNoneMatch As String, Optional ByVal ODataQuery As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId
            End If
    End Select
    If ODataQuery <> vbNullString Then strUrl = strUrl & "?" & ODataQuery
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If IfNoneMatch <> vbNullString Then .SetRequestHeader "if-none-match", IfNoneMatch
        .Send
        .WaitForResponse
        If .Status = 304 Then
            Call HandleError(.ResponseText)
        End If
        If .Status = 200 Then
            Set GetDriveItem = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function UpdateItemProperties(Destination As Destination, ByVal DriveItem As DriveItem, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfMatch As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send JsonConverter.ConvertToJson(DriveItem.ToJson)
        .WaitForResponse
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status = 200 Then
            Set UpdateItemProperties = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub CheckIn(Destination As Destination, ByVal CheckInParameters As CheckInParameters, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkin"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/checkin"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkin"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/checkin"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkin"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/checkin"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkin"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/checkin"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkin"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/checkin"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(CheckInParameters.ToJson)
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub CheckOut(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkout"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/checkout"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkout"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/checkout"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkout"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/checkout"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkout"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/checkout"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/checkout"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/checkout"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Sub Delete(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfMatch As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "DELETE", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send
        .WaitForResponse
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub

Public Function Move(Destination As Destination, DriveItem As DriveItem, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfMatch As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath)
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "PATCH", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        If IfMatch <> vbNullString Then .SetRequestHeader "if-match", IfMatch
        .Send JsonConverter.ConvertToJson(DriveItem.ToJson)
        .WaitForResponse
        If .Status = 412 Then Call HandleError(.ResponseText)
        If .Status = 200 Then
            Set Move = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Copy(Destination As Destination, ByVal DriveItem As DriveItem, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal QueryParameters As String) As String
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/copy"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/copy"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/copy"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/copy"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/copy"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/copy"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/copy"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/copy"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/copy"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/copy"
            End If
    End Select
    If QueryParameters <> vbNullString Then strUrl = strUrl & "?" & QueryParameters
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(DriveItem.ToJson)
        .WaitForResponse
        If .Status = 202 Then
            Copy = .GetResponseHeader("Destination")
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub ConvertContent(ByVal NewConvertedFileName As String, Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal IfNoneMatch As String)
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content?format=pdf"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/content?format=pdf"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content?format=pdf"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/content?format=pdf"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content?format=pdf"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/content?format=pdf"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content?format=pdf"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/content?format=pdf"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/content?format=pdf"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/content?format=pdf"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        If IfNoneMatch <> vbNullString Then .SetRequestHeader "if-none-match", IfNoneMatch
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objStream As ADODB.Stream
            Set objStream = New ADODB.Stream
            With objStream
                .Type = adTypeBinary
                .Open
                .Write objWinHttp.ResponseBody
                .SaveToFile NewConvertedFileName
                .Close
            End With
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Sub

Public Function ListVersions(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As Collection
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/versions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & ":/versions"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/versions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & ":/versions"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/versions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & ":/versions"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/versions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & ":/versions"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/versions"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & ":/versions"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "GET", strUrl, True
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .Send
        .WaitForResponse
        Dim colDriveItemVersions As New Collection
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objCreatedBy As IdentitySet
            Dim objLastModifiedBy As IdentitySet
            Dim objIdentity As Identity
            Dim objDriveItemVersion As DriveItemVersion
            Dim i As Long
            If objJson.Item("value").Count > 0 Then
                For i = 1 To objJson.Item("value").Count
                    Set objDriveItemVersion = New DriveItemVersion
                    With objDriveItemVersion
                        .Id = objJson.Item("value")(i)("id")
                        .Size = objJson.Item("value")(i)("size")
                        Set objCreatedBy = New IdentitySet
                        Set objIdentity = New Identity
                        .LastModifiedDateTime = objJson.Item("value")(i)("lastModifiedDateTime")
                        If objJson.Item("value")(i).Exists("lastModifiedBy") Then
                            Set objLastModifiedBy = New IdentitySet
                            Set objIdentity = New Identity
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("user") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("user")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("user")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("user")("id")
                                Set objLastModifiedBy.User = objIdentity
                            End If
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("application") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("application")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("application")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("application")("id")
                                Set objLastModifiedBy.Application = objIdentity
                            End If
                            If objJson.Item("value")(i)("lastModifiedBy").Exists("device") And IsEmpty(objJson.Item("value")(i)("lastModifiedBy")("device")) = False Then
                                objIdentity.DisplayName = objJson.Item("value")(i)("lastModifiedBy")("device")("displayName")
                                objIdentity.Id = objJson.Item("value")(i)("lastModifiedBy")("device")("id")
                                Set objLastModifiedBy.Device = objIdentity
                            End If
                        End If
                        Set .LastModifiedBy = objLastModifiedBy
                        colDriveItemVersions.Add objDriveItemVersion
                    End With
                Next
                Set ListVersions = colDriveItemVersions
            End If
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Follow(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String) As DriveItem
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/follow"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/follow"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/follow"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/follow"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/follow"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/follow"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/follow"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/follow"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/follow"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/follow"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Set Follow = FetchDriveItemInfo(.ResponseText)
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function GetItemAnalytics(ByVal DriveId As String, ByVal ItemId As String, Optional ByVal TimeRange As PredefinedTimeRange = AllTime, Optional ByVal ODataQuery As String) As ItemAnalytics
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    Dim strUrl As String
    strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/analytics/"
    With objWinHttp
        If TimeRange = AllTime Then
            strUrl = strUrl & "allTime"
        Else: strUrl = strUrl & "lastSevenDays"
        End If
        If ODataQuery <> vbNullString Then
            strUrl = strUrl & "?" & ODataQuery
        End If
        .Open "GET", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .Send
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Dim objIncompletedData As IncompleteData
            Dim objItemActionStat As ItemActionStat
            Dim objItemActivityStat As ItemActivityStat
            Dim objItemAnalytics As ItemAnalytics
            If objJson.Exists("incompleteData") Then
                Set objIncompletedData = New IncompleteData
                With objIncompletedData
                    .WasThrottled = CBool(objJson.Item("incompleteData")("wasThrottled"))
                End With
            End If
            If objJson.Exists("access") Then
                Set objItemActionStat = New ItemActionStat
                With objItemActionStat
                    .ActionCount = objJson.Item("access")("actionCount")
                    .ActorCount = objJson.Item("access")("actorCount")
                End With
            End If
            Set objItemActivityStat = New ItemActivityStat
            With objItemActivityStat
                .StartDateTime = objJson.Item("startDateTime")
                .EndDateTime = objJson.Item("endDateTime")
                .IsTrending = CBool(objJson.Item("isTrending"))
                If Not objIncompletedData Is Nothing Then Set .IncompletedData = objIncompletedData
                If Not objItemActionStat Is Nothing Then Set .Access = objItemActionStat
            End With
            Set objItemAnalytics = New ItemAnalytics
            With objItemAnalytics
                If TimeRange = AllTime Then
                    Set .AllTime = objItemActivityStat
                Else: Set .LastSevenDays = objItemActivityStat
                End If
            End With
            Set GetItemAnalytics = objItemAnalytics
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Function Preview(Destination As Destination, ByVal RequestParameters As Scripting.Dictionary, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal SiteId As String, Optional ByVal DriveId As String, Optional ByVal UserId As String, Optional ByVal GroupId As String, Optional ByVal ShareId As String) As Scripting.Dictionary
    Dim objWebUtilities As WebUtilities
    Set objWebUtilities = New WebUtilities
    Dim strUrl As String
    Select Case Destination
        Case DestinationDrive
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/preview"
            Else: strUrl = "https://graph.microsoft.com/v1.0/drives/" & DriveId & "/items/" & ItemId & "/preview"
            End If
        Case DestinationGroup
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/preview"
            Else: strUrl = "https://graph.microsoft.com/v1.0/groups/" & GroupId & "/drive/items/" & ItemId & "/preview"
            End If
        Case DestinationOtherUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/preview"
            Else: strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/" & ItemId & "/preview"
            End If
        Case DestinationSignedInUser
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/preview"
            Else: strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/" & ItemId & "/preview"
            End If
        Case DestinationSite
            If ItemPath <> vbNullString Then
                strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/root:/" & objWebUtilities.URLEncode(ItemPath) & ":/preview"
            Else: strUrl = "https://graph.microsoft.com/v1.0/sites/" & SiteId & "/drive/items/" & ItemId & "/preview"
            End If
        Case DestinationShare
            strUrl = "https://graph.microsoft.com/v1.0/shares/" & ShareId & "/driveItem/preview"
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open "POST", strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send JsonConverter.ConvertToJson(RequestParameters)
        .WaitForResponse
        If .Status = 200 Then
            Dim objJson As Scripting.Dictionary
            Set objJson = JsonConverter.ParseJson(.ResponseText)
            Set Preview = objJson
        Else: Call HandleError(.ResponseText)
        End If
    End With
End Function

Public Sub Unfollow(Destination As Destination, Optional ByVal ItemId As String, Optional ByVal ItemPath As String, Optional ByVal UserId As String)
    Dim strUrl As String, strMethod As String
    Select Case Destination
        Case DestinationSignedInUser
            If ItemPath = vbNullString Then
                strMethod = "DELETE"
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/following/" & ItemId
            Else
                strMethod = "POST"
                strUrl = "https://graph.microsoft.com/v1.0/me/drive/items/root:/" & ItemPath & ":/following"
            End If
        Case DestinationOtherUser
            If ItemPath = vbNullString Then
                strMethod = "DELETE"
                strUrl = "https://graph.microsoft.com/v1.0/users" & UserId & "/drive/following/" & ItemId
            Else
                strMethod = "POST"
                strUrl = "https://graph.microsoft.com/v1.0/users/" & UserId & "/drive/items/root:/" & ItemPath & ":/following"
            End If
    End Select
    Dim objWinHttp As WinHttp.WinHttpRequest
    Set objWinHttp = New WinHttp.WinHttpRequest
    With objWinHttp
        .Open strMethod, strUrl, True
        .SetRequestHeader "Authorization", "Bearer " & AccessToken
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        If .Status <> 204 Then Call HandleError(.ResponseText)
    End With
End Sub
