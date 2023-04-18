VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CTag As String
Private m_Description As String
Private m_ETag As String
Private m_WebDavUrl As String
Private m_Id As String
Private m_Name As String
Private m_WebUrl As String
Private m_Folder As DriveFolder
Private m_Size As Variant
Private m_Audio As Audio
Private m_Bundle As Bundle
Private m_CreatedBy As IdentitySet
Private m_CreatedDateTime As String
Private m_Deleted As Deleted
Private m_File As DriveFile
Private m_FileSystemInfo As FileSystemInfo
Private m_Image As Image
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Location As GeoCoordinates
Private m_Malware As Malware
Private m_Package As Package
Private m_ParentReference As ItemReference
Private m_PendingOperations As PendingOperations
Private m_Photo As Photo
Private m_PublicationFacet As PublicationFacet
Private m_RemoteItem As RemoteItem
Private m_Root As Root
Private m_SearchResult As SearchResult
Private m_Shared As ItemShared
Private m_SharepointIds As SharepointIds
Private m_SpecialFolder As SpecialFolder
Private m_Video As Video
Private m_DownloadUrl As String

Public Property Get DownloadUrl() As String
    DownloadUrl = m_DownloadUrl
End Property

Public Property Let DownloadUrl(Value As String)
    m_DownloadUrl = Value
End Property

Public Property Get CTag() As String
    CTag = m_CTag
End Property

Public Property Let CTag(Value As String)
    m_CTag = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get ETag() As String
    ETag = m_ETag
End Property

Public Property Let ETag(Value As String)
    m_ETag = Value
End Property

Public Property Get WebDavUrl() As String
    WebDavUrl = m_WebDavUrl
End Property

Public Property Let WebDavUrl(Value As String)
    m_WebDavUrl = Value
End Property

Public Property Get Audio() As Audio
    Set Audio = m_Audio
End Property

Public Property Set Audio(Value As Audio)
    Set m_Audio = Value
End Property

Public Property Get Bundle() As Bundle
    Set Bundle = m_Bundle
End Property

Public Property Set Bundle(Value As Bundle)
    Set m_Bundle = Value
End Property

Public Property Get CreatedBy() As IdentitySet
    Set CreatedBy = m_CreatedBy
End Property

Public Property Set CreatedBy(Value As IdentitySet)
    Set m_CreatedBy = Value
End Property

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get Deleted() As Deleted
    Set Deleted = m_Deleted
End Property

Public Property Set Deleted(Value As Deleted)
    Set m_Deleted = Value
End Property

Public Property Get File() As DriveFile
    Set File = m_File
End Property

Public Property Set File(Value As DriveFile)
    Set m_File = Value
End Property

Public Property Get FileSystemInfo() As FileSystemInfo
    Set FileSystemInfo = m_FileSystemInfo
End Property

Public Property Set FileSystemInfo(Value As FileSystemInfo)
    Set m_FileSystemInfo = Value
End Property

Public Property Get Image() As Image
    Set Image = m_Image
End Property

Public Property Set Image(Value As Image)
    Set m_Image = Value
End Property

Public Property Get LastModifiedBy() As IdentitySet
    Set LastModifiedBy = m_LastModifiedBy
End Property

Public Property Set LastModifiedBy(Value As IdentitySet)
    Set m_LastModifiedBy = Value
End Property

Public Property Get LastModifiedDateTime() As String
    LastModifiedDateTime = m_LastModifiedDateTime
End Property

Public Property Let LastModifiedDateTime(Value As String)
    m_LastModifiedDateTime = Value
End Property

Public Property Get Location() As GeoCoordinates
    Set Location = m_Location
End Property

Public Property Set Location(Value As GeoCoordinates)
    Set m_Location = Value
End Property

Public Property Get Malware() As Malware
    Set Malware = m_Malware
End Property

Public Property Set Malware(Value As Malware)
    Set m_Malware = Value
End Property

Public Property Get Package() As Package
    Set Package = m_Package
End Property

Public Property Set Package(Value As Package)
    Set m_Package = Value
End Property

Public Property Get ParentReference() As ItemReference
    Set ParentReference = m_ParentReference
End Property

Public Property Set ParentReference(Value As ItemReference)
    Set m_ParentReference = Value
End Property

Public Property Get PendingOperations() As PendingOperations
    Set PendingOperations = m_PendingOperations
End Property

Public Property Set PendingOperations(Value As PendingOperations)
    Set m_PendingOperations = Value
End Property

Public Property Get Photo() As Photo
    Set Photo = m_Photo
End Property

Public Property Set Photo(Value As Photo)
    Set m_Photo = Value
End Property

Public Property Get PublicationFacet() As PublicationFacet
    Set PublicationFacet = m_PublicationFacet
End Property

Public Property Set PublicationFacet(Value As PublicationFacet)
    Set m_PublicationFacet = Value
End Property

Public Property Get RemoteItem() As RemoteItem
    Set RemoteItem = m_RemoteItem
End Property

Public Property Set RemoteItem(Value As RemoteItem)
    Set m_RemoteItem = Value
End Property

Public Property Get Root() As Root
    Set Root = m_Root
End Property

Public Property Set Root(Value As Root)
    Set m_Root = Value
End Property

Public Property Get SearchResult() As SearchResult
    Set SearchResult = m_SearchResult
End Property

Public Property Set SearchResult(Value As SearchResult)
    Set m_SearchResult = Value
End Property

Public Property Get ItemShared() As ItemShared
    Set ItemShared = m_Shared
End Property

Public Property Set ItemShared(Value As ItemShared)
    Set m_Shared = Value
End Property

Public Property Get SharepointIds() As SharepointIds
    Set SharepointIds = m_SharepointIds
End Property

Public Property Set SharepointIds(Value As SharepointIds)
    Set m_SharepointIds = Value
End Property

Public Property Get SpecialFolder() As SpecialFolder
    Set SpecialFolder = m_SpecialFolder
End Property

Public Property Set SpecialFolder(Value As SpecialFolder)
    Set m_SpecialFolder = Value
End Property

Public Property Get Video() As Video
    Set Video = m_Video
End Property

Public Property Set Video(Value As Video)
    Set m_Video = Value
End Property

Public Property Get Id() As String
   Id = m_Id
End Property

Public Property Let Id(Value As String)
   m_Id = Value
End Property

Public Property Get Name() As String
   Name = m_Name
End Property

Public Property Let Name(Value As String)
   m_Name = Value
End Property

Public Property Get Folder() As DriveFolder
  Set Folder = m_Folder
End Property

Public Property Set Folder(Value As DriveFolder)
   Set m_Folder = Value
End Property

Public Property Get WebUrl() As String
   WebUrl = m_WebUrl
End Property

Public Property Let WebUrl(Value As String)
   m_WebUrl = Value
End Property

Public Property Get Size() As Variant
   Size = m_Size
End Property

Public Property Let Size(Value As Variant)
   m_Size = Value
End Property

Public Function ToJson() As Variant
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Name <> vbNullString Then .Add "name", Name
        If Not Folder Is Nothing Then .Add "folder", Folder.ToJson
        If Not File Is Nothing Then .Add "file", File.ToJson
        If Description <> vbNullString Then .Add "description", Description
        If IsEmpty(Size) = False Then .Add "size", Size
        If Not ParentReference Is Nothing Then .Add "parentReference", ParentReference.ToJson
    End With
    Set ToJson = objJson
End Function
