VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemoteItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_CreatedBy As IdentitySet
Private m_CreatedDateTime As String
Private m_File As DriveFile
Private m_FileSystemInfo As FileSystemInfo
Private m_Folder As DriveFolder
Private m_Image As Image
Private m_LastModifiedBy As IdentitySet
Private m_LastModifiedDateTime As String
Private m_Name As String
Private m_Package As Package
Private m_ParentReference As ItemReference
Private m_ItemShared As ItemShared
Private m_SharepointIds As SharepointIds
Private m_SpecialFolder As SpecialFolder
Private m_Size As Variant
Private m_Video As Video
Private m_WebDavUrl As String
Private m_WebUrl As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
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

Public Property Get Folder() As DriveFolder
    Set Folder = m_Folder
End Property

Public Property Set Folder(Value As Folder)
    Set m_Folder = Value
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

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
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

Public Property Get ItemShared() As ItemShared
    Set ItemShared = m_ItemShared
End Property

Public Property Set ItemShared(Value As ItemShared)
    Set m_ItemShared = Value
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

Public Property Get Size() As Variant
    Size = m_Size
End Property

Public Property Let Size(Value As Variant)
    m_Size = Value
End Property

Public Property Get Video() As Video
    Set Video = m_Video
End Property

Public Property Set Video(Value As Video)
    Set m_Video = Value
End Property

Public Property Get WebDavUrl() As String
    WebDavUrl = m_WebDavUrl
End Property

Public Property Let WebDavUrl(Value As String)
    m_WebDavUrl = Value
End Property

Public Property Get WebUrl() As String
    WebUrl = m_WebUrl
End Property

Public Property Let WebUrl(Value As String)
    m_WebUrl = Value
End Property


























