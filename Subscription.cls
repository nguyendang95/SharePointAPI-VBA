VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Subscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Resource As String
Private m_ChangeType As String
Private m_ClientState As String
Private m_NotificationUrl As String
Private m_ExpirationDateTime As String
Private m_ApplicationId As String
Private m_CreatorId As String
Private m_IncludeResourceData As Variant
Private m_LifecycleNotificationUrl As String
Private m_EncryptionCertificate As String
Private m_LatestSupportedTlsVersion As String
Private m_NotificationQueryOptions As String
Private m_NotificationUrlAppId As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Resource() As String
    Resource = m_Resource
End Property

Public Property Let Resource(Value As String)
    m_Resource = Value
End Property

Public Property Get ChangeType() As String
    ChangeType = m_ChangeType
End Property

Public Property Let ChangeType(Value As String)
    m_ChangeType = Value
End Property

Public Property Get ClientState() As String
    ClientState = m_ClientState
End Property

Public Property Let ClientState(Value As String)
    m_ClientState = Value
End Property

Public Property Get NotificationUrl() As String
    NotificationUrl = m_NotificationUrl
End Property

Public Property Let NotificationUrl(Value As String)
    m_NotificationUrl = Value
End Property

Public Property Get ExpirationDateTime() As String
    ExpirationDateTime = m_ExpirationDateTime
End Property

Public Property Let ExpirationDateTime(Value As String)
    m_ExpirationDateTime = Value
End Property

Public Property Get ApplicationId() As String
    ApplicationId = m_ApplicationId
End Property

Public Property Let ApplicationId(Value As String)
    m_ApplicationId = Value
End Property

Public Property Get CreatorId() As String
    CreatorId = m_CreatorId
End Property

Public Property Let CreatorId(Value As String)
    m_CreatorId = Value
End Property

Public Property Get IncludeResourceData() As Variant
    IncludeResourceData = m_IncludeResourceData
End Property

Public Property Let IncludeResourceData(Value As Variant)
    m_IncludeResourceData = Value
End Property

Public Property Get LifecycleNotificationUrl() As String
    LifecycleNotificationUrl = m_LifecycleNotificationUrl
End Property

Public Property Let LifecycleNotificationUrl(Value As String)
    m_LifecycleNotificationUrl = Value
End Property

Public Property Get EncryptionCertificate() As String
    EncryptionCertificate = m_EncryptionCertificate
End Property

Public Property Let EncryptionCertificate(Value As String)
    m_EncryptionCertificate = Value
End Property

Public Property Get LatestSupportedTlsVersion() As String
    LatestSupportedTlsVersion = m_LatestSupportedTlsVersion
End Property

Public Property Let LatestSupportedTlsVersion(Value As String)
    m_LatestSupportedTlsVersion = Value
End Property

Public Property Get NotificationQueryOptions() As String
    NotificationQueryOptions = m_NotificationQueryOptions
End Property

Public Property Let NotificationQueryOptions(Value As String)
    m_NotificationQueryOptions = Value
End Property

Public Property Get NotificationUrlAppId() As String
    NotificationUrlAppId = m_NotificationUrlAppId
End Property

Public Property Let NotificationUrlAppId(Value As String)
    m_NotificationUrlAppId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Resource <> vbNullString Then .Add "resource", Resource
        If ChangeType <> vbNullString Then .Add "changeType", ChangeType
        If ClientState <> vbNullString Then .Add "clientState", ClientState
        If NotificationUrl <> vbNullString Then .Add "notificationUrl", NotificationUrl
        If ExpirationDateTime <> vbNullString Then .Add "expirationDateTime", ExpirationDateTime
        If ApplicationId <> vbNullString Then .Add "applicationId", ApplicationId
        If CreatorId <> vbNullString Then .Add "creatorId", CreatorId
        If IsEmpty(IncludeResourceData) = False Then .Add "IncludeResourceData", IncludeResourceData
        If LifecycleNotificationUrl <> vbNullString Then .Add "lifecycleNotificationUrl", LifecycleNotificationUrl
        If EncryptionCertificate <> vbNullString Then .Add "encryptionCertificate", EncryptionCertificate
        If LatestSupportedTlsVersion <> vbNullString Then .Add "latestSupportedTlsVersion", LatestSupportedTlsVersion
        If NotificationQueryOptions <> vbNullString Then .Add "notificationQueryOptions", NotificationQueryOptions
        If NotificationUrlAppId <> vbNullString Then .Add "notificationUrlAppId", NotificationUrlAppId
    End With
    Set ToJson = objJson
End Function


