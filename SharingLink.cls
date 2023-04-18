VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharingLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Application As Identity
Private m_PreventsDownload As Variant
Private m_SharingLinkType As String
Private m_Scope As String
Private m_WebHtml As String
Private m_WebUrl As String

Public Property Get Application() As Identity
    Set Application = m_Application
End Property

Public Property Set Application(Value As Identity)
    Set m_Application = Value
End Property

Public Property Get PreventsDownload() As Variant
    PreventsDownload = m_PreventsDownload
End Property

Public Property Let PreventsDownload(Value As Variant)
    m_PreventsDownload = Value
End Property

Public Property Get SharingLinkType() As String
    SharingLinkType = m_SharingLinkType
End Property

Public Property Let SharingLinkType(Value As String)
    m_SharingLinkType = Value
End Property

Public Property Get Scope() As String
    Scope = m_Scope
End Property

Public Property Let Scope(Value As String)
    m_Scope = Value
End Property

Public Property Get WebHtml() As String
    WebHtml = m_WebHtml
End Property

Public Property Let WebHtml(Value As String)
    m_WebHtml = Value
End Property

Public Property Get WebUrl() As String
    WebUrl = m_WebUrl
End Property

Public Property Let WebUrl(Value As String)
    m_WebUrl = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Application Is Nothing Then .Add "application", Application.ToJson
        If IsEmpty(PreventsDownload) = False Then .Add "PreventsDownload", PreventsDownload
        If SharingLinkType <> vbNullString Then .Add "sharingLinkType", SharingLinkType
        If Scope <> vbNullString Then .Add "scope", Scope
        If WebHtml <> vbNullString Then .Add "webHtml", WebHtml
        If WebUrl <> vbNullString Then .Add "webUrl", WebUrl
    End With
    Set ToJson = objJson
End Function
