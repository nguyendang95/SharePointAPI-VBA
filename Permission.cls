VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Permission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_GrantedToV2 As SharePointIdentitySet
Private m_GrantedToIdentitiesV2 As New Collection
Private m_InheritedFrom As ItemReference
Private m_Invitation As SharingInvitation
Private m_Link As SharingLink
Private m_Roles As New Collection
Private m_ShareId As String
Private m_ExpirationDateTime As String
Private m_HasPassword As Variant

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get GrantedToV2() As SharePointIdentitySet
    Set GrantedToV2 = m_GrantedToV2
End Property

Public Property Set GrantedToV2(Value As SharePointIdentitySet)
    Set m_GrantedToV2 = Value
End Property

Public Property Get GrantedToIdentitiesV2() As Collection
    Set GrantedToIdentitiesV2 = m_GrantedToIdentitiesV2
End Property

Public Property Set GrantedToIdentitiesV2(Value As Collection)
    Set m_GrantedToIdentitiesV2 = Value
End Property

Public Property Get InheritedFrom() As ItemReference
    Set InheritedFrom = m_InheritedFrom
End Property

Public Property Set InheritedFrom(Value As ItemReference)
    Set m_InheritedFrom = Value
End Property

Public Property Get Invitation() As SharingInvitation
    Set Invitation = m_Invitation
End Property

Public Property Set Invitation(Value As SharingInvitation)
    Set m_Invitation = Value
End Property

Public Property Get Link() As SharingLink
    Set Link = m_Link
End Property

Public Property Set Link(Value As SharingLink)
    Set m_Link = Value
End Property

Public Property Get Roles() As Collection
    Set Roles = m_Roles
End Property

Public Property Set Roles(Value As Collection)
    Set m_Roles = Value
End Property

Public Property Get ShareId() As String
    ShareId = m_ShareId
End Property

Public Property Let ShareId(Value As String)
    m_ShareId = Value
End Property

Public Property Get ExpirationDateTime() As String
    ExpirationDateTime = m_ExpirationDateTime
End Property

Public Property Let ExpirationDateTime(Value As String)
    m_ExpirationDateTime = Value
End Property

Public Property Get HasPassword() As Variant
    HasPassword = m_HasPassword
End Property

Public Property Let HasPassword(Value As Variant)
    m_HasPassword = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If ExpirationDateTime <> vbNullString Then .Add "expirationDateTime", ExpirationDateTime
    End With
    Set ToJson = objJson
End Function

