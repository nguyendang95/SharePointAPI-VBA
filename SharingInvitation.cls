VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharingInvitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Email As String
Private m_InvitedBy As IdentitySet
Private m_SignInRequired As Variant

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(Value As String)
    m_Email = Value
End Property

Public Property Get InvitedBy() As IdentitySet
    Set InvitedBy = m_InvitedBy
End Property

Public Property Set InvitedBy(Value As IdentitySet)
    Set m_InvitedBy = Value
End Property

Public Property Get SignInRequired() As Variant
    SignInRequired = m_SignInRequired
End Property

Public Property Let SignInRequired(Value As Variant)
    m_SignInRequired = Value
End Property



