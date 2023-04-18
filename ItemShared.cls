VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemShared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Owner As IdentitySet
Private m_Scope As String
Private m_SharedBy As IdentitySet
Private m_SharedDateTime As String

Public Property Get Owner() As IdentitySet
    Set Owner = m_Owner
End Property

Public Property Set Owner(Value As IdentitySet)
    Set m_Owner = Value
End Property

Public Property Get Scope() As String
    Scope = m_Scope
End Property

Public Property Let Scope(Value As String)
    m_Scope = Value
End Property

Public Property Get SharedBy() As IdentitySet
    Set SharedBy = m_SharedBy
End Property

Public Property Set SharedBy(Value As IdentitySet)
    Set m_SharedBy = Value
End Property

Public Property Get SharedDateTime() As String
    SharedDateTime = m_SharedDateTime
End Property

Public Property Let SharedDateTime(Value As String)
    m_SharedDateTime = Value
End Property



