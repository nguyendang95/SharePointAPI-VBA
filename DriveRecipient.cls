VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveRecipient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Alias As String
Private m_Email As String
Private m_ObjectId As String

Public Property Get Alias() As String
    Alias = m_Alias
End Property

Public Property Let Alias(Value As String)
    m_Alias = Value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Email(Value As String)
    m_Email = Value
End Property

Public Property Get ObjectId() As String
    ObjectId = m_ObjectId
End Property

Public Property Let ObjectId(Value As String)
    m_ObjectId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Alias <> vbNullString Then .Add "alias", Alias
        If Email <> vbNullString Then .Add "email", Email
        If ObjectId <> vbNullString Then .Add "objectId", ObjectId
    End With
    Set ToJson = objJson
End Function


