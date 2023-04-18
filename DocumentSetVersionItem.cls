VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentSetVersionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ItemId As String
Private m_Title As String
Private m_VersionId As String

Public Property Get ItemId() As String
    ItemId = m_ItemId
End Property

Public Property Let ItemId(Value As String)
    m_ItemId = Value
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(Value As String)
    m_Title = Value
End Property

Public Property Get VersionId() As String
    VersionId = m_VersionId
End Property

Public Property Let VersionId(Value As String)
    m_VersionId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If ItemId <> vbNullString Then .Add "itemId", ItemId
        If Title <> vbNullString Then .Add "title", Title
        If VersionId <> vbNullString Then .Add "versionId", VersionId
    End With
    Set ToJson = objJson
End Function

