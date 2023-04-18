VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicationFacet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Level As String
Private m_VersionId As String

Public Property Get Level() As String
    Level = m_Level
End Property

Public Property Let Level(Value As String)
    m_Level = Value
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
        If Level <> vbNullString Then .Add "level", Level
        If VersionId <> vbNullString Then .Add "versionId", VersionId
    End With
    Set ToJson = objJson
End Function

