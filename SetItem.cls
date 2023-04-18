VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_LocalizedNames As New Collection
Private m_Description As String
Private m_CreatedDateTime As String
Private m_Properties As New Collection

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get LocalizedNames() As Collection
    Set LocalizedNames = m_LocalizedNames
End Property

Public Property Set LocalizedNames(Value As Collection)
    Set m_LocalizedNames = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get Properties() As Collection
    Set Properties = m_Properties
End Property

Public Property Set Properties(Value As Collection)
    Set m_Properties = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If LocalizedNames.Count > 0 Then .Add "localizedNames", LocalizedNames
        If Description <> vbNullString Then .Add "description", Description
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Properties.Count > 0 Then .Add "properties", Properties
    End With
    Set ToJson = objJson
End Function


