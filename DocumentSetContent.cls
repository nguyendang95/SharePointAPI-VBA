VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentSetContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ContentType As ContentTypeInfo
Private m_FileName As String
Private m_FolderName As String

Public Property Get ContentType() As ContentTypeInfo
    Set ContentType = m_ContentType
End Property

Public Property Set ContentType(Value As ContentTypeInfo)
    Set m_ContentType = Value
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(Value As String)
    m_FileName = Value
End Property

Public Property Get FolderName() As String
    FolderName = m_FolderName
End Property

Public Property Let FolderName(Value As String)
    m_FolderName = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not ContentType Is Nothing Then .Add "contentType", ContentType.ToJson
        If FileName <> vbNullString Then .Add "fileName", FileName
        If FolderName <> vbNullString Then .Add "folderName", FolderName
    End With
    Set ToJson = objJson
End Function


