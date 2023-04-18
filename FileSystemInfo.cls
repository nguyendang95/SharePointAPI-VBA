VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CreatedDateTime As String
Private m_LastAccessedDateTime As String
Private m_LastModifiedDateTime As String

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get LastAccessedDateTime() As String
    LastAccessedDateTime = m_LastAccessedDateTime
End Property

Public Property Let LastAccessedDateTime(Value As String)
    m_LastAccessedDateTime = Value
End Property

Public Property Get LastModifiedDateTime() As String
    LastModifiedDateTime = m_LastModifiedDateTime
End Property

Public Property Let LastModifiedDateTime(Value As String)
    m_LastModifiedDateTime = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If LastAccessedDateTime <> vbNullString Then .Add "lastAccessedDateTime", LastAccessedDateTime
        If LastModifiedDateTime <> vbNullString Then .Add "lastModifiedDateTime", LastModifiedDateTime
    End With
    Set ToJson = objJson
End Function


